VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCMISBankDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Deposit Data Entry"
   ClientHeight    =   8415
   ClientLeft      =   5280
   ClientTop       =   4020
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "BankDeposit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   15270
   Begin VB.PictureBox picLSV 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   4770
      ScaleHeight     =   6735
      ScaleWidth      =   9720
      TabIndex        =   79
      Top             =   750
      Width           =   9750
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   8940
         MouseIcon       =   "BankDeposit.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   5850
         Width           =   705
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "0.00"
         Top             =   5070
         Width           =   1815
      End
      Begin wizButton.cmd cmd1 
         Height          =   315
         Left            =   8550
         TabIndex        =   98
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         TX              =   "cmd1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "BankDeposit.frx":0D5A
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   94
         Text            =   "0.00"
         Top             =   5430
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "0.00"
         Top             =   5430
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   92
         Text            =   "0.00"
         Top             =   5070
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   390
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   6630
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   315
         Left            =   8190
         TabIndex        =   84
         Text            =   "0.00"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "BankDeposit.frx":0D76
         Left            =   60
         List            =   "BankDeposit.frx":0D83
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "BankDeposit.frx":0D9A
         Left            =   1530
         List            =   "BankDeposit.frx":0DA4
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   1200
         Width           =   4515
      End
      Begin MSComctlLib.ListView lsvTran 
         Height          =   3345
         Left            =   60
         TabIndex        =   80
         Top             =   1650
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bank name/ Customer Name"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cash/Chech Amount"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin wizButton.cmd cmd2 
         Height          =   315
         Left            =   9150
         TabIndex        =   99
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   556
         TX              =   "cmd1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "BankDeposit.frx":0DCE
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   795
         Left            =   8250
         MouseIcon       =   "BankDeposit.frx":0DEA
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":0F3C
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   5850
         Width           =   705
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OR Number :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   6645
         TabIndex        =   101
         Top             =   5550
         Width           =   1065
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   0
         Left            =   6360
         TabIndex        =   97
         Top             =   5190
         Width           =   1350
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Type :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   90
         TabIndex        =   96
         Top             =   5550
         Width           =   1080
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   105
         TabIndex        =   95
         Top             =   5190
         Width           =   1065
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Deposit  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   91
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   89
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   1530
         TabIndex        =   88
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6990
         TabIndex        =   87
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   8250
         TabIndex        =   86
         Top             =   960
         Width           =   1335
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   81
         Top             =   -30
         Width           =   12015
         _Version        =   655364
         _ExtentX        =   21193
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picBankDepo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   6675
      Left            =   2640
      ScaleHeight     =   6645
      ScaleWidth      =   9045
      TabIndex        =   23
      Top             =   840
      Width           =   9075
      Begin VB.TextBox txtBankDeposit 
         BackColor       =   &H00FFFFFF&
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
         Height          =   405
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   60
         Width           =   2115
      End
      Begin VB.CommandButton cmdDeleteBANKDEPO 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   240
         MouseIcon       =   "BankDeposit.frx":128C
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5670
         Width           =   705
      End
      Begin VB.ComboBox cboCheckTransactions 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "BankDeposit.frx":1709
         Left            =   1590
         List            =   "BankDeposit.frx":170B
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
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
         Height          =   915
         Left            =   210
         ScaleHeight     =   885
         ScaleWidth      =   8595
         TabIndex        =   29
         Top             =   4680
         Width           =   8625
         Begin VB.TextBox txtBankCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   -780
            Width           =   1455
         End
         Begin VB.TextBox txtTimeCreate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   -780
            Width           =   1455
         End
         Begin VB.TextBox txtCheckAmount 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   5280
            TabIndex        =   38
            Top             =   -390
            Width           =   1455
         End
         Begin VB.TextBox txtORNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   7050
            TabIndex        =   35
            Top             =   450
            Width           =   1455
         End
         Begin VB.TextBox txtCheckNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   7050
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtCheckType 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1170
            TabIndex        =   31
            Top             =   450
            Width           =   4605
         End
         Begin VB.TextBox txtCheckDte 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   60
            Width           =   1455
         End
         Begin VB.Label labTranID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
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
            Left            =   3150
            TabIndex        =   41
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Number  :"
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
            Left            =   5160
            TabIndex        =   37
            Top             =   540
            Width           =   1845
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number  :"
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
            Left            =   5160
            TabIndex        =   36
            Top             =   150
            Width           =   1845
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Type  :"
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
            Left            =   -720
            TabIndex        =   33
            Top             =   540
            Width           =   1845
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date  :"
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
            Left            =   -720
            TabIndex        =   32
            Top             =   150
            Width           =   1845
         End
      End
      Begin VB.ComboBox cboBankCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   1590
         TabIndex        =   5
         Text            =   "cboBankCode"
         Top             =   900
         Width           =   4485
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   900
         Width           =   1305
      End
      Begin VB.TextBox txtDeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7380
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtTimDeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   900
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid grdCheckCardTransactions 
         Height          =   3285
         Left            =   210
         TabIndex        =   9
         Top             =   1320
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   5794
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   14606302
         BackColorSel    =   14606302
         BackColorBkg    =   14606302
         FillStyle       =   1
         Appearance      =   0
         MousePointer    =   15
         FormatString    =   " Code         |   Bank Name                                                 |    Time        | Check Amount  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "BankDeposit.frx":170D
      End
      Begin MSComCtl2.DTPicker dtpDatDeposit 
         Height          =   405
         Left            =   2010
         TabIndex        =   2
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   47775745
         CurrentDate     =   38216
      End
      Begin VB.CommandButton cmdCancelBANKDEPO 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   8100
         MouseIcon       =   "BankDeposit.frx":1A27
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":1B79
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   5670
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveBANKDEPO 
         Caption         =   "&Save"
         Height          =   795
         Left            =   7410
         MouseIcon       =   "BankDeposit.frx":1EB7
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":2009
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5670
         Width           =   705
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10820
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Deposit :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   330
         TabIndex        =   42
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label labBankDepoID 
         Caption         =   "Label14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1980
         TabIndex        =   28
         Top             =   1020
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Deposit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7380
         TabIndex        =   27
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6150
         TabIndex        =   26
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   1620
         TabIndex        =   25
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   660
         Width           =   1305
      End
   End
   Begin VB.PictureBox picHEad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   30
      ScaleHeight     =   705
      ScaleWidth      =   11655
      TabIndex        =   75
      Top             =   30
      Width           =   11685
      Begin VB.ComboBox cboDeposit_To 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   2610
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   150
         Width           =   7995
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
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
         Left            =   10650
         TabIndex        =   77
         Top             =   180
         Width           =   975
      End
      Begin Crystal.CrystalReport rptBankDepo 
         Left            =   -30
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Daily Bank Deposit"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   735
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Width           =   11865
         _Version        =   655364
         _ExtentX        =   20929
         _ExtentY        =   1296
         _StockProps     =   14
         Caption         =   "    Bank Name :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.PictureBox fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   7605
      Left            =   30
      ScaleHeight     =   7575
      ScaleWidth      =   2475
      TabIndex        =   72
      Top             =   750
      Width           =   2505
      Begin VB.TextBox textSearch 
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
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   30
         MaxLength       =   35
         TabIndex        =   74
         Top             =   60
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstBANKDEPO 
         Height          =   6945
         Left            =   30
         TabIndex        =   73
         Top             =   540
         Width           =   2385
         _ExtentX        =   4207
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
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "BankDeposit.frx":2359
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date Of Deposit"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
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
      Height          =   870
      Left            =   7770
      ScaleHeight     =   870
      ScaleWidth      =   4245
      TabIndex        =   44
      Top             =   7530
      Width           =   4245
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   3510
         MouseIcon       =   "BankDeposit.frx":24BB
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":260D
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "BankDeposit.frx":2973
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":2AC5
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   795
         Left            =   2130
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "BankDeposit.frx":2E2B
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":2F7D
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Post this Transaction"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   1440
         MouseIcon       =   "BankDeposit.frx":32A2
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":33F4
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   750
         MouseIcon       =   "BankDeposit.frx":3750
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":38A2
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   60
         MouseIcon       =   "BankDeposit.frx":3BB5
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":3D07
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picVar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12330
      ScaleHeight     =   1635
      ScaleWidth      =   2535
      TabIndex        =   65
      Top             =   6810
      Width           =   2595
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OFF_HD ID"
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
         Index           =   3
         Left            =   0
         TabIndex        =   71
         Top             =   1140
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "BANKDEPO ID"
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
         Index           =   2
         Left            =   0
         TabIndex        =   70
         Top             =   570
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OR NUM"
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
         Index           =   1
         Left            =   0
         TabIndex        =   69
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label labHD_ID 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   68
         Top             =   1320
         Width           =   2205
      End
      Begin VB.Label labBANKDEPO_ID 
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   67
         Top             =   780
         Width           =   2205
      End
      Begin VB.Label labORNUM 
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   66
         Top             =   210
         Width           =   2205
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   855
      Left            =   2640
      ScaleHeight     =   825
      ScaleWidth      =   4890
      TabIndex        =   54
      Top             =   7500
      Width           =   4920
      Begin VB.TextBox txtCheckNum 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3540
         TabIndex        =   57
         Top             =   60
         Width           =   1305
      End
      Begin VB.TextBox txtCheckDate 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1200
         TabIndex        =   56
         Top             =   60
         Width           =   1305
      End
      Begin VB.TextBox txtTseklase 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1200
         TabIndex        =   55
         Top             =   450
         Width           =   3645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Type  :"
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
         Left            =   -690
         TabIndex        =   60
         Top             =   510
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check No.  :"
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
         Left            =   2535
         TabIndex        =   59
         Top             =   180
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Date  :"
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
         Left            =   -690
         TabIndex        =   58
         Top             =   150
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture6 
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
      Height          =   885
      Left            =   12930
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   51
      Top             =   8760
      Width           =   1980
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   750
         MouseIcon       =   "BankDeposit.frx":4001
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":4153
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "BankDeposit.frx":4491
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":45E3
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   585
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   9045
      TabIndex        =   12
      Top             =   840
      Width           =   9075
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8040
         Top             =   30
      End
      Begin VB.TextBox txtDatDeposit 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   60
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Deposit  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   90
         Width           =   2235
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
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
      Height          =   1275
      Left            =   2640
      ScaleHeight     =   1245
      ScaleWidth      =   9045
      TabIndex        =   14
      Top             =   6240
      Width           =   9075
      Begin VB.TextBox txtCardDeposit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2520
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCashAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   60
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCheckAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2520
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDepositedAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   525
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   630
         Width           =   2385
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Card Deposit  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   480
         TabIndex        =   43
         Top             =   930
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Deposit  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   465
         TabIndex        =   21
         Top             =   150
         Width           =   1950
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Deposit  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   345
         TabIndex        =   20
         Top             =   510
         Width           =   2070
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deposited Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   6225
         TabIndex        =   19
         Top             =   360
         Width           =   2280
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   4845
      Left            =   2640
      ScaleHeight     =   4815
      ScaleWidth      =   9045
      TabIndex        =   22
      Top             =   1410
      Width           =   9075
      Begin MSFlexGridLib.MSFlexGrid grdBankDepo 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   90
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8070
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   " Type          |   Bank Name                                                 |    Time        | Amount Deposit  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "BankDeposit.frx":4933
      End
   End
   Begin VB.Label lblBANKID 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   64
      Top             =   8040
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5430
      TabIndex        =   11
      Top             =   2580
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6330
      TabIndex        =   10
      Top             =   2550
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCMISBankDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBANKDEPO                                                        As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim TOTAL_CASH_DEPOSIT, TOTAL_CHECK_DEPOSIT, TOTAL_CARD_DEPOSIT       As Double
Attribute TOTAL_CHECK_DEPOSIT.VB_VarUserMemId = 1073938434
Attribute TOTAL_CARD_DEPOSIT.VB_VarUserMemId = 1073938434
Dim PREV_CASH_DEPOSIT, PREV_CHECK_DEPOSIT, PREV_CARD_DEPOSIT          As Double
Attribute PREV_CASH_DEPOSIT.VB_VarUserMemId = 1073938437
Attribute PREV_CHECK_DEPOSIT.VB_VarUserMemId = 1073938437
Attribute PREV_CARD_DEPOSIT.VB_VarUserMemId = 1073938437

Function SetCustomerName(XXX As Variant)
    Dim rsCustomer                                                    As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select CusNam from ALL_CUSMAS Where CusCde = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = rsCustomer!CusNam
    End If
    Set rsCustomer = Nothing
End Function

Function SetBankCode(XXX As Variant)
    Dim rsBankName                                                    As New ADODB.Recordset
    
    'UPDATE BY   : MJP 052209 0400PM
    'DESCRIPTION : TO GET FROM THE ALL_BANKS TABLE, BECAUSE CMIS_BANKS IS AN UNION OF TWO TABLE POSIBLE ERROR IS PROGRAM MY SET THE BANKS IN THE CMIS ONLY
    '    Set rsBankName = gconDMIS.Execute("Select BANKCODE from ALL_BANKS Where BANKNAME = " & N2Str2Null(XXX) & "")
    'UPDATE BY   : MJP 052209 0400PM
    
    
    Set rsBankName = gconDMIS.Execute("Select BANKCODE from CMIS_BANKS Where BANKNAME = " & N2Str2Null(XXX) & "")
    If Not (rsBankName.EOF And rsBankName.BOF) Then
        SetBankCode = Null2String(rsBankName!bankcode)
    End If
    Set rsBankName = Nothing
End Function

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                   As New ADODB.Recordset
    'UPDATE BY   : MJP 052209 0400PM
    'DESCRIPTION : TO GET FROM THE ALL_BANKS TABLE, BECAUSE CMIS_BANKS IS AN UNION OF TWO TABLE POSIBLE ERROR IS PROGRAM MY SET THE BANKS IN THE CMIS ONLY
    '    Set rsBankName = gconDMIS.Execute("Select BANKCODE from ALL_BANKS Where BANKCode = " & N2Str2Null(XXX) & "")
    'UPDATE BY   : MJP 052209 0400PM
    
    Set rsBankName = gconDMIS.Execute("Select BANKNAME from CMIS_BANKS Where BANKCode = " & N2Str2Null(XXX) & "")
    If Not (rsBankName.EOF And rsBankName.BOF) Then
        SetBankName = Null2String(rsBankName!BANKNAME)
    End If
    Set rsBankName = Nothing
End Function

Function SetCheckClass(XXX As Variant)
    Dim rsSBOOK                                                       As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'F' and CODE = '" & XXX & "'")
    If Not (rsSBOOK.EOF And rsSBOOK.BOF) Then
        SetCheckClass = rsSBOOK!DESCNAME
    End If
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                       As New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select CODE from CMIS_SBOOK Where Book = 'F' and DESCNAME = '" & XXX & "'")
    If Not (rsSBOOK.EOF And rsSBOOK.BOF) Then
        SetCheckClassCode = rsSBOOK!code
    End If
End Function

Sub ShowTransactionsGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                            As New ADODB.Recordset
    Dim rsOFF_Details                                                 As New ADODB.Recordset
    
    If cboType.Text = "CASH" Then
        Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_Off_hd Where ID = " & XXX)
        If Not (rsBANKDEPO_Details.EOF And rsBANKDEPO_Details.BOF) Then
            labTranID.Caption = rsBANKDEPO_Details!Id
            txtCheckDte.Text = ""
            txtCheckType.Text = ""
            txtCheckNumber.Text = ""
            txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)
            txtBankCode.Text = ""
            txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
            txtCheckAmount.Text = ""
            txtDeposit.Text = ToDoubleNumber(rsBANKDEPO_Details!CashAmount)
        Else
            labTranID.Caption = ""
            txtCheckDate.Text = "": txtCheckType.Text = ""
            txtCheckNumber.Text = "": txtORNumber.Text = ""
            txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
        End If
    ElseIf cboType.Text = "CARD" Then
        Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_Off_hd Where ID = " & XXX)
        If Not (rsBANKDEPO_Details.EOF And rsBANKDEPO_Details.BOF) Then
            labTranID.Caption = rsBANKDEPO_Details!Id
            txtCheckDte.Text = Null2String(rsBANKDEPO_Details!carddate)
            txtCheckType.Text = ""
            txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!cardnumber)
            txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)
            txtBankCode.Text = Null2String(rsBANKDEPO_Details!cardbnkcde)
            txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
            txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CARDAMOUNT))
            Set rsOFF_Details = New ADODB.Recordset
            Set rsOFF_Details = gconDMIS.Execute("Select SUM(DISCOUNT) AS TOTAL_DISCOUNT, SUM(TAX) AS TOTAL_TAX from CMIS_Off_Dt Where OR_NUM = " & N2Str2Null(rsBANKDEPO_Details!OR_NUM))
            If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
                txtCheckAmount.Text = ToDoubleNumber(NumericVal(txtCheckAmount.Text) - (N2Str2Zero(rsOFF_Details!TOTAL_TAX) + N2Str2Zero(rsOFF_Details!TOTAL_DISCOUNT)))
            End If
            txtDeposit.Text = ToDoubleNumber(txtCheckAmount.Text)
        Else
            labTranID.Caption = ""
            txtCheckDate.Text = "": txtCheckType.Text = ""
            txtCheckNumber.Text = "": txtORNumber.Text = ""
            txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
        End If
    End If
    
    If cboCheckTransactions.Text = "Cashier Collection" Then
        If cboType.Text = "CHECK" Then
            Set rsBANKDEPO_Details = New ADODB.Recordset
            Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_Off_hd Where ID = " & XXX)
            If Not (rsBANKDEPO_Details.EOF And rsBANKDEPO_Details.BOF) Then
                labTranID.Caption = rsBANKDEPO_Details!Id
                'txtBankDeposit = Null2Date(rsBANKDEPO_Details!datdeposit)
                txtCheckType.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
                'UPDATE NOV. 10, 2005
                'If Null2String(rsBANKDEPO_Details!TOF) = "3" Then
                '   txtCheckDte.Text = Null2String(rsBANKDEPO_Details!CARDDATE)
                '   txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!CARDNUMBER)
                '   txtBankCode.Text = Null2String(rsBANKDEPO_Details!CARDBNKCDE)
                '   txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CARDAmount))
                '   txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CARDAmount))
                'Else
                txtCheckDte.Text = Null2String(rsBANKDEPO_Details!CheckDate)
                txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!Tseke)
                txtBankCode.Text = Null2String(rsBANKDEPO_Details!bankcode)
                txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!ChkAmount))
                txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!ChkAmount))
                'End If
                txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)
                txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
                'if check amount less tax,discount (08/19/2005)
                'Set rsOFF_Details = New ADODB.Recordset
                'Set rsOFF_Details = gconDMIS.Execute("Select SUM(DISCOUNT) AS TOTAL_DISCOUNT, SUM(TAX) AS TOTAL_TAX from CMIS_Off_Dt Where OR_NUM = " & N2Str2Null(rsBANKDEPO_Details!OR_NUM))
                'If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
                '   txtCheckAmount.Text = ToDoubleNumber(NumericVal(txtCheckAmount.Text) - (N2Str2Zero(rsOFF_Details!TOTAL_TAX) + N2Str2Zero(rsOFF_Details!TOTAL_DISCOUNT)))
                'End If
                'txtDeposit.Text = ToDoubleNumber(txtCheckAmount.Text)
            Else
                labTranID.Caption = ""
                txtBankDeposit = LOGDATE
                txtCheckDate.Text = "": txtCheckType.Text = ""
                txtCheckNumber.Text = "": txtORNumber.Text = ""
                txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
            End If
        End If
        Set rsBANKDEPO_Details = Nothing
    ElseIf cboCheckTransactions.Text = "Check Encashment" Then
        If cboType.Text = "CHECK" Then
            Set rsBANKDEPO_Details = New ADODB.Recordset
            Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_InCash Where ID = " & XXX)
            If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
                labTranID.Caption = rsBANKDEPO_Details!Id
                'txtBankDeposit = Null2Date(rsBANKDEPO_Details!datdeposit)
                txtCheckDte.Text = Null2String(rsBANKDEPO_Details!CHKDATE)
                txtCheckType.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
                txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!CHKNUMBER)
                txtBankCode.Text = Null2String(rsBANKDEPO_Details!bankcode)
                txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
                txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!ChkAmount))
                txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!ChkAmount))
            Else
                labTranID.Caption = ""
                txtBankDeposit = LOGDATE
                txtCheckDate.Text = "": txtCheckType.Text = ""
                txtCheckNumber.Text = "":
                txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
            End If
        End If
        Set rsBANKDEPO_Details = Nothing
    End If
End Sub

Sub SetSelectedType()
    If cboType.Text = "" Then
        cboCheckTransactions.Visible = False
        'MODIFIED SEPT. 8, 2007
        grdCheckCardTransactions.Enabled = False
        'grdCheckCardTransactions.Enabled = True
        cboBankCode.Enabled = False
        cboDeposit_To.Enabled = False
        txtDeposit.Enabled = False
    Else
        If cboType.Text = "CASH" Then
            'MODIFIED SEPT. 8, 2007
            'cboCheckTransactions.Visible = False
            'grdCheckCardTransactions.Enabled = False
            'txtDeposit.Enabled = True
            cboCheckTransactions.Visible = True
            grdCheckCardTransactions.Enabled = True
            txtDeposit.Enabled = False
            cboBankCode.Enabled = False
            
            cboDeposit_To.Enabled = True
        Else
            If cboType.Text = "CHECK" Then
                cboCheckTransactions.Visible = True
                grdCheckCardTransactions.Enabled = True
                'cboBankCode.Enabled = True
                txtDeposit.Enabled = False
                
                cboDeposit_To.Enabled = True
            Else
                cboCheckTransactions.Visible = False
                grdCheckCardTransactions.Enabled = True
                cboBankCode.Enabled = False
                txtDeposit.Enabled = True
                
                cboDeposit_To.Enabled = True
            End If
        End If
        
        cboCheckTransactions.Text = "Cashier Collection"
        Call cboCheckTransactions_Click
    End If
End Sub

Sub rsRefresh()
    Set rsBANKDEPO = New ADODB.Recordset
    Set rsBANKDEPO = gconDMIS.Execute("Select DISTINCT DATDEPOSIT from CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' order by DATDEPOSIT desc")
End Sub

Sub StoreMemvars()
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
        txtDatDeposit.Text = Null2Date(rsBANKDEPO!datdeposit)
        StoreDetails
    Else
        txtDatDeposit.Text = LOGDATE
        'cmdAdd.Value = True
    End If
End Sub

Sub StoreDetails()
    Dim rsBANKDEPODet                                                 As New ADODB.Recordset
    Dim VTYPE                                                         As String
    Dim I                                                             As Long
    TOTAL_CASH_DEPOSIT = 0: TOTAL_CHECK_DEPOSIT = 0: TOTAL_CARD_DEPOSIT = 0: InitGrid: I = 0
    
    Set rsBANKDEPODet = gconDMIS.Execute("Select * from CMIS_BankDepo where DEPOSIT_TO = '" & SetBankCode(RTrim(LTrim(cboDeposit_To))) & "' AND DATDEPOSIT = '" & txtDatDeposit.Text & "' Order by TYPE, ID asc")
    If Not rsBANKDEPODet.EOF And Not rsBANKDEPODet.BOF Then
        rsBANKDEPODet.MoveFirst
        Do While Not rsBANKDEPODet.EOF
            I = I + 1
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                VTYPE = "CASH"
                grdBankDepo.AddItem VTYPE & Chr(9) & _
                    " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & Chr(9) & _
                    Null2String(rsBANKDEPODet!timdeposit) & Chr(9) & _
                    ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!Deposit)) & Chr(9) & _
                    rsBANKDEPODet!Id
                
                'grdBankDepo.AddItem vType & Chr(9) & " XXX" & Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!Deposit)) & Chr(9) & rsBANKDEPODet!Id
            End If
            If Null2String(rsBANKDEPODet!Type) = "2" Then
                VTYPE = "CHECK"
                grdBankDepo.AddItem VTYPE & Chr(9) & " " & SetBankName(Null2String(rsBANKDEPODet!bankcode)) & Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!Deposit)) & Chr(9) & rsBANKDEPODet!Id
            End If
            If Null2String(rsBANKDEPODet!Type) = "3" Then
                VTYPE = "CARD"
                grdBankDepo.AddItem VTYPE & Chr(9) & " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!Deposit)) & Chr(9) & rsBANKDEPODet!Id
            End If
            
            If I = 1 Then grdBankDepo.RemoveItem 1
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
            End If
            If Null2String(rsBANKDEPODet!Type) = "2" Then
                TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
            End If
            If Null2String(rsBANKDEPODet!Type) = "3" Then
                TOTAL_CARD_DEPOSIT = TOTAL_CARD_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
            End If
            rsBANKDEPODet.MoveNext
        Loop
    End If
    
    txtTotalCashAmt.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT)
    txtTotalCheckAmt.Text = ToDoubleNumber(TOTAL_CHECK_DEPOSIT)
    txtCardDeposit.Text = ToDoubleNumber(TOTAL_CARD_DEPOSIT)
    txtTotalDepositedAmount.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT + TOTAL_CHECK_DEPOSIT + TOTAL_CARD_DEPOSIT)
End Sub

Sub ShowGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                            As New ADODB.Recordset
    Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_BankDepo Where ID = " & XXX)
    If Not (rsBANKDEPO_Details.EOF And rsBANKDEPO_Details.BOF) Then
        txtBankDeposit = Null2Date(rsBANKDEPO_Details!datdeposit)
        txtCheckDate.Text = Null2String(rsBANKDEPO_Details!CheckDate)
        txtTseklase.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
        txtCheckNum.Text = Null2String(rsBANKDEPO_Details!CheckNum)
    Else
        txtCheckDate.Text = "": txtTseklase.Text = ""
        txtCheckNum.Text = ""
    End If
    Set rsBANKDEPO_Details = Nothing
End Sub

Sub StoreGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                            As ADODB.Recordset
    Set rsBANKDEPO_Details = New ADODB.Recordset
    Set rsBANKDEPO_Details = gconDMIS.Execute("Select * from CMIS_BankDepo Where ID = " & XXX)
    If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
        labBankDepoID.Caption = rsBANKDEPO_Details!Id
        If Null2String(rsBANKDEPO_Details!Type) = "1" Then
            cboType.Text = "CASH"
            PREV_CASH_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!Deposit)
            PREV_CHECK_DEPOSIT = 0
            PREV_CARD_DEPOSIT = 0
        End If
        If Null2String(rsBANKDEPO_Details!Type) = "2" Then
            cboType.Text = "CHECK"
            PREV_CASH_DEPOSIT = 0
            PREV_CHECK_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!Deposit)
            PREV_CARD_DEPOSIT = 0
        End If
        If Null2String(rsBANKDEPO_Details!Type) = "3" Then
            cboType.Text = "CARD"
            PREV_CASH_DEPOSIT = 0
            PREV_CHECK_DEPOSIT = 0
            PREV_CARD_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!Deposit)
        End If
        If Null2String(rsBANKDEPO_Details!bankcode) <> "" Then
            If SetBankName(Null2String(rsBANKDEPO_Details!bankcode)) <> "" Then
                cboBankCode.Text = SetBankName(Null2String(rsBANKDEPO_Details!bankcode))
            Else
                If cboBankCode.Text = SetCustomerName(Null2String(rsBANKDEPO_Details!bankcode)) <> "" Then
                    cboBankCode.Text = SetCustomerName(Null2String(rsBANKDEPO_Details!bankcode))
                Else
                    cboBankCode.ListIndex = -1
                End If
            End If
        Else
            cboBankCode.ListIndex = -1
        End If
        txtTimDeposit.Text = Null2String(rsBANKDEPO_Details!timdeposit)
        txtDatDeposit.Text = Null2String(rsBANKDEPO_Details!datdeposit)
        txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!Deposit))
        txtCheckDte.Text = Null2Date(rsBANKDEPO_Details!CheckDate)
        txtCheckType.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
        txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!CheckNum)
        txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)

        If Format(CDate(txtDatDeposit), "MM/DD/YYYY") = Format(CURRENT_CUTOFF_DATE, "MM/DD/YYYY") Then
            cmdDeleteBANKDEPO.Enabled = True
            cmdSaveBANKDEPO.Enabled = True
        Else
            cmdDeleteBANKDEPO.Enabled = False
            cmdSaveBANKDEPO.Enabled = False
        End If
    End If
    Set rsBANKDEPO_Details = Nothing
End Sub

Sub InitGrid()
    cleargrid grdBankDepo
    grdBankDepo.FormatString = " Type          |   Bank Name / Customer Name                       |    Time       | Amount Deposit "
    grdBankDepo.ColWidth(4) = 1
End Sub

Sub InitTransactionsGrid()
    cleargrid grdCheckCardTransactions
    grdCheckCardTransactions.FormatString = " Code         |   Bank Name                                        |    Time       | Check Amount  "
    grdCheckCardTransactions.ColWidth(4) = 1
End Sub

Sub InitCbo()
    cboType.Clear
    cboType.AddItem "CASH"
    cboType.AddItem "CHECK"
    cboType.AddItem "CARD"
    cboCheckTransactions.Clear
    cboCheckTransactions.AddItem "Cashier Collection"
    cboCheckTransactions.AddItem "Check Encashment"
    'cboCheckTransactions.AddItem "Petty Cash Fund Replenishment"
    'cboCheckTransactions.AddItem "LTO Registration Replenishment"
    'cboCheckTransactions.AddItem "Payment of Cash Advances"
    Dim rsBANK                                                        As ADODB.Recordset
    Set rsBANK = New ADODB.Recordset
    Set rsBANK = gconDMIS.Execute("Select BANKNAME from ALL_BANKS order by BANKNAME ASC")
    If Not rsBANK.EOF And Not rsBANK.BOF Then
        Combo_Loadval cboBankCode, rsBANK
    End If
    Set rsBANK = New ADODB.Recordset
    Set rsBANK = gconDMIS.Execute("Select BANKNAME from ALL_BANKS order by BANKNAME ASC")
    If Not rsBANK.EOF And Not rsBANK.BOF Then
        Combo_Loadval cboDeposit_To, rsBANK
    End If
    Set rsBANK = Nothing
End Sub

Sub initMemvars()
    If AddorEdit = "ADD" Then
        txtDatDeposit.Text = LOGDATE
    Else
        txtDatDeposit.Text = ""
    End If
    txtTotalCashAmt.Text = "0.00"
    txtTotalCheckAmt.Text = "0.00"
    txtCheckDate.Text = ""
    txtTseklase.Text = ""
    txtTotalDepositedAmount.Text = "0.00"
    txtCheckNum.Text = ""
End Sub

Sub InitBankDepoMemVars()
    txtBankDeposit = LOGDATE
    cboType.ListIndex = -1
    cboBankCode.Enabled = False
    cboBankCode.ListIndex = -1
    'cboDeposit_To.Enabled = False
    'cboDeposit_To.ListIndex = -1
    txtTimDeposit.Enabled = False
    txtTimDeposit.Text = ""
    txtDeposit.Enabled = False
    txtDeposit.Text = "0.00"

    InitTransactionsGrid
    labTranID.Caption = "": txtBankDeposit = LOGDATE
    txtCheckDate.Text = "": txtCheckType.Text = ""
    txtCheckNumber.Text = "": txtORNumber.Text = ""
    txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = ""
End Sub

Sub FillGrid()
    Dim BankDeposit                                                   As ADODB.Recordset
    lstBANKDEPO.Sorted = False: lstBANKDEPO.ListItems.Clear
    lstBANKDEPO.Enabled = False
    Set BankDeposit = New ADODB.Recordset
    Set BankDeposit = gconDMIS.Execute("select  DISTINCT top 50 DATDEPOSIT from CMIS_BankDepo Where DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' order by DATDEPOSIT desc")
    If Not (BankDeposit.EOF And BankDeposit.BOF) Then
        lstBANKDEPO.Enabled = True
        Listview_Loadval Me.lstBANKDEPO.ListItems, BankDeposit
        lstBANKDEPO.Refresh
        lstBANKDEPO.Enabled = True
    Else
        lstBANKDEPO.Enabled = False
    End If

    Set BankDeposit = Nothing
End Sub

Sub FillSearchGrid(XXX As Variant)
    Dim BankDeposit                                                   As New ADODB.Recordset
    lstBANKDEPO.Sorted = False: lstBANKDEPO.ListItems.Clear
    lstBANKDEPO.Enabled = False
    
    XXX = Repleys(LTrim(RTrim(XXX)))
    If XXX = "" Then
        Set BankDeposit = gconDMIS.Execute("select DISTINCT DATDEPOSIT from CMIS_BankDepo Where DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' order by DATDEPOSIT desc")
    Else
        Set BankDeposit = gconDMIS.Execute("select DISTINCT DATDEPOSIT from CMIS_BankDepo where DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' AND DATDEPOSIT like '" & XXX & "%' order by DATDEPOSIT desc")
    End If
    If Not (BankDeposit.EOF And BankDeposit.BOF) Then
        lstBANKDEPO.Enabled = True
        Listview_Loadval Me.lstBANKDEPO.ListItems, BankDeposit
        lstBANKDEPO.Refresh
        lstBANKDEPO.Enabled = True
    Else
        lstBANKDEPO.Enabled = False
    End If

    Set BankDeposit = Nothing
End Sub

Private Sub cboBankCode_GotFocus()
    VBComBoBoxDroppedDown cboBankCode
End Sub

Private Sub cboCheckTransactions_Click()
    Call InitTransactionsGrid
    Dim rsCHECKDet                                                    As New ADODB.Recordset
    Dim I                                                             As Long
    Dim ITEM                                                          As ListItem

    If cboType.Text = "CARD" Then
        Set rsCHECKDet = gconDMIS.Execute("Select * from CMIS_Off_hd where Deposit = 0 and CardAmount > 0 AND CANCEL = 0 Order by ID asc")
        If Not (rsCHECKDet.EOF And rsCHECKDet.BOF) Then
            rsCHECKDet.MoveFirst
            Call cleargrid(grdCheckCardTransactions)
            'grdCheckCardTransactions.FormatString = " OR NUM. | Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            'grdCheckCardTransactions.ColWidth(5) = 1

            grdCheckCardTransactions.FormatString = "  Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            grdCheckCardTransactions.ColWidth(4) = 1
            Do While Not rsCHECKDet.EOF
                I = I + 1
                'grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!OR_NUM) & Chr(9) & Null2String(rsCHECKDet!CUSCDE) & Chr(9) & " " & Null2String(rsCHECKDet!cusname) & Chr(9) & Null2String(rsCHECKDet!datecreate) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!cardamount)) & Chr(9) & rsCHECKDet!Id
                grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!CUSCDE) & _
                    Chr(9) & " " & Null2String(rsCHECKDet!CUSNAME) & _
                    Chr(9) & Null2String(rsCHECKDet!DATECREATE) & _
                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!CARDAMOUNT)) & _
                    Chr(9) & rsCHECKDet!Id

                If I = 1 Then grdCheckCardTransactions.RemoveItem 1
                'If Null2String(rsBANKDEPODet!Type) = "1" Then
                '   TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                'End If
                'If Null2String(rsBANKDEPODet!Type) = "2" Then
                '   TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                'End If
                rsCHECKDet.MoveNext
            Loop
            
            'UPDATE BY   : MJP 06012009 1056AM
            'DESCRIPTION : DISPLAY IN LISTVIEW
                rsCHECKDet.MoveFirst
                lsvTran.ListItems.Clear
                Do While Not rsCHECKDet.EOF
                    Set ITEM = lsvTran.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                    ITEM.SubItems(1) = Null2String(rsCHECKDet!CUSNAME)
                    ITEM.SubItems(2) = Null2String(rsCHECKDet!DATECREATE)
                    ITEM.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CARDAMOUNT))
                    ITEM.SubItems(4) = rsCHECKDet!Id
                    
                    rsCHECKDet.MoveNext
                Loop
            'DESCRIPTION : DISPLAY IN LISTVIEW
        End If
    ElseIf cboType.Text = "CASH" Then
        Set rsCHECKDet = gconDMIS.Execute("Select * from CMIS_Off_hd where Deposit = 0 and CashAmount > 0 AND CANCEL = 0 Order by ID asc")
        If Not (rsCHECKDet.EOF And rsCHECKDet.BOF) Then
            rsCHECKDet.MoveFirst
            Call cleargrid(grdCheckCardTransactions)
            'grdCheckCardTransactions.FormatString = " OR NUM. | Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            'grdCheckCardTransactions.ColWidth(5) = 1

            grdCheckCardTransactions.FormatString = " Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            grdCheckCardTransactions.ColWidth(4) = 1
            Do While Not rsCHECKDet.EOF
                I = I + 1
                'grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!OR_NUM) & Chr(9) & Null2String(rsCHECKDet!CUSCDE) & Chr(9) & " " & Null2String(rsCHECKDet!cusname) & Chr(9) & Null2String(rsCHECKDet!datecreate) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount)) & Chr(9) & rsCHECKDet!Id
                grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!CUSCDE) & _
                    Chr(9) & " " & Null2String(rsCHECKDet!CUSNAME) & _
                    Chr(9) & Null2String(rsCHECKDet!DATECREATE) & _
                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount)) & _
                    Chr(9) & rsCHECKDet!Id

                If I = 1 Then grdCheckCardTransactions.RemoveItem 1
                'If Null2String(rsBANKDEPODet!Type) = "1" Then
                '   TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                'End If
                'If Null2String(rsBANKDEPODet!Type) = "2" Then
                '   TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                'End If
                rsCHECKDet.MoveNext
            Loop
            
            'UPDATE BY   : MJP 06012009 1056AM
            'DESCRIPTION : DISPLAY IN LISTVIEW
                rsCHECKDet.MoveFirst
                lsvTran.ListItems.Clear
                Do While Not rsCHECKDet.EOF
                    Set ITEM = lsvTran.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                    ITEM.SubItems(1) = Null2String(rsCHECKDet!CUSNAME)
                    ITEM.SubItems(2) = Null2String(rsCHECKDet!DATECREATE)
                    ITEM.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount))
                    ITEM.SubItems(4) = rsCHECKDet!Id
                    
                    rsCHECKDet.MoveNext
                Loop
            'DESCRIPTION : DISPLAY IN LISTVIEW
        End If
    End If

    Set rsCHECKDet = New ADODB.Recordset
    If cboCheckTransactions.Text = "Cashier Collection" Then
        If cboType.Text = "CHECK" Then
            'cmdComputeCard.Visible = True

            Call cleargrid(grdCheckCardTransactions)
            'grdCheckCardTransactions.FormatString = " OR NUM. | Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            'grdCheckCardTransactions.ColWidth(5) = 1

            grdCheckCardTransactions.FormatString = " Cust. Code   |   Cust. Name                                           |    Date        | Card Amount   "
            grdCheckCardTransactions.ColWidth(4) = 1

            Set rsCHECKDet = gconDMIS.Execute("Select * from CMIS_Off_hd where Deposit = 0 and chkAmount > 0 AND CANCEL = 0 Order by ID asc")
            If Not (rsCHECKDet.EOF And rsCHECKDet.BOF) Then
                rsCHECKDet.MoveFirst
                Do While Not rsCHECKDet.EOF
                    I = I + 1
                    'grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!OR_NUM) & Chr(9) & Null2String(rsCHECKDet!bankcode) & Chr(9) & " " & SetBankName(Null2String(rsCHECKDet!bankcode)) & Chr(9) & Null2String(rsCHECKDet!TimeCreate) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!ChkAmount)) & Chr(9) & rsCHECKDet!Id
                    grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!bankcode) & _
                        Chr(9) & " " & SetBankName(Null2String(rsCHECKDet!bankcode)) & _
                        Chr(9) & Null2String(rsCHECKDet!TimeCreate) & _
                        Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!ChkAmount)) & _
                        Chr(9) & rsCHECKDet!Id

                    If I = 1 Then grdCheckCardTransactions.RemoveItem 1
                    'If Null2String(rsBANKDEPODet!Type) = "1" Then
                    '   TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                    'End If
                    'If Null2String(rsBANKDEPODet!Type) = "2" Then
                    '   TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                    'End If
                    rsCHECKDet.MoveNext
                Loop
                
                'UPDATE BY   : MJP 06012009 1056AM
                'DESCRIPTION : DISPLAY IN LISTVIEW
                    rsCHECKDet.MoveFirst
                    lsvTran.ListItems.Clear
                    Do While Not rsCHECKDet.EOF
                        Set ITEM = lsvTran.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                        ITEM.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                        ITEM.SubItems(2) = Null2String(rsCHECKDet!TimeCreate)
                        ITEM.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!ChkAmount))
                        ITEM.SubItems(4) = rsCHECKDet!Id
                        
                        rsCHECKDet.MoveNext
                    Loop
                'DESCRIPTION : DISPLAY IN LISTVIEW
            End If
        End If
    ElseIf cboCheckTransactions.Text = "Check Encashment" Then
        If cboType.Text = "CHECK" Then
            Set rsCHECKDet = gconDMIS.Execute("Select * from CMIS_InCash where Deposit = 0 and chkAmount > 0 Order by ID asc")
            If Not (rsCHECKDet.EOF And rsCHECKDet.BOF) Then
                rsCHECKDet.MoveFirst

                Call cleargrid(grdCheckCardTransactions)
                Do While Not rsCHECKDet.EOF
                    I = I + 1

                    grdCheckCardTransactions.AddItem " " & Null2String(rsCHECKDet!bankcode) & _
                        Chr(9) & " " & SetBankName(Null2String(rsCHECKDet!bankcode)) & _
                        Chr(9) & Null2String(rsCHECKDet!TimeCreate) & _
                        Chr(9) & ToDoubleNumber(N2Str2Zero(rsCHECKDet!ChkAmount)) & _
                        Chr(9) & rsCHECKDet!Id

                    If I = 1 Then grdCheckCardTransactions.RemoveItem 1
                    'If Null2String(rsBANKDEPODet!Type) = "1" Then
                    '   TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                    'End If
                    'If Null2String(rsBANKDEPODet!Type) = "2" Then
                    '   TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!Deposit)
                    'End If
                    rsCHECKDet.MoveNext
                Loop
                
                'UPDATE BY   : MJP 06012009 1056AM
                'DESCRIPTION : DISPLAY IN LISTVIEW
                    rsCHECKDet.MoveFirst
                    lsvTran.ListItems.Clear
                    Do While Not rsCHECKDet.EOF
                        Set ITEM = lsvTran.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                        ITEM.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                        ITEM.SubItems(2) = Null2String(rsCHECKDet!TimeCreate)
                        ITEM.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!ChkAmount))
                        ITEM.SubItems(4) = rsCHECKDet!Id
                        
                        rsCHECKDet.MoveNext
                    Loop
                'DESCRIPTION : DISPLAY IN LISTVIEW
            End If
        End If
    End If
End Sub

Private Sub cboCheckTransactions_GotFocus()
    VBComBoBoxDroppedDown cboCheckTransactions
End Sub

Private Sub cboDeposit_To_Click()
    cmdShow.Value = True
End Sub

Private Sub cboDeposit_To_GotFocus()
    VBComBoBoxDroppedDown cboDeposit_To
End Sub

Private Sub cboType_Change()
    Call SetSelectedType
End Sub

Private Sub cboType_Click()
    Call SetSelectedType
End Sub

Private Sub cboType_GotFocus()
    InitTransactionsGrid
    labTranID.Caption = ""
    txtCheckDate.Text = "": txtCheckType.Text = ""
    txtCheckNumber.Text = "": txtORNumber.Text = ""
    txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
    VBComBoBoxDroppedDown cboType
End Sub

Private Sub cmd1_Click()
    picLSV.ZOrder 0
End Sub

Private Sub cmd2_Click()
    picLSV.ZOrder 1
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub

    AddorEdit = "ADD"
    picBankDepo.Visible = True: picBankDepo.ZOrder 0
    cmdDeleteBANKDEPO.Enabled = False
    InitBankDepoMemVars
    Picture5.Enabled = False
    fraDetails.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    lstBANKDEPO.Enabled = True
    textSearch.Enabled = True
End Sub

Private Sub cmdCancelBANKDEPO_Click()
    AddorEdit = ""
    picBankDepo.Visible = False: picBankDepo.ZOrder 1
    'StoreMemvars
    'FillGrid
    lstBANKDEPO.Enabled = True
    fraDetails.Enabled = True
    Picture5.Enabled = True
End Sub

Private Sub cmdDeleteBANKDEPO_Click()
    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:
    If ShowConfirmDelete = True Then
        Dim rsJoyDeposit                                              As ADODB.Recordset
        Set rsJoyDeposit = New ADODB.Recordset
        Set rsJoyDeposit = gconDMIS.Execute("Select * from CMIS_BankDepo Where ID = " & labBankDepoID.Caption)
        If Not rsJoyDeposit.EOF And Not rsJoyDeposit.BOF Then
            gconDMIS.Execute ("delete from CMIS_BankDepo Where ID = " & labBankDepoID.Caption)
            gconDMIS.Execute ("update CMIS_Off_Hd Set Deposit = 0 Where OR_NUM = " & N2Str2Null(rsJoyDeposit!OR_NUM))
            If cboType.Text = "CASH" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                                " CASH = CASH + " & NumericVal(txtDeposit.Text) & "," & _
                                " CASHDEPO = CASHDEPO - " & NumericVal(txtDeposit.Text) & _
                                " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If cboType.Text = "CHECK" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                                " [CHECK] = [CHECK] + " & NumericVal(txtDeposit.Text) & "," & _
                                " CHECKDEPO = CHECKDEPO - " & NumericVal(txtDeposit.Text) & _
                                " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If cboType.Text = "CARD" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                                " CARD = CARD + " & NumericVal(txtDeposit.Text) & "," & _
                                " CARDDEPO = CARDDEPO - " & NumericVal(txtDeposit.Text) & _
                                " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            ShowDeletedMsg
        End If
    End If
    rsRefresh
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then rsBANKDEPO.MoveLast
    cmdCancelBANKDEPO_Click
    On Error Resume Next
    rsBANKDEPO.Find "DatDeposit = " & N2Date2Null(txtDatDeposit.Text)
    StoreMemvars

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
    grdBankDepo.Col = 4
    
    If grdBankDepo.Text <> "" Then
        AddorEdit = "EDIT"
        picBankDepo.Visible = True
        picBankDepo.ZOrder 0
        cmdDeleteBANKDEPO.Enabled = True
        
        Call StoreGridDetails(grdBankDepo.Text)
        Picture5.Enabled = False
        fraDetails.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
    Picture5.Enabled = True
    fraDetails.Enabled = True
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:
    'Exit Sub
    'ErrorCode:
    '    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:

    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    If IsDate(txtDatDeposit.Text) = False Then
        MsgBox "Pls click Date Deposited.", vbInformation, "Check Date"
        Exit Sub
    End If

    Screen.MousePointer = 11
    With rptBankDepo
        .Formulas(0) = "DEALER_NAME = '" & COMPANY_NAME & "'"
        .Formulas(1) = "DEALER_ADDRESS = '" & COMPANY_ADDRESS & "'"
        .Formulas(2) = "PREPAREDBY= '" & PreparedBy & "'"
        .Formulas(3) = "NOTEDBY= '" & NotedBy & "'"
        .Formulas(4) = "CHECKEDBY= '" & CheckedBy & "'"
        .Formulas(5) = "PRINTEDBY= " & N2Str2Null(LOGNAME)
    End With
    'original
    'PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit.rpt", "{BankDepo.DatDeposit} = Date(" & Year(txtDatDeposit.Text) & "," & Month(txtDatDeposit.Text) & "," & Day(txtDatDeposit.Text) & ") AND {BankDepo.DEPOSIT_TO} = '" & SetBankCode(cboDeposit_To) & "'", CMIS_REPORT_CONNECTION, 1
    
    'all bank deposited
    PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit.rpt", "{BankDepo.DatDeposit} = Date(" & Year(txtDatDeposit.Text) & "," & Month(txtDatDeposit.Text) & "," & Day(txtDatDeposit.Text) & ")", CMIS_REPORT_CONNECTION, 1
    
    'PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit.rpt", "{BankDepo.DatDeposit} = Date(" & Year(txtDatDeposit.Text) & "," & Month(txtDatDeposit.Text) & "," & Day(txtDatDeposit.Text) & ") AND {BankDepo.DEPOSIT_TO} = '" & SetBankCode(cboDeposit_To) & "'", CMIS_REPORT_CONNECTION, 1
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveBANKDEPO_Click()
    'On Error GoTo Errorcode
    Dim vBankCode                                                     As String
    Dim vTseklase                                                     As String
    Dim vDeposit                                                      As String
    Dim vDatDeposit                                                   As String
    Dim vTimDeposit                                                   As String
    Dim vWhoDeposit                                                   As String
    Dim VTYPE                                                         As String

    Dim vInCashChk                                                    As Integer
    Dim vCollectChk                                                   As Integer
    Dim vP_pay_Chk                                                    As Integer
    Dim vL_pay_Chk                                                    As Integer
    Dim vU_pay_Chk                                                    As Integer
    Dim vA_pay_Chk                                                    As Integer

    Dim vOR_NUM                                                       As String
    Dim vDeposit_To                                                   As String
    Dim vCheckDate                                                    As String
    Dim vCardDate                                                     As String
    Dim vCheckNum                                                     As String
    Dim vCardNumber                                                   As String

    Dim RSTMP As New ADODB.Recordset
    Dim xMODULENAME As String
    
    If cboType.Text = "CASH" Then
        vBankCode = "NULL"
        vTseklase = "NULL"
        VTYPE = "'1'"
        vCheckDate = "NULL"
        vCheckNum = "NULL"
        vCardDate = "NULL"
        vCardNumber = "NULL"
        vOR_NUM = "NULL"
        vDeposit = NumericVal(txtDeposit.Text)
        vDatDeposit = N2Date2Null(txtBankDeposit)
        vTimDeposit = N2Str2Null(txtTimDeposit.Text)
        vWhoDeposit = "'00005'"
    End If

    If cboType.Text = "CHECK" Then
        If cboCheckTransactions.Text = "Cashier Collection" Then
            If Trim(txtORNumber.Text) = "" Then
                MsgBox "Please select Check Collections to deposit", vbInformation, "Nothing to Deposit"
                Exit Sub
            End If
        End If
        vBankCode = N2Str2Null(txtBankCode.Text)
        VTYPE = "'2'"
        vTseklase = N2Str2Null(SetCheckClassCode(txtCheckType.Text))
        vCheckDate = N2Str2Null(txtCheckDte.Text)
        vCheckNum = N2Str2Null(txtCheckNumber.Text)
        vOR_NUM = N2Str2Null(txtORNumber.Text)
        'vDeposit = NumericVal(txtCheckAmount.Text)
        vDeposit = NumericVal(txtDeposit.Text)
        vDatDeposit = N2Date2Null(txtBankDeposit)
        vTimDeposit = N2Str2Null(txtTimeCreate.Text)
        vWhoDeposit = "'00005'"
        vCardDate = "NULL"
        vCardNumber = "NULL"
    End If

    If cboType.Text = "CARD" Then
        If Trim(txtORNumber.Text) = "" Then
            MsgBox "Please select Card Collections to deposit", vbInformation, "Nothing to Deposit"
            Exit Sub
        End If
        'vBankCode = N2Str2Null(SetBankCode(cboBankCode.Text))

        'If cboType.Text = "Cash" Then
        '   grdCheckCardTransactions.Col = 1: vBankCode = N2Str2Null(Trim(grdCheckCardTransactions.Text))
        'Else
        grdCheckCardTransactions.Col = 0: vBankCode = N2Str2Null(Trim(grdCheckCardTransactions.Text))
        'End If

        VTYPE = "'3'"
        vTseklase = "NULL"
        vCheckDate = "NULL"
        vCheckNum = "NULL"
        vCardDate = N2Str2Null(txtCheckDte.Text)
        vCardNumber = N2Str2Null(txtCheckNumber.Text)
        vOR_NUM = N2Str2Null(txtORNumber.Text)
        vDeposit = NumericVal(txtDeposit.Text)
        vDatDeposit = N2Date2Null(txtBankDeposit)
        vTimDeposit = N2Str2Null(txtTimeCreate.Text)
        vWhoDeposit = N2Str2Null(LOGCODE)
    End If

    If cboType.Text = "CASH" Then
        If Trim(txtORNumber.Text) = "" Then
            MsgBox "Please select Card Collections to deposit", vbInformation, "Nothing to Deposit"
            Exit Sub
        End If
        'vBankCode = N2Str2Null(SetBankCode(cboBankCode.Text))

        grdCheckCardTransactions.Col = 0: vBankCode = N2Str2Null(Trim(grdCheckCardTransactions.Text))
        'grdCheckCardTransactions.Col = 1: vBankCode = N2Str2Null(Trim(grdCheckCardTransactions.Text))

        VTYPE = "'1'"
        vTseklase = "NULL"
        vCheckDate = "NULL"
        vCheckNum = "NULL"
        vCardDate = "NULL"
        vCardNumber = "NULL"
        vOR_NUM = N2Str2Null(txtORNumber.Text)
        vDeposit = NumericVal(txtDeposit.Text)
        vDatDeposit = N2Date2Null(txtBankDeposit)
        vTimDeposit = N2Str2Null(txtTimeCreate.Text)
        vWhoDeposit = N2Str2Null(LOGCODE)
    End If

    vInCashChk = 0
    If cboCheckTransactions.Text = "Cashier Collection" Then
        vCollectChk = 1
    Else
        vCollectChk = 0
    End If
    vP_pay_Chk = 0
    vL_pay_Chk = 0
    vU_pay_Chk = 0
    vA_pay_Chk = 0

    If Trim(cboDeposit_To.Text) = "" Then
        MsgBox "Pls. Select where to deposit...", vbInformation, "Bank Not Selected"
        Exit Sub
    End If
    vDeposit_To = N2Str2Null(SetBankCode(cboDeposit_To.Text))

    If AddorEdit = "ADD" Then
         SQL_STATEMENT = "Insert into CMIS_BankDepo " & _
                        "(BankCode, Tseklase, Deposit, DatDeposit, TimDeposit, WhoDeposit, [Type], InCashChk, CollectChk, P_pay_Chk, L_pay_Chk, U_pay_Chk, A_pay_Chk, OR_Num, Deposit_To, CheckDate, CardDate, CheckNum, CardNumber)" & _
                        " values (" & vBankCode & _
                        "," & vTseklase & _
                        "," & vDeposit & _
                        "," & vDatDeposit & _
                        "," & vTimDeposit & _
                        "," & vWhoDeposit & _
                        "," & VTYPE & _
                        "," & vInCashChk & _
                        "," & vCollectChk & _
                        "," & vP_pay_Chk & _
                        "," & vL_pay_Chk & _
                        "," & vU_pay_Chk & _
                        "," & vA_pay_Chk & _
                        "," & vOR_NUM & _
                        "," & vDeposit_To & _
                        "," & vCheckDate & _
                        "," & vCardDate & _
                        "," & vCheckNum & _
                        "," & vCardNumber & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("A", "TRANSACTION BANKDEPOSIT", SQL_STATEMENT, lblBANKID, "", "OR NO: ", "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        'If cboCheckTransactions.Text = "Cashier Collection" Then
        'End If
        If cboCheckTransactions.Text = "Check Encashment" Then
            gconDMIS.Execute ("update CMIS_InCash Set Deposit = 1 Where ID = " & labTranID.Caption)
        Else
            If labTranID.Caption <> "" Then
                SQL_STATEMENT = "update CMIS_Off_Hd Set Deposit = 1 Where ID = " & labTranID.Caption
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------
                    Set RSTMP = gconDMIS.Execute("SELECT VAT,OR_NUM FROM CMIS_OFF_HD WHERE ID = " & labTranID.Caption & "")
                    If Not (RSTMP.BOF And RSTMP.EOF) Then
                        If Null2String(RSTMP!vat) = "1" Then xMODULENAME = "TRANSACTION O.R. WITH VAT"
                        If Null2String(RSTMP!vat) = "0" Then xMODULENAME = "TRANSACTION O.R. WITHOUT VAT"
                        
                        Call NEW_LogAudit("E", xMODULENAME, SQL_STATEMENT, lblBANKID, "", "OR NO: " & Null2String(RSTMP!OR_NUM), "", "")
                    End If
                    Set RSTMP = Nothing
                'NEW LOG AUDIT---------------------------------------------------------
            End If
        End If
        If cboType.Text = "CASH" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CASH = CASH - " & vDeposit & "," & _
                            " CASHDEPO = CASHDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If cboType.Text = "CHECK" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " [CHECK] = [CHECK] - " & vDeposit & "," & _
                            " CHECKDEPO = CHECKDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If cboType.Text = "CARD" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CARD = CARD - " & vDeposit & "," & _
                            " CARDDEPO = CARDDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        
        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CMIS_BankDepo Set " & _
                        " BankCode = " & vBankCode & "," & _
                        " Tseklase = " & vTseklase & "," & _
                        " Deposit = " & vDeposit & "," & _
                        " DatDeposit = " & vDatDeposit & "," & _
                        " TimDeposit = " & vTimDeposit & "," & _
                        " WhoDeposit = " & vWhoDeposit & "," & _
                        " Type = " & VTYPE & "," & _
                        " InCashChk = " & vInCashChk & "," & _
                        " CollectChk = " & vCollectChk & "," & _
                        " P_pay_Chk = " & vP_pay_Chk & "," & _
                        " L_pay_Chk = " & vL_pay_Chk & "," & _
                        " U_pay_Chk = " & vU_pay_Chk & "," & _
                        " A_pay_Chk = " & vA_pay_Chk & "," & _
                        " OR_Num = " & vOR_NUM & "," & _
                        " Deposit_To = " & vDeposit_To & "," & _
                        " CheckDate = " & vCheckDate & "," & _
                        " CardDate = " & vCardDate & "," & _
                        " CheckNum = " & vCheckNum & "," & _
                        " CardNumber = " & vCardNumber & _
                        " Where ID = " & labBankDepoID.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("E", "TRANSACTION BANKDEPOSIT", SQL_STATEMENT, lblBANKID, "", "OR NO: ", "", labBankDepoID.Caption)
        'NEW LOG AUDIT---------------------------------------------------------
        
        If cboCheckTransactions.Text = "Cashier Collection" Then
            If labTranID.Caption <> "" Then
                SQL_STATEMENT = "update CMIS_Off_Hd Set Deposit = 1 Where ID = " & labTranID.Caption
                gconDMIS.Execute SQL_STATEMENT
                
                'NEW LOG AUDIT---------------------------------------------------------
                    Set RSTMP = gconDMIS.Execute("SELECT VAT,OR_NUM FROM CMIS_OFF_HD WHERE ID = " & labTranID.Caption & "")
                    If Not (RSTMP.BOF And RSTMP.EOF) Then
                        If Null2String(RSTMP!vat) = "1" Then xMODULENAME = "TRANSACTION O.R. WITH VAT"
                        If Null2String(RSTMP!vat) = "0" Then xMODULENAME = "TRANSACTION O.R. WITHOUT VAT"
                        
                        Call NEW_LogAudit("E", xMODULENAME, SQL_STATEMENT, lblBANKID, "", "OR NO: " & Null2String(RSTMP!OR_NUM), "", "")
                    End If
                    Set RSTMP = Nothing
                'NEW LOG AUDIT---------------------------------------------------------
            End If
        End If
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                        " CASH = (CASH + " & PREV_CASH_DEPOSIT & ")," & _
                        " CARD = (CARD + " & PREV_CARD_DEPOSIT & ")," & _
                        " [CHECK] = ([CHECK] + " & PREV_CHECK_DEPOSIT & ")," & _
                        " CASHDEPO = (CASHDEPO - " & PREV_CASH_DEPOSIT & ")," & _
                        " CARDDEPO = (CARDDEPO - " & PREV_CARD_DEPOSIT & ")," & _
                        " CHECKDEPO = (CHECKDEPO - " & PREV_CHECK_DEPOSIT & _
                        ") where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                        
        If cboType.Text = "CASH" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CASH = CASH - " & vDeposit & "," & _
                            " CASHDEPO = CASHDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If cboType.Text = "CHECK" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " [CHECK] = [CHECK] - " & vDeposit & "," & _
                            " CHECKDEPO = CHECKDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If cboType.Text = "CARD" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CARD = CARD - " & vDeposit & "," & _
                            " CARDDEPO = CARDDEPO + " & vDeposit & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        
        Call ShowSuccessFullyUpdated
    End If

    Call rsRefresh
    
    On Error Resume Next
    Call StoreDetails
    
    rsBANKDEPO.Find "DATDEPOSIT = '" & txtBankDeposit & "'"
    Call cmdCancelBANKDEPO_Click
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdShow_Click()
    lblBANKID.Caption = FindTransactionID(N2Str2Null(cboDeposit_To), "BANKNAME", "ALL_BANKS")
    Call rsRefresh
    Call cmdCancelBANKDEPO_Click
    Call FillSearchGrid("")
    grdBankDepo.Rows = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            cmdAdd.Value = True
        Case vbKeyF11
            Shell "calc.exe"
        Case vbKeyEscape
            cmdCancelBANKDEPO_Click
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Not cboDeposit_To.Text = "" Then
                Unload frmALL_AuditInquiry
                 
                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION BANKDEPOSIT)"
                Call frmALL_AuditInquiry.DisplayHistory(lblBANKID, "TRANSACTION BANKDEPOSIT")
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Dim rsProfile                                                     As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    
    Set rsProfile = Nothing:
    CenterMe frmMain, Me, 1: initMemvars
    
    cmdPost.Enabled = False:    'cmdPrint.Enabled = False
    cmdPost.Caption = "": cmdPost.Picture = LoadPicture("")
    'cmdPrint.Caption = "": 'cmdPrint.Picture = LoadPicture("")
    
    textSearch.Text = "": InitCbo: InitGrid
    cboDeposit_To.ListIndex = 0
    picBankDepo.Visible = False: picBankDepo.ZOrder 1
    Screen.MousePointer = 0
End Sub

Private Sub grdBANKDEPO_Click()
    grdBankDepo.Col = 4
    If grdBankDepo.Text <> "" Then
        Call ShowGridDetails(grdBankDepo.Text)
        
        With grdBankDepo
            labBANKDEPO_ID = .TextMatrix(.MouseRow, 4)
        End With
    End If
End Sub

Private Sub grdBankDepo_DblClick()
    'COMMENT BY : MJP 05212009 0513PM
    'TO ADD AN DELETE FUNCTION IN BANK DEPOSIT
        'cmdEdit.Value = True
    'COMMENT BY : MJP 05212009 0513PM
    'Exit Sub
    
    'UPDATE BY   : MJP 05212009 0513PM
    'DESCRIPTION : TO HAVE AN DELETE FUNCTION IN BANK DEPOSIT
        If grdBankDepo.Rows = 1 Then Exit Sub
        
        If MsgBox("Delete this Bank Deposit Entry, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            
        Dim xCHECKTRAN As Integer
        Dim XTYPE As Integer
        Dim xDEPOSIT As Double
        Dim RSTMP As New ADODB.Recordset
        Set RSTMP = gconDMIS.Execute("SELECT DEPOSIT,TYPE, COLLECTCHK, OR_NUM FROM CMIS_BANKDEPO WHERE ID = " & labBANKDEPO_ID & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            labORNUM = Null2String(RSTMP!OR_NUM)
            xCHECKTRAN = RSTMP!COLLECTCHK
            XTYPE = RSTMP!Type
            xDEPOSIT = NumericVal(RSTMP!Deposit)
        End If
        Set RSTMP = Nothing
        Set RSTMP = gconDMIS.Execute("SELECT ID FROM CMIS_OFF_HD WHERE OR_NUM = " & N2Str2Null(labORNUM) & "")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            labHD_ID = RSTMP!Id
        End If
        Set RSTMP = Nothing
        
        SQL_STATEMENT = "DELETE FROM CMIS_BANKDEPO WHERE ID = " & labBANKDEPO_ID & ""
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("X", "TRANSACTION BANKDEPOSIT", SQL_STATEMENT, labBANKDEPO_ID, "", "OR NO: " & labORNUM & "", "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        
        If xCHECKTRAN = 0 Then
            gconDMIS.Execute ("update CMIS_InCash Set Deposit = 0 Where ID = " & labHD_ID.Caption)
        Else
            SQL_STATEMENT = "update CMIS_Off_Hd Set Deposit = 0 Where ID = " & labHD_ID.Caption
            gconDMIS.Execute SQL_STATEMENT
        End If
        
        If XTYPE = 1 Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CASH = CASH + " & xDEPOSIT & "," & _
                            " CASHDEPO = CASHDEPO - " & xDEPOSIT & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If XTYPE = 2 Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " [CHECK] = [CHECK] + " & xDEPOSIT & "," & _
                            " CHECKDEPO = CHECKDEPO - " & xDEPOSIT & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If XTYPE = 3 Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " CARD = CARD + " & xDEPOSIT & "," & _
                            " CARDDEPO = CARDDEPO - " & xDEPOSIT & _
                            " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        
        Call ShowDeletedMsg
        Call StoreDetails
    
    'DESCRIPTION : TO HAVE AN DELETE FUNCTION IN BANK DEPOSIT
    'UPDATE BY   : MJP 05212009 0513PM
End Sub

Private Sub grdCheckCardTransactions_Click()
    grdCheckCardTransactions.Col = 4


    'If cboCheckTransactions.Text = "Check Encashment" Then
    '    If cboType.Text = "Check" Then
    '        grdCheckCardTransactions.Col = 4
    '    End If
    'Else
    '    grdCheckCardTransactions.Col = 5
    'End If

    If grdCheckCardTransactions.Text <> "" Then
        Call ShowTransactionsGridDetails(grdCheckCardTransactions.Text)
        fraDetails.Enabled = False
        Picture5.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    If AddorEdit = "ADD" Then
        txtTimDeposit.Text = Time: DoEvents
    End If
End Sub

Private Sub txtDeposit_GotFocus()
    If NumericVal(txtDeposit.Text) = 0 Then txtDeposit.Text = "" Else txtDeposit.Text = NumericVal(txtDeposit.Text)
End Sub

Private Sub txtDeposit_LostFocus()
    txtDeposit.Text = ToDoubleNumber(txtDeposit.Text)
End Sub

'SEARCH MODULE
Private Sub lstBANKDEPO_GotFocus()
    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    'rsBANKDEPO.Bookmark = rsFind(rsBANKDEPO.Clone, "DATDEPOSIT", lstBANKDEPO.SelectedItem).Bookmark
    StoreDetails
End Sub

Private Sub lstBANKDEPO_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    'rsBANKDEPO.Bookmark = rsFind(rsBANKDEPO.Clone, "DATDEPOSIT", lstBANKDEPO.SelectedItem).Bookmark
    Call StoreDetails
End Sub

Private Sub lstBANKDEPO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstBANKDEPO
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lstBANKDEPO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    'If Trim(textSearch.Text) = "" Then
    '    FillGrid
    'Else
    Call FillSearchGrid(textSearch.Text)
    'End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        On Error Resume Next
        lstBANKDEPO.SetFocus
    End If
End Sub

