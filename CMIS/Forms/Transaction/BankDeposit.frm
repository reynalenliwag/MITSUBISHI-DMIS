VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISBankDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Deposit Data Entry"
   ClientHeight    =   8565
   ClientLeft      =   5280
   ClientTop       =   4020
   ClientWidth     =   12195
   ForeColor       =   &H8000000F&
   Icon            =   "BankDeposit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   12195
   Begin VB.PictureBox Pic_PrintBankDepo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3830
      Left            =   4800
      ScaleHeight     =   3795
      ScaleWidth      =   5430
      TabIndex        =   79
      Top             =   2040
      Visible         =   0   'False
      Width           =   5455
      Begin VB.ComboBox CboBankDepo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   240
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   1440
         Width           =   4995
      End
      Begin VB.PictureBox PicDateRange 
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   4905
         TabIndex        =   84
         Top             =   1860
         Width           =   4970
         Begin MSComCtl2.DTPicker dtp_Frm_BD 
            Height          =   405
            Left            =   870
            TabIndex        =   85
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98041857
            CurrentDate     =   38216
         End
         Begin MSComCtl2.DTPicker dtp_To_BD 
            Height          =   405
            Left            =   3090
            TabIndex        =   86
            Top             =   180
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98041857
            CurrentDate     =   38216
         End
         Begin VB.Label lblTo_BD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   2610
            TabIndex        =   88
            Top             =   240
            Width           =   435
         End
         Begin VB.Label lblFrm_BD 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.CommandButton cmdPrint_Bankdepo 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3690
         MouseIcon       =   "BankDeposit.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Print Report"
         Top             =   2760
         Width           =   765
      End
      Begin VB.CommandButton cmdClose_Bankdepo 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4440
         MouseIcon       =   "BankDeposit.frx":0EBB
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":100D
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Close Window"
         Top             =   2760
         Width           =   765
      End
      Begin VB.OptionButton OptPrintAsof 
         Caption         =   "Option1"
         Height          =   195
         Left            =   240
         TabIndex        =   81
         Top             =   1125
         Width           =   255
      End
      Begin VB.OptionButton OptPrint 
         Caption         =   "Option1"
         Height          =   195
         Left            =   240
         TabIndex        =   80
         Top             =   765
         Width           =   255
      End
      Begin VB.Label lbl_PrintBankDepoAsof 
         Caption         =   "Print Bank Deposit As of:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   92
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lbl_PrintBankDepo 
         Caption         =   "Print Bank Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   91
         Top             =   720
         Width           =   4575
      End
      Begin XtremeShortcutBar.ShortcutCaption Cap_BankDepo 
         Height          =   375
         Left            =   -360
         TabIndex        =   90
         Top             =   150
         Width           =   5785
         _Version        =   655364
         _ExtentX        =   10204
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Bank Deposit"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
   End
   Begin Crystal.CrystalReport rptBankDepo 
      Left            =   0
      Top             =   8070
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
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2650
      ScaleHeight     =   825
      ScaleWidth      =   6495
      TabIndex        =   57
      Top             =   7620
      Width           =   6525
      Begin VB.TextBox txtCheckNum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   4560
         TabIndex        =   60
         Top             =   60
         Width           =   1860
      End
      Begin VB.TextBox txtCheckDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1140
         TabIndex        =   59
         Top             =   60
         Width           =   2025
      End
      Begin VB.TextBox txtTseklase 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1140
         TabIndex        =   58
         Top             =   450
         Width           =   5280
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -750
         TabIndex        =   63
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Check No.  :"
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
         Height          =   255
         Left            =   2670
         TabIndex        =   62
         Top             =   90
         Width           =   1845
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -750
         TabIndex        =   61
         Top             =   90
         Width           =   1845
      End
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
      Left            =   11040
      TabIndex        =   67
      Top             =   150
      Width           =   975
   End
   Begin VB.ComboBox cboDeposit_To 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   450
      Left            =   2580
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   120
      Width           =   8355
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   9250
      ScaleHeight     =   870
      ScaleWidth      =   2880
      TabIndex        =   47
      Top             =   7620
      Width           =   2880
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H80000016&
         Caption         =   "E&xit"
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
         Left            =   2120
         MouseIcon       =   "BankDeposit.frx":1458
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":15AA
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H80000016&
         Caption         =   "&Print"
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
         Left            =   1430
         MouseIcon       =   "BankDeposit.frx":1910
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":1A62
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000016&
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
         Height          =   795
         Left            =   740
         MouseIcon       =   "BankDeposit.frx":1DC8
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":1F1A
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H80000016&
         Caption         =   "&Find"
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
         Left            =   50
         MouseIcon       =   "BankDeposit.frx":222D
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":237F
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   45
      TabIndex        =   12
      Top             =   900
      Width           =   2475
      Begin MSComctlLib.ListView lstBANKDEPO 
         Height          =   6945
         Left            =   60
         TabIndex        =   14
         Top             =   510
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "BankDeposit.frx":2679
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date of Deposit"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   60
         MaxLength       =   35
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   60
         Width           =   2385
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   12930
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   52
      Top             =   8760
      Width           =   1980
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
         Left            =   750
         MouseIcon       =   "BankDeposit.frx":27DB
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":292D
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Cancel"
         Top             =   30
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
         Left            =   60
         MouseIcon       =   "BankDeposit.frx":2C6B
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":2DBD
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2640
      ScaleHeight     =   1245
      ScaleWidth      =   9390
      TabIndex        =   17
      Top             =   6300
      Width           =   9425
      Begin VB.TextBox txtCardDeposit 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2040
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCashAmt 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   60
         Width           =   1815
      End
      Begin VB.TextBox txtTotalCheckAmt 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   2040
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtTotalDepositedAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   620
         Width           =   3465
      End
      Begin VB.Label lblTotalItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   9090
         TabIndex        =   78
         Top             =   30
         Width           =   105
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Card Deposit  :"
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
         Height          =   255
         Left            =   150
         TabIndex        =   46
         Top             =   870
         Width           =   2385
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Deposit  :"
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
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   90
         Width           =   2385
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Deposit  :"
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
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   480
         Width           =   2385
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deposited Amount"
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
         Left            =   7170
         TabIndex        =   22
         Top             =   315
         Width           =   2025
      End
   End
   Begin VB.PictureBox PicBankDepo 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   2640
      ScaleHeight     =   6585
      ScaleWidth      =   9405
      TabIndex        =   26
      Top             =   930
      Width           =   9435
      Begin MSComctlLib.ListView lsvDet 
         Height          =   3825
         Left            =   210
         TabIndex        =   72
         Top             =   1320
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   6747
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
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Bank Name / Customer Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Check Amount"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "OR Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "id"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtSelectedDate 
         Height          =   315
         Left            =   2160
         TabIndex        =   75
         Top             =   5250
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   86704129
         CurrentDate     =   40211
      End
      Begin VB.TextBox txtBankDeposit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   60
         Width           =   1815
      End
      Begin VB.TextBox txtORSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   69
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteBANKDEPO 
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
         Left            =   210
         MouseIcon       =   "BankDeposit.frx":310D
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":325F
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   5670
         Width           =   705
      End
      Begin VB.ComboBox cboCheckTransactions 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         ItemData        =   "BankDeposit.frx":358A
         Left            =   1590
         List            =   "BankDeposit.frx":358C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   960
         ScaleHeight     =   795
         ScaleWidth      =   6795
         TabIndex        =   32
         Top             =   5640
         Width           =   6825
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
            Top             =   -390
            Width           =   1455
         End
         Begin VB.TextBox txtORNumber 
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
            Height          =   345
            Left            =   5160
            TabIndex        =   38
            Top             =   420
            Width           =   1575
         End
         Begin VB.TextBox txtCheckNumber 
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
            Height          =   345
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   30
            Width           =   1575
         End
         Begin VB.TextBox txtCheckType 
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
            Height          =   345
            Left            =   1080
            TabIndex        =   34
            Top             =   420
            Width           =   2685
         End
         Begin VB.TextBox txtCheckDte 
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
            Height          =   345
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   30
            Width           =   1455
         End
         Begin VB.Label labTranID 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
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
            Height          =   210
            Left            =   90
            TabIndex        =   44
            Top             =   990
            Width           =   585
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Number  :"
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
            Height          =   210
            Left            =   4065
            TabIndex        =   40
            Top             =   510
            Width           =   1050
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number  :"
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
            Height          =   210
            Left            =   3930
            TabIndex        =   39
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Type  :"
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
            Height          =   210
            Left            =   45
            TabIndex        =   36
            Top             =   510
            Width           =   990
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date  :"
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
            Height          =   210
            Left            =   75
            TabIndex        =   35
            Top             =   120
            Width           =   960
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
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7680
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtTimDeposit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         Width           =   1425
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
         MouseIcon       =   "BankDeposit.frx":358E
      End
      Begin VB.CommandButton cmdCancelBANKDEPO 
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
         Left            =   8520
         MouseIcon       =   "BankDeposit.frx":38A8
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":39FA
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5670
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveBANKDEPO 
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
         Left            =   7830
         MouseIcon       =   "BankDeposit.frx":3D38
         MousePointer    =   99  'Custom
         Picture         =   "BankDeposit.frx":3E8A
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5670
         Width           =   705
      End
      Begin MSComCtl2.DTPicker txtBankDeposit2 
         Height          =   405
         Left            =   2010
         TabIndex        =   71
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   86769665
         CurrentDate     =   40033
      End
      Begin MSComCtl2.DTPicker dtpDatDeposit 
         Height          =   405
         Left            =   2010
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
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
         Format          =   86769665
         CurrentDate     =   38216
      End
      Begin VB.CheckBox chkSelect 
         Caption         =   "&Select"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   73
         Top             =   5250
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10820
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   6960
         TabIndex        =   77
         Top             =   5280
         Width           =   465
      End
      Begin VB.Shape Shape1 
         Height          =   405
         Left            =   210
         Top             =   5190
         Width           =   8985
      End
      Begin VB.Label Label19 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1620
         TabIndex        =   76
         Top             =   5300
         Width           =   405
      End
      Begin VB.Label lblTotal 
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7590
         TabIndex        =   74
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "O.R. Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6240
         TabIndex        =   70
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Deposit :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   210
         TabIndex        =   45
         Top             =   150
         Width           =   1695
      End
      Begin VB.Label labBankDepoID 
         Height          =   195
         Left            =   3390
         TabIndex        =   31
         Top             =   660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Deposit"
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
         Height          =   255
         Left            =   7680
         TabIndex        =   30
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Height          =   255
         Left            =   6150
         TabIndex        =   29
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Height          =   255
         Left            =   1620
         TabIndex        =   28
         Top             =   660
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   660
         Width           =   1305
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   4785
      Left            =   2580
      ScaleHeight     =   4785
      ScaleWidth      =   9495
      TabIndex        =   25
      Top             =   1530
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdBankDepo 
         Height          =   4575
         Left            =   55
         TabIndex        =   1
         Top             =   120
         Width           =   9400
         _ExtentX        =   16563
         _ExtentY        =   8070
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   15
         FormatString    =   " Type          |   Bank Name                                                 |    Time        | Amount Deposit  | OR Number"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "BankDeposit.frx":41DA
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2640
      ScaleHeight     =   570
      ScaleWidth      =   9405
      TabIndex        =   15
      Top             =   930
      Width           =   9435
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8550
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
         TabIndex        =   16
         Top             =   90
         Width           =   2235
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
      Left            =   270
      TabIndex        =   68
      Top             =   3870
      Visible         =   0   'False
      Width           =   1875
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   735
      Left            =   0
      TabIndex        =   56
      Top             =   -30
      Width           =   12105
      _Version        =   655364
      _ExtentX        =   21352
      _ExtentY        =   1296
      _StockProps     =   14
      Caption         =   " Bank Name :"
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
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   5430
      TabIndex        =   11
      Top             =   2580
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
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
Dim rsBANKDEPO                                                      As ADODB.Recordset
Dim AddorEdit                                                       As String
Dim TOTAL_CASH_DEPOSIT                                              As Double
Dim TOTAL_CHECK_DEPOSIT                                             As Double
Dim TOTAL_CARD_DEPOSIT                                              As Double
Dim PREV_CASH_DEPOSIT                                               As Double
Dim PREV_CHECK_DEPOSIT                                              As Double
Dim PREV_CARD_DEPOSIT                                               As Double
Dim VTYPE                                                           As String
Dim error_msg                                                       As String
Dim icnt                                                            As Integer
Dim LOOP_CNT                                                        As Integer
Dim str_MSG                                                         As String

Function SetCustomerName(XXX As Variant)
    Dim rsCustomer                                                  As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("SELECT CusNam FROM ALL_CUSMAS WHERE CusCde = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = rsCustomer!CusNam
    End If
    Set rsCustomer = Nothing
End Function

Function SetBankCode(XXX As Variant)
    Dim rsBankName                                                  As New ADODB.Recordset
    'UPDATE BY   : MJP 052209 0400PM
    'DESCRIPTION : TO GET FROM THE ALL_BANKS TABLE, BECAUSE CMIS_BANKS IS AN UNION OF TWO TABLE POSIBLE ERROR IS PROGRAM MY SET THE BANKS IN THE CMIS ONLY
    Set rsBankName = gconDMIS.Execute("SELECT BANKCODE FROM ALL_BANKDEPOSITS WHERE BANKNAME = " & N2Str2Null(XXX) & "")
    'SET rsBankName = gconDMIS.Execute("SELECT BANKCODE FROM CMIS_BANKS WHERE BANKNAME = " & N2Str2Null(XXX) & "")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankCode = rsBankName!bankcode
    End If
End Function

Function SetBankName(XXX As Variant)
    Dim rsBankName                                                  As ADODB.Recordset
    Set rsBankName = New ADODB.Recordset
    Set rsBankName = gconDMIS.Execute("SELECT BANKNAME FROM CMIS_BANKS WHERE BANKCode = '" & XXX & "'")
    If Not rsBankName.EOF And Not rsBankName.BOF Then
        SetBankName = rsBankName!BANKNAME
    End If
End Function

Function SetCheckClass(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT DESCNAME FROM CMIS_SBOOK WHERE Book = 'F' AND CODE = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClass = rsSBOOK!DESCNAME
    End If
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT CODE FROM CMIS_SBOOK WHERE Book = 'F' AND DESCNAME = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
End Function

Sub SetSELECTedType()
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
                cboDeposit_To.Enabled = True
                txtDeposit.Enabled = False
            Else
                cboCheckTransactions.Visible = False
                grdCheckCardTransactions.Enabled = True
                cboBankCode.Enabled = False
                cboDeposit_To.Enabled = True
                txtDeposit.Enabled = True
            End If
        End If
        chkSelect.Value = False
        cboCheckTransactions.Text = "Cashier Collection"
        cboCheckTransactions_Click
    End If
End Sub

Sub rsRefresh()
    Set rsBANKDEPO = New ADODB.Recordset
    Set rsBANKDEPO = gconDMIS.Execute("SELECT DISTINCT DATDEPOSIT FROM CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' ORDER BY DATDEPOSIT DESC")
End Sub

Sub StoreMemVars()
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
        txtDatDeposit.Text = Null2Date(rsBANKDEPO!DATDEPOSIT)
        StoreDetails
    Else
        txtDatDeposit.Text = LOGDATE
        'cmdAdd.Value = True
    End If
End Sub

Sub StoreDetails()
    Dim rsBANKDEPODet                                               As ADODB.Recordset
    Dim VTYPE                                                       As String
    Dim i                                                           As Long

    TOTAL_CASH_DEPOSIT = 0
    TOTAL_CHECK_DEPOSIT = 0
    TOTAL_CARD_DEPOSIT = 0
    InitGrid
    i = 0
    
    Set rsBANKDEPODet = New ADODB.Recordset
    Set rsBANKDEPODet = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(RTrim(LTrim(cboDeposit_To))) & "' AND DATDEPOSIT = '" & txtDatDeposit.Text & "' ORDER BY TYPE, ID ASC")
    If Not rsBANKDEPODet.EOF And Not rsBANKDEPODet.BOF Then
        rsBANKDEPODet.MoveFirst
        Do While Not rsBANKDEPODet.EOF
            i = i + 1
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                VTYPE = "CASH"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            ElseIf Null2String(rsBANKDEPODet!Type) = "2" Then
                VTYPE = "CHECK"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetBankName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            ElseIf Null2String(rsBANKDEPODet!Type) = "3" Then
                VTYPE = "CARD"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            End If
            
            If i = 1 Then grdBANKDEPO.RemoveItem 1
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            ElseIf Null2String(rsBANKDEPODet!Type) = "2" Then
                TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            ElseIf Null2String(rsBANKDEPODet!Type) = "3" Then
                TOTAL_CARD_DEPOSIT = TOTAL_CARD_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            End If
            rsBANKDEPODet.MoveNext
        Loop
    End If
    txtTotalCashAmt.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT)
    txtTotalCheckAmt.Text = ToDoubleNumber(TOTAL_CHECK_DEPOSIT)
    txtCardDeposit.Text = ToDoubleNumber(TOTAL_CARD_DEPOSIT)
    txtTotalDepositedAmount.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT + TOTAL_CHECK_DEPOSIT + TOTAL_CARD_DEPOSIT)
    
    'MITSUBISHI ONLY
    Dim rsItemCount                                                 As ADODB.Recordset
    Set rsItemCount = New ADODB.Recordset
    rsItemCount.Open "SELECT COUNT(*) as TotalItem FROM CMIS_BANKDEPO WHERE DEPOSIT_TO = '" & SetBankCode(RTrim(LTrim(cboDeposit_To))) & "' AND DATDEPOSIT = '" & txtDatDeposit.Text & "'", gconDMIS, adOpenForwardOnly
    If Not rsItemCount.EOF And Not rsItemCount.BOF Then
        lblTotalItem.Caption = "Total Item: " + Null2String(rsItemCount!TotalItem)
    End If
    Set rsItemCount = Nothing
End Sub

Sub ShowGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                          As New ADODB.Recordset
    Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE ID = " & XXX)
    If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
        txtBankDeposit = Null2Date(rsBANKDEPO_Details!DATDEPOSIT)
        txtCheckDate.Text = Null2String(rsBANKDEPO_Details!CheckDate)
        txtTseklase.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
        txtCheckNum.Text = Null2String(rsBANKDEPO_Details!CheckNum)
    Else
        txtCheckDate.Text = ""
        txtTseklase.Text = ""
        txtCheckNum.Text = ""
    End If
    Set rsBANKDEPO_Details = Nothing
End Sub

Sub StoreGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                          As ADODB.Recordset
    Set rsBANKDEPO_Details = New ADODB.Recordset
    Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE ID = " & XXX)
    If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
        labBankDepoID.Caption = rsBANKDEPO_Details!Id
        If Null2String(rsBANKDEPO_Details!Type) = "1" Then
            cboType.Text = "CASH"
            PREV_CASH_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!DEPOSIT)
            PREV_CHECK_DEPOSIT = 0
            PREV_CARD_DEPOSIT = 0
        ElseIf Null2String(rsBANKDEPO_Details!Type) = "2" Then
            cboType.Text = "CHECK"
            PREV_CASH_DEPOSIT = 0
            PREV_CHECK_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!DEPOSIT)
            PREV_CARD_DEPOSIT = 0
        ElseIf Null2String(rsBANKDEPO_Details!Type) = "3" Then
            cboType.Text = "CARD"
            PREV_CASH_DEPOSIT = 0
            PREV_CHECK_DEPOSIT = 0
            PREV_CARD_DEPOSIT = N2Str2Zero(rsBANKDEPO_Details!DEPOSIT)
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
        txtDatDeposit.Text = Null2String(rsBANKDEPO_Details!DATDEPOSIT)
        txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!DEPOSIT))
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
    cleargrid grdBANKDEPO
    grdBANKDEPO.FormatString = " Type          |   Bank Name / Customer Name                                    |    Time       | Amount Deposit | id | OR Number    "
    grdBANKDEPO.ColWidth(4) = 1
End Sub

Sub InitTransactionsGrid()
    cleargrid grdCheckCardTransactions
    grdCheckCardTransactions.FormatString = " Code         |   Bank Name                                        |    Time       | Check Amount  | OR Number | "
    grdCheckCardTransactions.ColWidth(5) = 1
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
    Dim rsBANK                                              As ADODB.Recordset
    Set rsBANK = New ADODB.Recordset
    Set rsBANK = gconDMIS.Execute("SELECT BANKNAME FROM ALL_BANKDEPOSITS ORDER BY BANKNAME ASC")
    If Not rsBANK.EOF And Not rsBANK.BOF Then
        Combo_Loadval cboBankCode, rsBANK
    End If
    
    Set rsBANK = New ADODB.Recordset
    Set rsBANK = gconDMIS.Execute("SELECT BANKNAME FROM ALL_BANKDEPOSITS ORDER BY BANKNAME ASC")
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

    If COMPANY_CODE = "DJM" Then
        txtBankDeposit = CURRENT_CUTOFF_DATE
        txtBankDeposit.Enabled = True
        txtBankDeposit.Locked = False
    Else
        txtBankDeposit = CURRENT_CUTOFF_DATE
    End If
    
    cboType.ListIndex = -1
    cboBankCode.Enabled = False
    cboBankCode.ListIndex = -1
    'cboDeposit_To.Enabled = False
    'cboDeposit_To.ListIndex = -1
    txtTimDeposit.Enabled = False
    txtTimDeposit.Text = ""
    txtDeposit.Enabled = False
    txtDeposit.Text = "0.00"
    lsvDet.ListItems.Clear
    InitTransactionsGrid
    labTranID.Caption = ""
    
    If COMPANY_CODE = "DJM" Then
        txtBankDeposit = txtBankDeposit.Text
    Else
        txtBankDeposit = LOGDATE
    End If
    
    dtSelectedDate.Value = LOGDATE
    txtCheckDate.Text = ""
    txtCheckType.Text = ""
    txtCheckNumber.Text = ""
    txtORNumber.Text = ""
    txtORSearch.Text = ""
    txtBankCode.Text = ""
    txtTimeCreate.Text = ""
    txtCheckAmount.Text = ""
End Sub

Sub FillGrid()
    Dim BankDeposit                                                 As ADODB.Recordset
    
    'grdBankDepo.Clear
    
    lstBANKDEPO.Sorted = False
    lstBANKDEPO.ListItems.Clear
    lstBANKDEPO.Enabled = False
    
    Set BankDeposit = New ADODB.Recordset
    Set BankDeposit = gconDMIS.Execute("SELECT DISTINCT DATDEPOSIT FROM CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' ORDER BY DATDEPOSIT DESC")
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
    Dim BankDeposit                                                 As ADODB.Recordset
    
    lstBANKDEPO.Sorted = False: lstBANKDEPO.ListItems.Clear
    lstBANKDEPO.Enabled = False
    XXX = Repleys(LTrim(RTrim(XXX)))
    
    Set BankDeposit = New ADODB.Recordset
    Set BankDeposit = gconDMIS.Execute("SELECT DISTINCT DATDEPOSIT FROM CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(cboDeposit_To.Text) & "' AND DATDEPOSIT LIKE '" & XXX & "%' ORDER BY DATDEPOSIT DESC")
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
    Dim rsCHECKDet                                                  As ADODB.Recordset
    Dim i                                                           As Long
    Dim xList                                                       As ListItem
    
    InitTransactionsGrid
    lsvDet.ListItems.Clear
    lblTotal.Caption = "0.00"
    
    If cboType.Text = "CARD" Then
        If COMPANY_CODE = M_COMPANY_CODE Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit3 = 0 AND CardAmount > 0 AND CANCEL = 0 ORDER BY ID ASC")
        Else
            If COMPANY_CODE = "HGC" Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
            Else
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND LEFT(OR_NUM,3) <> 'SOA' ORDER BY ID ASC")
            End If
        End If
        If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
            rsCHECKDet.MoveFirst
            Do While Not rsCHECKDet.EOF
                Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CardAmount))
                xList.SubItems(4) = rsCHECKDet!OR_NUM
                xList.SubItems(5) = rsCHECKDet!Id
                lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CardAmount))
                rsCHECKDet.MoveNext
            Loop
            chkSelect.Enabled = True
        Else
            chkSelect.Enabled = False
        End If
    ElseIf cboType.Text = "CASH" Then
        If COMPANY_CODE = M_COMPANY_CODE Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit1 = 0 AND CashAmount > 0 AND CANCEL = 0 ORDER BY ID ASC")
        Else
            If COMPANY_CODE = "HGC" Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
            Else
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND LEFT(OR_NUM,3) <> 'SOA' ORDER BY ID ASC")
            End If
        End If
        If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
            rsCHECKDet.MoveFirst
            Do While Not rsCHECKDet.EOF
                Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount))
                xList.SubItems(4) = rsCHECKDet!OR_NUM
                xList.SubItems(5) = rsCHECKDet!Id
                lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CashAmount))
                rsCHECKDet.MoveNext
            Loop
            chkSelect.Enabled = True
        Else
            chkSelect.Enabled = False
        End If
    End If

    If cboCheckTransactions.Text = "Cashier Collection" Then
        Set rsCHECKDet = New ADODB.Recordset
        If cboType.Text = "CHECK" Then
            If COMPANY_CODE = M_COMPANY_CODE Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE ISNULL(Deposit2,0) = 0 AND chkAmount > 0 AND CANCEL = 0 ORDER BY ID ASC")
            Else
                If COMPANY_CODE = "HGC" Then
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
                Else
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND LEFT(OR_NUM,3) <> 'SOA' ORDER BY ID ASC")
                End If
            End If
            If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                rsCHECKDet.MoveFirst
                Do While Not rsCHECKDet.EOF
                    Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                    xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                    xList.SubItems(4) = rsCHECKDet!OR_NUM
                    xList.SubItems(5) = rsCHECKDet!Id
                    lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CHKAMOUNT))
                    rsCHECKDet.MoveNext
                Loop
                chkSelect.Enabled = True
            Else
                chkSelect.Enabled = False
            End If
        End If
    ElseIf cboCheckTransactions.Text = "Check Encashment" Then
        Set rsCHECKDet = New ADODB.Recordset
        If cboType.Text = "CHECK" Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE Deposit = 0 AND chkAmount > 0 AND LEFT(OR_NUM,3) <> 'SOA' ORDER BY ID ASC")
            If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                rsCHECKDet.MoveFirst
                Do While Not rsCHECKDet.EOF
                    Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                    xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                    xList.SubItems(4) = rsCHECKDet!OR_NUM
                    xList.SubItems(5) = rsCHECKDet!Id
                    rsCHECKDet.MoveNext
                Loop
                chkSelect.Enabled = True
            Else
                chkSelect.Enabled = False
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
    Call SetSELECTedType
End Sub

Private Sub cboType_Click()
    Call SetSELECTedType
End Sub

Private Sub cboType_GotFocus()
    InitTransactionsGrid
    labTranID.Caption = ""
    txtCheckDate.Text = ""
    txtCheckType.Text = ""
    txtCheckNumber.Text = ""
    txtORNumber.Text = ""
    txtBankCode.Text = ""
    txtTimeCreate.Text = ""
    txtCheckAmount.Text = "0.00"
    VBComBoBoxDroppedDown cboType
End Sub

Private Sub chkSelect_Click()
    Dim iCount                                              As Integer
    
    If lblTotal.Caption <> 0 Then lblTotal.Caption = "0.00"
    If chkSelect.Value = 1 Then
        For iCount = 1 To lsvDet.ListItems.Count
            lsvDet.ListItems(iCount).Checked = True
            lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(lsvDet.ListItems(iCount).SubItems(3)))
        Next
    Else
        For iCount = 1 To lsvDet.ListItems.Count
            lsvDet.ListItems(iCount).Checked = False
            lblTotal.Caption = "0.00"
        Next
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub

    AddorEdit = "ADD"
    PicBankDepo.Visible = True
    PicBankDepo.ZOrder 0
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
    PicBankDepo.Visible = False
    PicBankDepo.ZOrder 1
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
        Dim rsJoyDeposit                                    As ADODB.Recordset
        Set rsJoyDeposit = New ADODB.Recordset
        Set rsJoyDeposit = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE ID = " & labBankDepoID.Caption)
        If Not rsJoyDeposit.EOF And Not rsJoyDeposit.BOF Then
            gconDMIS.Execute ("DELETE FROM CMIS_BankDepo WHERE ID = " & labBankDepoID.Caption)
            
            If COMPANY_CODE = M_COMPANY_CODE Then
                gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit1 = 0 WHERE OR_NUM = " & N2Str2Null(rsJoyDeposit!OR_NUM))
                gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit2 = 0 WHERE OR_NUM = " & N2Str2Null(rsJoyDeposit!OR_NUM))
                gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit3 = 0 WHERE OR_NUM = " & N2Str2Null(rsJoyDeposit!OR_NUM))
            Else
                gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit = 0 WHERE OR_NUM = " & N2Str2Null(rsJoyDeposit!OR_NUM))
            End If
            
            If cboType.Text = "CASH" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CASH = CASH + " & NumericVal(txtDeposit.Text) & "," & _
                                  " CASHDEPO = CASHDEPO - " & NumericVal(txtDeposit.Text) & _
                                  " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf cboType.Text = "CHECK" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " [CHECK] = [CHECK] + " & NumericVal(txtDeposit.Text) & "," & _
                                  " CHECKDEPO = CHECKDEPO - " & NumericVal(txtDeposit.Text) & _
                                  " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf cboType.Text = "CARD" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CARD = CARD + " & NumericVal(txtDeposit.Text) & "," & _
                                  " CARDDEPO = CARDDEPO - " & NumericVal(txtDeposit.Text) & _
                                  " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            ShowDeletedMsg
        End If
    End If
    rsRefresh
    
    If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then rsBANKDEPO.MoveLast
    cmdCancelBANKDEPO_Click
    On Error Resume Next
    rsBANKDEPO.Find "DatDeposit = " & N2Date2Null(txtDatDeposit.Text)
    StoreMemVars
    Exit Sub
    
ErrorCode:
    ShowVBError
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

End Sub

'Private Sub cmdPost_Click()
'    If Function_Access(LOGID, "Acess_Post", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
'End Sub

Private Sub cmdPrint_Click()
''''''''''    If Function_Access(LOGID, "Acess_Print", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
''''''''''    'updating code:    JAA - 07112007
''''''''''    'On Error GoTo ErrorCode:
''''''''''
''''''''''    'txtDatDeposit.Text = lstBANKDEPO.SelectedItem
''''''''''    If IsDate(txtDatDeposit.Text) = False Then
''''''''''        MsgBox "Please click Date Deposited.", vbInformation, "Check Date"
''''''''''        Exit Sub
''''''''''    End If
''''''''''
''''''''''    Screen.MousePointer = 11
''''''''''    With rptBankDepo
''''''''''        .Formulas(0) = "DEALER_NAME = '" & COMPANY_NAME & "'"
''''''''''        .Formulas(1) = "DEALER_ADDRESS = '" & COMPANY_ADDRESS & "'"
''''''''''        .Formulas(2) = "PREPAREDBY= '" & PreparedBy & "'"
''''''''''        .Formulas(3) = "NOTEDBY= '" & NotedBy & "'"
''''''''''        .Formulas(4) = "CHECKEDBY= '" & CheckedBy & "'"
''''''''''        .Formulas(5) = "PRINTEDBY= " & N2Str2Null(LOGNAME)
''''''''''    End With
''''''''''    PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit.rpt", "{BankDepo.DatDeposit} = Date(" & Year(txtDatDeposit.Text) & "," & Month(txtDatDeposit.Text) & "," & Day(txtDatDeposit.Text) & ") AND {BankDepo.DEPOSIT_TO} = '" & SetBankCode(cboDeposit_To) & "'", CMIS_REPORT_CONNECTION, 1
''''''''''    Screen.MousePointer = 0
''''''''''    Exit Sub
''''''''''
''''''''''ErrorCode:
''''''''''    ShowVBError




    Pic_PrintBankDepo.ZOrder 0
    Pic_PrintBankDepo.Visible = True

    If txtDatDeposit <> "" And cboDeposit_To <> "" Then
        OptPrint.Value = True
    ElseIf txtDatDeposit = "" And cboDeposit_To <> "" Then
        CboBankDepo.Text = cboDeposit_To.Text
        OptPrintAsof.Value = True
    Else
        OptPrintAsof.Value = True
    End If

    lstBANKDEPO.Enabled = False
    textSearch.Enabled = False
    cboDeposit_To.Enabled = False
    cmdShow.Enabled = False
End Sub

Private Sub cmdSaveBANKDEPO_Click()
    On Error GoTo ErrorCode
    
    If Trim(cboDeposit_To.Text) = "" Then
        MsgBox "Please select where to deposit...", vbInformation, "Bank not Selected"
        Exit Sub
    End If

    If cboType.Text = "" Then
        MsgBox "Please select type", vbInformation, "Select Type"
        cboType.SetFocus
        Exit Sub
    End If
    
    icnt = 0
    For LOOP_CNT = 1 To lsvDet.ListItems.Count
        If lsvDet.ListItems.Item(LOOP_CNT).Checked = True Then
            icnt = icnt + 1
         End If
    Next
    If icnt = 0 Then MsgBox "No item selected!", vbInformation + vbOKOnly: Exit Sub
       
    If AddorEdit = "ADD" Then
        If CheckCutoff(CDate(txtBankDeposit)) = True And UCase(LOGNAME) <> "NETSPEED" Then
            MsgBox "Deposit not allowed. Please check Cut-off Date.", vbInformation, "Check Date"
            Exit Sub
        End If
    End If
    
    'JJE 05042013 5:00PM
    'To roolback transaction if error occured
    gconDMIS.BeginTrans
    If Transave = False Then
            str_MSG = "Error Appear During Saving." & vbCrLf
            str_MSG = str_MSG & "Description: "
            str_MSG = str_MSG & " " & error_msg
            str_MSG = str_MSG & " " & vbCrLf
            str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
            str_MSG = str_MSG & "Telephone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            MsgBox str_MSG, vbCritical, "Posting Error"
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
    End If
    gconDMIS.CommitTrans
                                
    rsRefresh
    On Error Resume Next
    
    rsBANKDEPO.Find "DATDEPOSIT = '" & txtBankDeposit & "'"
    cmdCancelBANKDEPO_Click
    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    StoreDetails2
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Function Transave()
    On Error GoTo ErrorCode
    
    Dim rsTMP                                                       As New ADODB.Recordset
    Dim vBankCode                                                   As String
    Dim vTseklase                                                   As String
    Dim vDeposit                                                    As String
    Dim vDatDeposit                                                 As String
    Dim vTimDeposit                                                 As String
    Dim vWhoDeposit                                                 As String
    Dim vInCashChk                                                  As Integer
    Dim vCollectChk                                                 As Integer
    Dim vP_pay_Chk                                                  As Integer
    Dim vL_pay_Chk                                                  As Integer
    Dim vU_pay_Chk                                                  As Integer
    Dim vA_pay_Chk                                                  As Integer
    Dim vOR_NUM                                                     As String
    Dim vDeposit_To                                                 As String
    Dim vCheckDate                                                  As String
    Dim vCardDate                                                   As String
    Dim vCheckNum                                                   As String
    Dim vCardNumber                                                 As String
    Dim xMODULENAME                                                 As String
    
    VTYPE = ""
    For LOOP_CNT = 1 To lsvDet.ListItems.Count
        If lsvDet.ListItems.Item(LOOP_CNT).Checked = False Then GoTo NEXT_ITEM
        Call ShowTransactionsGridDetails(lsvDet.ListItems(LOOP_CNT).SubItems(5))
        
        If cboType.Text = "CASH" Then
            VTYPE = "'1'"
            vCheckDate = "NULL"
            vCheckNum = "NULL"
            vCardDate = "NULL"
            vCardNumber = "NULL"
            grdCheckCardTransactions.Col = 0
            vBankCode = N2Str2Null(Trim(lsvDet.ListItems(LOOP_CNT).Text))
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
        ElseIf cboType.Text = "CHECK" Then
            vBankCode = N2Str2Null(txtBankCode.Text)
            VTYPE = "'2'"
            vTseklase = N2Str2Null(SetCheckClassCode(txtCheckType.Text))
            vCheckDate = N2Str2Null(txtCheckDte.Text)
            vCheckNum = N2Str2Null(txtCheckNumber.Text)
            vOR_NUM = N2Str2Null(txtORNumber.Text)
            vDeposit = NumericVal(txtDeposit.Text)
            vDatDeposit = N2Date2Null(txtBankDeposit)
            vTimDeposit = N2Str2Null(txtTimeCreate.Text)
            vWhoDeposit = "'00005'"
            vCardDate = "NULL"
            vCardNumber = "NULL"
        ElseIf cboType.Text = "CARD" Then
            grdCheckCardTransactions.Col = 0
            vBankCode = N2Str2Null(Trim(lsvDet.ListItems(LOOP_CNT).Text))
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
        vDeposit_To = N2Str2Null(SetBankCode(cboDeposit_To.Text))
        
        If AddorEdit = "ADD" Then
            SQL_STATEMENT = "INSERT INTO CMIS_BankDepo " & _
                                "(BankCode,Tseklase,Deposit,DatDeposit,TimDeposit,WhoDeposit,[Type],InCashChk,CollectChk,P_pay_Chk,L_pay_Chk,U_pay_Chk,A_pay_Chk,OR_Num,Deposit_To,CheckDate,CardDate,CheckNum,CardNumber)" & _
                                " VALUES (" & vBankCode & _
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
                
                If cboCheckTransactions.Text = "Check Encashment" Then
                    gconDMIS.Execute ("UPDATE CMIS_InCash SET Deposit = 1 WHERE ID = " & labTranID.Caption)
                Else
                    If labTranID.Caption <> "" Then
                        If COMPANY_CODE = M_COMPANY_CODE Then
                            If VTYPE = "'1'" Then
                                SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit1 = 1 WHERE ID = " & labTranID.Caption
                                gconDMIS.Execute SQL_STATEMENT
                            ElseIf VTYPE = "'2'" Then
                                SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit2 = 1 WHERE ID = " & labTranID.Caption
                                gconDMIS.Execute SQL_STATEMENT
                            ElseIf VTYPE = "'3'" Then
                                SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit3 = 1 WHERE ID = " & labTranID.Caption
                                gconDMIS.Execute SQL_STATEMENT
                            End If
                        Else
                            SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit = 1 WHERE ID = " & labTranID.Caption
                            gconDMIS.Execute SQL_STATEMENT
                        End If
    
                        'NEW LOG AUDIT---------------------------------------------------------
                        Set rsTMP = gconDMIS.Execute("SELECT VAT,OR_NUM FROM CMIS_OFF_HD WHERE ID = " & labTranID.Caption & "")
                        If Not (rsTMP.BOF And rsTMP.EOF) Then
                            If Null2String(rsTMP!vat) = "1" Then xMODULENAME = "TRANSACTION O.R. WITH VAT"
                            If Null2String(rsTMP!vat) = "0" Then xMODULENAME = "TRANSACTION O.R. WITHOUT VAT"
    
                            Call NEW_LogAudit("E", xMODULENAME, SQL_STATEMENT, lblBANKID, "", "OR NO: " & Null2String(rsTMP!OR_NUM), "", "")
                        End If
                        Set rsTMP = Nothing
                        'NEW LOG AUDIT---------------------------------------------------------
                    End If
                End If
                Call SaveCashPosition(cboType, NumericVal(vDeposit))
        Else
            SQL_STATEMENT = "UPDATE CMIS_BankDepo SET " & _
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
                            " WHERE ID = " & labBankDepoID.Caption
            gconDMIS.Execute SQL_STATEMENT
            'NEW LOG AUDIT---------------------------------------------------------
            Call NEW_LogAudit("E", "TRANSACTION BANKDEPOSIT", SQL_STATEMENT, lblBANKID, "", "OR NO: ", "", labBankDepoID.Caption)
            'NEW LOG AUDIT---------------------------------------------------------
            
            If cboCheckTransactions.Text = "Cashier Collection" Then
                If labTranID.Caption <> "" Then
                    If VTYPE = "'1'" Then
                        SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit1 = 1 WHERE ID = " & labTranID.Caption
                        gconDMIS.Execute SQL_STATEMENT
                    ElseIf VTYPE = "'2'" Then
                        SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit2 = 1 WHERE ID = " & labTranID.Caption
                        gconDMIS.Execute SQL_STATEMENT
                    ElseIf VTYPE = "'3'" Then
                        SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Deposit3 = 1 WHERE ID = " & labTranID.Caption
                        gconDMIS.Execute SQL_STATEMENT
                    End If

                    'NEW LOG AUDIT---------------------------------------------------------
                    Set rsTMP = gconDMIS.Execute("SELECT VAT,OR_NUM FROM CMIS_OFF_HD WHERE ID = " & labTranID.Caption & "")
                    If Not (rsTMP.BOF And rsTMP.EOF) Then
                        If Null2String(rsTMP!vat) = "1" Then xMODULENAME = "TRANSACTION O.R. WITH VAT"
                        If Null2String(rsTMP!vat) = "0" Then xMODULENAME = "TRANSACTION O.R. WITHOUT VAT"

                        Call NEW_LogAudit("E", xMODULENAME, SQL_STATEMENT, lblBANKID, "", "OR NO: " & Null2String(rsTMP!OR_NUM), "", "")
                    End If
                    Set rsTMP = Nothing
                    'NEW LOG AUDIT---------------------------------------------------------
                End If
            End If
'            gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
'                              " CASH = (CASH + " & PREV_CASH_DEPOSIT & ")," & _
'                              " CARD = (CARD + " & PREV_CARD_DEPOSIT & ")," & _
'                              " [CHECK] = ([CHECK] + " & PREV_CHECK_DEPOSIT & ")," & _
'                              " CASHDEPO = (CASHDEPO - " & PREV_CASH_DEPOSIT & ")," & _
'                              " CARDDEPO = (CARDDEPO - " & PREV_CARD_DEPOSIT & ")," & _
'                              " CHECKDEPO = (CHECKDEPO - " & PREV_CHECK_DEPOSIT & ") WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")

            If cboType.Text = "CASH" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CASH = CASH - " & vDeposit & "," & _
                                  " CASHDEPO = CASHDEPO + " & vDeposit & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf cboType.Text = "CHECK" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " [CHECK] = [CHECK] - " & vDeposit & "," & _
                                  " CHECKDEPO = CHECKDEPO + " & vDeposit & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf cboType.Text = "CARD" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CARD = CARD - " & vDeposit & "," & _
                                  " CARDDEPO = CARDDEPO + " & vDeposit & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        End If
        
NEXT_ITEM:
    Next
    Transave = True
    Exit Function
    
ErrorCode:
    Transave = False
    error_msg = error
End Function

Private Sub cmdShow_Click()
    lblBANKID.Caption = FindTransactionID(N2Str2Null(cboDeposit_To), "BANKNAME", "ALL_BANKDEPOSITS")
    rsRefresh
    cmdCancelBANKDEPO_Click
    FillGrid
End Sub

Private Sub dtSelectedDate_Change()
    Dim rsCHECKDet                                                  As ADODB.Recordset
    Dim i                                                           As Long
    Dim xList                                                       As ListItem
    
    InitTransactionsGrid
    lsvDet.ListItems.Clear
    lblTotal.Caption = "0.00"
    
    If cboType.Text = "CARD" Then
        If COMPANY_CODE = M_COMPANY_CODE Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit3 = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
        Else
            If COMPANY_CODE = "HGC" Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
            Else
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
            End If
        End If
        If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
            rsCHECKDet.MoveFirst
            Do While Not rsCHECKDet.EOF
                Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CardAmount))
                xList.SubItems(4) = rsCHECKDet!Id
                lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CardAmount))
                rsCHECKDet.MoveNext
            Loop
            chkSelect.Enabled = True
        Else
            chkSelect.Enabled = False
        End If
    End If

    If cboType.Text = "CASH" Then
        If COMPANY_CODE = M_COMPANY_CODE Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit1 = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
        Else
            If COMPANY_CODE = "HGC" Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
            Else
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
            End If
        End If
        If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
            rsCHECKDet.MoveFirst
            Do While Not rsCHECKDet.EOF
                Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount))
                xList.SubItems(4) = rsCHECKDet!Id
                lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CashAmount))
                rsCHECKDet.MoveNext
            Loop
            chkSelect.Enabled = True
        Else
            chkSelect.Enabled = False
        End If
    End If

    If cboCheckTransactions.Text = "Cashier Collection" Then
        Set rsCHECKDet = New ADODB.Recordset
        If cboType.Text = "CHECK" Then
            If COMPANY_CODE = M_COMPANY_CODE Then
                Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit2 = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
            Else
                If COMPANY_CODE = "HGC" Then
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_DATE >= '2/1/2010' AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
                Else
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_DATE = '" & dtSelectedDate.Value & "' ORDER BY ID ASC")
                End If
            End If
            If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                rsCHECKDet.MoveFirst
                Do While Not rsCHECKDet.EOF
                    Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                    xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                    xList.SubItems(4) = rsCHECKDet!Id
                    lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CHKAMOUNT))
                    rsCHECKDet.MoveNext
                Loop
                chkSelect.Enabled = True
            Else
                chkSelect.Enabled = False
            End If
        End If
    End If

    If cboCheckTransactions.Text = "Check Encashment" Then
        Set rsCHECKDet = New ADODB.Recordset
        If cboType.Text = "CHECK" Then
            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE Deposit = 0 AND chkAmount > 0 ORDER BY ID ASC")
            If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                rsCHECKDet.MoveFirst
                Do While Not rsCHECKDet.EOF
                    Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                    xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                    xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                    xList.SubItems(4) = rsCHECKDet!Id
                    lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(rsCHECKDet!CHKAMOUNT))
                    rsCHECKDet.MoveNext
                Loop
                chkSelect.Enabled = True
            Else
                chkSelect.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        cmdAdd.Value = True
    Case vbKeyF11
        Shell "calc.exe"
    Case vbKeyEscape
        cmdCancelBANKDEPO_Click
        '    Case vbKeyDelete
        '    If txtDatDeposit <> "" Then
        '        If CheckCutoff(txtDatDeposit) = True Then
        '            MsgBox "Action not permited, Cut Off has been processed.", vbInformation, "Message"
        '            Exit Sub
        '        Else
        '            Dim xCustomer                                   As String
        '            Dim rsBANKDEPO                                  As ADODB.Recordset
        '            SET rsBANKDEPO = New ADODB.Recordset
        '            rsBANKDEPO.Open "SELECT * FROM CMIS_BankDepo WHERE ID='" & grdBankDepo.Text & "'", gconDMIS, adOpenForwardOnly
        '            If Not rsBANKDEPO.EOF AND Not rsBANKDEPO.BOF Then
        '                If Null2String(rsBANKDEPO!Type) = "1" Then
        '                    xCustomer = SetCustomerName(Null2String(rsBANKDEPO!bankcode)) & _
                             '                                Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
        '                ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
        '                    xCustomer = SetBankName(Null2String(rsBANKDEPO!bankcode)) & _
                             '                                Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
        '                ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
        '                    xCustomer = SetCustomerName(Null2String(rsBANKDEPO!bankcode)) & _
                             '                                Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
        '                End If
        '                If MsgBox("Are you sure you want to DELETE this entry?" & vbCrLf & xCustomer, vbQuestion + vbYesNo, "DELETE") = vbYes Then
        '                    gconDMIS.Execute ("DELETE FROM CMIS_BankDepo WHERE ID='" & grdBankDepo.Text & "'")
        '                    If Null2String(rsBANKDEPO!Type) = "1" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit1 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
        '                    ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit2 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
        '                    ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit3 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
        '                    End If
        '                    StoreDetails
        '
        '                    If Null2String(rsBANKDEPO!Type) = "1" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                 '                                          " CASH = CASH + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                 '                                          " CASHDEPO = CASHDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                 '                                          " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
        '                    ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                 '                                          " [CHECK] = [CHECK] + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                 '                                          " CHECKDEPO = CHECKDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                 '                                          " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
        '                    ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
        '                        gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                 '                                          " CARD = CARD + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                 '                                          " CARDDEPO = CARDDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                 '                                          " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
        '                    End If
        '                Else
        '                    Exit Sub
        '                End If
        '            End If
        '        End If
        '    End If
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
    Dim rsProfile                                                   As ADODB.Recordset
    
    Screen.MousePointer = 11
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("SELECT * FROM ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing:

    CenterMe frmMain, Me, 1
    initMemvars

    textSearch.Text = ""
    InitCbo
    InitGrid

    cboDeposit_To.ListIndex = 0
    Call textSearch_Change

    PicBankDepo.Visible = False
    PicBankDepo.ZOrder 1
    Screen.MousePointer = 0
End Sub

Private Sub grdBANKDEPO_Click()
    grdBANKDEPO.Col = 4
    If grdBANKDEPO.Text <> "" Then
        ShowGridDetails grdBANKDEPO.Text
    End If
End Sub

Private Sub grdBankDepo_DblClick()
'    cmdEdit.Value = True
End Sub

Private Sub grdBankDepo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        'JJE 05042013 5:28PM
        'To rollback transaction if error occured
        If txtDatDeposit <> "" Then
            If CheckCutoff(CDate(txtDatDeposit)) = True And UCase(LOGNAME) <> "NETSPEED" Then
                MsgBox "Action not permitted, Cut Off has been processed.", vbInformation, "Message"
                Exit Sub
            End If
            gconDMIS.BeginTrans
            If TranDete = False Then
                str_MSG = "Error Appear During Deleting." & vbCrLf
                str_MSG = str_MSG & "Description: "
                str_MSG = str_MSG & " " & error_msg
                str_MSG = str_MSG & " " & vbCrLf
                str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
                str_MSG = str_MSG & "Telephone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
                str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
                MsgBox str_MSG, vbCritical, "Posting Error"
                gconDMIS.RollbackTrans
                Screen.MousePointer = 0
            End If
            gconDMIS.CommitTrans
        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Function TranDete() As Boolean
    On Error GoTo ErrorCode
    
    If txtDatDeposit <> "" Then
        Dim xCustomer                                               As String
        Dim rsBANKDEPO                                              As ADODB.Recordset
        Set rsBANKDEPO = New ADODB.Recordset
        rsBANKDEPO.Open "SELECT * FROM CMIS_BankDepo WHERE ID='" & grdBANKDEPO.Text & "'", gconDMIS, adOpenForwardOnly
        If Not rsBANKDEPO.EOF And Not rsBANKDEPO.BOF Then
            If Null2String(rsBANKDEPO!Type) = "1" Then
                xCustomer = SetCustomerName(Null2String(rsBANKDEPO!bankcode)) & _
                            Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
            ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
                xCustomer = SetBankName(Null2String(rsBANKDEPO!bankcode)) & _
                            Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
            ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
                xCustomer = SetCustomerName(Null2String(rsBANKDEPO!bankcode)) & _
                            Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPO!DEPOSIT))
            End If
            If MsgBox("Are you sure you want to delete this entry?" & vbCrLf & xCustomer, vbQuestion + vbYesNo, "DELETE") = vbYes Then
                gconDMIS.Execute ("DELETE FROM CMIS_BankDepo WHERE ID='" & grdBANKDEPO.Text & "'")
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If Null2String(rsBANKDEPO!Type) = "1" Then
                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit1 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
                    ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit2 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
                    ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
                        gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit3 = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
                    End If
                Else
                    gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit = 0 WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO!OR_NUM))
                End If
                StoreDetails
                
                'Updated by JJE 11/05/2012 2PM
                If Null2String(rsBANKDEPO!Type) = "1" Then
                    gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                      " CASH = CASH + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                      " CASHDEPO = CASHDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                      " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
'                                      " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
                ElseIf Null2String(rsBANKDEPO!Type) = "2" Then
                    gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                      " [CHECK] = [CHECK] + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                      " CHECKDEPO = CHECKDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                      " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
'                                      " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
                ElseIf Null2String(rsBANKDEPO!Type) = "3" Then
                    gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                      " CARD = CARD + " & NumericVal(rsBANKDEPO!DEPOSIT) & "," & _
                                      " CARDDEPO = CARDDEPO - " & NumericVal(rsBANKDEPO!DEPOSIT) & _
                                      " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
'                                      " WHERE CUTDATE = '" & Format(CDate(txtDatDeposit), "MM/DD/YYYY") & "'")
                End If
                'JJE
            Else
                Exit Function
            End If
        End If
    End If
    TranDete = True
    Exit Function
    
ErrorCode:
    TranDete = False
    error_msg = error
End Function

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
        ShowTransactionsGridDetails grdCheckCardTransactions.Text
        fraDetails.Enabled = False
        Picture5.Enabled = False
    End If
End Sub

Private Sub lsvDet_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim nCnt                                                        As Integer
    If lblTotal.Caption <> "0.00" Then lblTotal.Caption = "0.00"
    For nCnt = 1 To lsvDet.ListItems.Count
        If lsvDet.ListItems.Item(nCnt).Checked = True Then
            lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(lsvDet.ListItems.Item(nCnt).SubItems(3)))
        End If
    Next
End Sub

Private Sub lsvDet_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.ListSubItems(4).Text <> "" Then
        Call ShowTransactionsGridDetails(Item.ListSubItems(5).Text)
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
    'rsBANKDEPO.Bookmark = rsFind(rsBANKDEPO.Clone, "DATDEPOSIT", lstBANKDEPO.SELECTedItem).Bookmark
    StoreDetails
End Sub

Private Sub lstBANKDEPO_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    'rsBANKDEPO.Bookmark = rsFind(rsBANKDEPO.Clone, "DATDEPOSIT", lstBANKDEPO.SELECTedItem).Bookmark
    StoreDetails
End Sub

Private Sub lstBANKDEPO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstBANKDEPO
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

Private Sub lstBANKDEPO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        On Error Resume Next
        lstBANKDEPO.SetFocus
    End If
End Sub

Sub StoreDetails2()
    Dim rsBANKDEPODet                                               As ADODB.Recordset
    Dim VTYPE                                                       As String
    Dim i                                                           As Long

    TOTAL_CASH_DEPOSIT = 0
    TOTAL_CHECK_DEPOSIT = 0
    TOTAL_CARD_DEPOSIT = 0
    InitGrid
    i = 0
    
    Set rsBANKDEPODet = New ADODB.Recordset
    Set rsBANKDEPODet = gconDMIS.Execute("SELECT * FROM CMIS_BankDepo WHERE DEPOSIT_TO = '" & SetBankCode(RTrim(LTrim(cboDeposit_To))) & "' AND DATDEPOSIT = '" & txtDatDeposit.Text & "' ORDER BY TYPE, ID ASC")
    If Not rsBANKDEPODet.EOF And Not rsBANKDEPODet.BOF Then
        rsBANKDEPODet.MoveFirst
        Do While Not rsBANKDEPODet.EOF
            i = i + 1
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                VTYPE = "CASH"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            ElseIf Null2String(rsBANKDEPODet!Type) = "2" Then
                VTYPE = "CHECK"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetBankName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            ElseIf Null2String(rsBANKDEPODet!Type) = "3" Then
                VTYPE = "CARD"
                grdBANKDEPO.AddItem VTYPE & _
                                    Chr(9) & " " & SetCustomerName(Null2String(rsBANKDEPODet!bankcode)) & _
                                    Chr(9) & Null2String(rsBANKDEPODet!timdeposit) & _
                                    Chr(9) & ToDoubleNumber(N2Str2Zero(rsBANKDEPODet!DEPOSIT)) & _
                                    Chr(9) & rsBANKDEPODet!Id & _
                                    Chr(9) & rsBANKDEPODet!OR_NUM
            End If
            If i = 1 Then grdBANKDEPO.RemoveItem 1
            
            If Null2String(rsBANKDEPODet!Type) = "1" Then
                TOTAL_CASH_DEPOSIT = TOTAL_CASH_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            ElseIf Null2String(rsBANKDEPODet!Type) = "2" Then
                TOTAL_CHECK_DEPOSIT = TOTAL_CHECK_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            ElseIf Null2String(rsBANKDEPODet!Type) = "3" Then
                TOTAL_CARD_DEPOSIT = TOTAL_CARD_DEPOSIT + N2Str2Zero(rsBANKDEPODet!DEPOSIT)
            End If
            rsBANKDEPODet.MoveNext
        Loop
    End If
    txtTotalCashAmt.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT)
    txtTotalCheckAmt.Text = ToDoubleNumber(TOTAL_CHECK_DEPOSIT)
    txtCardDeposit.Text = ToDoubleNumber(TOTAL_CARD_DEPOSIT)
    txtTotalDepositedAmount.Text = ToDoubleNumber(TOTAL_CASH_DEPOSIT + TOTAL_CHECK_DEPOSIT + TOTAL_CARD_DEPOSIT)
End Sub

Function CheckCutoff(xCutoffDate) As Boolean
    Dim rsProcessCutOff                                             As ADODB.Recordset
    Set rsProcessCutOff = New ADODB.Recordset
    rsProcessCutOff.Open "SELECT DISTINCT CUTDATE FROM CMIS_OFF_HD WHERE CUTDATE IN (SELECT CUTDATE FROM CMIS_CASH_POS WHERE CUTDATE='" & CDate(xCutoffDate) & "') AND CUTDATE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsProcessCutOff.EOF And Not rsProcessCutOff.BOF Then
        CheckCutoff = True
    End If
End Function

Sub ShowTransactionsGridDetails(XXX As Long)
    Dim rsBANKDEPO_Details                                          As New ADODB.Recordset
    Dim rsOFF_Details                                               As New ADODB.Recordset

    If cboType.Text = "CASH" Then
        Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE ID = " & XXX)
        If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
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
            txtCheckDate.Text = ""
            txtCheckType.Text = ""
            txtCheckNumber.Text = ""
            txtORNumber.Text = ""
            txtBankCode.Text = ""
            txtTimeCreate.Text = ""
            txtCheckAmount.Text = "0.00"
        End If
    ElseIf cboType.Text = "CARD" Then
        Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE ID = " & XXX)
        If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
            labTranID.Caption = rsBANKDEPO_Details!Id
            txtCheckDte.Text = Null2String(rsBANKDEPO_Details!carddate)
            txtCheckType.Text = ""
            txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!cardnumber)
            txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)
            txtBankCode.Text = Null2String(rsBANKDEPO_Details!cardbnkcde)
            txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
            txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CardAmount))
            Set rsOFF_Details = New ADODB.Recordset
            Set rsOFF_Details = gconDMIS.Execute("SELECT SUM(DISCOUNT) AS TOTAL_DISCOUNT, SUM(TAX) AS TOTAL_TAX FROM CMIS_Off_Dt WHERE OR_NUM = " & N2Str2Null(rsBANKDEPO_Details!OR_NUM))
            If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
                txtCheckAmount.Text = ToDoubleNumber(NumericVal(txtCheckAmount.Text) - (N2Str2Zero(rsOFF_Details!TOTAL_TAX) + N2Str2Zero(rsOFF_Details!TOTAL_DISCOUNT)))
            End If
            txtDeposit.Text = ToDoubleNumber(NumericVal(txtCheckAmount.Text))
        Else
            labTranID.Caption = ""
            txtCheckDate.Text = "": txtCheckType.Text = ""
            txtCheckNumber.Text = "": txtORNumber.Text = ""
            txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
        End If
    ElseIf cboType.Text = "CHECK" Then
        Set rsBANKDEPO_Details = New ADODB.Recordset
        If cboCheckTransactions.Text = "Cashier Collection" Then
            Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE ID = " & XXX)
            If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
                labTranID.Caption = rsBANKDEPO_Details!Id
                txtCheckType.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
                txtCheckDte.Text = Null2String(rsBANKDEPO_Details!CheckDate)
                txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!Tseke)
                txtBankCode.Text = Null2String(rsBANKDEPO_Details!bankcode)
                txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CHKAMOUNT))
                txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CHKAMOUNT))
                txtORNumber.Text = Null2String(rsBANKDEPO_Details!OR_NUM)
                txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
            Else
                labTranID.Caption = ""
                txtBankDeposit = LOGDATE
                dtSelectedDate.Value = LOGDATE
                txtCheckDate.Text = "": txtCheckType.Text = ""
                txtCheckNumber.Text = "": txtORNumber.Text = ""
                txtBankCode.Text = "": txtTimeCreate.Text = "": txtCheckAmount.Text = "0.00"
            End If
            Set rsBANKDEPO_Details = Nothing
        ElseIf cboCheckTransactions.Text = "Check Encashment" Then
            Set rsBANKDEPO_Details = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE ID = " & XXX)
            If Not rsBANKDEPO_Details.EOF And Not rsBANKDEPO_Details.BOF Then
                labTranID.Caption = rsBANKDEPO_Details!Id
                txtCheckDte.Text = Null2String(rsBANKDEPO_Details!CHKDATE)
                txtCheckType.Text = SetCheckClass(Null2String(rsBANKDEPO_Details!Tseklase))
                txtCheckNumber.Text = Null2String(rsBANKDEPO_Details!CHKNUMBER)
                txtBankCode.Text = Null2String(rsBANKDEPO_Details!bankcode)
                txtTimeCreate.Text = Null2String(rsBANKDEPO_Details!TimeCreate)
                txtCheckAmount.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CHKAMOUNT))
                txtDeposit.Text = ToDoubleNumber(N2Str2Zero(rsBANKDEPO_Details!CHKAMOUNT))
            Else
                labTranID.Caption = ""
                txtBankDeposit = LOGDATE
                dtSelectedDate.Value = LOGDATE
                txtCheckDate.Text = ""
                txtCheckType.Text = ""
                txtCheckNumber.Text = ""
                txtBankCode.Text = ""
                txtTimeCreate.Text = ""
                txtCheckAmount.Text = "0.00"
            End If
            Set rsBANKDEPO_Details = Nothing
        End If
    End If
End Sub

Private Sub txtORSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'InitTransactionsGrid
            Dim rsCHECKDet                                              As ADODB.Recordset
            Dim i                                                       As Long
            Dim xList                                                   As ListItem
            lsvDet.ListItems.Clear
            If cboType.Text = "CARD" Then
                If COMPANY_CODE = M_COMPANY_CODE Then
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit3 = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                Else
                    If COMPANY_CODE = "HGC" Then
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
                    Else
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CardAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                    End If
                End If
                If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                    rsCHECKDet.MoveFirst
                    Do While Not rsCHECKDet.EOF
                        Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                        xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                        xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                        xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CardAmount))
                        xList.SubItems(4) = rsCHECKDet!Id
                        rsCHECKDet.MoveNext
                    Loop
                End If
            End If
    
            If cboType.Text = "CASH" Then
                If COMPANY_CODE = M_COMPANY_CODE Then
                    Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit1 = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                Else
                    If COMPANY_CODE = "HGC" Then
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
                    Else
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND CashAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                    End If
                End If
                If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                    rsCHECKDet.MoveFirst
                    Do While Not rsCHECKDet.EOF
                        Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!CUSCDE))
                        xList.SubItems(1) = Null2String(rsCHECKDet!cusname)
                        xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                        xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CashAmount))
                        xList.SubItems(4) = rsCHECKDet!Id
                        rsCHECKDet.MoveNext
                    Loop
                End If
            End If
    
            If cboCheckTransactions.Text = "Cashier Collection" Then
                Set rsCHECKDet = New ADODB.Recordset
                If cboType.Text = "CHECK" Then
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit2 = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                    Else
                        If COMPANY_CODE = "HGC" Then
                            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
                        Else
                            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                        End If
                    End If
                    If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                        rsCHECKDet.MoveFirst
                        Do While Not rsCHECKDet.EOF
                            Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                            xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                            xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                            xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                            xList.SubItems(4) = rsCHECKDet!Id
                            rsCHECKDet.MoveNext
                        Loop
                    End If
                End If
            End If
    
            If cboCheckTransactions.Text = "Check Encashment" Then
                Set rsCHECKDet = New ADODB.Recordset
                If cboType.Text = "CHECK" Then
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE Deposit2 = 0 AND chkAmount > 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                    Else
                        If COMPANY_CODE = "HGC" Then
                            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE Deposit = 0 AND chkAmount > 0 AND CANCEL = 0 AND OR_Num LIKE '" & txtORSearch.Text & "' AND OR_DATE >= '2/1/2010' ORDER BY ID ASC")
                        Else
                            Set rsCHECKDet = gconDMIS.Execute("SELECT * FROM CMIS_InCash WHERE Deposit = 0 AND chkAmount > 0 AND OR_Num LIKE '" & txtORSearch.Text & "' ORDER BY ID ASC")
                        End If
                    End If
                    If Not rsCHECKDet.EOF And Not rsCHECKDet.BOF Then
                        rsCHECKDet.MoveFirst
                        Do While Not rsCHECKDet.EOF
                            Set xList = lsvDet.ListItems.Add(, , Null2String(rsCHECKDet!bankcode))
                            xList.SubItems(1) = SetBankName(Null2String(rsCHECKDet!bankcode))
                            xList.SubItems(2) = Null2String(rsCHECKDet!OR_DATE)
                            xList.SubItems(3) = ToDoubleNumber(N2Str2Zero(rsCHECKDet!CHKAMOUNT))
                            xList.SubItems(4) = rsCHECKDet!Id
                            rsCHECKDet.MoveNext
                        Loop
                    End If
                End If
            End If
            txtORSearch.SetFocus
    End If
End Sub

Private Sub cmdPrint_Bankdepo_Click()
    If dtp_Frm_BD.Value > dtp_To_BD.Value Then
            MsgBox "Invalid date range.", vbInformation, ""
            dtp_Frm_BD.Value = LOGDATE
            dtp_To_BD.Value = LOGDATE
            On Error Resume Next
            dtp_Frm_BD.SetFocus
            Exit Sub
    End If
    
    
    'If Function_Access(LOGID, "Acess_Print", "TRANSACTION BANKDEPOSIT") = False Then Exit Sub
    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:

    '    txtDatDeposit.Text = lstBANKDEPO.SelectedItem
    '    If IsDate(txtDatDeposit.Text) = False Then
    '        MsgBox "Please click Date Deposited.", vbInformation, "Check Date"
    '        Exit Sub
    '    End If

    Screen.MousePointer = 11
    With rptBankDepo
        .Formulas(0) = "DEALER_NAME = '" & COMPANY_NAME & "'"
        .Formulas(1) = "DEALER_ADDRESS = '" & COMPANY_ADDRESS & "'"
'        .Formulas(2) = "PREPAREDBY= '" & PreparedBy & "'"
'        .Formulas(3) = "NOTEDBY= '" & NotedBy & "'"
'        .Formulas(4) = "CHECKEDBY= '" & CheckedBy & "'"
'        .Formulas(5) = "PRINTEDBY= " & N2Str2Null(LOGNAME)
    End With
    
    If OptPrintAsof.Value = True Then
        If CboBankDepo.Text = "ALL" Then
            PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit_All.rpt", "{CMIS_BankDepo.DATDEPOSIT} >= Date(" & Year(dtp_Frm_BD) & "," & Month(dtp_Frm_BD) & "," & Day(dtp_Frm_BD) & ") AND {CMIS_BankDepo.DATDEPOSIT} <= Date(" & Year(dtp_To_BD) & "," & Month(dtp_To_BD) & "," & Day(dtp_To_BD) & ") ", CMIS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit_Bank.rpt", "{CMIS_BANKDEPO.DATDEPOSIT} >= Date(" & Year(dtp_Frm_BD) & "," & Month(dtp_Frm_BD) & "," & Day(dtp_Frm_BD) & ") AND {CMIS_BANKDEPO.DATDEPOSIT} <= Date(" & Year(dtp_To_BD) & "," & Month(dtp_To_BD) & "," & Day(dtp_To_BD) & ") AND {ALL_BANKDEPOSITS.BANKNAME} = '" & CboBankDepo.Text & "'", CMIS_REPORT_CONNECTION, 1
        End If
    Else
        PrintSQLReport rptBankDepo, CMIS_REPORT_PATH & "BankDeposit.rpt", "{CMIS_BankDepo.DatDeposit} = Date(" & Year(txtDatDeposit.Text) & "," & Month(txtDatDeposit.Text) & "," & Day(txtDatDeposit.Text) & ") AND {CMIS_BankDepo.DEPOSIT_TO} = '" & SetBankCode(cboDeposit_To) & "'", CMIS_REPORT_CONNECTION, 1
    End If
    
    Screen.MousePointer = 0
    cmdClose_Bankdepo.Value = True
    Exit Sub


ErrorCode:
    ShowVBError
End Sub

Private Sub cmdClose_Bankdepo_Click()
    Pic_PrintBankDepo.ZOrder 1
    lstBANKDEPO.Enabled = True
    textSearch.Enabled = True
    cboDeposit_To.Enabled = True
    cmdShow.Enabled = True
End Sub

Private Sub OptPrint_Click()
    If txtDatDeposit.Text <> "" And cboDeposit_To.Text <> "" Then
        If OptPrint.Value = True Then
            CboBankDepo.Enabled = False
            PicDateRange.Enabled = False
            lbl_PrintBankDepoAsof.Enabled = False
            lbl_PrintBankDepo.Enabled = True
            CboBankDepo.Clear
        End If
    Else
    OptPrintAsof.Value = True
    End If
End Sub

Private Sub OptPrintAsof_Click()
    If OptPrintAsof.Value = True Then
        CboBankDepo.Enabled = True
        PicDateRange.Enabled = True
        lbl_PrintBankDepo.Enabled = False
        lbl_PrintBankDepoAsof.Enabled = True
        Call getPrintBank
        dtp_Frm_BD.Value = LOGDATE
        dtp_To_BD.Value = LOGDATE
        If cboDeposit_To.Text = "" Then
            CboBankDepo.Text = "ALL"
        Else
            CboBankDepo.Text = cboDeposit_To.Text
        End If
    End If
End Sub

Sub getPrintBank()
    Dim rsgetPrintBank As ADODB.Recordset
    Set rsgetPrintBank = New ADODB.Recordset
    Set rsgetPrintBank = gconDMIS.Execute("SELECT BANKNAME FROM ALL_BankDepositS ORDER BY BANKNAME ASC")
    If Not rsgetPrintBank.EOF And Not rsgetPrintBank.BOF Then
        Combo_Loadval CboBankDepo, rsgetPrintBank
        CboBankDepo.AddItem "ALL"
    End If
    Set rsgetPrintBank = Nothing
End Sub
