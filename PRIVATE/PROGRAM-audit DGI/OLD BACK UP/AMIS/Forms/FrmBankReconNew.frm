VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form FrmBankReconNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation Transaction"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FrmBankReconNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11745
      TabIndex        =   107
      Top             =   8040
      Width           =   11745
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10410
         Picture         =   "FrmBankReconNew.frx":030A
         TabIndex        =   110
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton cmdReconHistory 
         Caption         =   "Recon History"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7170
         TabIndex        =   109
         Top             =   30
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdview 
         Caption         =   "&View"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6030
         Picture         =   "FrmBankReconNew.frx":0670
         TabIndex        =   112
         ToolTipText     =   "View"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton cmdOpening 
         Caption         =   "&Opening"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4830
         Picture         =   "FrmBankReconNew.frx":0AB2
         TabIndex        =   113
         ToolTipText     =   "Openning"
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "&Adjustment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3630
         Picture         =   "FrmBankReconNew.frx":0EF4
         TabIndex        =   114
         ToolTipText     =   "Adjust"
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Re&fresh"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2430
         Picture         =   "FrmBankReconNew.frx":3266
         TabIndex        =   115
         ToolTipText     =   "Refresh"
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "R&eports"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         Picture         =   "FrmBankReconNew.frx":3591
         TabIndex        =   116
         ToolTipText     =   "Report"
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "&Reload"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   30
         Picture         =   "FrmBankReconNew.frx":38AA
         TabIndex        =   117
         ToolTipText     =   "Reload"
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9270
         Picture         =   "FrmBankReconNew.frx":3D02
         TabIndex        =   111
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton Command 
         Height          =   315
         Left            =   8760
         TabIndex        =   108
         Top             =   60
         Visible         =   0   'False
         Width           =   345
      End
   End
   Begin VB.PictureBox picOutstanding 
      BackColor       =   &H00FF8080&
      Height          =   1845
      Left            =   9270
      ScaleHeight     =   1785
      ScaleWidth      =   2325
      TabIndex        =   72
      Top             =   8700
      Visible         =   0   'False
      Width           =   2385
      Begin VB.PictureBox Picture 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   0
         ScaleHeight     =   285
         ScaleWidth      =   4755
         TabIndex        =   75
         Top             =   0
         Width           =   4785
         Begin VB.CommandButton cmdCloseO 
            Caption         =   "X"
            Height          =   315
            Left            =   1890
            TabIndex        =   76
            Top             =   30
            Width           =   405
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Printing Option"
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
            TabIndex        =   77
            Top             =   0
            Width           =   1605
         End
      End
      Begin VB.OptionButton optDetailed 
         BackColor       =   &H00FF8080&
         Caption         =   "Detailed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   74
         Top             =   1110
         Width           =   2475
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H00FF8080&
         Caption         =   "By Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   73
         Top             =   660
         Width           =   2475
      End
   End
   Begin VB.Frame frameSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   60
      TabIndex        =   32
      Top             =   6210
      Width           =   6945
      Begin VB.CommandButton cmdViewMonthly 
         Caption         =   "&View Recon Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         Picture         =   "FrmBankReconNew.frx":4052
         TabIndex        =   106
         ToolTipText     =   "View"
         Top             =   1290
         Width           =   1785
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View Unreconciled Only"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   4140
         TabIndex        =   57
         Top             =   1020
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.TextBox txtLed 
         Height          =   315
         Left            =   750
         TabIndex        =   36
         Top             =   1320
         Width           =   2445
      End
      Begin VB.OptionButton optCheckNOR 
         Caption         =   "By Check No"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   35
         Top             =   990
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton optCheckNOR 
         Caption         =   "By OR No"
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   34
         Top             =   990
         Width           =   1425
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3270
         TabIndex        =   33
         Top             =   1290
         Width           =   1785
      End
      Begin VB.Label lblClearedWithdrawals 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4290
         TabIndex        =   105
         ToolTipText     =   "Cleared Withdrawals"
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label lblClearedDeposits 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1500
         TabIndex        =   104
         ToolTipText     =   "Cleared Deposits"
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cleared Withdrawals"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   4470
         TabIndex        =   103
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cleared Deposits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   102
         Top             =   300
         Width           =   1470
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   690
         TabIndex        =   101
         Top             =   570
         Width           =   735
      End
      Begin VB.Label lblOpening 
         Height          =   225
         Left            =   4590
         TabIndex        =   79
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblBankID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4200
         TabIndex        =   78
         Top             =   1320
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label27 
         Caption         =   "Find"
         Height          =   315
         Left            =   180
         TabIndex        =   37
         Top             =   1350
         Width           =   1065
      End
      Begin VB.Shape Shape 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   735
         Index           =   0
         Left            =   90
         Top             =   180
         Width           =   6765
      End
      Begin VB.Shape Shape 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   735
         Index           =   3
         Left            =   60
         Top             =   210
         Width           =   6765
      End
   End
   Begin VB.Frame frameView 
      Caption         =   "View Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   60
      TabIndex        =   58
      Top             =   6240
      Visible         =   0   'False
      Width           =   6945
      Begin VB.Frame Frame2 
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
         Height          =   645
         Left            =   2880
         TabIndex        =   68
         Top             =   150
         Width           =   3915
         Begin VB.TextBox txtSearchCheck 
            Height          =   315
            Left            =   1470
            TabIndex        =   69
            Text            =   "Text2"
            Top             =   240
            Width           =   2355
         End
         Begin VB.Label Label13 
            Caption         =   "Check No."
            Height          =   315
            Left            =   420
            TabIndex        =   70
            Top             =   270
            Width           =   1065
         End
      End
      Begin VB.OptionButton optLclearedDeposit 
         Caption         =   "Cleared Deposits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2250
         TabIndex        =   64
         Top             =   1110
         Width           =   2145
      End
      Begin VB.OptionButton optLunClearedDeposit 
         Caption         =   "Uncleared Deposits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4470
         TabIndex        =   63
         Top             =   1110
         Width           =   2265
      End
      Begin VB.OptionButton optLclearedWithdrawals 
         Caption         =   "Cleared Withdrawals"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4470
         TabIndex        =   62
         Top             =   1410
         Width           =   2445
      End
      Begin VB.OptionButton optLOutstandingCheck 
         Caption         =   "Outstanding Checks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2250
         TabIndex        =   61
         Top             =   1410
         Width           =   2325
      End
      Begin VB.OptionButton otpLall 
         Caption         =   "All Ledger/Acct No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   60
         Top             =   1080
         Width           =   2115
      End
      Begin VB.OptionButton optLStaledcheck 
         Caption         =   "Staled Checks"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   59
         Top             =   1380
         Width           =   2145
      End
      Begin VB.Label Label10 
         Caption         =   "N - entered/Outstanding"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   210
         TabIndex        =   67
         Top             =   210
         Width           =   2625
      End
      Begin VB.Label Label11 
         Caption         =   "C - Cleared Check"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   210
         TabIndex        =   66
         Top             =   750
         Width           =   2745
      End
      Begin VB.Label Label12 
         Caption         =   "S - Staled Check"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   210
         TabIndex        =   65
         Top             =   480
         Width           =   1845
      End
   End
   Begin VB.TextBox txtEndingBal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      ToolTipText     =   "Statement Balance"
      Top             =   6390
      Width           =   1935
   End
   Begin XtremeSuiteControls.TabControl SSTab1 
      Height          =   5370
      Left            =   0
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   810
      Width           =   11595
      _Version        =   655364
      _ExtentX        =   20452
      _ExtentY        =   9472
      _StockProps     =   64
      AllowReorder    =   -1  'True
      Appearance      =   4
      Color           =   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.HotTracking=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.MinTabWidth=   100
      ItemCount       =   2
      Item(0).Caption =   "Bank Reconciliation"
      Item(0).Tooltip =   "Bank Reconciliation"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "chkClear"
      Item(0).Control(1)=   "grdRecon"
      Item(1).Caption =   "Inquiry "
      Item(1).Tooltip =   "Inquiry "
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "txtdebitL"
      Item(1).Control(1)=   "txtcreditL"
      Item(1).Control(2)=   "Label23"
      Item(1).Control(3)=   "grdBankrecon"
      Begin FlexCell.Grid grdBankrecon 
         Height          =   4455
         Left            =   -69880
         TabIndex        =   119
         Top             =   390
         Visible         =   0   'False
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7858
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   30
      End
      Begin FlexCell.Grid grdRecon 
         Height          =   4905
         Left            =   60
         TabIndex        =   118
         Top             =   390
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   8652
         Cols            =   5
         DefaultFontSize =   8.25
         GridColor       =   12632256
         Rows            =   30
      End
      Begin VB.CheckBox chkClear 
         Height          =   225
         Left            =   9990
         TabIndex        =   71
         Top             =   60
         Width           =   225
      End
      Begin VB.TextBox txtcreditL 
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
         Left            =   -60580
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   55
         Top             =   4920
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtdebitL 
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
         Left            =   -62440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   54
         Top             =   4920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Crystal.CrystalReport rptBankRecon 
         Left            =   10110
         Top             =   750
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label23 
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
         Left            =   -63610
         TabIndex        =   56
         Top             =   4980
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.PictureBox picRecon 
      BackColor       =   &H00FFFFFF&
      Height          =   5475
      Left            =   2550
      ScaleHeight     =   5415
      ScaleWidth      =   7065
      TabIndex        =   13
      Top             =   810
      Visible         =   0   'False
      Width           =   7125
      Begin VB.TextBox txtBankAdjustment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2940
         TabIndex        =   94
         Text            =   "0.00"
         Top             =   5520
         Width           =   1875
      End
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6480
         Top             =   5610
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5910
         TabIndex        =   31
         Top             =   4920
         Width           =   1035
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4395
         Left            =   150
         ScaleHeight     =   4365
         ScaleWidth      =   6735
         TabIndex        =   16
         Top             =   480
         Width           =   6765
         Begin VB.TextBox txtBankCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "0.00"
            Top             =   2820
            Width           =   1875
         End
         Begin VB.TextBox txtUBankCharges 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            TabIndex        =   92
            Text            =   "0.00"
            Top             =   3540
            Width           =   1875
         End
         Begin VB.TextBox txtUDeposit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            TabIndex        =   91
            Text            =   "0.00"
            Top             =   3180
            Width           =   1875
         End
         Begin VB.TextBox txtUnadjustedBook 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "0.00"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.TextBox txtInterest 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "0.00"
            Top             =   2460
            Width           =   1875
         End
         Begin VB.TextBox txtDepinTransit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   1440
            Width           =   1875
         End
         Begin VB.TextBox txtAdjustedBook 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   3960
            Width           =   1875
         End
         Begin VB.TextBox txtAdjustedBank 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   3960
            Width           =   1875
         End
         Begin VB.TextBox txtUnadjustedBank 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "0.00"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.TextBox txtOutstanding 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1770
            Width           =   1875
         End
         Begin VB.Label lblUBankCharges 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4800
            TabIndex        =   100
            Top             =   3570
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblUDeposit 
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4800
            TabIndex        =   99
            Top             =   3240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Charges"
            Height          =   285
            Left            =   210
            TabIndex        =   97
            Top             =   2820
            Width           =   1695
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Interest"
            Height          =   285
            Left            =   210
            TabIndex        =   96
            Top             =   2490
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unidentified Bank Charges"
            Height          =   210
            Left            =   210
            TabIndex        =   88
            Top             =   3540
            Width           =   2505
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Unidentified Deposit"
            Height          =   405
            Left            =   210
            TabIndex        =   87
            Top             =   3180
            Width           =   2025
         End
         Begin VB.Label lblDateAsOf 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "As of "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2790
            TabIndex        =   30
            Top             =   60
            Width           =   3945
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00008000&
            X1              =   0
            X2              =   6720
            Y1              =   3930
            Y2              =   3930
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustments"
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Adjusted Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   120
            TabIndex        =   28
            Top             =   3990
            Width           =   1905
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding checks"
            Height          =   405
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   1905
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Deposits-in-transit"
            Height          =   405
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "Unadjusted Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   120
            TabIndex        =   25
            Top             =   1110
            Width           =   2265
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Book"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   4890
            TabIndex        =   24
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2910
            TabIndex        =   23
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "As of "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   330
            TabIndex        =   22
            Top             =   90
            Width           =   1605
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00008000&
            X1              =   4740
            X2              =   4740
            Y1              =   510
            Y2              =   4380
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            X1              =   -30
            X2              =   6750
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00008000&
            X1              =   -30
            X2              =   6750
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00008000&
            X1              =   2760
            X2              =   2760
            Y1              =   0
            Y2              =   4380
         End
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   7035
         TabIndex        =   14
         Top             =   0
         Width           =   7065
         Begin VB.CommandButton cmdMin 
            Caption         =   "-"
            Height          =   210
            Left            =   6810
            TabIndex        =   86
            Top             =   -60
            Width           =   240
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Reconciliation"
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
            Height          =   345
            Left            =   30
            TabIndex        =   15
            Top             =   60
            Width           =   6975
         End
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Adjustment"
         Height          =   405
         Left            =   330
         TabIndex        =   95
         Top             =   5520
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Label lblNote 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NOTE: Press ENTER key to apply unidentified bank charges"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   150
         TabIndex        =   93
         Top             =   5040
         Visible         =   0   'False
         Width           =   5805
      End
   End
   Begin VB.PictureBox PicDateRange 
      BackColor       =   &H8000000B&
      Height          =   2205
      Left            =   2580
      ScaleHeight     =   2145
      ScaleWidth      =   4335
      TabIndex        =   44
      Top             =   5385
      Width           =   4395
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3405
         MouseIcon       =   "FrmBankReconNew.frx":4494
         MousePointer    =   99  'Custom
         Picture         =   "FrmBankReconNew.frx":45E6
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Close Window"
         Top             =   1260
         Width           =   885
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   4365
         TabIndex        =   46
         Top             =   0
         Width           =   4395
         Begin VB.Label lblPicRange 
            BackStyle       =   0  'Transparent
            Caption         =   "Date Range"
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
            Left            =   120
            TabIndex        =   47
            Top             =   60
            Width           =   4005
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2520
         MouseIcon       =   "FrmBankReconNew.frx":4A31
         MousePointer    =   99  'Custom
         Picture         =   "FrmBankReconNew.frx":4B83
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Print Report"
         Top             =   1260
         Width           =   885
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   405
         Left            =   570
         TabIndex        =   48
         Top             =   630
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
         Format          =   20316161
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   405
         Left            =   2580
         TabIndex        =   49
         Top             =   630
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
         Format          =   20316161
         CurrentDate     =   38216
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   52
         Top             =   690
         Width           =   525
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2190
         TabIndex        =   51
         Top             =   690
         Width           =   435
      End
   End
   Begin VB.PictureBox picoption 
      BackColor       =   &H8000000B&
      Height          =   2550
      Left            =   60
      ScaleHeight     =   2490
      ScaleWidth      =   2475
      TabIndex        =   38
      Top             =   5385
      Width           =   2535
      Begin VB.CommandButton cmdReconciled 
         Caption         =   "Reconciled"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2070
         Width           =   2085
      End
      Begin VB.CommandButton cmdStaled 
         Caption         =   "Staled Checks"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1740
         Width           =   2085
      End
      Begin VB.CommandButton cmdOutstanding 
         Caption         =   "Outstanding Checks"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1410
         Width           =   2085
      End
      Begin VB.CommandButton cmdUncleared 
         Caption         =   "Deposits in Transit"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   1080
         Width           =   2085
      End
      Begin VB.CommandButton cmdClearedWithdrawals 
         Caption         =   "Cleared Withdrawals"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   750
         Width           =   2085
      End
      Begin VB.CommandButton cmdClearedDeposits 
         Caption         =   "Cleared Deposits"
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
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   420
         Width           =   2085
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -60
         ScaleHeight     =   345
         ScaleWidth      =   4755
         TabIndex        =   39
         Top             =   -30
         Width           =   4785
         Begin VB.CommandButton cmdCloseOption 
            Caption         =   "X"
            Height          =   315
            Left            =   2130
            TabIndex        =   40
            Top             =   0
            Width           =   405
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Printing Option"
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
            Left            =   150
            TabIndex        =   41
            Top             =   60
            Width           =   1605
         End
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   30
         TabIndex        =   42
         Top             =   3150
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         Picture         =   "FrmBankReconNew.frx":5022
         BarPicture      =   "FrmBankReconNew.frx":503E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label labCPB 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
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
         Height          =   225
         Left            =   30
         TabIndex        =   43
         Top             =   3030
         Width           =   5835
      End
   End
   Begin VB.Label lblAccount 
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   6390
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Label lblDeposit 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   315
      Left            =   9510
      TabIndex        =   11
      ToolTipText     =   "Deposits"
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label lblOutstanding 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   315
      Left            =   9510
      TabIndex        =   10
      ToolTipText     =   "Outstanding Checks"
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   315
      Left            =   9510
      TabIndex        =   9
      ToolTipText     =   "Balance"
      Top             =   7260
      Width           =   1935
   End
   Begin VB.Label lblDifference 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   9510
      TabIndex        =   8
      ToolTipText     =   "Unreconcilled Difference"
      Top             =   7530
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7290
      TabIndex        =   6
      Top             =   7320
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Ending Balance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7260
      TabIndex        =   5
      Top             =   6420
      Width           =   3030
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding Checks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7290
      TabIndex        =   4
      Top             =   7020
      Width           =   1710
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deposits in Transit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7290
      TabIndex        =   3
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unreconciled Difference"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   7290
      TabIndex        =   2
      Top             =   7590
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lblBank 
      BackStyle       =   0  'Transparent
      Caption         =   "Reconcile Account - Asia United Bank"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   90
      Width           =   8205
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmBankReconNew.frx":505A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   360
      Width           =   10965
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   795
      Index           =   1
      Left            =   30
      Top             =   30
      Width           =   11505
   End
   Begin VB.Shape Shape 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   1635
      Index           =   2
      Left            =   7110
      Top             =   6300
      Width           =   4455
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous Month"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next Month"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FrmBankReconNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsload4recon2                                 As ADODB.Recordset
Dim ctl                                           As Control
Dim xOpening                                      As Double
Dim xDeposit                                      As Double
Dim xOutstanding                                  As Double
Dim START_DEBIT                                   As Double
Dim START_CREDIT                                  As Double
Dim START_DEBIT_C                                 As Double
Dim START_CREDIT_C                                As Double
Dim xdebit                                        As Double
Dim xcredit                                       As Double
Dim xBALANCE                                      As Double
Dim xReference                                    As String
Dim Options                                       As String
Dim AdjustType                                    As String
Dim Reconstatus                                   As String
Dim GridNo                                        As Integer
Dim xReconMonth                                   As Date
Dim Search_mode                                   As Boolean
Dim Prev_Recon                                    As Boolean

Private Sub chkClear_Click()
    xdebit = NumericVal(lblDeposit)
    xcredit = NumericVal(lblOutstanding)
    Dim xSkip                                     As Boolean
    If chkClear.Value = 1 Then
        With grdRecon
            'cmdReload_Click
            For GridNo = 1 To .Rows - 1
                If NumericVal(.Cell(GridNo, 6).Text) < 1 Then
                    .Cell(GridNo, 6).Text = 1
                    xSkip = False
                Else
                    xSkip = True
                End If
                If NumericVal(.Cell(GridNo, 6).Text) >= 1 Then
                    If xSkip = False Then
                        xdebit = xdebit - NumericVal(.Cell(GridNo, 4).Text)
                        xcredit = xcredit - NumericVal(.Cell(GridNo, 5).Text)
                    End If
                End If
                START_DEBIT = xdebit
                START_CREDIT = xcredit
                lblDeposit = ToDoubleNumber(Round(START_DEBIT, 2))
                txtDepinTransit = lblDeposit
                lblOutstanding = ToDoubleNumber(Round(START_CREDIT, 2))
                txtOutstanding = ToDoubleNumber(lblOutstanding)
                '                lblClearedDeposits = ToDoubleNumber(NumericVal(lblClearedDeposits.Caption) + NumericVal(xDebit))
                '                lblClearedWithdrawals = ToDoubleNumber(NumericVal(lblClearedWithdrawals.Caption) + NumericVal(xCredit))
                txtAdjustedBank.Text = ToDoubleNumber(NumericVal(lblBalance.Caption) + NumericVal(xOpening))
                txtAdjustedBook.Text = ToDoubleNumber(NumericVal(txtUnadjustedBank.Text) + NumericVal(lblDeposit.Caption) - NumericVal(lblOutstanding.Caption))
                'ComputeEndBalance
                'xReference = (grdRecon.Cell(grdRecon.ActiveCell.Row, 3).Text)
            Next
        End With
    Else
        With grdRecon
            For GridNo = 1 To .Rows - 1
                If NumericVal(.Cell(GridNo, 6).Text) >= 1 Then
                    .Cell(GridNo, 6).Text = ""
                    xSkip = False
                Else
                    xSkip = True
                End If
                If NumericVal(.Cell(GridNo, 6).Text) < 1 Then
                    If xSkip = False Then
                        xdebit = xdebit + NumericVal(.Cell(GridNo, 4).Text)
                        xcredit = xcredit + NumericVal(.Cell(GridNo, 5).Text)
                    End If
                End If
                START_DEBIT = xdebit
                START_CREDIT = xcredit
                lblDeposit = ToDoubleNumber(Round(START_DEBIT, 2))
                txtDepinTransit = lblDeposit
                lblOutstanding = ToDoubleNumber(Round(START_CREDIT, 2))
                txtOutstanding = ToDoubleNumber(lblOutstanding)
                txtAdjustedBank.Text = ToDoubleNumber(NumericVal(lblBalance.Caption) + NumericVal(xOpening))
                txtAdjustedBook.Text = ToDoubleNumber(NumericVal(txtUnadjustedBank.Text) + NumericVal(lblDeposit.Caption) - NumericVal(lblOutstanding.Caption))
            Next
        End With
        '    cmdReload_Click
    End If

    If Function_Access(LOGID, "Acess_Edit", "BANK RECONCILIATION") = False Then Exit Sub

    Dim xVOUCHERNO, xJType, xCheckNo              As String
    Dim xJdate                                    As Date
    Dim X                                         As Long
    Screen.MousePointer = 11
    On Error Resume Next
    For X = 1 To grdRecon.Rows - 1
        xVOUCHERNO = grdRecon.Cell(X, 8).Text
        xJType = grdRecon.Cell(X, 9).Text
        xCheckNo = grdRecon.Cell(X, 3).Text
        xJdate = grdRecon.Cell(X, 1).Text
        If NumericVal(grdRecon.Cell(X, 6).Text) > 0 Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'C' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
            ' Update BY BTT
            gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon,BankID) Values('" & xVOUCHERNO & _
                             "'," & N2Str2Null(lblDateAsOf) & ",'" & xJType & "','C'," & N2Str2Null(xJdate) & "," & N2Str2Null(lblBankID.Caption) & ")"

        Else
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'N' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"

            gconDMIS.Execute "delete AMIS_reconstatus " & _
                             " where VoucherNo = '" & xVOUCHERNO & "' and JType = '" & xJType & "'"

        End If
        If NumericVal(grdRecon.Cell(X, 7).Text) > 0 Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'S' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
        End If
    Next X
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdjust_Click()
    If Module_Access(LOGID, "GENERAL JOURNAL", "TRANSACTION") = False Then Exit Sub
    JOURNALTYPE = "GJ"
    On Error Resume Next
    Unload frmAMISJournalEntry
    frmAMISJournalEntry.Show
End Sub

Private Sub cmdCancel_Click()
    PicDateRange.Visible = False
    picoption.Visible = True
End Sub

Private Sub cmdClearedDeposits_Click()
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    PicDateRange.ZOrder 0
    PicDateRange.Visible = True
    Options = "Cleared Deposits"
    lblPicRange.Caption = "Date Range: " & cmdClearedDeposits.Caption
    '    On Error GoTo Errorcode:
    '    Dim filter                                         As String
    '    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    '    Dim Ans                                            As String
    '    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    '    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    '    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    '
    '    Screen.MousePointer = 11
    '    'If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
    '        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='C' and {RECON.Jtype}='DRJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    '    'End If
    '
    '    Screen.MousePointer = 0
    '    LogAudit "V", "BANK RECONCILIATION", lblAccount
    '    Exit Sub
    'Errorcode:
    '    ShowVBError
End Sub

Private Sub cmdClearedWithdrawals_Click()
'Clear Widrawals
'reconstatus='C' and credit > 0 and jtype='CDJ'
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    PicDateRange.ZOrder 0
    PicDateRange.Visible = True
    Options = "Cleared Withdrawals"
    lblPicRange.Caption = "Date Range: " & cmdClearedWithdrawals.Caption
    '    On Error GoTo Errorcode:
    '    Dim filter                                         As String
    '    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    '    Dim Ans                                            As String
    '    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    '    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    '    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    '
    '    Screen.MousePointer = 11
    '    'If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
    '        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='C' and {RECON.Jtype}='CDJ' and {recon.credit}  > 0 ", DMIS_REPORT_Connection, 1
    '    'End If
    '
    '    Screen.MousePointer = 0
    '    LogAudit "V", "BANK RECONCILIATION", lblAccount
    '    Exit Sub
    'Errorcode:
    '    ShowVBError
End Sub

Private Sub cmdCloseO_Click()
    picOutstanding.Visible = False
End Sub

Private Sub cmdCloseOption_Click()
    picoption.Visible = False
    Picture1.Enabled = True
    frameSearch.Enabled = True
    SSTab1.Enabled = True
End Sub

Private Sub cmdEdit_Click()
    Unload Me
    Unload frmReconcileAccount
End Sub

Private Sub cmdMin_Click()
    picRecon.Visible = False
End Sub

Private Sub cmdOK_Click()
    picRecon.Visible = False
    AdjustType = ""
    'txtInterest.Text = "0.00"
    'txtBankAdjustment.Text = "0.00"
    'cmdSave.SetFocus
    'lblDifference = ToDoubleNumber(NumericVal(txtEndingBal.Text) - NumericVal(xOpening) + NumericVal(xOutstanding1) - NumericVal(xDeposit1))
End Sub

Private Sub cmdOpening_Click()
    FormExistsShow frmAMISbanksOpening
End Sub

Private Sub cmdOutstanding_Click()
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    PicDateRange.ZOrder 0
    PicDateRange.Visible = True
    Options = "Outstanding Checks"
    lblPicRange.Caption = "Date Range: " & cmdOutstanding.Caption
    'picOutstanding.Visible = True
    'optType.Value = False
    'optDetailed.Value = False
    'Options = "Outstanding"
    'optType.Caption = "Type"
    'optDetailed.Caption = "Detailed"
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    Dim rsPrint                                   As ADODB.Recordset
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & Format(dtpFrom, "mmmm dd, yyyy") & " - " & Format(dtpTo, "mmmm dd, yyyy") & "'"
    'rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalances)
    Set rsPrint = New ADODB.Recordset
    Screen.MousePointer = 11
    If Options = "Cleared Deposits" Then
        rsPrint.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where Jdate Between '" & CDate(dtpFrom) & "' and '" & CDate(dtpTo) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'C' and JType ='DRJ' and Debit > 0 Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPrint.EOF And Not rsPrint.BOF Then
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconTransit.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and {RECON.ReconStatus}='C' and {RECON.Jtype}='DRJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
        Else
            MsgBox "No record to print!", vbExclamation, "Check Date"
        End If
    ElseIf Options = "Cleared Withdrawals" Then
        rsPrint.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where Jdate Between '" & CDate(dtpFrom) & "' and '" & CDate(dtpTo) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'C' and JType ='CDJ' and Credit > 0 Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPrint.EOF And Not rsPrint.BOF Then
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconOCW.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and {RECON.ReconStatus}='C' and {RECON.Jtype}='CDJ' and {recon.credit}  > 0 ", DMIS_REPORT_Connection, 1
        Else
            MsgBox "No record to print!", vbExclamation, "Check Date"
        End If
    ElseIf Options = "Deposits in Transit" Then
        rsPrint.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where Jdate Between '" & CDate(dtpFrom) & "' and '" & CDate(dtpTo) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and JType <>'GJ' and Debit > 0 Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPrint.EOF And Not rsPrint.BOF Then
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconTransit.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and {RECON.ReconStatus}='N' and {RECON.Jtype}<>'GJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
        Else
            MsgBox "No record to print!", vbExclamation, "Check Date"
        End If
    ElseIf Options = "Outstanding Checks" Then
        rsPrint.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where Jdate Between '" & CDate(dtpFrom) & "' and '" & CDate(dtpTo) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and (JType ='CDJ' or JType ='BOB') Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPrint.EOF And Not rsPrint.BOF Then
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconOCW.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and {RECON.ReconStatus}='N' and ({recon.jtype} ='CDJ' or {recon.jtype} ='BOB')", DMIS_REPORT_Connection, 1
        Else
            MsgBox "No record to print!", vbExclamation, "Check Date"
        End If
    ElseIf Options = "Staled" Then
        rsPrint.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where Jdate Between '" & CDate(dtpFrom) & "' and '" & CDate(dtpTo) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'S' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPrint.EOF And Not rsPrint.BOF Then
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconOCW.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")) and {RECON.ReconStatus}='S'", DMIS_REPORT_Connection, 1
        Else
            MsgBox "No record to print!", vbExclamation, "Check Date"
        End If
    End If
    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdReconHistory_Click()
    frmReconHistory.Show
End Sub

Private Sub cmdRefresh_Click()
    frmReconcileAccount.ZOrder 0
    'lblBankID.Caption = ""
    'lblBank.Caption = ""
    Screen.MousePointer = 11
    For Each ctl In ControlS
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next ctl
    grdRecon.Rows = 2
    grdRecon.Cell(1, 1).Text = ""
    grdRecon.Cell(1, 2).Text = ""
    grdRecon.Cell(1, 3).Text = ""
    grdRecon.Cell(1, 4).Text = ""
    grdRecon.Cell(1, 5).Text = ""
    grdRecon.Cell(1, 6).Text = ""
    grdRecon.Cell(1, 7).Text = ""
    grdRecon.Cell(1, 8).Text = ""
    grdRecon.Cell(1, 9).Text = ""
    grdRecon.Cell(1, 10).Text = ""
    '    lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
    lblDeposit.Caption = "0.00"
    lblOutstanding.Caption = "0.00"
    lblBalance.Caption = "0.00"
    chkClear.Value = 0
    Screen.MousePointer = 0
    cmdRefresh.Enabled = False
    cmdReport.Enabled = False
End Sub

Private Sub cmdReload_Click()
    Dim rsLoad4Recon                              As ADODB.Recordset
    Dim X, e                                      As Integer
    START_DEBIT = 0
    START_CREDIT = 0
    START_DEBIT_C = 0
    START_CREDIT_C = 0
    xOutstanding = 0
    xDeposit = 0
    '    txtStartBal = "0.00"
    lblDeposit = "0.00"
    lblOutstanding = "0.00"
    lblDifference = "0.00"
    lblBalance = "0.00"
    txtInterest.Text = "0.00"
    'txtBankAdjustment.Text = "0.00"
    txtUDeposit.Text = "0.00"
    txtUBankCharges.Text = "0.00"
    lblClearedDeposits = "0.00"
    lblClearedWithdrawals = "0.00"
    Prev_Recon = False
    lblDateAsOf.Caption = frmReconcileAccount.dtCurrent.Value
    txtEndingBal.Text = ToDoubleNumber(rEndingBalance)
    Dim xx                                        As Integer
    grdRecon.Rows = 1: xx = 0
    ' rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount.Caption) & "' AND ReconStatus = 'N' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Dim varReference                              As String
    Dim vReconStatus                              As Byte
    Set rsLoad4Recon = New ADODB.Recordset
    grdRecon.AutoRedraw = False
    '    BackRecon
    '    Label21.Caption = "Computing Data.."
    If Search_mode = False Then
        If Check1.Value = 1 Then
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,CheckDate,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and Status <> 'C' AND JTYPE <> 'OPB' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,CheckDate,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and Status <> 'C' AND JTYPE <> 'OPB' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        ' Search mode = totoo
    Else

        If optCheckNOR(0).Value = True Then                ' By Check no
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' AND JTYPE <> 'OPB' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND JTYPE <> 'OPB' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            'By OR
        Else
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,CheckDate,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' AND JTYPE <> 'OPB' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND JTYPE <> 'OPB' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        txtLed.Text = ""
        Search_mode = False
    End If

    If Not rsLoad4Recon.EOF And Not rsLoad4Recon.EOF Then
        grdRecon.Rows = 1
        Do Until rsLoad4Recon.EOF
            If Trim(Null2String(rsLoad4Recon![Reconstatus])) = "C" Then
                vReconStatus = 1
            Else
                START_DEBIT = START_DEBIT + N2Str2Zero(rsLoad4Recon![DEBIT])
                START_CREDIT = START_CREDIT + N2Str2Zero(rsLoad4Recon![CREDIT])
                lblDeposit = ToDoubleNumber(START_DEBIT)
                lblOutstanding = ToDoubleNumber(START_CREDIT)
                vReconStatus = 0
            End If
            If Null2String(rsLoad4Recon!jtype) = "CDJ" Then
                varReference = "CHK#" & Null2String(rsLoad4Recon!CheckNo)
            ElseIf Null2String(rsLoad4Recon!jtype) = "DRJ" Then
                varReference = "OR#" & Null2String(rsLoad4Recon!INVOICENO)
            Else
                varReference = Null2String(rsLoad4Recon![ReferenceNo])
            End If

            grdRecon.AddItem Format(rsLoad4Recon![JDate], "mm/dd/yyyy") & vbTab & _
                             StrConv(rsLoad4Recon![remarks], vbProperCase) & vbTab & _
                             Format(Null2Date(rsLoad4Recon![CheckDate]), "mm/dd/yyyy") & vbTab & _
                             varReference & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![DEBIT]) & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![CREDIT]) & vbTab & _
                             vReconStatus & vbTab & _
                             "" & vbTab & _
                             rsLoad4Recon![VOUCHERNO] & vbTab & _
                             rsLoad4Recon![jtype] & vbTab & _
                             False
            rsLoad4Recon.MoveNext
            DoEvents
            '            PROGBAR.Value = PROGBAR.Value + 1
            '            Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
            '            Label22 = varRefirence
        Loop
    End If
    grdRecon.AutoRedraw = True
    grdRecon.Refresh
    lblDeposit = ToDoubleNumber(lblDeposit)
    lblOutstanding = ToDoubleNumber(lblOutstanding)

    Dim rsBegBalance                              As ADODB.Recordset
    Set rsBegBalance = New ADODB.Recordset
    rsBegBalance.Open "SELECT SUM(DEBIT) AS DEPOSIT,SUM(CREDIT) AS OUTSTANDING,SUM(DEBIT)-SUM(CREDIT) AS BALANCE,ACCT_CODE from AMIS_vw_Recondata where JTYPE <> 'OPB' AND jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' GROUP BY ACCT_CODE", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBegBalance.EOF And Not rsBegBalance.BOF Then
        xBALANCE = NumericVal(rsBegBalance!BALANCE)
    End If
    Set rsBegBalance = Nothing

    Dim rsBalance                                 As ADODB.Recordset
    Set rsBalance = New ADODB.Recordset
    'SELECT SUM(DEBIT) AS DEPOSIT,SUM(CREDIT) AS OUTSTANDING,SUM(DEBIT)-SUM(CREDIT) AS BALANCE,ACCT_CODE FROM AMIS_VW_RECONDATA WHERE JDATE <= '6/30/2008' AND BANKACCTNO = '03833-2000-172' GROUP BY ACCT_CODE

    rsBalance.Open "SELECT SUM(DEBIT) AS DEPOSIT,SUM(CREDIT) AS OUTSTANDING from AMIS_vw_Recondata where ReconStatus <> 'C' and jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and Status <> 'C'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBalance.EOF And Not rsBalance.BOF Then
        xOutstanding = N2Str2Zero(rsBalance!Outstanding)
        xDeposit = N2Str2Zero(rsBalance!DEPOSIT)
    End If
    Set rsBalance = Nothing
    lblBalance = ToDoubleNumber(Round(((NumericVal(LTrim(xDeposit))) - NumericVal(LTrim(xOutstanding))), 2))

    lblDifference = ToDoubleNumber(Round((NumericVal(txtEndingBal.Text) + lblDeposit.Caption - lblOutstanding.Caption) - NumericVal(lblBalance), 2))

    Dim rsOpening1                                As ADODB.Recordset
    Set rsOpening1 = New ADODB.Recordset
    rsOpening1.Open "select Starting_Balance AS OPENING from ALL_BANKS where BankAcctno = '" & Trim(lblAccount) & "' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsOpening1.EOF And Not rsOpening1.BOF Then
        xOpening = N2Str2Zero(rsOpening1!Opening)
    End If
    Set rsOpening1 = Nothing

    LogAudit "R", "BANK RECONCILIATION", lblAccount
    Screen.MousePointer = 0
    cmdRefresh.Enabled = True
    cmdReport.Enabled = True
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdReport_Click()
    picoption.Visible = True
    picoption.ZOrder 0
    Picture1.Enabled = False
    frameSearch.Enabled = False
    SSTab1.Enabled = False
End Sub

Private Sub cmdSave_Click()
    If NumericVal(lblDifference) <> 0 Then
        MsgBox "Difference must be equal to 0", vbInformation, "Bank Reconcillation"
    Else

        On Error GoTo ErrorCode:

        Dim rsOpening                             As ADODB.Recordset
        Set rsOpening = New ADODB.Recordset
        rsOpening.Open "select SUM(DEBIT)-SUM(CREDIT) AS OPENING from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'C'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsOpening.EOF And Not rsOpening.BOF Then
            lblOpening.Caption = rsOpening!Opening
        End If
        Set rsOpening = Nothing
        gconDMIS.Execute ("Update ALL_BANKS SET " & _
                          " STARTING_BALANCE = " & NumericVal(txtEndingBal.Text) & "," & _
                          " ENDING_BALANCE = " & NumericVal(txtAdjustedBank.Text) & "," & _
                          " BOOK_BALANCE = " & NumericVal(txtUnadjustedBook.Text) & "," & _
                          " BANK_BALANCE = " & NumericVal(txtUnadjustedBank.Text) & "," & _
                          " LASTDATE_RECON = " & N2Str2Null(CDate(lblDateAsOf.Caption)) & _
                          " WHERE BANKACCTNO = '" & Trim(lblAccount.Caption) & "'")

        'gconDMIS.Execute ("insert into AMIS_RECONHISTORY (BankID,ReconDate,Bank,Book,Adjusted) values (" & N2Str2Null(lblBankID.Caption) & ", " & N2Str2Null(lblDateAsOf.Caption) & "," & NumericVal(txtUnadjustedBank.Text) & "," & NumericVal(txtUnadjustedBook.Text) & ", " & NumericVal(txtAdjustedBank.Text) & ")")
        '    " STARTING_DIFFERENCE = " & NumericVal(txt3.Text) & "," & _
             '                    " ENDING_DIFFERENCE = " & NumericVal(txt4.Text) & "," & _
             'gconDMIS.Execute ("Update All_Banks Set Starting_Balance='" & NumericVal(lblEndingBal.Caption) & "',Ending_Balance = '" & NumericVal(txtEndingBal.Text) & "' where BankAcctNo ='" & lblaccount & "'")
        Screen.MousePointer = 0
        cmdSave.Enabled = False

        Dim GridNo                                As Integer
        Dim xlApp                                 As Excel.Application
        Dim xlBook                                As Excel.Workbook
        Dim xlSheet                               As Excel.Worksheet
        Dim xlRange                               As Excel.Range
        Dim xCounter                              As Integer
        Dim rsLoad4Recon                          As ADODB.Recordset

        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "JOURNALS\BankRecon.XLT")
        Set xlSheet = xlBook.Worksheets(1)
        xlSheet.Cells(1, "A") = COMPANY_NAME
        xlSheet.Cells(2, "A") = COMPANY_ADDRESS
        xlSheet.Cells(5, "A") = "As of: " + lblDateAsOf.Caption
        xlSheet.Cells(8, "A") = "Unadjusted Balance"
        'xlSheet.Cells(8, "A").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet.Cells(8, "A").Font.Bold = True
        xlSheet.Cells(8, "C") = txtUnadjustedBook
        xlSheet.Cells(8, "C").Font.Bold = True
        xlSheet.Cells(8, "D") = rEndingBalance
        xlSheet.Cells(8, "D").Font.Bold = True

        xCounter = 9
        With grdRecon
            Dim ReconImport                       As Integer
            xlSheet.Cells(xCounter, "A") = "Deposits in Transit"
            xlSheet.Cells(xCounter, "A").Font.Bold = True
            For GridNo = 1 To .Rows - 1
                If NumericVal(.Cell(GridNo, 6).Text) = 0 Then
                    If NumericVal(.Cell(GridNo, 4).Text) > 0 Then
                        xlSheet.Cells(xCounter, "D") = (.Cell(GridNo, 1).Text)
                        'xlSheet.Cells(xCounter, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                        xlSheet.Cells(xCounter, "E") = (.Cell(GridNo, 3).Text)
                        xlSheet.Cells(xCounter, "F") = NumericVal(.Cell(GridNo, 4).Text)
                        xCounter = xCounter + 1
                    End If
                    DoEvents
                End If
            Next GridNo
            xCounter = xCounter + 1
            xlSheet.Cells(xCounter, "A") = "Outstanding Checks"
            xlSheet.Cells(xCounter, "A").Font.Bold = True
            For GridNo = 1 To .Rows - 1
                If NumericVal(.Cell(GridNo, 6).Text) = 0 Then
                    If NumericVal(.Cell(GridNo, 5).Text) > 0 Then
                        xlSheet.Cells(xCounter, "D") = (.Cell(GridNo, 1).Text)
                        xlSheet.Cells(xCounter, "E") = (.Cell(GridNo, 3).Text)
                        xlSheet.Cells(xCounter, "F") = NumericVal(.Cell(GridNo, 5).Text)
                        xCounter = xCounter + 1
                    End If
                End If
                DoEvents
            Next GridNo
        End With
        xlSheet.Cells(xCounter + 1, "A") = "Adjustments"
        xlSheet.Cells(xCounter + 1, "A").Font.Bold = True
        xlSheet.Cells(xCounter + 2, "B") = "Interest"
        xlSheet.Cells(xCounter + 3, "B") = "Bank Charges"
        xlSheet.Cells(xCounter + 4, "D") = "Unidentified Deposit"
        xlSheet.Cells(xCounter + 5, "D") = "Unidentified Bank Charges"
        xlSheet.Cells(xCounter + 7, "A") = "Adjusted Book Balance"
        xlSheet.Cells(xCounter + 7, "A").Font.Bold = True
        xlSheet.Cells(xCounter + 7, "C") = txtAdjustedBook.Text
        xlSheet.Cells(xCounter + 7, "C").Font.Bold = True
        xlSheet.Cells(xCounter + 7, "D") = "Adjusted Bank Balance"
        xlSheet.Cells(xCounter + 7, "D").Font.Bold = True
        xlSheet.Cells(xCounter + 7, "F") = txtAdjustedBank.Text
        xlSheet.Cells(xCounter + 7, "F").Font.Bold = True
        xlApp.Visible = True
        Set xlApp = Nothing

        MsgBox "Data Successfully updated", vbInformation, "Saved..."
        'cmdRefresh.Value = True
        LogAudit "A", "BANK RECONCILIATION", lblAccount
        Exit Sub
ErrorCode:
        ShowVBError
    End If
End Sub

Private Sub cmdSearch_Click()
    Dim Search_ChkOR                              As String
    Dim xLength                                   As Integer
    Dim GridNo, SearchGrid                        As Integer
    If txtLed.Text = "" Then
        MsgBox "Please input a reference!", vbInformation, "Information"
        txtLed.SetFocus
        Exit Sub
    End If
    Search_mode = True
    Search_ChkOR = txtLed.Text
    If Len(Search_ChkOR) = 0 Then Exit Sub

    Search_ChkOR = LCase$(Search_ChkOR)

    With grdRecon
        For GridNo = 1 To .Rows - 1
            xLength = Len(.Cell(GridNo, 3).Text)
            If optCheckNOR(0).Value = True Then
                If Len(Search_ChkOR) = 0 Then Exit Sub
                If Mid(.Cell(GridNo, 3).Text, 5, xLength) = Search_ChkOR Then
                    SearchGrid = GridNo
                    .Cell(GridNo, 3).BackColor = QBColor(14)
                    Search_ChkOR = ""
                End If
            Else
                If Len(Search_ChkOR) = 0 Then Exit Sub
                If Mid(.Cell(GridNo, 3).Text, 4, xLength) = Search_ChkOR Then
                    SearchGrid = GridNo
                    .Cell(GridNo, 3).BackColor = QBColor(14)
                    Search_ChkOR = ""
                End If
            End If
            .TopRow = SearchGrid
        Next
        If SearchGrid = 0 Then
            MsgBox "No reference found. Kindly check your entry...", vbInformation, "Search Result"
        End If
    End With
    'cmdReload_Click
End Sub

Private Sub cmdStaled_Click()
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    PicDateRange.ZOrder 0
    PicDateRange.Visible = True
    Options = "Staled"
    lblPicRange.Caption = "Date Range: " & cmdStaled.Caption
    '    On Error GoTo Errorcode:
    '    Dim filter                                         As String
    '    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    '    Dim Ans                                            As String
    '    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    '    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'as of : " & lblDateAsOf & "'"
    '    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    '
    '    Screen.MousePointer = 11
    '        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='S'", DMIS_REPORT_Connection, 1
    '
    '    Screen.MousePointer = 0
    '    LogAudit "V", "BANK RECONCILIATION", lblAccount
    '    Exit Sub
    'Errorcode:
    '    ShowVBError
End Sub

Private Sub cmdUncleared_Click()
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    PicDateRange.ZOrder 0
    PicDateRange.Visible = True
    Options = "Deposits in Transit"
    lblPicRange.Caption = "Date Range: " & cmdUncleared.Caption
    '    On Error GoTo Errorcode:
    '    Dim filter                                         As String
    '    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    '    Dim Ans                                            As String
    '    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    '    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    '    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    '
    '    Screen.MousePointer = 11
    '        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N' and {RECON.Jtype}='GJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    '
    '    Screen.MousePointer = 0
    '    LogAudit "V", "BANK RECONCILIATION", lblAccount
    '    Exit Sub
    'Errorcode:
    '    ShowVBError
End Sub

Private Sub cmdview_Click()
    Dim rsAdjustment                              As ADODB.Recordset
    Dim rsUnidentified                            As ADODB.Recordset
    Dim xdebit, zDebit                            As Double
    Dim xcredit, zCredit                          As Double
    Dim xBook_Balance                             As Double

    picRecon.ZOrder 0
    Set rsAdjustment = New ADODB.Recordset
    rsAdjustment.Open "SELECT RS.VOUCHERNO,RS.DATE_CLEARED,RS.JTYPE,RS.RECON_STATUS,RS.DATE_BEFORE_RECON,RS.BANKID,RS.ADJUSTTYPE,RS.JTYPE,RD.DEBIT,RD.CREDIT FROM AMIS_RECONSTATUS RS INNER JOIN AMIS_VW_RECONDATA RD ON RS.VOUCHERNO=RD.VOUCHERNO AND RS.JTYPE=RD.JTYPE WHERE RS.DATE_CLEARED='" & lblDateAsOf & "' AND RS.BANKID='" & lblBankID & "' and RS.JType ='GJ'", gconDMIS, adOpenForwardOnly
    If Not rsAdjustment.EOF And Not rsAdjustment.BOF Then
        Do While Not rsAdjustment.EOF
            '        If rsAdjustment!AdjustType = "A" Then
            '            zDebit = zDebit + NumericVal(rsAdjustment!DEBIT)
            '            zCredit = zCredit + NumericVal(rsAdjustment!CREDIT)
            '        ElseIf rsAdjustment!AdjustType = "B" Then
            xdebit = xdebit + NumericVal(rsAdjustment!DEBIT)
            xcredit = xcredit + NumericVal(rsAdjustment!CREDIT)
            '        End If
            rsAdjustment.MoveNext
        Loop
        '        txtInterest.Text = ToDoubleNumber(xDebit)
        '        txtBankCharges.Text = ToDoubleNumber(xCredit)
        'txtBankAdjustment.Text = ToDoubleNumber(zDebit - zCredit)
    Else
        txtInterest = "0.00"
        txtBankCharges = "0.00"
    End If

    Set rsUnidentified = New ADODB.Recordset
    rsUnidentified.Open "Select UDeposit,UBankCharges,Book_Balance from All_Banks where ID = '" & lblBankID & "'", gconDMIS, adOpenForwardOnly
    If Not rsUnidentified.EOF And Not rsUnidentified.BOF Then
        txtUDeposit = ToDoubleNumber(NumericVal(rsUnidentified!UDeposit))
        txtUBankCharges = ToDoubleNumber(NumericVal(rsUnidentified!UBankCharges))
        xBook_Balance = ToDoubleNumber(NumericVal(rsUnidentified!Book_Balance))
    End If

    If Prev_Recon = True Then
        lblDateAsOf = xReconMonth
    Else
        lblDateAsOf.Caption = frmReconcileAccount.dtCurrent.Value
    End If
    picRecon.Visible = True
    txtDepinTransit.Text = ToDoubleNumber(lblDeposit.Caption)
    txtOutstanding.Text = ToDoubleNumber(lblOutstanding.Caption)
    txtUnadjustedBank.Text = ToDoubleNumber(txtEndingBal.Text)
    'lblBalance.Caption = ToDoubleNumber((NumericVal(xDeposit) - NumericVal(xOutstanding)) + NumericVal(txtInterest.Text) - NumericVal(txtBankCharges.Text))
    '    txtUnadjustedBook.Text = ToDoubleNumber(NumericVal(lblBalance.Caption) + NumericVal(xBook_Balance))
    txtUnadjustedBook.Text = ToDoubleNumber(xBALANCE)
    'txtUnadjustedBook.Text = ToDoubleNumber((NumericVal(txtUnadjustedBank.Text) + NumericVal(txtDepinTransit) + NumericVal(txtBankCharges.Text) - NumericVal(txtUDeposit.Text) + NumericVal(txtUBankCharges.Text)) - (NumericVal(txtOutstanding.Text) + NumericVal(txtInterest.Text)))
    txtAdjustedBank.Text = ToDoubleNumber((NumericVal(txtUnadjustedBank.Text) + NumericVal(txtDepinTransit)) - NumericVal(txtOutstanding.Text) - NumericVal(txtUDeposit.Text) + NumericVal(txtUBankCharges.Text))
    txtAdjustedBook.Text = ToDoubleNumber(NumericVal(txtUnadjustedBook.Text) + NumericVal(txtInterest) - NumericVal(txtBankCharges))
    lblDifference.Caption = ToDoubleNumber(NumericVal(txtAdjustedBank.Text) - NumericVal(txtAdjustedBook.Text))
    Set rsAdjustment = Nothing
    Set rsUnidentified = Nothing
End Sub

Private Sub cmdViewMonthly_Click()
    Dim rsLoad4Recon                              As ADODB.Recordset
    grdRecon.Rows = 1
    Dim vReconStatus                              As Byte
    Dim varReference                              As String
    Set rsLoad4Recon = New ADODB.Recordset
    grdRecon.AutoRedraw = False
    START_DEBIT = 0
    START_CREDIT = 0
    START_DEBIT_C = 0
    START_CREDIT_C = 0
    xOutstanding = 0
    xDeposit = 0
    lblDeposit = "0.00"
    lblOutstanding = "0.00"
    Prev_Recon = False
    '    lblClearedDeposits = "0.00"
    '    lblClearedWithdrawals = "0.00"
    If Search_mode = False Then
        If Check1.Value = 1 Then
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(lblDateAsOf)) & "' and jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(lblDateAsOf)) & "' and jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        xReconMonth = CDate(lblDateAsOf)
        'xReconMonth = xReconMonth - Day(xReconMonth)
        ' Search mode = totoo
    Else

        If optCheckNOR(0).Value = True Then                ' By Check no
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            'By OR
        Else
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        txtLed.Text = ""
        Search_mode = False
    End If

    If Not rsLoad4Recon.EOF And Not rsLoad4Recon.EOF Then
        grdRecon.Rows = 1
        Do Until rsLoad4Recon.EOF
            If Trim(Null2String(rsLoad4Recon![Reconstatus])) = "C" Then
                'START_DEBIT_C = START_DEBIT_C + N2Str2Zero(rsLoad4Recon![DEBIT])
                'START_CREDIT_C = START_CREDIT_C + N2Str2Zero(rsLoad4Recon![CREDIT])
                '                With grdRecon
                '                For gridno = 1 To .Rows - 1
                '                    If NumericVal(.Cell(gridno, 6).Text) >= 1 Then
                '                        START_DEBIT_C = START_DEBIT_C + NumericVal(.Cell(gridno, 4).Text)
                '                        START_CREDIT_C = START_CREDIT_C + NumericVal(.Cell(gridno, 5).Text)
                '                    End If
                '                Next gridno
                '                End With
                '                lblClearedDeposits = ToDoubleNumber(START_DEBIT_C)
                '                lblClearedWithdrawals = ToDoubleNumber(START_CREDIT_C)
                vReconStatus = 1
            Else
                START_DEBIT = START_DEBIT + N2Str2Zero(rsLoad4Recon![DEBIT])
                START_CREDIT = START_CREDIT + N2Str2Zero(rsLoad4Recon![CREDIT])
                lblDeposit = ToDoubleNumber(START_DEBIT)
                lblOutstanding = ToDoubleNumber(START_CREDIT)
                vReconStatus = 0
            End If
            If Null2String(rsLoad4Recon!jtype) = "CDJ" Then
                varReference = "CHK#" & Null2String(rsLoad4Recon!CheckNo)
            ElseIf Null2String(rsLoad4Recon!jtype) = "DRJ" Then
                varReference = "OR#" & Null2String(rsLoad4Recon!INVOICENO)
            Else
                varReference = Null2String(rsLoad4Recon![ReferenceNo])
            End If

            grdRecon.AddItem Format(rsLoad4Recon![JDate], "mm/dd/yyyy") & vbTab & _
                             StrConv(rsLoad4Recon![remarks], vbProperCase) & vbTab & _
                             varReference & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![DEBIT]) & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![CREDIT]) & vbTab & _
                             vReconStatus & vbTab & _
                             "" & vbTab & _
                             rsLoad4Recon![VOUCHERNO] & vbTab & _
                             rsLoad4Recon![jtype] & vbTab & _
                             False
            rsLoad4Recon.MoveNext
            DoEvents
            '            PROGBAR.Value = PROGBAR.Value + 1
            '            Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
            '            Label22 = varRefirence
        Loop
    End If
    With grdRecon
        For GridNo = 1 To .Rows - 1
            If NumericVal(.Cell(GridNo, 6).Text) >= 1 Then
                START_DEBIT_C = START_DEBIT_C + NumericVal(.Cell(GridNo, 4).Text)
                START_CREDIT_C = START_CREDIT_C + NumericVal(.Cell(GridNo, 5).Text)
            End If
        Next GridNo
    End With
    lblClearedDeposits = ToDoubleNumber(START_DEBIT_C)
    lblClearedWithdrawals = ToDoubleNumber(START_CREDIT_C)
    grdRecon.AutoRedraw = True
    grdRecon.Refresh
End Sub

Private Sub cmdViewMonthly_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuView
    End If
End Sub

Private Sub Command1_Click()
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    optDetailed.Value = False
    picOutstanding.Visible = False
    Screen.MousePointer = 11

    rptBankRecon.Formulas(7) = "Bankstatement= " & NumericVal(txtEndingBal)
    PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconStatement.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N' and ({recon.jtype} ='DRJ' or {recon.jtype} ='CDJ' or {recon.jtype} ='BOB')", DMIS_REPORT_Connection, 1

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
End Sub

Private Sub Command_Click()
    Dim GridNo                                    As Integer
    Dim xlApp                                     As Excel.Application
    Dim xlBook                                    As Excel.Workbook
    Dim xlSheet                                   As Excel.Worksheet
    Dim xlRange                                   As Excel.Range
    Dim xCounter                                  As Integer
    Dim rsLoad4Recon                              As ADODB.Recordset

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "JOURNALS\BankRecon.XLT")
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Cells(1, "A") = COMPANY_NAME
    xlSheet.Cells(2, "A") = COMPANY_ADDRESS
    xlSheet.Cells(5, "A") = "As of: " + lblDateAsOf.Caption
    xlSheet.Cells(8, "A") = "Unadjusted Balance"
    'xlSheet.Cells(8, "A").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet.Cells(8, "A").Font.Bold = True
    xlSheet.Cells(8, "C") = txtUnadjustedBook
    xlSheet.Cells(8, "C").Font.Bold = True
    xlSheet.Cells(8, "D") = rEndingBalance
    xlSheet.Cells(8, "D").Font.Bold = True

    xCounter = 9
    With grdRecon
        Dim ReconImport                           As Integer
        xlSheet.Cells(xCounter, "A") = "Deposits in Transit"
        xlSheet.Cells(xCounter, "A").Font.Bold = True
        For GridNo = 1 To .Rows - 1
            If NumericVal(.Cell(GridNo, 6).Text) = 0 Then
                If NumericVal(.Cell(GridNo, 4).Text) > 0 Then
                    xlSheet.Cells(xCounter, "D") = (.Cell(GridNo, 1).Text)
                    'xlSheet.Cells(xCounter, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet.Cells(xCounter, "E") = (.Cell(GridNo, 3).Text)
                    xlSheet.Cells(xCounter, "F") = NumericVal(.Cell(GridNo, 4).Text)
                    xCounter = xCounter + 1
                End If
                DoEvents
            End If
        Next GridNo
        xCounter = xCounter + 1
        xlSheet.Cells(xCounter, "A") = "Outstanding Checks"
        xlSheet.Cells(xCounter, "A").Font.Bold = True
        For GridNo = 1 To .Rows - 1
            If NumericVal(.Cell(GridNo, 6).Text) = 0 Then
                If NumericVal(.Cell(GridNo, 5).Text) > 0 Then
                    xlSheet.Cells(xCounter, "D") = (.Cell(GridNo, 1).Text)
                    xlSheet.Cells(xCounter, "E") = (.Cell(GridNo, 3).Text)
                    xlSheet.Cells(xCounter, "F") = NumericVal(.Cell(GridNo, 5).Text)
                    xCounter = xCounter + 1
                End If
            End If
            DoEvents
        Next GridNo
    End With
    xlSheet.Cells(xCounter + 1, "A") = "Adjustments"
    xlSheet.Cells(xCounter + 1, "A").Font.Bold = True
    xlSheet.Cells(xCounter + 2, "B") = "Interest"
    xlSheet.Cells(xCounter + 3, "B") = "Bank Charges"
    xlSheet.Cells(xCounter + 4, "D") = "Unidentified Deposit"
    xlSheet.Cells(xCounter + 5, "D") = "Unidentified Bank Charges"
    xlSheet.Cells(xCounter + 7, "A") = "Adjusted Book Balance"
    xlSheet.Cells(xCounter + 7, "A").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "C") = txtAdjustedBook.Text
    xlSheet.Cells(xCounter + 7, "C").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "D") = "Adjusted Bank Balance"
    xlSheet.Cells(xCounter + 7, "D").Font.Bold = True
    xlSheet.Cells(xCounter + 7, "F") = txtAdjustedBank.Text
    xlSheet.Cells(xCounter + 7, "F").Font.Bold = True
    xlApp.Visible = True
    Set xlApp = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        picRecon.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    cmdRefresh_Click
    '    optViewAll.Value = True
    picoption.Visible = False
    PicDateRange.Visible = False
    '    Picture4.Visible = False
    Search_mode = False
    'Set rsload4recon2 = New ADODB.Recordset
    'rsload4recon2.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(frmReconcileAccount.dtCurrent) & "' and BankAcctno = '" & Trim(frmReconcileAccount.cboBank) & "' AND ReconStatus = 'N' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    With grdRecon
        .Cols = 12: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = " Journal Date "
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Check Date"
        .Cell(0, 4).Text = "Reference"
        .Cell(0, 5).Text = "Deposits"
        .Cell(0, 6).Text = "Withdrawals"
        .Cell(0, 7).Text = "Cleared"
        .Cell(0, 8).Text = "Staled"
        .Cell(0, 9).Text = "CV#"
        .Cell(0, 10).Text = "Type"
        .Cell(0, 11).Text = "Date Cleared"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:                 '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:                 '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellCheckBox
        .Column(7).CellType = cellCheckBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox
        .Column(10).CellType = cellTextBox
        .Column(11).CellType = cellCalendar

        .Column(0).Width = 18
        .Column(1).Width = 80: .Column(1).Locked = True
        .Column(2).Width = 295: .Column(2).Locked = True
        .Column(3).Width = 80: .Column(3).Locked = True
        .Column(4).Width = 90: .Column(4).Locked = True
        .Column(5).Width = 80: .Column(5).Locked = True: .Column(5).Alignment = cellRightGeneral
        .Column(6).Width = 80: .Column(6).Locked = True: .Column(6).Alignment = cellRightGeneral
        .Column(7).Width = 55:
        .Column(8).Width = 45:
        .Column(9).Width = 0: .Column(9).Locked = True
        .Column(10).Width = 0: .Column(10).Locked = True
        .Column(11).Width = 85:                            '.Column(9).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 10, .Rows - 1, 10).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Private Sub ShortcutCaption2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmReconcileAccount
End Sub

Private Sub grdRecon_Click()
    On Error Resume Next
    If grdRecon.Rows <= 1 Then
        Exit Sub
    Else
        xdebit = NumericVal(lblDeposit)
        xcredit = NumericVal(lblOutstanding)
        'For X = 1 To grdRecon.Rows - 1
        If grdRecon.ActiveCell.Col = 6 Then
            If NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 6).Text) >= 1 Then
                xdebit = xdebit - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text)
                xcredit = xcredit - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text)
                lblClearedDeposits = ToDoubleNumber(NumericVal(lblClearedDeposits.Caption) + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text))
                lblClearedWithdrawals = ToDoubleNumber(NumericVal(lblClearedWithdrawals.Caption) + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text))
            Else
                xdebit = xdebit + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text)
                xcredit = xcredit + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text)
                lblClearedDeposits = ToDoubleNumber(NumericVal(lblClearedDeposits.Caption) - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text))
                lblClearedWithdrawals = ToDoubleNumber(NumericVal(lblClearedWithdrawals.Caption) - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text))
            End If
            'Next X

            START_DEBIT = xdebit
            START_CREDIT = xcredit
            lblDeposit = ToDoubleNumber(Round(START_DEBIT, 2))
            txtDepinTransit = lblDeposit
            lblOutstanding = ToDoubleNumber(Round(START_CREDIT, 2))
            txtOutstanding = ToDoubleNumber(lblOutstanding)
            txtAdjustedBank.Text = ToDoubleNumber((NumericVal(txtUnadjustedBank.Text) + NumericVal(txtDepinTransit)) - NumericVal(txtOutstanding.Text))
            txtAdjustedBook.Text = ToDoubleNumber(NumericVal(txtUnadjustedBook.Text) + NumericVal(txtInterest))
            'ComputeEndBalance
            xReference = (grdRecon.Cell(grdRecon.ActiveCell.Row, 3).Text)
            lblBalance = ToDoubleNumber(Round(NumericVal(lblDeposit - lblOutstanding), 2))
            lblDifference = ToDoubleNumber(Round((NumericVal(txtEndingBal.Text) + lblDeposit.Caption - lblOutstanding.Caption) - NumericVal(lblBalance), 2))
            'VerifyBackRecon (xReference)
            'txt4 = ToDoubleNumber(Round(NumericVal(txt1) - NumericVal(txtEndBal), 2))
        End If

        If grdRecon.ActiveCell.Col = 7 Then
            If NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 7).Text) >= 1 Then
                xdebit = xdebit - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text)
                xcredit = xcredit - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text)
                lblBalance = ToDoubleNumber(NumericVal(lblBalance) + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text))
                txtUnadjustedBook.Text = ToDoubleNumber(lblBalance.Caption)
            Else
                xdebit = xdebit + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 4).Text)
                xcredit = xcredit + NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text)
                lblBalance = ToDoubleNumber(NumericVal(lblBalance) - NumericVal(grdRecon.Cell(grdRecon.ActiveCell.Row, 5).Text))
                txtUnadjustedBook.Text = ToDoubleNumber(lblBalance.Caption)
            End If
            'Next X
            START_DEBIT = xdebit
            START_CREDIT = xcredit
            lblOutstanding = ToDoubleNumber(Round(START_CREDIT, 2))
            txtOutstanding = ToDoubleNumber(lblOutstanding)
            'ComputeEndBalance
            xReference = (grdRecon.Cell(grdRecon.ActiveCell.Row, 3).Text)
            VerifyBackRecon (xReference)
            lblDifference = ToDoubleNumber(Round((NumericVal(txtEndingBal.Text) + lblDeposit.Caption - lblOutstanding.Caption) - NumericVal(lblBalance), 2))
            'txt4 = ToDoubleNumber(Round(NumericVal(txt1) - NumericVal(txtEndBal), 2))
        End If

        If grdRecon.ActiveCell.Col = 10 Then
            grdRecon.ActiveCell.CellType = cellCalendar
        End If

        If Function_Access(LOGID, "Acess_Edit", "BANK RECONCILIATION") = False Then Exit Sub

        Dim xVOUCHERNO, xJType, xCheckNo          As String
        Dim xJdate                                As Date
        Dim X                                     As Long
        Screen.MousePointer = 11

        '    For x = 1 To grdRecon.Rows - 1

        With grdRecon
            xVOUCHERNO = .Cell(.ActiveCell.Row, 8).Text
            xJType = .Cell(.ActiveCell.Row, 9).Text
            xCheckNo = .Cell(.ActiveCell.Row, 3).Text
            xJdate = .Cell(.ActiveCell.Row, 1).Text
            If NumericVal(.Cell(.ActiveCell.Row, 6).Text) > 0 Then
                gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                 " ReconStatus = 'C' " & "" & _
                                 " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
                ' Update BY BTT
                If Prev_Recon = False Then
                    gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon,BankID) Values('" & xVOUCHERNO & _
                                     "'," & N2Str2Null(lblDateAsOf) & ",'" & xJType & "','C'," & N2Str2Null(xJdate) & "," & N2Str2Null(lblBankID.Caption) & ")"
                Else
                    gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon,BankID) Values('" & xVOUCHERNO & _
                                     "'," & N2Str2Null(xReconMonth) & ",'" & xJType & "','C'," & N2Str2Null(xJdate) & "," & N2Str2Null(lblBankID.Caption) & ")"
                End If
                '            If AdjustType = "A" Then
                '                gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon,BankID,AdjustType) Values('" & xVoucherNo & _
                                 '                                 "'," & N2Str2Null(lblDateAsOf) & ",'" & xJtype & "','C'," & N2Str2Null(xJdate) & "," & N2Str2Null(lblBankID.Caption) & ",'" & AdjustType & "')"
                '            ElseIf AdjustType = "B" Then
                '                gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon,BankID,AdjustType) Values('" & xVoucherNo & _
                                 '                                 "'," & N2Str2Null(lblDateAsOf) & ",'" & xJtype & "','C'," & N2Str2Null(xJdate) & "," & N2Str2Null(lblBankID.Caption) & ",'" & AdjustType & "')"
                '            End If

            Else
                gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                 " ReconStatus = 'N' " & "" & _
                                 " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"

                gconDMIS.Execute "delete AMIS_reconstatus " & _
                                 " where VoucherNo = '" & xVOUCHERNO & "' and JType = '" & xJType & "'"


            End If
            If NumericVal(.Cell(.ActiveCell.Row, 7).Text) > 0 Then
                gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                 " ReconStatus = 'S' " & "" & _
                                 " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
            End If
        End With
    End If
    Screen.MousePointer = 0

    '    Next x
End Sub

Function VerifyBackRecon(xReferenceNo As String) As Boolean
'Update By BTT 1022008
    Dim Ans                                       As String
    Dim finalAns                                  As String
    Dim SQL                                       As String
    Dim TheReferenceNo                            As String
    Dim theJtype                                  As String
    Dim temp                                      As String
    Dim RS                                        As New ADODB.Recordset

    TheReferenceNo = Right(xReferenceNo, 6)
    temp = Left(xReferenceNo, 3)

    'SQL = "select * from AMIS_reconStatus where voucherno='" & TheReferenceNo & "' and jtype='" & thejtype & "'"
    SQL = "select * from AMIS_reconStatus where voucherno='" & TheReferenceNo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        Ans = MsgBox("This has been already cleared.Do you want to ovewrite the transaction?", vbInformation + vbYesNo)
        If Ans = vbYes Then
            'Update the Data
            finalAns = MsgBox("Are you sure do you want to overwite this transaction", vbInformation + vbYesNo)
            If finalAns = vbYes Then
                gconDMIS.Execute "delete AMIS_reconstatus " & _
                                 " where VoucherNo = '" & TheReferenceNo & "'"
                LogAudit "X", "BANK RECONCILIATION", lblAccount
            End If
        Else
            ' Do nothing
        End If
    Else
        'Do nothing
    End If
    Set RS = Nothing
End Function

Sub BackRecon()
'Update By BTT 1022008
    Dim LookUpRecon                               As ADODB.Recordset
    Dim SQL                                       As String
    '    Label21.Caption = "Checking data.."
    '    If SSTab1.SelectedItem = 1 Then
    '        Picture4.Visible = True
    '    End If
    Dim RsCutDate                                 As New ADODB.Recordset
    SQL = "SELECT * from AMIS_reconstatus where date_before_recon < = '" & CDate(lblDateAsOf) & "'"

    Set RsCutDate = New ADODB.Recordset
    Set RsCutDate = gconDMIS.Execute(SQL)
    '    PROGBAR.Value = 0
    Do While Not RsCutDate.EOF

        Set LookUpRecon = New ADODB.Recordset
        LookUpRecon.Open "select * from AMIS_journal_hd where jdate <= '" & CDate(lblDateAsOf) & "' and voucherno='" & (RsCutDate!VOUCHERNO) & "' and jtype='" & (RsCutDate!jtype) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly

        '        PROGBAR.Max = RsCutDate.RecordCount

        If Not LookUpRecon.BOF And Not LookUpRecon.EOF Then
            If lblDateAsOf < RsCutDate!date_cleared Then
                gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                 " ReconStatus = 'N' " & "" & _
                                 " where VoucherNo = '" & (RsCutDate!VOUCHERNO) & "' AND JType = '" & (RsCutDate!jtype) & "'"
            Else
                gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                 " ReconStatus = 'C' " & "" & _
                                 " where VoucherNo = '" & (RsCutDate!VOUCHERNO) & "' AND JType = '" & (RsCutDate!jtype) & "'"
            End If
        End If
        RsCutDate.MoveNext
        DoEvents
        '        PROGBAR.Value = PROGBAR.Value + 1
        '        Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
        '        Label22 = Null2String((LookUpRecon!VOUCHERNO))
    Loop

    Set RsCutDate = Nothing
    Set LookUpRecon = Nothing
End Sub

Private Sub grdRecon_DblClick()
    If grdRecon.Rows <= 1 Then
        Exit Sub
    Else
        Dim VARVOUCHERNO                          As String
        If Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "APJ" Then
            JOURNALTYPE = "APJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "CDJ" Then
            JOURNALTYPE = "CDJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 2) = "SJ" Then
            JOURNALTYPE = "SJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "CRJ" Then
            JOURNALTYPE = "CRJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 2) = "GJ" Then
            JOURNALTYPE = "GJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "ADJ" Then
            JOURNALTYPE = "ADJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "PDJ" Then
            JOURNALTYPE = "PDJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "CLO" Then
            JOURNALTYPE = "CLO"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "DRJ" Then
            JOURNALTYPE = "DRJ"
        ElseIf Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 10).Text, 3) = "BOB" Then
            JOURNALTYPE = "BOB"
        Else
            JOURNALTYPE = "OPB"                            '
        End If
        'JOURNALTYPE = Left(grdRecon.Cell(grdRecon.ActiveCell.Row, 3).Text, 3)
        VARVOUCHERNO = Right(grdRecon.Cell(grdRecon.ActiveCell.Row, 9).Text, 6)
        Screen.MousePointer = 11
        On Error Resume Next
        If JOURNALTYPE = "DRJ" Then
            Unload frmAMISJournalEntry_DRJ
            frmAMISJournalEntry_DRJ.Show
            Call frmAMISJournalEntry_DRJ.StoreSearch(VARVOUCHERNO)

        ElseIf JOURNALTYPE = "BOB" Then
            Unload frmAMISbanksOpening
            frmAMISbanksOpening.Show
            Call frmAMISbanksOpening.StoreSearch(VARVOUCHERNO)

        Else
            Unload frmAMISJournalEntry
            frmAMISJournalEntry.Show
            Call frmAMISJournalEntry.StoreSearch(VARVOUCHERNO)
        End If
    End If
    Screen.MousePointer = 0
End Sub

Private Sub lblBank_Change()
    Dim rsAllBank                                 As ADODB.Recordset
    Set rsAllBank = gconDMIS.Execute("select * from All_Banks where BANKACCTNO = '" & lblAccount.Caption & "'")
    If Not rsAllBank.EOF And Not rsAllBank.BOF Then
        lblBankID.Caption = rsAllBank!ID
    End If
    cmdReload.Enabled = True
    cmdReload_Click
End Sub

Private Sub optAll_Click()

End Sub

Private Sub OptCD_Click()
'and jtype='GJ' and Debit>0
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='C' and {RECON.Jtype}='DRJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub mnuPrevious_Click()
    Dim rsLoad4Recon                              As ADODB.Recordset
    Set rsLoad4Recon = New ADODB.Recordset
    Dim vReconStatus                              As Byte
    Dim varReference                              As String
    grdRecon.Rows = 1
    grdRecon.AutoRedraw = False
    START_DEBIT = 0
    START_CREDIT = 0
    START_DEBIT_C = 0
    START_CREDIT_C = 0
    xOutstanding = 0
    xDeposit = 0
    lblDeposit = "0.00"
    lblOutstanding = "0.00"
    '    lblClearedDeposits = "0.00"
    '    lblClearedWithdrawals = "0.00"
    Prev_Recon = True
    xReconMonth = xReconMonth - Day(xReconMonth)
    If Search_mode = False Then
        If Check1.Value = 1 Then
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If

        ' Search mode = totoo
    Else

        If optCheckNOR(0).Value = True Then                ' By Check no
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            'By OR
        Else
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select [Reconstatus],[DEBIT],[CREDIT],JType,CheckNo,INVOICENO,[ReferenceNo],JDate,Remarks,VoucherNo from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' AND ReconStatus = 'N' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate >= '" & CDate(firstDay(xReconMonth)) & "' and jdate <= '" & CDate(xReconMonth) & "' and BankAcctno = '" & Trim(lblAccount) & "' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        txtLed.Text = ""
        Search_mode = False
    End If

    If Not rsLoad4Recon.EOF And Not rsLoad4Recon.EOF Then
        grdRecon.Rows = 1
        Do Until rsLoad4Recon.EOF
            If Trim(Null2String(rsLoad4Recon![Reconstatus])) = "C" Then
                vReconStatus = 1
            Else
                START_DEBIT = START_DEBIT + N2Str2Zero(rsLoad4Recon![DEBIT])
                START_CREDIT = START_CREDIT + N2Str2Zero(rsLoad4Recon![CREDIT])
                lblDeposit = ToDoubleNumber(START_DEBIT)
                lblOutstanding = ToDoubleNumber(START_CREDIT)
                vReconStatus = 0
            End If
            If Null2String(rsLoad4Recon!jtype) = "CDJ" Then
                varReference = "CHK#" & Null2String(rsLoad4Recon!CheckNo)
            ElseIf Null2String(rsLoad4Recon!jtype) = "DRJ" Then
                varReference = "OR#" & Null2String(rsLoad4Recon!INVOICENO)
            Else
                varReference = Null2String(rsLoad4Recon![ReferenceNo])
            End If

            grdRecon.AddItem Format(rsLoad4Recon![JDate], "mm/dd/yyyy") & vbTab & _
                             StrConv(rsLoad4Recon![remarks], vbProperCase) & vbTab & _
                             varReference & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![DEBIT]) & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![CREDIT]) & vbTab & _
                             vReconStatus & vbTab & _
                             "" & vbTab & _
                             rsLoad4Recon![VOUCHERNO] & vbTab & _
                             rsLoad4Recon![jtype] & vbTab & _
                             False
            rsLoad4Recon.MoveNext
            DoEvents
            '            PROGBAR.Value = PROGBAR.Value + 1
            '            Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
            '            Label22 = varRefirence
        Loop
    End If
    With grdRecon
        For GridNo = 1 To .Rows - 1
            If NumericVal(.Cell(GridNo, 6).Text) >= 1 Then
                START_DEBIT_C = START_DEBIT_C + NumericVal(.Cell(GridNo, 4).Text)
                START_CREDIT_C = START_CREDIT_C + NumericVal(.Cell(GridNo, 5).Text)
            End If
        Next GridNo
    End With
    lblClearedDeposits = ToDoubleNumber(START_DEBIT_C)
    lblClearedWithdrawals = ToDoubleNumber(START_CREDIT_C)
    grdRecon.AutoRedraw = True
    grdRecon.Refresh
End Sub

Private Sub optCheckNOR_Click(Index As Integer)
    txtLed.Text = ""
    txtLed.SetFocus
End Sub

Private Sub OptCW_Click()


End Sub

Private Sub optDetailed_Click()
    Dim filter                                    As String
    On Error GoTo ErrorCode:
    If Options = "All Ledger" Then
        If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
        rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
        rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)

        Screen.MousePointer = 11

        'If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        'PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "'" & filter, DMIS_REPORT_Connection, 1
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ")", DMIS_REPORT_Connection, 1
        'End If
        Screen.MousePointer = 0
        LogAudit "V", "BANK RECONCILIATION", lblAccount
        Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
        rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
        rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
        optDetailed.Value = False
        picOutstanding.Visible = False
        Screen.MousePointer = 11

        rptBankRecon.Formulas(7) = "Bankstatement= " & NumericVal(txtEndingBal)
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BasnkReconDetail.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N' and ({recon.jtype} ='CDJ' or {recon.jtype} ='BOB')", DMIS_REPORT_Connection, 1

        Screen.MousePointer = 0
        LogAudit "V", "BANK RECONCILIATION", lblAccount
        Exit Sub
    End If
ErrorCode:
    ShowVBError
End Sub

Private Sub Option2_Click()
'and {recon.status}='N' and {recon.jtype}='GJ' and {recon.debit} > 0
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N' and {RECON.Jtype}='GJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub OptStaledC_Click()
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'as of : " & lblDateAsOf & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='S'", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub optLclearedDeposit_Click()
    fillGridRecon
End Sub

Private Sub optLclearedWithdrawals_Click()
    fillGridRecon
End Sub

Private Sub optLOutstandingCheck_Click()
    fillGridRecon
End Sub

Private Sub optLStaledcheck_Click()
    fillGridRecon
End Sub

Private Sub optLunClearedDeposit_Click()
    fillGridRecon
End Sub

Private Sub optType_Click()
    Dim filter                                    As String
    On Error GoTo ErrorCode:
    If Options = "All Ledger" Then
        If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
        rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
        rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
        optType.Value = False
        picOutstanding.Visible = False
        Screen.MousePointer = 11

        dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
        dtpTo = LOGDATE
        PicDateRange.Visible = True
        'picoption.Visible = False
        'PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECON.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ")", DMIS_REPORT_Connection, 1

        Screen.MousePointer = 0
        LogAudit "V", "BANK RECONCILIATION", lblAccount
        Exit Sub
    ElseIf Options = "Outstanding" Then
        If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
        rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
        rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
        optType.Value = False
        picOutstanding.Visible = False
        Screen.MousePointer = 11

        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconGroup.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N'", DMIS_REPORT_Connection, 1

        Screen.MousePointer = 0
        LogAudit "V", "BANK RECONCILIATION", lblAccount
        Exit Sub
    End If
ErrorCode:
    ShowVBError
End Sub

Private Sub otpLall_Click()
    fillGridRecon
End Sub

Sub fillGridRecon()
'Update By BTT :
    Screen.MousePointer = 11
    Dim RS                                        As New ADODB.Recordset
    Dim Reference                                 As String
    'cleargrid grdBankrecon:
    InitGridRecon
    Dim TOTAL_CREDIT                              As Double
    Dim TOTAL_DEBIT                               As Double
    Dim cnt                                       As Integer

    TOTAL_CREDIT = 0
    TOTAL_DEBIT = 0
    grdBankrecon.Rows = 1
    With RS
        If otpLall.Value = True Then                       'All Ledger
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'N' order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        ElseIf optLclearedDeposit.Value = True Then        'Cleared Deposits
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'C' and JType = 'DRJ' and Debit > 0 order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        ElseIf optLunClearedDeposit.Value = True Then      'Uncleared Deposits
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'N' and (JType = 'GJ' or JType='DRJ') and Debit > 0 order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        ElseIf optLStaledcheck.Value = True Then           'Staled Checks
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'S' order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        ElseIf optLOutstandingCheck.Value = True Then      'Outstanding Checks
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'N' and (JType = 'CDJ' or JType='BOB') order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        ElseIf optLclearedWithdrawals.Value = True Then    'Cleared Withdrawals
            .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(lblDateAsOf) & "' and BankAcctno = '" & Trim(lblAccount) & "' and ReconStatus = 'C' and JType = 'CDJ' and Credit > 0 order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If

        If Not .EOF And Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1
                If Null2String(RS!jtype) = "DRJ" Then
                    Reference = Null2String(RS!AcctName)
                Else
                    Reference = Null2String(RS!nameofvendor)
                End If

                grdBankrecon.AddItem (RS!JDate) & vbTab & (RS!jtype) & vbTab & _
                                     (RS!VOUCHERNO) & vbTab & Reference & vbTab & _
                                     (RS!CheckDate) & vbTab & (RS!CheckNo) & vbTab & _
                                     (RS!DEBIT) & vbTab & (RS!CREDIT) & vbTab & (RS!Reconstatus) & vbTab & _
                                     (RS!remarks)
                TOTAL_CREDIT = TOTAL_CREDIT + NumericVal(RS!CREDIT)
                TOTAL_DEBIT = TOTAL_DEBIT + NumericVal(RS!DEBIT)
                .MoveNext
                DoEvents
            Loop
        End If
        'If cnt > 0 Then grdBankrecon.RemoveItem 1
        txtcreditL.Text = ToDoubleNumber(TOTAL_CREDIT)
        txtdebitL.Text = ToDoubleNumber(TOTAL_DEBIT)
    End With
    Screen.MousePointer = 0
    Set RS = Nothing
End Sub

Sub InitGridRecon()
    With grdBankrecon
        .Cols = 11: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        grdBankrecon.Rows = 2
        grdBankrecon.Cell(1, 1).Text = ""
        grdBankrecon.Cell(1, 2).Text = ""
        grdBankrecon.Cell(1, 3).Text = ""
        grdBankrecon.Cell(1, 4).Text = ""
        grdBankrecon.Cell(1, 5).Text = ""
        grdBankrecon.Cell(1, 6).Text = ""
        grdBankrecon.Cell(1, 7).Text = ""
        grdBankrecon.Cell(1, 8).Text = ""
        grdBankrecon.Cell(1, 9).Text = ""


        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Trandate"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "CV#"
        .Cell(0, 4).Text = "Customer/Vendor"
        .Cell(0, 5).Text = "Checkdate"
        .Cell(0, 6).Text = "Check No"
        .Cell(0, 7).Text = "Debit"
        .Cell(0, 8).Text = "Credit"
        .Cell(0, 9).Text = "Status"
        .Cell(0, 10).Text = "Remarks"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:                 '.Column(2).MaxLength = 50
        .Column(3).CellType = cellTextBox:                 '.Column(3).MaxLength = 50
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox
        .Column(10).CellType = cellTextBox

        .Column(0).Width = 18
        .Column(1).Width = 70: .Column(1).Locked = True: .Column(1).Alignment = cellCenterGeneral
        .Column(2).Width = 50: .Column(2).Locked = True: .Column(2).Alignment = cellCenterGeneral
        .Column(3).Width = 60: .Column(3).Locked = True: .Column(3).Alignment = cellCenterGeneral
        .Column(4).Width = 260: .Column(4).Locked = True: .Column(4).Alignment = cellLeftGeneral
        .Column(5).Width = 70: .Column(5).Locked = True: .Column(5).Alignment = cellCenterGeneral
        .Column(6).Width = 70: .Column(6).Locked = True: .Column(6).Alignment = cellLeftGeneral
        .Column(7).Width = 75: .Column(7).Locked = True: .Column(7).Alignment = cellRightGeneral
        .Column(8).Width = 75: .Column(8).Locked = True: .Column(8).Alignment = cellRightGeneral
        .Column(9).Width = 75: .Column(9).Locked = True
        .Column(10).Width = 375: .Column(9).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 10, .Rows - 1, 10).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Private Sub otpOut_Click()
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & lblDateAsOf & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(lblBalance)
    Screen.MousePointer = 11

    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconGroup.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N'", DMIS_REPORT_Connection, 1
    Else
        rptBankRecon.Formulas(7) = "Bankstatement= " & NumericVal(txtEndingBal)
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconDetail.RPT", "{RECON.BANKACCTNO}='" & lblAccount & "' and {RECON.JDATE} <= Date(" & Year(lblDateAsOf) & "," & Month(lblDateAsOf) & "," & Day(lblDateAsOf) & ") and {RECON.ReconStatus}='N' and ({recon.jtype} ='CDJ' or {recon.jtype} ='BOB')", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", lblAccount
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub SSTab1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    SEARCH_TAB = SSTab1.SelectedItem
    Select Case SEARCH_TAB
    Case 0
        frameSearch.Visible = True
        frameView.Visible = False
    Case 1
        frameSearch.Visible = False
        frameView.Visible = True
        'otpLall.Value = True
        'otpLall_Click
        InitGridRecon
        optLclearedDeposit.Value = False
        optLunClearedDeposit.Value = False
        optLStaledcheck.Value = False
        optLOutstandingCheck.Value = False
        optLclearedWithdrawals.Value = False
    End Select
End Sub

Private Sub Timer_Timer()
    If lblNote.Visible = True Then
        lblNote.Visible = False
    Else
        lblNote.Visible = True
    End If
End Sub

Private Sub txtInterest_GotFocus()
'Adjustment = True
'txtInterest.BackColor = &HC0FFFF
'AdjustType = "B"
End Sub

Private Sub txtInterest_LostFocus()
    txtInterest.BackColor = &HFFFFFF
End Sub

Private Sub txtBankAdjustment_GotFocus()
    txtBankAdjustment.BackColor = &HC0FFFF
    AdjustType = "A"
    'Adjustment = True
End Sub

Private Sub txtBankAdjustment_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        txtBankAdjustment.Text = ToDoubleNumber(txtBankAdjustment.Text)
        txtAdjustedBank.Text = ToDoubleNumber(NumericVal(txtAdjustedBank.Text) + NumericVal(txtBankAdjustment.Text))
        cmdOK.SetFocus
    End If
End Sub

Private Sub txtBankAdjustment_LostFocus()
    txtBankAdjustment.BackColor = &HFFFFFF
End Sub

Private Sub txtLed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtLed.Text = "" Then
            MsgBox "Please input a reference!", vbInformation, "Information"
            txtLed.SetFocus
        Else
            cmdSearch.SetFocus
        End If
    End If
End Sub

Private Sub txtSearchCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim Search_Chk                            As String
        Dim GridNo, SearchGrid                    As Integer
        If txtSearchCheck.Text = "" Then
            MsgBox "Please input a reference!", vbInformation, "Information"
            Exit Sub
        End If
        Search_Chk = txtSearchCheck.Text
        If Len(Search_Chk) = 0 Then Exit Sub

        Search_Chk = LCase$(Search_Chk)
        With grdBankrecon
            For GridNo = 1 To .Rows - 1
                If Len(Search_Chk) = 0 Then Exit Sub
                If .Cell(GridNo, 6).Text = Search_Chk Then
                    SearchGrid = GridNo
                    .Cell(GridNo, 6).BackColor = QBColor(14)
                    Search_Chk = ""
                End If
                .TopRow = SearchGrid
            Next
            If SearchGrid = 0 Then
                MsgBox "No reference found. Kindly check your entry...", vbInformation, "Search Result"
            End If
        End With
    End If
End Sub

Private Sub txtUBankCharges_GotFocus()
    txtUBankCharges.Text = NumericVal(txtUBankCharges.Text)
    lblUBankCharges = txtUBankCharges
    Timer.Enabled = True
    lblNote.Visible = True
    lblNote.Caption = "NOTE: Press ENTER key to apply unidentified bank charges"
End Sub

Private Sub txtUBankCharges_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtUBankCharges.Text = "" Then
            MsgBox "Invalid amount", vbExclamation, "Message"
        Else
            If MsgBox("Apply this unidentified bank charges?", vbQuestion + vbYesNo, "Unindentifed Bank Charges") = vbYes Then
                If txtUBankCharges.Text >= 0 Then
                    txtAdjustedBank.Text = NumericVal(txtAdjustedBank.Text) + NumericVal(txtUBankCharges.Text)
                    txtAdjustedBank.Text = ToDoubleNumber(txtAdjustedBank.Text)
                    txtUBankCharges.Text = ToDoubleNumber(txtUBankCharges.Text)
                    gconDMIS.Execute "UPDATE All_Banks Set UBankCharges = '" & NumericVal(txtUBankCharges.Text) & "' where ID = '" & lblBankID & "'"
                    lblUBankCharges = txtUBankCharges
                End If
            End If
        End If
    End If
End Sub

Private Sub txtUBankCharges_LostFocus()
    txtUBankCharges.Text = ToDoubleNumber(txtUBankCharges.Text)
    txtUBankCharges = ToDoubleNumber(lblUBankCharges)
    lblUBankCharges = 0
    lblNote.Visible = False
    Timer.Enabled = False
End Sub

Private Sub txtUDeposit_GotFocus()
    txtUDeposit.Text = NumericVal(txtUDeposit.Text)
    lblUDeposit = txtUDeposit
    Timer.Enabled = True
    lblNote.Visible = True
    lblNote.Caption = "NOTE: Press ENTER key to apply unidentified deposits"
End Sub

Private Sub txtUDeposit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
    If KeyAscii = 13 Then
        If txtUDeposit.Text = "" Then
            MsgBox "Invalid amount", vbExclamation, "Message"
        Else
            If MsgBox("Apply this unidentified deposit?", vbQuestion + vbYesNo, "Unindentifed Deposit") = vbYes Then
                If NumericVal(txtUDeposit.Text) >= 0 Then
                    txtAdjustedBank.Text = NumericVal(txtAdjustedBank.Text) - NumericVal(txtUDeposit.Text)
                    txtAdjustedBank.Text = ToDoubleNumber(txtAdjustedBank.Text)
                    txtUDeposit.Text = ToDoubleNumber(txtUDeposit.Text)
                    gconDMIS.Execute "UPDATE All_Banks Set UDeposit = '" & NumericVal(txtUDeposit.Text) & "' where ID = '" & lblBankID & "'"
                    lblUDeposit = txtUDeposit
                End If
            End If
        End If
    End If
End Sub

Private Sub txtUDeposit_LostFocus()
    txtUDeposit.Text = ToDoubleNumber(txtUDeposit.Text)
    txtUDeposit = ToDoubleNumber(lblUDeposit)
    lblUDeposit = 0
    lblNote.Visible = False
    Timer.Enabled = False
End Sub
