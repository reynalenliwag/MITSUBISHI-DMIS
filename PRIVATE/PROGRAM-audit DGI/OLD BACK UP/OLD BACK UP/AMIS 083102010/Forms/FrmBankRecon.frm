VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form FrmBankRecon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Reconciliation Transaction"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
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
   Icon            =   "FrmBankRecon.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11670
   Begin VB.CommandButton cmdview 
      Caption         =   "&View"
      Height          =   765
      Left            =   5940
      Picture         =   "FrmBankRecon.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   8520
      Width           =   1125
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2820
      ScaleHeight     =   4485
      ScaleWidth      =   6465
      TabIndex        =   84
      Top             =   3120
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtBalanceJ 
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
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   102
         Text            =   "0.00"
         Top             =   1620
         Width           =   1905
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   6495
         TabIndex        =   100
         Top             =   0
         Width           =   6525
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
            Left            =   90
            TabIndex        =   101
            Top             =   30
            Width           =   6405
         End
      End
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3405
         Left            =   150
         ScaleHeight     =   3375
         ScaleWidth      =   6165
         TabIndex        =   86
         Top             =   480
         Width           =   6195
         Begin VB.TextBox txtDebit 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "0.00"
            Top             =   1410
            Width           =   1875
         End
         Begin VB.TextBox txtCredit 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   103
            Text            =   "0.00"
            Top             =   1770
            Width           =   1875
         End
         Begin VB.TextBox txtbank 
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
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   89
            Text            =   "0.00"
            Top             =   1140
            Width           =   1875
         End
         Begin VB.TextBox txtbook 
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
            Left            =   2220
            Locked          =   -1  'True
            TabIndex        =   88
            Text            =   "0.00"
            Top             =   2970
            Width           =   1935
         End
         Begin VB.TextBox txttotalBook 
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
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "0.00"
            Top             =   2970
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00008000&
            X1              =   2190
            X2              =   2190
            Y1              =   0
            Y2              =   3360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00008000&
            X1              =   2190
            X2              =   6480
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00008000&
            X1              =   2190
            X2              =   6450
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00008000&
            X1              =   4170
            X2              =   4170
            Y1              =   510
            Y2              =   3390
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
            TabIndex        =   99
            Top             =   90
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
            Left            =   2370
            TabIndex        =   98
            Top             =   630
            Width           =   1605
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
            Left            =   4350
            TabIndex        =   97
            Top             =   630
            Width           =   1605
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
            TabIndex        =   96
            Top             =   1110
            Width           =   2265
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Deposit-in-transasit"
            Height          =   405
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   1905
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "Outstanding checks"
            Height          =   405
            Left            =   120
            TabIndex        =   94
            Top             =   1800
            Width           =   1905
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Other Uncleared "
            Height          =   405
            Left            =   270
            TabIndex        =   93
            Top             =   2130
            Width           =   1905
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
            Left            =   210
            TabIndex        =   92
            Top             =   2970
            Width           =   1905
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Book Adjustment"
            Height          =   405
            Left            =   240
            TabIndex        =   91
            Top             =   2490
            Width           =   1695
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00008000&
            X1              =   2190
            X2              =   6450
            Y1              =   2910
            Y2              =   2910
         End
         Begin VB.Label Label28 
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
            Left            =   2190
            TabIndex        =   90
            Top             =   60
            Width           =   3945
         End
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   465
         Left            =   4920
         TabIndex        =   85
         Top             =   3930
         Width           =   1425
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Openning"
      Height          =   765
      Left            =   4770
      Picture         =   "FrmBankRecon.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Refresh"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton cmdAdjust 
      Caption         =   "&Adjust"
      Height          =   765
      Left            =   3600
      Picture         =   "FrmBankRecon.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Refresh"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.PictureBox PicDateRange 
      BackColor       =   &H00FF8080&
      Height          =   1845
      Left            =   180
      ScaleHeight     =   1785
      ScaleWidth      =   4335
      TabIndex        =   48
      Top             =   6690
      Width           =   4395
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
         MouseIcon       =   "FrmBankRecon.frx":2F00
         MousePointer    =   99  'Custom
         Picture         =   "FrmBankRecon.frx":3052
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Print Report"
         Top             =   900
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
         TabIndex        =   55
         Top             =   0
         Width           =   4395
         Begin VB.Label Label19 
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
            TabIndex        =   56
            Top             =   60
            Width           =   1605
         End
      End
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   405
         Left            =   570
         TabIndex        =   51
         Top             =   420
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
         Format          =   51511297
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   405
         Left            =   2580
         TabIndex        =   52
         Top             =   420
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
         Format          =   51511297
         CurrentDate     =   38216
      End
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
         MouseIcon       =   "FrmBankRecon.frx":34F1
         MousePointer    =   99  'Custom
         Picture         =   "FrmBankRecon.frx":3643
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Close Window"
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2190
         TabIndex        =   54
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         TabIndex        =   53
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.PictureBox picoption 
      BackColor       =   &H00FF8080&
      Height          =   2445
      Left            =   120
      ScaleHeight     =   2385
      ScaleWidth      =   4425
      TabIndex        =   64
      Top             =   6120
      Width           =   4485
      Begin VB.OptionButton optAll 
         BackColor       =   &H00FF8080&
         Caption         =   "All Ledger/Acct No.(Crystal Report)"
         Height          =   345
         Left            =   30
         TabIndex        =   73
         Top             =   390
         Width           =   4365
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -60
         ScaleHeight     =   345
         ScaleWidth      =   4755
         TabIndex        =   70
         Top             =   -30
         Width           =   4785
         Begin VB.CommandButton cmdCloseOption 
            Caption         =   "X"
            Height          =   315
            Left            =   4020
            TabIndex        =   71
            Top             =   30
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
            TabIndex        =   72
            Top             =   60
            Width           =   1605
         End
      End
      Begin VB.OptionButton OptStaledC 
         BackColor       =   &H00FF8080&
         Caption         =   "Staled Checks "
         Height          =   345
         Left            =   30
         TabIndex        =   69
         Top             =   720
         Width           =   4365
      End
      Begin VB.OptionButton otpOut 
         BackColor       =   &H00FF8080&
         Caption         =   "Outstanding Checks "
         Height          =   345
         Left            =   30
         TabIndex        =   68
         Top             =   1050
         Width           =   4365
      End
      Begin VB.OptionButton OptCD 
         BackColor       =   &H00FF8080&
         Caption         =   "Cleared Deposits "
         Height          =   345
         Left            =   30
         TabIndex        =   67
         Top             =   1380
         Width           =   4365
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF8080&
         Caption         =   "Un-Cleared Deposits "
         Height          =   345
         Left            =   30
         TabIndex        =   66
         Top             =   1710
         Width           =   4365
      End
      Begin VB.OptionButton OptCW 
         BackColor       =   &H00FF8080&
         Caption         =   "Cleared Withdrawals "
         Height          =   345
         Left            =   30
         TabIndex        =   65
         Top             =   2040
         Width           =   4365
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   30
         TabIndex        =   74
         Top             =   3150
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   556
         Picture         =   "FrmBankRecon.frx":3A8E
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "FrmBankRecon.frx":3AAA
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
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
         TabIndex        =   75
         Top             =   3030
         Width           =   5835
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3330
      ScaleHeight     =   945
      ScaleWidth      =   5625
      TabIndex        =   57
      Top             =   3780
      Width           =   5625
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   5805
         TabIndex        =   62
         Top             =   0
         Width           =   5805
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   30
            TabIndex        =   63
            Top             =   0
            Width           =   3105
         End
      End
      Begin MSComctlLib.ProgressBar PROGBAR 
         Height          =   405
         Left            =   30
         TabIndex        =   58
         Top             =   480
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "OR"
         Height          =   255
         Left            =   510
         TabIndex        =   60
         Top             =   240
         Width           =   4365
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   60
      TabIndex        =   4
      Top             =   -30
      Width           =   11565
      Begin VB.CheckBox Check1 
         Caption         =   "View Un-Reconciled Only"
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
         Left            =   5970
         TabIndex        =   42
         Top             =   300
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.ComboBox cboBank 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   345
         Left            =   4440
         TabIndex        =   7
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   51511299
         CurrentDate     =   38946
      End
      Begin MSComCtl2.DTPicker dtTO 
         Height          =   345
         Left            =   10080
         TabIndex        =   8
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   51511299
         CurrentDate     =   38946
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Statement Date:"
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
         Left            =   8670
         TabIndex        =   9
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Account No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   2505
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Exit"
      Height          =   795
      Left            =   10470
      Picture         =   "FrmBankRecon.frx":3AC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   8505
      Width           =   1185
   End
   Begin XtremeSuiteControls.TabControl SSTab1 
      Height          =   7080
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1350
      Width           =   11595
      _Version        =   655364
      _ExtentX        =   20452
      _ExtentY        =   12488
      _StockProps     =   64
      AllowReorder    =   -1  'True
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.MinTabWidth=   100
      ItemCount       =   3
      Item(0).Caption =   "Inquiry "
      Item(0).Tooltip =   "Inquiry "
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Bank Reconciliation"
      Item(1).Tooltip =   "Bank Reconciliation"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Ledger"
      Item(2).Tooltip =   "Ledger"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Picture6"
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   6435
         Left            =   -69940
         ScaleHeight     =   6435
         ScaleWidth      =   11505
         TabIndex        =   106
         Top             =   600
         Visible         =   0   'False
         Width           =   11505
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
            Left            =   7560
            MaxLength       =   20
            TabIndex        =   111
            Top             =   5910
            Width           =   1815
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
            Left            =   9420
            MaxLength       =   20
            TabIndex        =   110
            Top             =   5910
            Width           =   1785
         End
         Begin VB.Frame Frame5 
            Caption         =   "Option"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   120
            TabIndex        =   109
            Top             =   60
            Width           =   9675
            Begin VB.OptionButton optLclearedDeposit 
               Caption         =   "View Cleared Deposits"
               Height          =   210
               Left            =   3270
               TabIndex        =   118
               Top             =   360
               Width           =   2685
            End
            Begin VB.OptionButton optLunClearedDeposit 
               Caption         =   "View Un-Cleared Deposits"
               Height          =   210
               Left            =   6480
               TabIndex        =   117
               Top             =   360
               Width           =   2835
            End
            Begin VB.OptionButton optLclearedWithdrawals 
               Caption         =   "View Cleared Withdrawals"
               Height          =   225
               Left            =   6480
               TabIndex        =   116
               Top             =   690
               Width           =   3015
            End
            Begin VB.OptionButton optLOutstandingCheck 
               Caption         =   "View Outstanding Checks"
               Height          =   210
               Left            =   3270
               TabIndex        =   115
               Top             =   690
               Width           =   2865
            End
            Begin VB.OptionButton otpLall 
               Caption         =   "View All Ledger/Acct No."
               Height          =   210
               Left            =   180
               TabIndex        =   114
               Top             =   330
               Value           =   -1  'True
               Width           =   2775
            End
            Begin VB.OptionButton optLStaledcheck 
               Caption         =   "View Staled Checks"
               Height          =   210
               Left            =   180
               TabIndex        =   113
               Top             =   660
               Width           =   2805
            End
         End
         Begin VB.CommandButton cmdReloadrecongrd 
            Caption         =   "Load"
            Height          =   435
            Left            =   9870
            TabIndex        =   108
            Top             =   210
            Width           =   1575
         End
         Begin MSFlexGridLib.MSFlexGrid grdBankrecon 
            Height          =   4635
            Left            =   0
            TabIndex        =   107
            Top             =   1230
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   8176
            _Version        =   393216
            Cols            =   8
            FixedCols       =   0
            BackColorFixed  =   14737632
            BackColorSel    =   16777088
            BackColorBkg    =   14737632
            TextStyleFixed  =   3
            HighLight       =   2
            FillStyle       =   1
            SelectionMode   =   1
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
            Left            =   6390
            TabIndex        =   112
            Top             =   5970
            Width           =   1395
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   6450
         Left            =   30
         TabIndex        =   12
         Top             =   600
         Width           =   11535
         _Version        =   655364
         _ExtentX        =   20346
         _ExtentY        =   11377
         _StockProps     =   0
         Begin VB.Frame Frame1 
            Height          =   6525
            Left            =   30
            TabIndex        =   13
            Top             =   90
            Width           =   11355
            Begin VB.OptionButton optStaled 
               Caption         =   "View Staled Checks"
               Height          =   210
               Left            =   4710
               TabIndex        =   22
               Top             =   510
               Width           =   2805
            End
            Begin VB.OptionButton optViewAll 
               Caption         =   "View All Ledger/Acct No."
               Height          =   210
               Left            =   4710
               TabIndex        =   21
               Top             =   240
               Width           =   2775
            End
            Begin VB.OptionButton optOutstanding 
               Caption         =   "View Outstanding Checks"
               Height          =   210
               Left            =   4710
               TabIndex        =   20
               Top             =   810
               Width           =   2865
            End
            Begin VB.OptionButton optClearedWithdrawal 
               Caption         =   "View Cleared Withdrawals"
               Height          =   225
               Left            =   8070
               TabIndex        =   19
               Top             =   810
               Width           =   3015
            End
            Begin VB.OptionButton optUnclearedDep 
               Caption         =   "View Un-Cleared Deposits"
               Height          =   210
               Left            =   8070
               TabIndex        =   18
               Top             =   540
               Width           =   2835
            End
            Begin VB.OptionButton optClearedDep 
               Caption         =   "View Cleared Deposits"
               Height          =   210
               Left            =   8070
               TabIndex        =   17
               Top             =   240
               Width           =   2685
            End
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
               Left            =   90
               TabIndex        =   14
               Top             =   150
               Width           =   4365
               Begin VB.TextBox txtSearchCheck 
                  Height          =   315
                  Left            =   1890
                  TabIndex        =   15
                  Text            =   "Text2"
                  Top             =   210
                  Width           =   2355
               End
               Begin VB.Label Label13 
                  Caption         =   "Check No."
                  Height          =   315
                  Left            =   840
                  TabIndex        =   16
                  Top             =   240
                  Width           =   1065
               End
            End
            Begin MSComctlLib.ListView lstQuiry 
               Height          =   4575
               Left            =   90
               TabIndex        =   23
               Top             =   1680
               Width           =   11145
               _ExtentX        =   19659
               _ExtentY        =   8070
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
               MouseIcon       =   "FrmBankRecon.frx":3E2C
               NumItems        =   10
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Tran.Date"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Type"
                  Object.Width           =   1587
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "CV #"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Vendor/Customer"
                  Object.Width           =   4410
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Check Date"
                  Object.Width           =   2293
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Check #"
                  Object.Width           =   1940
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "Debit"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Credit"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   8
                  Text            =   "Status"
                  Object.Width           =   1764
               EndProperty
               BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   9
                  Text            =   "Remarks"
                  Object.Width           =   4410
               EndProperty
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
               TabIndex        =   26
               Top             =   1140
               Width           =   1845
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
               TabIndex        =   25
               Top             =   1410
               Width           =   2745
            End
            Begin VB.Label Label10 
               Caption         =   "N - entered/out-Standing"
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
               TabIndex        =   24
               Top             =   870
               Width           =   2625
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   6450
         Left            =   -69970
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   11535
         _Version        =   655364
         _ExtentX        =   20346
         _ExtentY        =   11377
         _StockProps     =   0
         Begin VB.Frame Frame4 
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1725
            Left            =   5370
            TabIndex        =   78
            Top             =   240
            Width           =   4005
            Begin VB.CommandButton Command3 
               Caption         =   "Search"
               Height          =   435
               Left            =   2490
               TabIndex        =   83
               Top             =   1170
               Width           =   1425
            End
            Begin VB.OptionButton optCheckNOR 
               Caption         =   "By OR No"
               Height          =   255
               Index           =   1
               Left            =   1950
               TabIndex        =   82
               Top             =   330
               Width           =   1425
            End
            Begin VB.OptionButton optCheckNOR 
               Caption         =   "By Check No"
               Height          =   255
               Index           =   0
               Left            =   210
               TabIndex        =   81
               Top             =   330
               Value           =   -1  'True
               Width           =   2025
            End
            Begin VB.TextBox txtLed 
               Height          =   315
               Left            =   750
               TabIndex        =   79
               Top             =   720
               Width           =   3135
            End
            Begin VB.Label Label27 
               Caption         =   "Find"
               Height          =   315
               Left            =   180
               TabIndex        =   80
               Top             =   750
               Width           =   1065
            End
         End
         Begin Crystal.CrystalReport rptBankRecon 
            Left            =   10890
            Top             =   1170
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.TextBox txt1 
            Alignment       =   1  'Right Justify
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
            Left            =   2910
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   450
            Width           =   2300
         End
         Begin FlexCell.Grid grdRecom 
            Height          =   3645
            Left            =   60
            TabIndex        =   29
            Top             =   2520
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   6429
            BackColorBkg    =   -2147483645
            Cols            =   5
            DefaultFontName =   "Arial"
            DefaultFontSize =   9.75
            Rows            =   30
         End
         Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
            Height          =   3765
            Left            =   1170
            TabIndex        =   30
            Top             =   6930
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   6641
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00808080&
            Enabled         =   0   'False
            Height          =   2115
            Left            =   60
            ScaleHeight     =   2055
            ScaleWidth      =   5175
            TabIndex        =   34
            Top             =   360
            Width           =   5235
            Begin VB.TextBox txt4 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   2820
               TabIndex        =   40
               Text            =   "0.00"
               Top             =   1620
               Width           =   2300
            End
            Begin VB.TextBox txtEndBal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
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
               TabIndex        =   38
               Text            =   "0.00"
               Top             =   1200
               Width           =   2300
            End
            Begin VB.TextBox txtClrdWtdwal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
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
               TabIndex        =   37
               Text            =   "0.00"
               Top             =   810
               Width           =   2300
            End
            Begin VB.TextBox txtClrdDep 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
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
               TabIndex        =   36
               Text            =   "0.00"
               Top             =   420
               Width           =   2300
            End
            Begin VB.TextBox txtStartBal 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000F&
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
               Left            =   2610
               Locked          =   -1  'True
               TabIndex        =   35
               Text            =   "0.00"
               Top             =   -420
               Width           =   2300
            End
            Begin VB.TextBox txt3 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   7470
               TabIndex        =   39
               Text            =   "0.00"
               Top             =   30
               Width           =   2300
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Statement Ending Balance"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   465
               Left            =   90
               TabIndex        =   47
               Top             =   60
               Width           =   3255
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unreconciled Difference"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   30
               TabIndex        =   46
               Top             =   1680
               Width           =   2505
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1680
               TabIndex        =   45
               Top             =   1260
               Width           =   1065
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Outstanding Checks"
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   150
               TabIndex        =   44
               Top             =   870
               Width           =   2625
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Deposits"
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   330
               TabIndex        =   43
               Top             =   480
               Width           =   2445
            End
         End
         Begin VB.TextBox txt2 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   1590
            Width           =   2300
         End
         Begin VB.Label Label7 
            Caption         =   "Ending Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2115
            TabIndex        =   33
            Top             =   1620
            Width           =   2085
         End
         Begin VB.Label Label8 
            Caption         =   "General Ledger"
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
            Index           =   0
            Left            =   3390
            TabIndex        =   32
            Top             =   30
            Width           =   1965
         End
         Begin VB.Label Label8 
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
            Height          =   405
            Index           =   1
            Left            =   975
            TabIndex        =   31
            Top             =   360
            Width           =   1065
         End
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   795
      Left            =   9300
      Picture         =   "FrmBankRecon.frx":3F8E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save Entry"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Re&fresh"
      Enabled         =   0   'False
      Height          =   765
      Left            =   2430
      Picture         =   "FrmBankRecon.frx":42DE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Refresh"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "R&eports"
      Enabled         =   0   'False
      Height          =   765
      Left            =   1260
      Picture         =   "FrmBankRecon.frx":4609
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Report"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "&Reload"
      Height          =   765
      Left            =   90
      Picture         =   "FrmBankRecon.frx":4922
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Reload"
      Top             =   8520
      Width           =   1185
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLS SELECT BANK ACCOUNT NUMBER!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   240
      TabIndex        =   119
      Top             =   810
      Width           =   11295
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   255
      Left            =   90
      TabIndex        =   61
      Top             =   90
      Width           =   2655
      _Version        =   655364
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   14
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      GradientColorLight=   16576
      GradientColorDark=   12632064
      ForeColor       =   8421504
   End
End
Attribute VB_Name = "FrmBankRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctl                                           As Control
Dim START_DEBIT                                   As Double
Dim START_CREDIT                                  As Double
Dim Reconstatus                                   As String
Dim Search_mode                                   As Boolean

Function SetAccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select * from ALL_Banks where BankAcctNo = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        SetAccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Sub InitGrid()
    With grdRecom
        .Cols = 11: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = " Date "
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "Reference"
        .Cell(0, 4).Text = "Deposits"
        .Cell(0, 5).Text = "Withdrawals"
        .Cell(0, 6).Text = "Clear"
        .Cell(0, 7).Text = "Staled"
        .Cell(0, 8).Text = "CV#"
        .Cell(0, 9).Text = "Type"
        .Cell(0, 10).Text = "Date Cleared"

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

        .Column(0).Width = 18
        .Column(1).Width = 80: .Column(1).Locked = True
        .Column(2).Width = 295: .Column(2).Locked = True
        .Column(3).Width = 90: .Column(3).Locked = True
        .Column(4).Width = 80: .Column(4).Locked = True
        .Column(5).Width = 80: .Column(5).Locked = True
        .Column(6).Width = 45
        .Column(7).Width = 45
        .Column(8).Width = 0: .Column(8).Locked = True
        .Column(9).Width = 0: .Column(9).Locked = True
        .Column(10).Width = 0: .Column(9).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 10, .Rows - 1, 10).ForeColor = RGB(0, 0, 128)
    End With
End Sub

Sub ComputeEndBalance()
    txtClrdDep = ToDoubleNumber(txtClrdDep)
    txtClrdWtdwal = ToDoubleNumber(txtClrdWtdwal)
    'txtEndBal = ToDoubleNumber(((NumericVal(LTrim(txtStartBal)) + NumericVal(LTrim(txtClrdDep))) - NumericVal(LTrim(txtClrdWtdwal))))
    'txt3 = ToDoubleNumber(NumericVal(txt1.Text) - NumericVal(txtStartBal.Text))
    txt4 = ToDoubleNumber(NumericVal(txt1.Text) - NumericVal(LTrim(txtClrdDep)) + NumericVal(LTrim(txtClrdWtdwal)) - NumericVal(txtEndBal.Text))
    'txt2 = ToDoubleNumber(((NumericVal(txt1) + NumericVal(txtClrdDep)) - NumericVal(txtClrdWtdwal)))
End Sub

Private Sub cboBank_Change()
    Screen.MousePointer = 11
    Dim rsALL_BANKS                               As ADODB.Recordset
    Set rsALL_BANKS = New ADODB.Recordset
    Set rsALL_BANKS = gconDMIS.Execute("Select * from ALL_BANKS Where BANKACCTNO = '" & cboBank.Text & "'")
    If Not rsALL_BANKS.EOF And Not rsALL_BANKS.BOF Then
        'Label9.Caption = " *** " & Null2String(rsALL_BANKS!bankname) & " *** "
        Label14.Caption = " *** " & Null2String(rsALL_BANKS!BankName) & " *** "
        txt1.Text = ToDoubleNumber(N2Str2Zero(rsALL_BANKS!STARTING_BALANCE))
        txt2.Text = ToDoubleNumber(N2Str2Zero(rsALL_BANKS!ENDING_BALANCE))
        If Null2Date(rsALL_BANKS!LASTDATE_RECON) = "" Then
            dtTo = LOGDATE
        Else
            dtTo = lastDay(Null2Date(rsALL_BANKS!LASTDATE_RECON))
        End If
        cmdReload.Enabled = True
        cmdReload_Click
    Else
        Screen.MousePointer = 0
        'Label9.Caption = "PLS SELECT BANK ACCOUNT NUMBER!"
        Label14.Caption = "PLS SELECT BANK ACCOUNT NUMBER!"
        cmdReload.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub cbobank_Click()
    Screen.MousePointer = 11
    Dim rsALL_BANKS                               As ADODB.Recordset
    Set rsALL_BANKS = New ADODB.Recordset
    Set rsALL_BANKS = gconDMIS.Execute("Select * from ALL_BANKS Where BANKACCTNO = '" & cboBank.Text & "'")
    If Not rsALL_BANKS.EOF And Not rsALL_BANKS.BOF Then
        'Label9.Caption = " *** " & Null2String(rsALL_BANKS!bankname) & " *** "
        Label14.Caption = " *** " & Null2String(rsALL_BANKS!BankName) & " *** "
        txt1.Text = ToDoubleNumber(N2Str2Zero(rsALL_BANKS!STARTING_BALANCE))
        txt2.Text = ToDoubleNumber(N2Str2Zero(rsALL_BANKS!ENDING_BALANCE))
        If Null2Date(rsALL_BANKS!LASTDATE_RECON) = "" Then
            dtTo = LOGDATE
        Else
            dtTo = lastDay(Null2Date(rsALL_BANKS!LASTDATE_RECON))
        End If
        cmdReload.Enabled = True
        cmdReload_Click
    Else
        Screen.MousePointer = 0
        'Label9.Caption = "PLS SELECT BANK ACCOUNT NUMBER!"
        Label14.Caption = "PLS SELECT BANK ACCOUNT NUMBER!"
        cmdReload.Enabled = False
        Exit Sub
    End If

End Sub

Private Sub Check1_Click()
    cmdReload.Value = True
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

Private Sub cmdCloseOption_Click()
    picoption.Visible = False
End Sub

Private Sub cmdEdit_Click()
    Unload Me
End Sub

Private Sub cmdExtract_Click()

End Sub

Private Sub cmdOK_Click()
    Picture7.Visible = False

End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtpTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and ({RECON.JDATE} >= Date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") and {RECON.JDATE} <= Date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", DMIS_REPORT_Connection, 1
    Else
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECON.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")", DMIS_REPORT_Connection, 1
    End If
    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = 11
    cboBank.ListIndex = -1
    For Each ctl In ControlS
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        End If
    Next ctl
    grdRecom.Rows = 2
    grdRecom.Cell(1, 1).Text = ""
    grdRecom.Cell(1, 2).Text = ""
    grdRecom.Cell(1, 3).Text = ""
    grdRecom.Cell(1, 4).Text = ""
    grdRecom.Cell(1, 5).Text = ""
    grdRecom.Cell(1, 6).Text = ""
    grdRecom.Cell(1, 7).Text = ""
    grdRecom.Cell(1, 8).Text = ""
    grdRecom.Cell(1, 9).Text = ""
    lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
    Screen.MousePointer = 0
    cmdRefresh.Enabled = False
    cmdReport.Enabled = False
End Sub

'Upating Code       : AXP-0713200714:34
Private Sub cmdReload_Click()
    Screen.MousePointer = 11
    Dim rsLoad4Recon                              As ADODB.Recordset
    lstQuiry.Sorted = False: lstQuiry.ListItems.Clear: InitGrid
    START_DEBIT = 0
    START_CREDIT = 0
    txtStartBal = "0.00"
    txtClrdDep = "0.00"
    txtClrdWtdwal = "0.00"
    txtEndBal = "0.00"
    Dim xx                                        As Integer
    grdRecom.Rows = 1: xx = 0
    grdRecom.AutoRedraw = False

    'Update By BTT : 09292008
    If cboBank.Text = "" Then
        MsgBox "Please select bank account no!", vbInformation, "Information"
        Exit Sub
    End If

    If optViewAll.Value = True Then                        ' all ledger
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If
    If optStaled.Value = True Then                         ' staled check
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='S'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If
    If optOutstanding.Value = True Then                    ' outstanding check
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='N' and (jtype='CDJ' or jtype='BOB')", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If
    If optClearedWithdrawal.Value = True Then              ' cleared withdrawals
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='C' and credit > 0 and jtype='CDJ'", gconDMIS, adOpenForwardOnly, adLockReadOnly    '
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If
    If optUnclearedDep.Value = True Then                   ' uncleared deposit
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='N' and jtype='GJ' and Debit>0", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If
    If optClearedDep.Value = True Then                     ' cleared deposit
        Set rsLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        rsLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='C' and jtype='CRJ' and Debit >0", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsLoad4Recon.EOF And Not rsLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, rsLoad4Recon
        End If
    End If

    Dim varRefirence                              As String
    'Set rsLoad4Recon = New ADODB.Recordset
    'rsLoad4Recon.Open "select SUM(DEBIT) - SUM(CREDIT) AS BEG_BALANCE from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTO) & "'and BankAcctno = '" & Trim(cboBank) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'If Not rsLoad4Recon.EOF And Not rsLoad4Recon.EOF Then
    '    txtStartBal = ToDoubleNumber(N2Str2Zero(rsLoad4Recon!BEG_BALANCE))
    'End If
    ComputeJournalEntry
    Dim vReconStatus                              As Byte
    Set rsLoad4Recon = New ADODB.Recordset
    grdRecom.AutoRedraw = False
    BackRecon
    Label21.Caption = "Computing Data.."
    ' Update by BTT : 10182008
    If Search_mode = False Then
        If Check1.Value = 1 Then
            rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' AND ReconStatus = 'N' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        Else
            rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        ' Search mode = totoo
    Else

        If optCheckNOR(0).Value = True Then                ' By Check no
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' AND ReconStatus = 'N' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and checkno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            'By OR
        Else
            If Check1.Value = 1 Then
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' AND ReconStatus = 'N' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            Else
                rsLoad4Recon.Open "select * from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and invoiceno like '%" & txtLed.Text & "%' Order by JDate", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        txtLed.Text = ""
        Search_mode = False
    End If

    If Not rsLoad4Recon.EOF And Not rsLoad4Recon.EOF Then
        grdRecom.Rows = 1
        If SSTab1.SelectedItem = 1 Then
            Picture4.Visible = True
        End If
        PROGBAR.Value = 0
        PROGBAR.Max = rsLoad4Recon.RecordCount
        Do Until rsLoad4Recon.EOF
            If Trim(Null2String(rsLoad4Recon![Reconstatus])) = "C" Then
                vReconStatus = 1
            Else
                START_DEBIT = START_DEBIT + N2Str2Zero(rsLoad4Recon![DEBIT])
                START_CREDIT = START_CREDIT + N2Str2Zero(rsLoad4Recon![CREDIT])
                txtClrdDep = ToDoubleNumber(START_DEBIT)
                txtClrdWtdwal = ToDoubleNumber(START_CREDIT)
                vReconStatus = 0
            End If
            If Null2String(rsLoad4Recon!jtype) = "CDJ" Then
                varRefirence = "CHK#" & Null2String(rsLoad4Recon!CheckNo)
            ElseIf Null2String(rsLoad4Recon!jtype) = "DRJ" Then
                varRefirence = "OR#" & Null2String(rsLoad4Recon!INVOICENO)
            Else
                varRefirence = Null2String(rsLoad4Recon![ReferenceNo])
            End If
            grdRecom.AddItem rsLoad4Recon![JDate] & vbTab & _
                             rsLoad4Recon![remarks] & vbTab & _
                             varRefirence & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![DEBIT]) & vbTab & _
                             ToDoubleNumber(rsLoad4Recon![CREDIT]) & vbTab & _
                             vReconStatus & vbTab & _
                             "" & vbTab & _
                             rsLoad4Recon![VOUCHERNO] & vbTab & _
                             rsLoad4Recon![jtype] & vbTab & _
                             False
            rsLoad4Recon.MoveNext
            DoEvents
            PROGBAR.Value = PROGBAR.Value + 1
            Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
            Label22 = varRefirence
        Loop
    End If
    grdRecom.AutoRedraw = True
    grdRecom.Refresh
    txtClrdDep = ToDoubleNumber(txtClrdDep)
    txtClrdWtdwal = ToDoubleNumber(txtClrdWtdwal)
    'txtEndBal = ToDoubleNumber(((NumericVal(LTrim(txtStartBal)) + NumericVal(LTrim(txtClrdDep))) - NumericVal(LTrim(txtClrdWtdwal))))
    txtEndBal = ToDoubleNumber(txtStartBal)
    ComputeEndBalance
    LogAudit "R", "BANK RECONCILIATION", cboBank
    Screen.MousePointer = 0
    cmdRefresh.Enabled = True
    cmdReport.Enabled = True
    Picture4.Visible = False
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdReloadrecongrd_Click()
    fillGridRecon
End Sub

'Upating Code       : AXP-0713200714:34
Private Sub cmdReport_Click()
'    On Error GoTo Errorcode:
'    Dim filter                                                        As String
'    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
'
'    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
'    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
'    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTO & "'"
'    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)
'
'    Screen.MousePointer = 11    '
'    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
'        'PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "'" & filter, DMIS_REPORT_Connection, 1
'        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ")", DMIS_REPORT_Connection, 1
'    Else
'        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECON.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ")", DMIS_REPORT_Connection, 1
'    End If
'    Screen.MousePointer = 0
'    LogAudit "V", "BANK RECONCILIATION", cboBank
'    Exit Sub
'Errorcode:
'    ShowVBError
    picoption.Visible = True
End Sub

'Upating Code       : AXP-0713200714:34
Private Sub cmdSave_Click()
'On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", "BANK RECONCILIATION") = False Then Exit Sub

    Dim xVOUCHERNO, xJType, xCheckNo              As String
    Dim xJdate                                    As Date
    Dim X                                         As Long
    Screen.MousePointer = 11

    For X = 1 To grdRecom.Rows - 1
        xVOUCHERNO = grdRecom.Cell(X, 8).Text
        xJType = grdRecom.Cell(X, 9).Text
        xCheckNo = grdRecom.Cell(X, 3).Text
        xJdate = grdRecom.Cell(X, 1).Text
        If NumericVal(grdRecom.Cell(X, 6).Text) > 0 Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'C' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
            ' Update BY BTT
            gconDMIS.Execute "Insert into AMIS_reconstatus(Voucherno,Date_cleared,jtype,Recon_Status,date_before_recon) Values('" & xVOUCHERNO & _
                             "'," & N2Str2Null(dtTo) & ",'" & xJType & "','C'," & N2Str2Null(xJdate) & ")"

        Else
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'N' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"

            gconDMIS.Execute "delete AMIS_reconstatus " & _
                             " where VoucherNo = '" & xVOUCHERNO & "' and JType = '" & xJType & "'"


        End If
        If NumericVal(grdRecom.Cell(X, 7).Text) > 0 Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReconStatus = 'S' " & "" & _
                             " where VoucherNo = '" & xVOUCHERNO & "' AND JType = '" & xJType & "'"
        End If
    Next X
    gconDMIS.Execute ("Update ALL_BANKS SET " & _
                      " STARTING_BALANCE = " & NumericVal(txt1.Text) & "," & _
                      " ENDING_BALANCE = " & NumericVal(txt2.Text) & "," & _
                      " STARTING_DIFFERENCE = " & NumericVal(txt3.Text) & "," & _
                      " ENDING_DIFFERENCE = " & NumericVal(txt4.Text) & "," & _
                      " BOOK_BALANCE = " & NumericVal(txtEndBal.Text) & "," & _
                      " BANK_BALANCE = " & NumericVal(txt2.Text) & "," & _
                      " LASTDATE_RECON = " & N2Str2Null(dtTo) & _
                      " WHERE BANKACCTNO = '" & Trim(cboBank.Text) & "'")
    Screen.MousePointer = 0
    MsgBox "Data Successfully updated", vbInformation, "Saved..."
    'cmdRefresh.Value = True
    LogAudit "A", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdview_Click()
    Picture7.Visible = True
    txtBank = txt1.Text
    txtBook.Text = ToDoubleNumber(NumericVal(txtBank) - NumericVal(txtCredit))
    txttotalBook = ToDoubleNumber(txtBalanceJ)
    Label28 = dtTo.Value
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    frmAMISbanksOpening.Show
End Sub

Private Sub Command2_Click()
    ComputeJournalEntry
End Sub

Private Sub Command3_Click()
    If txtLed.Text = "" Then
        MsgBox "Please input a criteria!", vbInformation, "Information"
        Exit Sub
    End If
    Search_mode = True
    cmdReload_Click
End Sub

Private Sub dtFrom_Click()
'cmdReload_Click
    grdRecom.Rows = 1:
    grdRecom.AutoRedraw = False
End Sub

Private Sub dtTO_Change()
    grdRecom.Rows = 1:
    grdRecom.AutoRedraw = False
    cmdRefresh.Enabled = False
    cmdReport.Enabled = False
End Sub

Private Sub dtTO_Click()
    grdRecom.Rows = 1:
    grdRecom.AutoRedraw = False
    cmdRefresh.Enabled = False
    cmdReport.Enabled = False
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitGrid
    InitGridRecon
    Dim rsLoadBank                                As ADODB.Recordset
    Set rsLoadBank = New ADODB.Recordset
    rsLoadBank.Open "select bank from ALL_BANKNAME", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsLoadBank.EOF And Not rsLoadBank.EOF Then
        cboBank.Clear
        Do Until rsLoadBank.EOF
            cboBank.AddItem Null2String(rsLoadBank![Bank])
            rsLoadBank.MoveNext
        Loop
    End If
    cmdRefresh_Click
    optViewAll.Value = True
    picoption.Visible = False
    PicDateRange.Visible = False
    Picture4.Visible = False
    Search_mode = False
End Sub

Private Sub grdRecom_Click()
    Dim Ans                                       As String
    Dim xdebit, xcredit                           As Double
    Dim xReference                                As String
    xdebit = NumericVal(txtClrdDep)
    xcredit = NumericVal(txtClrdWtdwal)
    txt1 = ToDoubleNumber(txt1)
    'For X = 1 To grdRecom.Rows - 1
    If grdRecom.ActiveCell.Col = 6 Then
        If NumericVal(grdRecom.Cell(grdRecom.ActiveCell.Row, 6).Text) >= 1 Then
            xdebit = xdebit - NumericVal(grdRecom.Cell(grdRecom.ActiveCell.Row, 4).Text)
            xcredit = xcredit - NumericVal(grdRecom.Cell(grdRecom.ActiveCell.Row, 5).Text)
        Else
            xdebit = xdebit + NumericVal(grdRecom.Cell(grdRecom.ActiveCell.Row, 4).Text)
            xcredit = xcredit + NumericVal(grdRecom.Cell(grdRecom.ActiveCell.Row, 5).Text)
        End If
        'Next X
        START_DEBIT = xdebit
        START_CREDIT = xcredit
        txtClrdDep = ToDoubleNumber(Round(START_DEBIT, 2))
        txtClrdWtdwal = ToDoubleNumber(Round(START_CREDIT, 2))
        ComputeEndBalance
        txt2 = ToDoubleNumber(Round(NumericVal(txt1) + xdebit - xcredit, 2))
        xReference = (grdRecom.Cell(grdRecom.ActiveCell.Row, 3).Text)
        VerifyBackRecon (xReference)



        'txt4 = ToDoubleNumber(Round(NumericVal(txt1) - NumericVal(txtEndBal), 2))
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.SelectedItem = 1 Then
        cmdSave.Enabled = False
    Else
        cmdSave.Enabled = True
    End If
End Sub

Private Sub grdRecom_DblClick()
    Dim VARVOUCHERNO                              As String
    If Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "APJ" Then
        JOURNALTYPE = "APJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "CDJ" Then
        JOURNALTYPE = "CDJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 2) = "SJ" Then
        JOURNALTYPE = "SJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "CRJ" Then
        JOURNALTYPE = "CRJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 2) = "GJ" Then
        JOURNALTYPE = "GJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "ADJ" Then
        JOURNALTYPE = "ADJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "PDJ" Then
        JOURNALTYPE = "PDJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "CLO" Then
        JOURNALTYPE = "CLO"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "DRJ" Then
        JOURNALTYPE = "DRJ"
    ElseIf Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 9).Text, 3) = "BOB" Then
        JOURNALTYPE = "BOB"
    Else
        JOURNALTYPE = "OPB"                                '
    End If
    'JOURNALTYPE = Left(grdRecom.Cell(grdRecom.ActiveCell.Row, 3).Text, 3)
    VARVOUCHERNO = Right(grdRecom.Cell(grdRecom.ActiveCell.Row, 8).Text, 6)
    Screen.MousePointer = 11
    On Error Resume Next
    If JOURNALTYPE = "DRJ" Then
        Unload frmAMISJournalEntry_DRJ
        frmAMISJournalEntry_DRJ.Show
        Call frmAMISJournalEntry_DRJ.StoreSearch(VARVOUCHERNO)

    ElseIf JOURNALTYPE = "BOB" Then
        Unload frmAMISbanksOpening
        frmAMISbanksOpening.Show
        frmAMISbanksOpening.StoreSearch (VARVOUCHERNO)

    Else
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub optAll_Click()
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11

    Ans = MsgBox("Print By Date Range?.", vbQuestion + vbYesNo, "Information")

    If Ans = vbYes Then
        dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
        dtpTo = LOGDATE
        PicDateRange.Visible = True
        picoption.Visible = False
    Else
        If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
            'PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "'" & filter, DMIS_REPORT_Connection, 1
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECON.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")", DMIS_REPORT_Connection, 1
        End If
    End If
    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub OptCD_Click()
'and jtype='GJ' and Debit>0
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='C' and {RECON.Jtype}='DRJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub OptCW_Click()
'Clear Widrawals
'reconstatus='C' and credit > 0 and jtype='CDJ'
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='C' and {RECON.Jtype}='CDJ' and {recon.credit}  > 0 ", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
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
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='N' and {RECON.Jtype}='GJ' and {recon.debit}  > 0 ", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
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
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'as of : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)

    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BANKRECONGROUP.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='S'", DMIS_REPORT_Connection, 1
    End If

    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub otpOut_Click()
    On Error GoTo ErrorCode:
    Dim filter                                    As String
    If Function_Access(LOGID, "Acess_Print", "BANK RECONCILIATION") = False Then Exit Sub
    Dim Ans                                       As String
    rptBankRecon.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptBankRecon.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    rptBankRecon.Formulas(2) = "FORTHEMONTH = 'AS OF : " & dtTo & "'"
    rptBankRecon.Formulas(3) = "BalanceperLedger = " & NumericVal(txtEndBal)
    Screen.MousePointer = 11
    If MsgBox("Print Detailed?", vbQuestion + vbYesNo, "NO will default sorted Printing") = vbYes Then
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconGroup.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='N'", DMIS_REPORT_Connection, 1
    Else
        rptBankRecon.Formulas(7) = "Bankstatement= " & NumericVal(txt1)
        PrintSQLReport rptBankRecon, AMIS_REPORT_PATH & "JOURNALS\BankReconDetail.RPT", "{RECON.BANKACCTNO}='" & cboBank & "' and {RECON.JDATE} <= Date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {RECON.ReconStatus}='N' and ({recon.jtype} ='CDJ' or {recon.jtype} ='BOB')", DMIS_REPORT_Connection, 1
    End If
    Screen.MousePointer = 0
    LogAudit "V", "BANK RECONCILIATION", cboBank
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub txt1_GotFocus()
    txt1.Text = NumericVal(txt1.Text)
End Sub

Private Sub txt1_LostFocus()
    ComputeEndBalance
End Sub

Private Sub txt2_Change()
    ComputeEndBalance
End Sub

Private Sub txt2_GotFocus()
    txt2.Text = NumericVal(txt2.Text)
End Sub
Sub SearchCheckNo()
'Update By BTT : 09292008
    Dim NardLoad4Recon                            As New ADODB.Recordset
    Dim keyword                                   As String

    keyword = Trim(txtSearchCheck.Text)

    If optViewAll.Value = True Then                        ' all ledger
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and Checkno like '" & keyword & "%'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
    If optStaled.Value = True Then                         ' staled check
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='S'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
    If optOutstanding.Value = True Then                    ' outstanding check
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='N' and jtype='CDJ'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
    If optClearedWithdrawal.Value = True Then              ' cleared withdrawals
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='C' and credit > 0 and jtype='CDJ'", gconDMIS, adOpenForwardOnly, adLockReadOnly    '
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
    If optUnclearedDep.Value = True Then                   ' uncleared deposit
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='N' and jtype='GJ' and Debit>0", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
    If optClearedDep.Value = True Then                     ' cleared deposit
        Set NardLoad4Recon = New ADODB.Recordset
        lstQuiry.Sorted = False: lstQuiry.ListItems.Clear
        NardLoad4Recon.Open "select jdate,jtype,VoucherNo,nameofVendor,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "' and reconstatus='C' and jtype='CRJ' and Debit>0", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not NardLoad4Recon.EOF And Not NardLoad4Recon.BOF Then
            Listview_Loadval Me.lstQuiry.ListItems, NardLoad4Recon
        End If
    End If
End Sub

Private Sub txtSearchCheck_Change()
    SearchCheckNo
End Sub
Sub BackRecon()
'Update By BTT 1022008
    Dim LookUpRecon                               As ADODB.Recordset
    Dim SQL                                       As String
    Label21.Caption = "Checking data.."
    If SSTab1.SelectedItem = 1 Then
        Picture4.Visible = True
    End If
    Dim RsCutDate                                 As New ADODB.Recordset
    SQL = "SELECT * from AMIS_reconstatus where date_before_recon < = '" & CDate(dtTo) & "'"

    Set RsCutDate = New ADODB.Recordset
    Set RsCutDate = gconDMIS.Execute(SQL)
    PROGBAR.Value = 0
    Do While Not RsCutDate.EOF

        Set LookUpRecon = New ADODB.Recordset
        LookUpRecon.Open "select * from AMIS_journal_hd where jdate <= '" & CDate(dtTo) & "' and voucherno='" & (RsCutDate!VOUCHERNO) & "' and jtype='" & (RsCutDate!jtype) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly



        PROGBAR.Max = RsCutDate.RecordCount

        If Not LookUpRecon.BOF And Not LookUpRecon.EOF Then
            If dtTo < RsCutDate!date_cleared Then
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
        PROGBAR.Value = PROGBAR.Value + 1
        Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
        Label22 = Null2String((LookUpRecon!VOUCHERNO))
    Loop

    Set RsCutDate = Nothing
    Set LookUpRecon = Nothing
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
        Ans = MsgBox("This has been allready cleared.Do you want to ovewrite the transaction?", vbInformation + vbYesNo)
        If Ans = vbYes Then
            'Update the Data
            finalAns = MsgBox("Are you sure do you want to overwite this transaction", vbInformation + vbYesNo)
            If finalAns = vbYes Then
                gconDMIS.Execute "delete AMIS_reconstatus " & _
                                 " where VoucherNo = '" & TheReferenceNo & "'"
                LogAudit "X", "BANK RECONCILIATION", cboBank
            End If
        Else
            ' Do nothing
        End If
    Else
        'Do nothing
    End If
    Set RS = Nothing
End Function

Sub ComputeJournalEntry()
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT sum(debit) as XTotal_debit, sum(Credit) as XTotal_credit from AMIS_vw_recondata where jdate <= '" & CDate(dtTo) & "'and BankAcctno = '" & Trim(cboBank) & "' and reconstatus <>'C'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtDebit = ToDoubleNumber(NumericVal(RS!xTOTAL_DEBIT))
        txtCredit = ToDoubleNumber(NumericVal(RS!xTOTAL_CREDIT))
        txtBalanceJ = ToDoubleNumber(NumericVal(txtDebit - txtCredit))
        txtStartBal = txtBalanceJ
    End If
    Set RS = Nothing
End Sub
Sub InitGridRecon()
    With grdBankrecon
        .ColWidth(0) = 1200
        .ColWidth(1) = 500
        .ColWidth(2) = 800
        .ColWidth(3) = 4500
        .ColWidth(4) = 1200
        .ColWidth(5) = 900
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .Row = 0
        .Col = 0: .Text = "Trandate"
        .Col = 1: .Text = "Type"
        .Col = 2: .Text = "CV#"
        .Col = 3: .Text = "Customer/Vendor"
        .Col = 4: .Text = "Checkdate"
        .Col = 5: .Text = "Checkno"
        .Col = 6: .Text = "Debit"
        .Col = 7: .Text = "Credit"
    End With
End Sub
Sub fillGridRecon()
'Update By BTT :
    Screen.MousePointer = 11
    Dim RS                                        As New ADODB.Recordset
    Dim Reference                                 As String
    cleargrid grdBankrecon: InitGrid
    Dim TOTAL_CREDIT                              As Double
    Dim TOTAL_DEBIT                               As Double
    Dim cnt                                       As Integer

    TOTAL_CREDIT = 0
    TOTAL_DEBIT = 0
    With RS
        .Open "select jdate,jtype,VoucherNo,nameofVendor,acctname,CheckDate,Checkno,Debit,Credit,ReconStatus,Remarks from AMIS_vw_RECONDATA where jdate <= '" & CDate(dtTo) & "' and BankAcctno = '" & Trim(cboBank) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not .EOF And Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                cnt = cnt + 1
                If Null2String(RS!jtype) = "DRJ" Then
                    Reference = Null2String(RS!AcctName)
                Else
                    Reference = Null2String(RS!nameofvendor)
                End If

                grdBankrecon.AddItem (RS!JDate) & Chr(9) & (RS!jtype) & Chr(9) & _
                                     (RS!VOUCHERNO) & Chr(9) & Reference & Chr(9) & _
                                     (RS!CheckDate) & Chr(9) & (RS!CheckNo) & Chr(9) & _
                                     (RS!DEBIT) & Chr(9) & (RS!CREDIT)
                TOTAL_CREDIT = TOTAL_CREDIT + NumericVal(RS!CREDIT)
                TOTAL_DEBIT = TOTAL_DEBIT + NumericVal(RS!DEBIT)
                .MoveNext
            Loop
        End If
        If cnt > 0 Then grdBankrecon.RemoveItem 1
        txtcreditL.Text = ToDoubleNumber(TOTAL_CREDIT)
        txtdebitL.Text = ToDoubleNumber(TOTAL_DEBIT)
    End With
    Screen.MousePointer = 0
    Set RS = Nothing
End Sub
