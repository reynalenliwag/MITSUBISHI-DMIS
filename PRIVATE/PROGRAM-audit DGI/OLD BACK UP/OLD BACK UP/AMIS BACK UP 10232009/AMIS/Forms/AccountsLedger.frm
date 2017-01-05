VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAMISLEDGERAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts General Ledger"
   ClientHeight    =   7980
   ClientLeft      =   720
   ClientTop       =   435
   ClientWidth     =   14145
   ForeColor       =   &H00FFFFFF&
   Icon            =   "AccountsLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   14145
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   5040
      ScaleHeight     =   945
      ScaleWidth      =   5625
      TabIndex        =   38
      Top             =   3210
      Width           =   5625
      Begin VB.PictureBox Picture5 
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   5805
         TabIndex        =   39
         Top             =   0
         Width           =   5805
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Loading Data.."
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
            TabIndex        =   40
            Top             =   0
            Width           =   5415
         End
      End
      Begin MSComctlLib.ProgressBar PROGBAR 
         Height          =   405
         Left            =   30
         TabIndex        =   41
         Top             =   480
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Label11"
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
         Height          =   285
         Left            =   600
         TabIndex        =   43
         Top             =   240
         Width           =   5085
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
         TabIndex        =   42
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2640
      TabIndex        =   6
      Top             =   0
      Width           =   11145
      Begin VB.TextBox txtBeginningBalance 
         Alignment       =   1  'Right Justify
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
         Left            =   9390
         MaxLength       =   10
         TabIndex        =   20
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txtAcctType 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   18
         Top             =   570
         Width           =   1665
      End
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
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   9
         Top             =   180
         Width           =   1665
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
         Left            =   2790
         MaxLength       =   3
         TabIndex        =   14
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
         TabIndex        =   12
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
         Left            =   1620
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "000"
         Top             =   180
         Width           =   435
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   3270
         TabIndex        =   16
         Top             =   180
         Width           =   7785
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Beg. Balance"
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
         Left            =   7830
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         TabIndex        =   17
         Top             =   600
         Width           =   1455
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
         TabIndex        =   15
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
         TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   7
         Top             =   180
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         TabIndex        =   8
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5985
      Left            =   2655
      TabIndex        =   21
      Top             =   945
      Width           =   11475
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show Ledger"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8955
         MouseIcon       =   "AccountsLedger.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   24
         ToolTipText     =   "Show Ledger"
         Top             =   180
         Width           =   1470
      End
      Begin MSFlexGridLib.MSFlexGrid grdAccountsLedger 
         Height          =   4785
         Left            =   90
         TabIndex        =   28
         Top             =   600
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   8440
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         ForeColor       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
         BackColorSel    =   16711680
         ForeColorSel    =   16777215
         BackColorBkg    =   14737632
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
         MouseIcon       =   "AccountsLedger.frx":045C
      End
      Begin Crystal.CrystalReport rptGeneralLedger 
         Left            =   90
         Top             =   4320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "G E N E R A L  L E D G E R"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtTotalBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   9450
         MaxLength       =   20
         TabIndex        =   31
         Top             =   5490
         Width           =   1725
      End
      Begin VB.TextBox txtTotalDebit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5970
         MaxLength       =   20
         TabIndex        =   29
         Top             =   5490
         Width           =   1695
      End
      Begin VB.TextBox txtTotalCredit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7710
         MaxLength       =   20
         TabIndex        =   32
         Top             =   5490
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   315
         Left            =   4215
         TabIndex        =   25
         Top             =   180
         Width           =   2025
         _ExtentX        =   3572
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
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   48168963
         CurrentDate     =   38148
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   315
         Left            =   6780
         TabIndex        =   23
         Top             =   180
         Width           =   2055
         _ExtentX        =   3625
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
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   48168963
         CurrentDate     =   38148
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   6330
         TabIndex        =   27
         Top             =   210
         Width           =   405
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Left            =   3600
         TabIndex        =   26
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Select Journals Date Range:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   90
         TabIndex        =   22
         Top             =   180
         Width           =   3345
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
         Left            =   4860
         TabIndex        =   30
         Top             =   5520
         Width           =   1395
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7875
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      Begin VB.OptionButton optCode 
         Caption         =   "By &Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "By &Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   660
         Width           =   1725
      End
      Begin VB.TextBox TextSearch 
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   990
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   6390
         Left            =   60
         TabIndex        =   5
         Top             =   1380
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   11271
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
         MouseIcon       =   "AccountsLedger.frx":0776
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ACCOUNTS"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Search by:"
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
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   1455
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
      Left            =   13080
      MouseIcon       =   "AccountsLedger.frx":08D8
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":0A2A
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Exit Window"
      Top             =   7035
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
      Left            =   12390
      MouseIcon       =   "AccountsLedger.frx":0D90
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":0EE2
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Print this Record"
      Top             =   7035
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
      Left            =   11700
      MouseIcon       =   "AccountsLedger.frx":1248
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":139A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Find a Record"
      Top             =   7035
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
      Left            =   11010
      MouseIcon       =   "AccountsLedger.frx":1694
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":17E6
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Move to Next Record"
      Top             =   7035
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
      Left            =   10320
      MouseIcon       =   "AccountsLedger.frx":1B3E
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":1C90
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Move to Previous Record"
      Top             =   7035
      Width           =   705
   End
End
Attribute VB_Name = "frmAMISLEDGERAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSCHARTACCOUNT                                     As ADODB.Recordset
Dim RSACCTYPE                                          As ADODB.Recordset
Dim RSJOURNAL_HD                                       As ADODB.Recordset
Dim RSJOURNAL_HDDET                                    As ADODB.Recordset
Dim RSPROFILE                                          As ADODB.Recordset
Dim ADDOREDIT                                          As String
Dim ORDER_BY                                           As String
Dim TUTAL_DEBIT                                        As Double
Dim TUTAL_CREDIT                                       As Double
Dim TUTAL_BALANCE                                      As Double
Dim BEGINNING_BALANCE                                  As Double

Function GetAccountType(XXX As String) As String
    Dim rsChartAcctType                                               As ADODB.Recordset
    Set rsChartAcctType = New ADODB.Recordset
    Set rsChartAcctType = gconDMIS.Execute("Select HeaderCode from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsChartAcctType.EOF And Not rsChartAcctType.BOF Then
        GetAccountType = Null2String(rsChartAcctType!HeaderCode)
    End If
    Set rsChartAcctType = Nothing
End Function

Function SetAccType(Acc As String) As String
    Set RSACCTYPE = New ADODB.Recordset
    RSACCTYPE.Open "select * from AMIS_Acctype where code = " & N2Str2Null(Acc), gconDMIS
    If Not RSACCTYPE.EOF And Not RSACCTYPE.BOF Then
        SetAccType = Null2String(RSACCTYPE!Description)
    Else
        SetAccType = "Not Defined"
    End If
End Function

Function SetCustomerName(VVV As Variant)
    Dim rsCUSTOMER                                                    As ADODB.Recordset
    Set rsCUSTOMER = New ADODB.Recordset
    'Set rsCustomer = gconDMIS.Execute("Select custcode,AcctName from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(VVV))
    Set rsCUSTOMER = gconDMIS.Execute("Select custcode,AcctName,custname from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(VVV))
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then SetCustomerName = Null2String(rsCUSTOMER!CUSTNAME) Else SetCustomerName = ""
End Function

Function SetVendorName(VVV As Variant)
    Dim rsVENDOR                                                      As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV))
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then SetVendorName = Null2String(rsVENDOR!Nameofvendor) Else SetVendorName = ""
End Function

Function ReturnGJInformation(XXX As String) As String
    Dim SQL                                                           As String
    Dim rs                                                            As New ADODB.Recordset

    SQL = "SELECT refno from AMIS_journal_HD where voucherno=" & XXX & " and jtype='GJ'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        ReturnGJInformation = Null2String(rs!refno)
    End If
    Set rs = Nothing
End Function

Sub FillGrid2()
    Dim rsChartAccounts                                               As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    Set rsChartAccounts = gconDMIS.Execute("select Description,acctcode from AMIS_ChartAccount order by Description asc")
    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Sub FillGrids()
    Dim OUTBALANCE                                                    As Double
    Dim Reference, REFERENCE_NAME                                     As String
    Dim theReferenceInvoice As String
    Dim cnt                                                           As Integer
    cleargrid grdAccountsLedger: InitGrid
    Set RSJOURNAL_HDDET = New ADODB.Recordset
    '    rsJournal_HDDet.Open "select AMIS_Journal_Det.ID,AMIS_Journal_Det.JNo,AMIS_Journal_Det.JDate,AMIS_Journal_Det.JType,AMIS_Journal_Det.Debit,AMIS_Journal_Det.Credit,AMIS_Journal_Det.VoucherNo,AMIS_Journal_Hd.VendorCode,AMIS_Journal_Hd.CustomerCode,AMIS_Journal_Hd.JNo,AMIS_Journal_Hd.Remarks from AMIS_Journal_Det inner Join AMIS_Journal_Hd on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo where AMIS_Journal_Det.Jdate >= '" & dtFrom & "' and AMIS_Journal_Det.Jdate <= '" & dtTo & "' and AMIS_Journal_Det.Status='P' and AMIS_Journal_Det.Acct_Code = '" & txtCode.Text & "' order by AMIS_Journal_Det.JDate asc,AMIS_Journal_Det.ID asc", gconDMIS
    Set RSJOURNAL_HDDET = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & txtCode.Text & "'")
    TUTAL_BALANCE = 0: TUTAL_BALANCE = TUTAL_BALANCE + N2Str2Zero(RSCHARTACCOUNT!BeginningBalance): cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0: BEGINNING_BALANCE = 0
    If Not RSJOURNAL_HDDET.EOF And Not RSJOURNAL_HDDET.BOF Then
        If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or GetAccountType(txtCode.Text) = "8" Then
            OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(RSJOURNAL_HDDET!TOTAL_CREDIT) - N2Str2Zero(RSJOURNAL_HDDET!TOTAL_DEBIT)), 2)
            BEGINNING_BALANCE = OUTBALANCE
            grdAccountsLedger.AddItem Format(dtFrom, "mm/dd/yyyy") & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "BEGINNING BALANCE" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & "" & Chr(9)
        Else
            OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(RSJOURNAL_HDDET!TOTAL_DEBIT) - N2Str2Zero(RSJOURNAL_HDDET!TOTAL_CREDIT)), 2)
            BEGINNING_BALANCE = OUTBALANCE
            grdAccountsLedger.AddItem Format(dtFrom, "mm/dd/yyyy") & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "BEGINNING BALANCE" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & "" & Chr(9)
        End If
    End If
    Set RSJOURNAL_HDDET = New ADODB.Recordset
    RSJOURNAL_HDDET.Open "select * from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB'and JTYPE <>'BOB') AND Jdate >= '" & CDate(dtFrom) & "' and Jdate <= '" & CDate(dtTo) & "' and Acct_Code = '" & txtCode.Text & "'  Order by Jdate asc,ID asc", gconDMIS
    '    rsJournal_HDDet.Open "select * from AMIS_vw_GLEDGER where Jdate >= '" & dtFrom & "' and Jdate <= '" & dtTO & "' and Acct_Code = '" & txtCode.Text & "'", gconDMIS
    If Not RSJOURNAL_HDDET.EOF And Not RSJOURNAL_HDDET.BOF Then
        RSJOURNAL_HDDET.MoveFirst
        Screen.MousePointer = 11:
        Picture4.Visible = True
        PROGBAR.Value = 0
        PROGBAR.Max = RSJOURNAL_HDDET.RecordCount
        Do While Not RSJOURNAL_HDDET.EOF
            cnt = cnt + 1
            If Null2String(RSJOURNAL_HDDET!jtype) = "APJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "VPJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "VDJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "VCJ" Then
                Reference = Null2String(RSJOURNAL_HDDET!jtype) & "-" & Null2String(RSJOURNAL_HDDET!VOUCHERNO)
                REFERENCE_NAME = SetVendorName(Null2String(RSJOURNAL_HDDET!VendorCode))
            ElseIf Null2String(RSJOURNAL_HDDET!jtype) = "CDJ" Then
                Reference = "CDJ-" & Null2String(RSJOURNAL_HDDET!VOUCHERNO)
                REFERENCE_NAME = SetVendorName(Null2String(RSJOURNAL_HDDET!VendorCode))
            ElseIf Null2String(RSJOURNAL_HDDET!jtype) = "SJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "CSJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "CCM" Then
                Reference = Null2String(RSJOURNAL_HDDET!jtype) & "-" & Null2String(RSJOURNAL_HDDET!VOUCHERNO)
                REFERENCE_NAME = SetCustomerName(Null2String(RSJOURNAL_HDDET!CustomerCode))
            ElseIf Null2String(RSJOURNAL_HDDET!jtype) = "CRJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "DRJ" Then
                Reference = Null2String(RSJOURNAL_HDDET!jtype) & "-" & Null2String(RSJOURNAL_HDDET!VOUCHERNO)
                REFERENCE_NAME = SetCustomerName(Null2String(RSJOURNAL_HDDET!CustomerCode))
            Else
                Reference = Null2String(RSJOURNAL_HDDET!jtype) & "-" & Null2String(RSJOURNAL_HDDET!VOUCHERNO)
                If Null2String(RSJOURNAL_HDDET!jtype) = "GJ" Then
                    REFERENCE_NAME = ReturnGJInformation(Null2String(RSJOURNAL_HDDET!VOUCHERNO))
                Else
                    REFERENCE_NAME = SetCustomerName(Null2String(RSJOURNAL_HDDET!CustomerCode))
                End If
            End If
            'Update by BTT:12/4/2008
            If Null2String(RSJOURNAL_HDDET!jtype) = "CSJ" Or Null2String(RSJOURNAL_HDDET!jtype) = "CCM" Then
                theReferenceInvoice = getRefNo(Null2String(RSJOURNAL_HDDET!jtype), Null2String(RSJOURNAL_HDDET!VOUCHERNO))
                Else
                theReferenceInvoice = IIf(Null2String(RSJOURNAL_HDDET!InvoiceType) = "", Null2String(RSJOURNAL_HDDET!INVOICENO), Null2String(RSJOURNAL_HDDET!InvoiceType) & "-" & Null2String(RSJOURNAL_HDDET!INVOICENO))
            End If
            If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or GetAccountType(txtCode.Text) = "8" Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(RSJOURNAL_HDDET!CREDIT) - N2Str2Zero(RSJOURNAL_HDDET!DEBIT)), 2)
            Else
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(RSJOURNAL_HDDET!DEBIT) - N2Str2Zero(RSJOURNAL_HDDET!CREDIT)), 2)
            End If
            grdAccountsLedger.AddItem Format(Null2String(RSJOURNAL_HDDET!jdate), "mm/dd/yyyy") & Chr(9) & _
                                     Reference & Chr(9) & _
                                      theReferenceInvoice & Chr(9) & _
                                      REFERENCE_NAME & Chr(9) & _
                                      ToDoubleNumber(N2Str2Zero(RSJOURNAL_HDDET!DEBIT)) & Chr(9) & _
                                      ToDoubleNumber(N2Str2Zero(RSJOURNAL_HDDET!CREDIT)) & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & Trim(Null2String(RSJOURNAL_HDDET!Remarks)) & Chr(9) & RSJOURNAL_HDDET!ID
            
            'grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!Jdate) & Chr(9) & _
            '                         Reference & Chr(9) & _
            '                          IIf(Null2String(rsJournal_HDDet!InvoiceType) = "", Null2String(rsJournal_HDDet!InvoiceNo), Null2String(rsJournal_HDDet!InvoiceType) & "-" & Null2String(rsJournal_HDDet!InvoiceNo)) & Chr(9) & _
            '                          REFERENCE_NAME & Chr(9) & _
            '                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!DEBIT)) & Chr(9) & _
            '                          ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!CREDIT)) & Chr(9) & _
            '                          Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & Trim(Null2String(rsJournal_HDDet!Remarks)) & Chr(9) & rsJournal_HDDet!ID
            TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(RSJOURNAL_HDDET!DEBIT)
            TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(RSJOURNAL_HDDET!CREDIT)
            RSJOURNAL_HDDET.MoveNext
            ' Update by BTT : 1072008
            DoEvents
            PROGBAR.Value = PROGBAR.Value + 1
            Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
            Label11.Caption = Reference + "-" + REFERENCE_NAME
        Loop
        If cnt > 0 Then grdAccountsLedger.RemoveItem 1
        txtTotalDebit.Text = Format(TUTAL_DEBIT, "###,###,###,##0.00")
        txtTotalCredit.Text = Format(TUTAL_CREDIT, "###,###,###,##0.00")
        txtTotalBalance.Text = Format(TUTAL_BALANCE + OUTBALANCE, "###,###,###,##0.00")
        Screen.MousePointer = 0: grdAccountsLedger.MousePointer = flexCustom
    Else
        If BEGINNING_BALANCE > 0 Then
            grdAccountsLedger.RemoveItem 1
            txtTotalDebit.Text = Format(TUTAL_DEBIT, "###,###,###,##0.00")
            txtTotalCredit.Text = Format(TUTAL_CREDIT, "###,###,###,##0.00")
            txtTotalBalance.Text = Format(TUTAL_BALANCE + OUTBALANCE, "###,###,###,##0.00")
            Screen.MousePointer = 0: grdAccountsLedger.MousePointer = flexCustom
        Else
            txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO: txtTotalBalance.Text = ZERO
            cleargrid grdAccountsLedger
        End If
    End If
    Set RSJOURNAL_HDDET = Nothing
    Picture4.Visible = False
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccounts                                               As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsChartAccounts = gconDMIS.Execute("select AcctCode from AMIS_ChartAccount where AcctCode like'" & XXX & "%' order by AcctCode asc")

    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsChartAccounts                                               As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsChartAccounts = gconDMIS.Execute("select Description,acctcode from AMIS_ChartAccount where Description like'" & XXX & "%' order by Description asc")
    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Sub InitGrid()
    With grdAccountsLedger
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1300: .ColWidth(2) = 1300
        .ColWidth(3) = 2000: .ColWidth(4) = 1750: .ColWidth(5) = 1750: .ColWidth(6) = 1750
        .ColWidth(7) = 25000: .ColWidth(8) = 1: .Row = 0
        .Col = 0: .Text = "DOCDATE"
        .Col = 1: .Text = "REFERENCE"
        .Col = 2: .Text = "REF. INV."
        .Col = 3: .Text = "REFERENCE NAME"
        .Col = 4: .Text = "DEBIT"
        .Col = 5: .Text = "CREDIT"
        .Col = 6: .Text = "BALANCE"
        .Col = 7: .Text = "PARTICULARS"
        .Col = 8: .Text = "ID"
    End With
End Sub

Sub InitMemVars()
    Frame1.Enabled = True
    txtCode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
    txtDescription.Text = "": txtAcctType.Text = ""
    txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
    txtTotalBalance.Text = ZERO: txtBeginningBalance.Text = ZERO
    dtTo.Value = Now
End Sub

Sub rsRefresh()
    Set RSCHARTACCOUNT = New ADODB.Recordset
    RSCHARTACCOUNT.Open "select * from AMIS_ChartAccount order by AcctCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemvars()
    If Not RSCHARTACCOUNT.EOF And Not RSCHARTACCOUNT.BOF Then
        Frame1.Enabled = False
        labID.Caption = RSCHARTACCOUNT!ID
        txtCode.Text = Null2String(RSCHARTACCOUNT!acctcode)
        txtCode1.Text = Mid(Null2String(RSCHARTACCOUNT!acctcode), 1, 3)
        txtCode2.Text = Mid(Null2String(RSCHARTACCOUNT!acctcode), 5, 2)
        txtCode3.Text = Mid(Null2String(RSCHARTACCOUNT!acctcode), 8, 3)
        txtDescription.Text = Null2String(RSCHARTACCOUNT!Description)
        txtAcctType.Text = SetAccType(Null2String(RSCHARTACCOUNT!AcctType))
        txtBeginningBalance.Text = ToDoubleNumber(N2Str2Zero(RSCHARTACCOUNT!BeginningBalance))
        Set RSJOURNAL_HDDET = New ADODB.Recordset
        RSJOURNAL_HDDET.Open "select MIN(AMIS_Journal_Det.JDate) AS MinimumDate, MAX(AMIS_Journal_Det.JDate) AS MaximumDate from AMIS_Journal_Det inner Join AMIS_Journal_Hd on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo where AMIS_Journal_Det.Status='P' and AMIS_Journal_Det.Acct_Code = '" & txtCode.Text & "'", gconDMIS
        If Not RSJOURNAL_HDDET.EOF And Not RSJOURNAL_HDDET.BOF Then
            If IsNull(RSJOURNAL_HDDET!MinimumDate) = True Then
                cmdShow.Enabled = False
                dtFrom.Enabled = False
                dtTo.Enabled = False
            Else
                dtFrom.Enabled = True: dtTo.Enabled = True: cmdShow.Enabled = True
                'dtFrom = Null2Date(rsJournal_HDDet!MinimumDate)
                'Set rsProfile = New ADODB.Recordset
                'Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
                'If Not rsProfile.EOF And Not rsProfile.BOF Then dtFrom = DateSerial(Null2String(rsProfile!periodyear), Null2String(rsProfile!periodmonth), "1")
                'dtFrom = firstDay(Null2Date(RSJOURNAL_HDDET!MaximumDate))
                'dtTo = Null2Date(RSJOURNAL_HDDET!MaximumDate)
            End If
        Else
            dtFrom = LOGDATE: dtTo = LOGDATE: cmdShow.Enabled = False
            dtFrom.Enabled = False: dtTo.Enabled = False
        End If
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Frame2.ZOrder 0
    On Error Resume Next
    TextSearch.SetFocus
End Sub

'Upating Code       : AXP-0713200713:33
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    RSCHARTACCOUNT.MoveNext
    If RSCHARTACCOUNT.EOF Then
        RSCHARTACCOUNT.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:33
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    RSCHARTACCOUNT.MovePrevious
    If RSCHARTACCOUNT.BOF Then
        RSCHARTACCOUNT.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()


    If MsgBox("Print General Ledger for this Account?", vbYesNo + vbQuestion, "Print: " & txtDescription.Text) = vbYes Then
        rptGeneralLedger.Reset
        rptGeneralLedger.Formulas(3) = "BEG_DATE = '" & Format(dtFrom, "MMM-DD-YYYY") & "'"
        rptGeneralLedger.Formulas(4) = "BEGINNING = " & BEGINNING_BALANCE
        rptGeneralLedger.Formulas(5) = "REPORTDATE = '" & Format(dtTo, "LONG DATE") & "'"
        rptGeneralLedger.ReportTitle = "G E N E R A L  L E D G E R"
        Dim RSPROFILE                                                 As ADODB.Recordset
        Set RSPROFILE = New ADODB.Recordset
        Set RSPROFILE = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (RSPROFILE.EOF And RSPROFILE.BOF) Then

            rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(RSPROFILE!CompanyName) & "'"
            rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(RSPROFILE!Companyaddress) & "'"

            'ShowReport "AccountGeneralLedger", "Ledgers", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", "G E N E R A L  L E D G E R", "FROM: " & dtFrom & " TO: " & dtTO, True

        End If
        If grdAccountsLedger.TextMatrix(1, 2) = "BEGINNING BALANCE" And grdAccountsLedger.Rows = 2 Then

            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger_Beg.Rpt", "{ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
        End If
        LogAudit "V", "ACCOUNTS GENERAL LEDGER", txtCode
    End If
End Sub

Private Sub cmdShow_Click()
    FillGrids
End Sub

Private Sub FillGrid()
    Dim rsChartAccounts                                               As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    Set rsChartAccounts = gconDMIS.Execute("select AcctCode from AMIS_ChartAccount order by AcctCode asc")
    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Private Sub Command1_Click()
'Dim I
'Dim RSDET As ADODB.Recordset
'Dim TOTALBAL As Double
'Dim XTYPE As String
'    For I = 1 To grdAccountsLedger.Rows - 1
'            grdAccountsLedger.RowSel = I
'
'        If Left(grdAccountsLedger.TextMatrix(I, 1), 3) = "CRJ" Then
'
'            Set RSDET = gconDMIS.Execute("SELECT INVOICENO, INVOICETYPE FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO=" & N2Str2Null(Replace(grdAccountsLedger.TextMatrix(I, 1), "CRJ", "")))
'
'
'            If Not (RSDET.EOF Or RSDET.BOF) Then
'                XTYPE = getJTYPE(RSDET!InvoiceNo, RSDET!InvoiceType)
'
'                If XTYPE = "SJ" Then
'                    If rsCHECKINVOICENOandTYPE(RSDET!InvoiceType, RSDET!InvoiceNo, "", "SJ") = False Then
'                       grdAccountsLedger.Col = 1
'                        grdAccountsLedger.Row = I
'                        grdAccountsLedger.CellFontBold = True
'                        grdAccountsLedger.CellForeColor = vbBlue
'                       TOTALBAL = TOTALBAL + grdAccountsLedger.TextMatrix(I, 5)
'                    End If
'                ElseIf XTYPE = "COB" Then
'                    If rsCHECKINVOICENOandTYPE(RSDET!InvoiceType, RSDET!InvoiceNo, "", "SJ") = False Then
'                       grdAccountsLedger.Col = 1
'                        grdAccountsLedger.Row = I
'                        grdAccountsLedger.CellFontBold = True
'                        grdAccountsLedger.CellForeColor = vbBlue
'                       TOTALBAL = TOTALBAL + grdAccountsLedger.TextMatrix(I, 5)
'                    End If
'                End If
'
'            Else
'                grdAccountsLedger.Col = 1
'                    grdAccountsLedger.Row = I
'                    grdAccountsLedger.CellFontBold = True
'                    grdAccountsLedger.CellForeColor = vbBlue
'                   TOTALBAL = TOTALBAL + grdAccountsLedger.TextMatrix(I, 5)
'            End If
'
'        End If
'    Next
'    MsgBox TOTALBAL
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyC Then
            optCode.Value = True
            optCode_Click
            On Error Resume Next
            TextSearch.SetFocus
        End If
        If KeyCode = vbKeyD Then
            optDescription.Value = True
            optDescription_Click
            On Error Resume Next
            TextSearch.SetFocus
        End If
    End If
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    TextSearch.Text = ""
    InitGrid
    InitMemVars
    StoreMemvars
    Picture4.Visible = False
    Screen.MousePointer = 0
End Sub

'Function SetCustomerName(VVV As Variant)
'Dim rsCustomer As ADODB.Recordset
'Set rsCustomer = New ADODB.Recordset
'Set rsCustomer = gconDMIS.Execute("Select custcode,AcctName from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(VVV))
'If Not rsCustomer.EOF And Not rsCustomer.BOF Then SetCustomerName = Null2String(rsCustomer!AcctName) Else SetCustomerName = ""
'End Function

Private Sub grdAccountsLedger_DblClick()
    grdAccountsLedger.Row = grdAccountsLedger.Row
    grdAccountsLedger.Col = 1
    Dim VARVOUCHERNO                                                  As String
    If Left(grdAccountsLedger.Text, 3) = "APJ" Then
        JOURNALTYPE = "APJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CDJ" Then
        JOURNALTYPE = "CDJ"
    ElseIf Left(grdAccountsLedger.Text, 2) = "SJ" Then
        JOURNALTYPE = "SJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CRJ" Then
        JOURNALTYPE = "CRJ"
    ElseIf Left(grdAccountsLedger.Text, 2) = "GJ" Then
        JOURNALTYPE = "GJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "ADJ" Then
        JOURNALTYPE = "ADJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "PDJ" Then
        JOURNALTYPE = "PDJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CLO" Then
        JOURNALTYPE = "CLO"
    ElseIf Left(grdAccountsLedger.Text, 3) = "DRJ" Then
        JOURNALTYPE = "DRJ"
    Else
        JOURNALTYPE = "OPB"    '
    End If
    'JOURNALTYPE = Left(grdAccountsLedger.Text, 3)
    VARVOUCHERNO = Right(grdAccountsLedger.Text, 6)
    Screen.MousePointer = 11
    On Error Resume Next
    If JOURNALTYPE = "DRJ" Then
        Unload frmAMISDRJJournalEntry
        frmAMISDRJJournalEntry.Show
        frmAMISDRJJournalEntry.StoreSearch (VARVOUCHERNO)
    Else
        Unload frmAMISJournalEntry
        frmAMISJournalEntry.Show
        frmAMISJournalEntry.StoreSearch (VARVOUCHERNO)
    End If
    Screen.MousePointer = 0
End Sub

Private Sub lstAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstAccounts
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next

    If optCode.Value = True Then
        RSCHARTACCOUNT.Bookmark = rsFind(RSCHARTACCOUNT.Clone, "acctcode", lstAccounts.SelectedItem).Bookmark
    Else
        RSCHARTACCOUNT.Bookmark = rsFind(RSCHARTACCOUNT.Clone, "acctcode", lstAccounts.SelectedItem.SubItems(1)).Bookmark
    End If
    cleargrid grdAccountsLedger
    InitGrid
    StoreMemvars
End Sub

Private Sub lstAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        TextSearch.SetFocus
    End If
End Sub

Private Sub optCode_Click()
    If TextSearch = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
    On Error Resume Next
    TextSearch.SetFocus
End Sub

Private Sub optDescription_Click()
    If TextSearch = "" Then FillGrid2 Else FillSearchGrid2 (TextSearch.Text)
    On Error Resume Next
    TextSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optCode.Value = True Then
        If Trim(TextSearch.Text) = "" Then FillGrid Else FillSearchGrid (TextSearch.Text)
    Else
        If Trim(TextSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (TextSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Frame2.ZOrder 0
    If KeyCode = vbKeyDown Then
        If lstAccounts.ListItems.Count > 0 And lstAccounts.Enabled = True Then
            lstAccounts.SetFocus
        End If
    End If
End Sub
Function getRefNo(XXX As String, YYY As String)
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT REFNO FROM AMIS_JOURNAL_HD WHERE JTYPE ='" & XXX & "' AND VOUCHERNO = '" & YYY & "'")
    If Not (rs.EOF And rs.BOF) Then
        getRefNo = Null2String(rs!refno)
    End If
    Set rs = Nothing
End Function

