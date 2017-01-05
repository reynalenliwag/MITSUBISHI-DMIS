VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAMISLEDGERAccounts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts General Ledger"
   ClientHeight    =   8115
   ClientLeft      =   720
   ClientTop       =   435
   ClientWidth     =   14145
   ForeColor       =   &H00FFFFFF&
   Icon            =   "AccountsLedger.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   14145
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
      Left            =   12600
      MouseIcon       =   "AccountsLedger.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":045C
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Print this Record"
      Top             =   7200
      Width           =   705
   End
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
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loading Data"
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
            Height          =   195
            Left            =   30
            TabIndex        =   40
            Top             =   0
            Width           =   1275
         End
      End
      Begin MSComctlLib.ProgressBar PROGBAR 
         Height          =   405
         Left            =   0
         TabIndex        =   51
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
         Height          =   8355
         Left            =   600
         TabIndex        =   43
         Top             =   240
         Width           =   14235
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   30
         TabIndex        =   42
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2650
      TabIndex        =   6
      Top             =   0
      Width           =   11475
      Begin VB.TextBox txtBeginningBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   9180
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   570
         Width           =   2235
      End
      Begin VB.TextBox txtAcctType 
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   18
         Top             =   570
         Width           =   1665
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
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
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   8145
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Left            =   7680
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
      Height          =   6180
      Left            =   2655
      TabIndex        =   21
      Top             =   945
      Width           =   11475
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
         MouseIcon       =   "AccountsLedger.frx":07C2
      End
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
         MouseIcon       =   "AccountsLedger.frx":0ADC
         MousePointer    =   99  'Custom
         TabIndex        =   24
         ToolTipText     =   "Show Ledger"
         Top             =   180
         Width           =   1470
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
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Height          =   360
         Left            =   9600
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         Top             =   5520
         Width           =   1785
      End
      Begin VB.TextBox txtTotalDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Height          =   360
         Left            =   6000
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   29
         Top             =   5520
         Width           =   1755
      End
      Begin VB.TextBox txtTotalCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Height          =   360
         Left            =   7800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   32
         Top             =   5520
         Width           =   1755
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
         Format          =   132120579
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
         Format          =   132120579
         CurrentDate     =   38148
      End
      Begin MSComctlLib.ProgressBar PRB 
         Height          =   255
         Left            =   2400
         TabIndex        =   52
         Top             =   5520
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   46
         Top             =   5880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblMax 
         Height          =   255
         Index           =   2
         Left            =   3960
         TabIndex        =   47
         Top             =   5880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblOf 
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   48
         Top             =   5880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblPRB 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label9 
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
         Caption         =   "Select Journals Date Range:"
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
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Height          =   240
         Left            =   5415
         TabIndex        =   30
         Top             =   5520
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7995
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
         Appearance      =   0  'Flat
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
         TabIndex        =   50
         Top             =   1440
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
         MouseIcon       =   "AccountsLedger.frx":0C2E
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
      Left            =   13320
      MouseIcon       =   "AccountsLedger.frx":0D90
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":0EE2
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Exit Window"
      Top             =   7200
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
      Left            =   11880
      MouseIcon       =   "AccountsLedger.frx":1248
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":139A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Find a Record"
      Top             =   7200
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
      Left            =   11160
      MouseIcon       =   "AccountsLedger.frx":1694
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":17E6
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Move to Next Record"
      Top             =   7200
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
      Left            =   10440
      MouseIcon       =   "AccountsLedger.frx":1B3E
      MousePointer    =   99  'Custom
      Picture         =   "AccountsLedger.frx":1C90
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Move to Previous Record"
      Top             =   7200
      Width           =   705
   End
   Begin VB.PictureBox picPrintopt 
      BackColor       =   &H00E0E0E0&
      Height          =   1365
      Left            =   6840
      ScaleHeight     =   1305
      ScaleWidth      =   3870
      TabIndex        =   44
      Top             =   3285
      Width           =   3930
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Height          =   1275
         Left            =   45
         TabIndex        =   45
         Top             =   0
         Width           =   3795
         Begin VB.OptionButton optPrintAll 
            BackColor       =   &H8000000A&
            Caption         =   "Print all account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   225
            Width           =   3570
         End
         Begin VB.OptionButton optPrintbyAccount 
            BackColor       =   &H8000000B&
            Caption         =   "Print by account"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   720
            Width           =   3570
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Height          =   1275
         Left            =   45
         TabIndex        =   53
         Top             =   0
         Width           =   3795
         Begin VB.OptionButton optDetailed 
            BackColor       =   &H8000000B&
            Caption         =   "Print Detailed"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   720
            Width           =   3570
         End
         Begin VB.OptionButton optSummary 
            BackColor       =   &H8000000A&
            Caption         =   "Print Summary"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   225
            Width           =   3570
         End
      End
   End
End
Attribute VB_Name = "frmAMISLEDGERAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Dim rsChartAccount, rsAccType                               As ADODB.Recordset
Dim rsJournal_HD, rsJournal_HDDet, rsProfile                As ADODB.Recordset
Dim AddorEdit, ORDER_BY                                     As String

Dim rsLOAD_VENDOR                                           As ADODB.Recordset
Dim rsLOAD_AP                                               As ADODB.Recordset
Dim REC                                                     As XtremeReportControl.ReportRecord
Dim TOTAL_DEBIT                                             As Double
Dim TOTAL_CREDIT                                            As Double
Dim xBALANCE                                                As Double
Dim FWD_BALANCE                                             As Double
Dim FWD_DEBIT                                               As Double
Dim FWD_CREDIT                                              As Double
Dim rsREF                                                   As ADODB.Recordset
Dim xlApplication                                           As Excel.Application
Dim xlWorkbook                                              As Excel.Workbook
Dim xlWorksheet                                             As Excel.Worksheet
Dim xlRange                                                 As Excel.Range
Dim xCounter                                                As Integer
Dim xAcctCode                                               As String

Dim TUTAL_DEBIT, TUTAL_CREDIT, TUTAL_BALANCE, BEGINNING_BALANCE As Double

Function GetAccountType(XXX As String) As String
    Dim rsChartAcctType                                     As ADODB.Recordset
    Set rsChartAcctType = New ADODB.Recordset
    Set rsChartAcctType = gconDMIS.Execute("Select HeaderCode from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsChartAcctType.EOF And Not rsChartAcctType.BOF Then
        GetAccountType = Null2String(rsChartAcctType!HeaderCode)
    End If
    Set rsChartAcctType = Nothing
End Function

Function GetHeaders(XXX As String) As String
    Dim rsGetHeaders                                    As ADODB.Recordset
    Set rsGetHeaders = New ADODB.Recordset
    Set rsGetHeaders = gconDMIS.Execute("Select Headers from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsGetHeaders.EOF And Not rsGetHeaders.BOF Then
        GetHeaders = Null2String(rsGetHeaders!HEADERS)
    End If
    Set rsGetHeaders = Nothing
End Function

Function GetTitleCode(XXX As String) As String
    Dim rsGetTitleCode                                   As ADODB.Recordset
    Set rsGetTitleCode = New ADODB.Recordset
    Set rsGetTitleCode = gconDMIS.Execute("Select TitleCode from AMIS_ChartAccount Where AcctCode = '" & XXX & "'")
    If Not rsGetTitleCode.EOF And Not rsGetTitleCode.BOF Then
        GetTitleCode = Null2String(rsGetTitleCode!TitleCode)
    End If
    Set rsGetTitleCode = Nothing
End Function

Function SetAccType(Acc As String) As String
    Set rsAccType = New ADODB.Recordset
    rsAccType.Open "select * from AMIS_Acctype where code = " & N2Str2Null(Acc), gconDMIS
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        SetAccType = Null2String(rsAccType!DESCRIPTION)
    Else
        SetAccType = "Not Defined"
    End If
End Function

Function SetCustomerName(VVV As Variant)
    Dim rsCustomer                                          As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    'OLD QUERY 'Set rsCustomer = gconDMIS.Execute("Select custcode,AcctName,custname from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(VVV))
    
    'SJR 082714
    Set rsCustomer = gconDMIS.Execute("Select custcode,AcctName,custname from " & _
                                    "(Select custcode,AcctName,custname from ALL_CUSTMASTER_AMIS " & _
                                    "Union " & _
                                    "select bankcode as custcode, bankname As AcctName,bankname As custname from ALL_BANKDEPOSITS) " & _
                                    "as tbl where tbl.custcode = " & N2Str2Null(VVV))
    'SJR 082714
    
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then SetCustomerName = Null2String(rsCustomer!CUSTNAME) Else SetCustomerName = ""
End Function

Function SetVendorName(VVV As Variant)
    Dim rsVENDOR                                            As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV))
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then SetVendorName = Null2String(rsVENDOR!nameofvendor) Else SetVendorName = ""
End Function

Function ReturnGJInformation(XXX As String) As String
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    SQL = "SELECT refno from AMIS_journal_HD where voucherno=" & XXX & " and jtype='GJ'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        ReturnGJInformation = Null2String(RS!REFNO)
    End If
    Set RS = Nothing
End Function

Sub FillGrid2()
    Dim rsChartAccounts                                     As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    Set rsChartAccounts = gconDMIS.Execute("select Description,acctcode from AMIS_ChartAccount WHERE DESCRIPTION <> '-' order by Description asc")
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
    Dim OUTBALANCE                                          As Double
    Dim Reference, REFERENCE_NAME, ENTITY_CLASS             As String
    Dim theReferenceInvoice                                 As String
    Dim cnt                                                 As Integer

    cleargrid grdAccountsLedger: initGrid
    Set rsJournal_HDDet = New ADODB.Recordset
    Set rsJournal_HDDet = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & txtCode.Text & "'")
    TUTAL_BALANCE = 0: TUTAL_BALANCE = TUTAL_BALANCE + N2Str2Zero(rsChartAccount!BeginningBalance): cnt = 0: TUTAL_DEBIT = 0: TUTAL_CREDIT = 0: OUTBALANCE = 0: BEGINNING_BALANCE = 0
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        'OLD
        '01/08/2015
        'If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or GetAccountType(txtCode.Text) = "8" Then
        If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or (GetAccountType(txtCode.Text) = "8" And COMPANY_CODE <> "DJM") Or (GetAccountType(txtCode.Text) = "7" And COMPANY_CODE = "DJM") Then
            OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT)), 2)
            BEGINNING_BALANCE = OUTBALANCE
            grdAccountsLedger.AddItem dtFrom & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "BEGINNING BALANCE" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & "" & Chr(9)
        Else
            OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT)), 2)
            BEGINNING_BALANCE = OUTBALANCE
            grdAccountsLedger.AddItem dtFrom & Chr(9) & _
                                      "" & Chr(9) & _
                                      "" & Chr(9) & _
                                      "BEGINNING BALANCE" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      "0.00" & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & "" & Chr(9)
        End If
    End If
    
    Set rsJournal_HDDet = New ADODB.Recordset
        'SJR 082714
        If COMPANY_CODE <> "HCA" And COMPANY_CODE <> "DJM" Then
            rsJournal_HDDet.Open "select * from AMIS_vw_vLEDGER where JTYPE <> 'VPJ' and JTYPE <> 'COB'and JTYPE <> 'BOB' AND Jdate >= '" & CDate(dtFrom) & "' and Jdate <= '" & CDate(dtTo) & "' and Acct_Code = '" & txtCode.Text & "' Order by Jdate asc, voucherno asc", gconDMIS
        ElseIf COMPANY_CODE = "DJM" Then
            rsJournal_HDDet.Open "select * from AMIS_vw_vLEDGER where JTYPE <> 'DRJ' and JTYPE <> 'VPJ' and JTYPE <> 'COB'and JTYPE <> 'BOB' AND Jdate >= '" & CDate(dtFrom) & "' and Jdate <= '" & CDate(dtTo) & "' and Acct_Code = '" & txtCode.Text & "' Order by Jdate asc,ID asc", gconDMIS
        Else
            rsJournal_HDDet.Open "SELECT TOP (100) PERCENT " & _
                                    "AMIS_JOURNAL_DET.ID, " & _
                                    "AMIS_JOURNAL_DET.JNO, " & _
                                    "AMIS_JOURNAL_DET.JDATE," & _
                                    "AMIS_JOURNAL_DET.JTYPE," & _
                                    "(CASE AMIS_JOURNAL_DET.JTYPE " & _
                                    "WHEN 'DRJ' THEN AMIS_JOURNAL_DET.INVOICENO " & _
                                    "Else AMIS_JOURNAL_HD.INVOICENO END)INVOICENO, " & _
                                    "AMIS_JOURNAL_DET.DEBIT, AMIS_JOURNAL_DET.CREDIT, " & _
                                    "AMIS_JOURNAL_DET.VOUCHERNO, AMIS_JOURNAL_HD.VENDORCODE, " & _
                                    "AMIS_JOURNAL_HD.CUSTOMERCODE, AMIS_JOURNAL_HD.REMARKS, " & _
                                    "AMIS_JOURNAL_DET.ACCT_CODE, AMIS_JOURNAL_DET.STATUS, " & _
                                    "AMIS_JOURNAL_HD.INVOICETYPE, AMIS_JOURNAL_HD.ENTITY_CLASS " & _
                                    "FROM AMIS_JOURNAL_DET INNER JOIN AMIS_JOURNAL_HD " & _
                                    "ON AMIS_JOURNAL_DET.JNO = AMIS_JOURNAL_HD.JNO " & _
                                    "AND AMIS_JOURNAL_DET.JTYPE = AMIS_JOURNAL_HD.JTYPE " & _
                                    "AND AMIS_JOURNAL_DET.VOUCHERNO = AMIS_JOURNAL_HD.VOUCHERNO " & _
                                    "WHERE (AMIS_JOURNAL_DET.STATUS = 'P') " & _
                                    "AND AMIS_JOURNAL_DET.JTYPE <> 'VPJ' " & _
                                    "AND AMIS_JOURNAL_DET.JTYPE <> 'COB' " & _
                                    "AND AMIS_JOURNAL_DET.JTYPE <> 'BOB' " & _
                                    "AND AMIS_JOURNAL_DET.JDATE >= '" & CDate(dtFrom) & "' " & _
                                    "AND AMIS_JOURNAL_DET.JDATE <= '" & CDate(dtTo) & "' " & _
                                    "AND AMIS_JOURNAL_DET.ACCT_CODE = '" & txtCode.Text & "' " & _
                                    "ORDER BY AMIS_JOURNAL_DET.JDATE, AMIS_JOURNAL_DET.VOUCHERNO", gconDMIS
        End If
        'SJR 082714
        
    If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
        rsJournal_HDDet.MoveFirst
        Screen.MousePointer = 11:
        Picture4.Visible = True
        PROGBAR.Value = 0
        PROGBAR.Max = rsJournal_HDDet.RecordCount
        Do While Not rsJournal_HDDet.EOF
            cnt = cnt + 1
            If Null2String(rsJournal_HDDet!JTYPE) = "APJ" Or Null2String(rsJournal_HDDet!JTYPE) = "VPJ" Or Null2String(rsJournal_HDDet!JTYPE) = "VDJ" Or Null2String(rsJournal_HDDet!JTYPE) = "VCJ" Then
                Reference = Null2String(rsJournal_HDDet!JTYPE) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
                ENTITY_CLASS = Null2String(rsJournal_HDDet!ENTITY_CLASS)
                If ENTITY_CLASS = "C" Then
                    REFERENCE_NAME = SetCustomerName(Null2String(rsJournal_HDDet!VendorCode))
                Else
                    REFERENCE_NAME = SetVendorName(Null2String(rsJournal_HDDet!VendorCode))
                End If
            ElseIf Null2String(rsJournal_HDDet!JTYPE) = "CDJ" Then
                Reference = "CDJ-" & Null2String(rsJournal_HDDet!VOUCHERNO)
                ENTITY_CLASS = Null2String(rsJournal_HDDet!ENTITY_CLASS)
                If ENTITY_CLASS = "C" Then
                    REFERENCE_NAME = SetCustomerName(Null2String(rsJournal_HDDet!VendorCode))
                Else
                    REFERENCE_NAME = SetVendorName(Null2String(rsJournal_HDDet!VendorCode))
                End If
            ElseIf Null2String(rsJournal_HDDet!JTYPE) = "SJ" Or Null2String(rsJournal_HDDet!JTYPE) = "CSJ" Or Null2String(rsJournal_HDDet!JTYPE) = "CCM" Then
                Reference = Null2String(rsJournal_HDDet!JTYPE) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
                REFERENCE_NAME = SetCustomerName(Null2String(rsJournal_HDDet!CustomerCode))
            ElseIf Null2String(rsJournal_HDDet!JTYPE) = "CRJ" Or Null2String(rsJournal_HDDet!JTYPE) = "DRJ" Then
                Reference = Null2String(rsJournal_HDDet!JTYPE) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
                REFERENCE_NAME = SetCustomerName(Null2String(rsJournal_HDDet!CustomerCode))
            Else
                Reference = Null2String(rsJournal_HDDet!JTYPE) & "-" & Null2String(rsJournal_HDDet!VOUCHERNO)
                If Null2String(rsJournal_HDDet!JTYPE) = "GJ" Then
                    REFERENCE_NAME = FIND_NAME_GJ(rsJournal_HDDet!ID)
                Else
                    REFERENCE_NAME = SetCustomerName(Null2String(rsJournal_HDDet!CustomerCode))
                End If
            End If
            
            'Update by BTT:12/4/2008
            If Null2String(rsJournal_HDDet!JTYPE) = "CSJ" Or Null2String(rsJournal_HDDet!JTYPE) = "CCM" Then
                theReferenceInvoice = getRefNo(Null2String(rsJournal_HDDet!JTYPE), Null2String(rsJournal_HDDet!VOUCHERNO))
            ElseIf Null2String(rsJournal_HDDet!JTYPE) = "GJ" Then
                theReferenceInvoice = FIND_GJ_REF(rsJournal_HDDet!ID)
            Else
                theReferenceInvoice = IIf(Null2String(rsJournal_HDDet!INVOICETYPE) = "", Null2String(rsJournal_HDDet!INVOICENO), Null2String(rsJournal_HDDet!INVOICETYPE) & "-" & Null2String(rsJournal_HDDet!INVOICENO))
            End If
            
            'OLD
            '01/08/2015
            'If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or GetAccountType(txtCode.Text) = "8" Then
            If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or (GetAccountType(txtCode.Text) = "8" And COMPANY_CODE <> "DJM") Or (GetAccountType(txtCode.Text) = "7" And COMPANY_CODE = "DJM") Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!Credit) - N2Str2Zero(rsJournal_HDDet!Debit)), 2)
            Else
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!Debit) - N2Str2Zero(rsJournal_HDDet!Credit)), 2)
            End If
            
            grdAccountsLedger.AddItem Null2String(rsJournal_HDDet!JDATE) & Chr(9) & _
                                      Reference & Chr(9) & _
                                      theReferenceInvoice & Chr(9) & _
                                      REFERENCE_NAME & Chr(9) & _
                                      ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!Debit)) & Chr(9) & _
                                      ToDoubleNumber(N2Str2Zero(rsJournal_HDDet!Credit)) & Chr(9) & _
                                      Format(OUTBALANCE, "###,###,###,##0.00") & Chr(9) & Trim(Null2String(rsJournal_HDDet!remarks)) & Chr(9) & rsJournal_HDDet!ID

            TUTAL_DEBIT = TUTAL_DEBIT + N2Str2Zero(rsJournal_HDDet!Debit)
            TUTAL_CREDIT = TUTAL_CREDIT + N2Str2Zero(rsJournal_HDDet!Credit)
            rsJournal_HDDet.MoveNext
            
            ' Update by BTT : 1072008
            DoEvents
            If PROGBAR.Value = PROGBAR.Max Then
                PROGBAR.Enabled = False
            Else
                PROGBAR.Value = PROGBAR.Value + 1
                Label20 = Round((PROGBAR.Value / PROGBAR.Max * 100), 0) & "%"
                Label11.Caption = Reference + "-" + REFERENCE_NAME
            End If
        Loop
        If cnt > 0 Then grdAccountsLedger.RemoveItem 1
        txtTotalDebit.Text = Format(TUTAL_DEBIT, "###,###,###,##0.00")
        txtTotalCredit.Text = Format(TUTAL_CREDIT, "###,###,###,##0.00")
        txtTotalBalance.Text = Format(TUTAL_BALANCE + OUTBALANCE, "###,###,###,##0.00")
        Screen.MousePointer = 0: grdAccountsLedger.MousePointer = flexCustom
    Else
        If BEGINNING_BALANCE <> 0 Then
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
    Set rsJournal_HDDet = Nothing
    Picture4.Visible = False
End Sub

Function FIND_NAME_GJ(xID As Variant) As String
'UPDATED BY: JUN
'DATE UPDATED: 07282009
'DESCRIPTION: GET THE CUSTOMER CODE AMIS_JOURNAL_HD AND GET THE CORRESPONDING NAME IN THE  ALL_ENTITY
    Dim rsFIND_NAME_GJ                                      As ADODB.Recordset
    Dim rsGet_Name                                          As ADODB.Recordset
    Dim checkEntityClass
    
    Set rsFIND_NAME_GJ = New ADODB.Recordset
    rsFIND_NAME_GJ.Open "Select * from Amis_journal_DET where ID = '" & xID & "'", gconDMIS, adOpenKeyset
    If Not rsFIND_NAME_GJ.EOF And Not rsFIND_NAME_GJ.BOF Then
        Set rsGet_Name = New ADODB.Recordset
        
        'Updated by SJR 072414
        checkEntityClass = rsFIND_NAME_GJ!ENTITY
        If Left(checkEntityClass, 1) = "E" Then
               checkEntityClass = Right(checkEntityClass, Len(checkEntityClass) - 1)
               rsGet_Name.Open "Select AccountName from All_ENTITY WHERE CODE = '" & Right(Null2String(checkEntityClass), 3) & "' AND ENTITYCODE = '" & Left(Null2String(rsFIND_NAME_GJ!ENTITY), 1) & "'", gconDMIS, adOpenKeyset
        Else
               rsGet_Name.Open "Select AccountName from All_ENTITY WHERE CODE = '" & Right(Null2String(rsFIND_NAME_GJ!ENTITY), 6) & "' AND ENTITYCODE = '" & Left(Null2String(rsFIND_NAME_GJ!ENTITY), 1) & "'", gconDMIS, adOpenKeyset
        End If
        'Updated by SJR 072414
        
        If Not rsGet_Name.EOF And Not rsGet_Name.BOF Then
            FIND_NAME_GJ = Null2String(rsGet_Name!ACCOUNTNAME)
            If Left(Null2String(rsFIND_NAME_GJ!ENTITY), 1) = "C" Then
                gconDMIS.Execute "UPDATE AMIS_JOURNAL_HD SET CUSTOMERCODE= " & N2Str2Null(Right(rsFIND_NAME_GJ!ENTITY, 6)) & ",VENDORCODE='999999'  WHERE VOUCHERNO ='" & Null2String(rsFIND_NAME_GJ!VOUCHERNO) & "' and JTYPE='GJ'"
            Else
                gconDMIS.Execute "UPDATE AMIS_JOURNAL_HD SET VENDORCODE= " & N2Str2Null(Right(rsFIND_NAME_GJ!ENTITY, 6)) & ",CUSTOMERCODE='999999'  WHERE VOUCHERNO ='" & Null2String(rsFIND_NAME_GJ!VOUCHERNO) & "' and JTYPE='GJ'"
            End If
        End If
    End If
    Set rsFIND_NAME_GJ = Nothing
End Function

Function FIND_GJ_REF(xID As Variant) As String
'UPDATED BY: JUN
'DATE UPDATED: 07282009
'DESCRIPTION: GET THE THE REFERENCE INVOICENO AND INVOICE TYPE FROM AMIS_JOURNAL_DET BY MEANS OF ID
    Dim rsFIND_GJ_REF                                       As ADODB.Recordset
    Set rsFIND_GJ_REF = New ADODB.Recordset
    rsFIND_GJ_REF.Open "Select InvoiceNo,Invoicetype from Amis_journal_det where ID ='" & xID & "'", gconDMIS, adOpenKeyset
    If Not rsFIND_GJ_REF.EOF And Not rsFIND_GJ_REF.BOF Then
        FIND_GJ_REF = Null2String(rsFIND_GJ_REF!INVOICETYPE) & "-" & Null2String(rsFIND_GJ_REF!INVOICENO)
    End If
    Set rsFIND_GJ_REF = Nothing
End Function

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccounts                                     As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    'XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsChartAccounts = gconDMIS.Execute("select AcctCode from AMIS_ChartAccount where AcctCode like'" & XXX & "%' AND DESCRIPTION <> '-' order by AcctCode asc")

    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsChartAccounts                                     As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    'XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsChartAccounts = gconDMIS.Execute("select Description,acctcode from AMIS_ChartAccount where Description like'" & XXX & "%' AND DESCRIPTION <> '-' order by Description asc")
    If Not (rsChartAccounts.EOF And rsChartAccounts.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccounts
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Sub initGrid()
    With grdAccountsLedger
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1300: .ColWidth(2) = 1300
        .ColWidth(3) = 2000: .ColWidth(4) = 1750: .ColWidth(5) = 1750: .ColWidth(6) = 1750
        .ColWidth(7) = 2327600: .ColWidth(8) = 1: .Row = 0
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

Sub initMemvars()
    Frame1.Enabled = True
    txtCode.Text = "": txtCode1.Text = "": txtCode2.Text = "": txtCode3.Text = ""
    txtDescription.Text = "": txtAcctType.Text = ""
    txtTotalDebit.Text = ZERO: txtTotalCredit.Text = ZERO
    txtTotalBalance.Text = ZERO: txtBeginningBalance.Text = ZERO
    dtTo.Value = Now
End Sub

Sub rsRefresh()
    Set rsChartAccount = New ADODB.Recordset
    rsChartAccount.Open "select * from AMIS_ChartAccount order by AcctCode asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        '        Frame1.Enabled = False
        labID.Caption = rsChartAccount!ID
        txtCode.Text = Null2String(rsChartAccount!AcctCode)
        txtCode1.Text = Mid(Null2String(rsChartAccount!AcctCode), 1, 3)
        txtCode2.Text = Mid(Null2String(rsChartAccount!AcctCode), 5, 2)
        txtCode3.Text = Mid(Null2String(rsChartAccount!AcctCode), 8, 3)
        txtDescription.Text = Null2String(rsChartAccount!DESCRIPTION)
        txtAcctType.Text = SetAccType(Null2String(rsChartAccount!ACCTTYPE))
        'txtBeginningBalance.Text = ToDoubleNumber(N2Str2Zero(rsChartAccount!BeginningBalance))
        txtBeginningBalance.Text = ToDoubleNumber(OpeningBalance(txtCode.Text))
        Set rsJournal_HDDet = New ADODB.Recordset
        rsJournal_HDDet.Open "select MIN(AMIS_Journal_Det.JDate) AS MinimumDate, MAX(AMIS_Journal_Det.JDate) AS MaximumDate from AMIS_Journal_Det inner Join AMIS_Journal_Hd on AMIS_Journal_Det.JNo = AMIS_Journal_Hd.JNo where AMIS_Journal_Det.Status='P' and AMIS_Journal_Det.Acct_Code = '" & txtCode.Text & "'", gconDMIS
        If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
            If IsNull(rsJournal_HDDet!MinimumDate) = True Then
                cmdShow.Enabled = False
                dtFrom.Enabled = False
                dtTo.Enabled = False
            Else
                dtFrom.Enabled = True: dtTo.Enabled = True: cmdShow.Enabled = True
                'dtFrom = Null2Date(rsJournal_HDDet!MinimumDate)
                'Set rsProfile = New ADODB.Recordset
                'Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
                'If Not rsProfile.EOF And Not rsProfile.BOF Then dtFrom = DateSerial(Null2String(rsProfile!periodyear), Null2String(rsProfile!periodmonth), "1")
                dtFrom = firstDay(Null2Date(rsJournal_HDDet!MaximumDate))
                dtTo = Null2Date(rsJournal_HDDet!MaximumDate)
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
    On Error GoTo ErrorCode:

    rsChartAccount.MoveNext
    If rsChartAccount.EOF Then
        rsChartAccount.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:33
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsChartAccount.MovePrevious
    If rsChartAccount.BOF Then
        rsChartAccount.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrint_Click()
    If MsgBox("Print General Ledger for this Account?", vbYesNo + vbQuestion, "Print: " & txtDescription.Text) = vbYes Then
         picPrintopt.Visible = True
         picPrintopt.ZOrder 0
'        rptGeneralLedger.Reset
'        rptGeneralLedger.Formulas(3) = "BEG_DATE = '" & Format(dtFrom, "MMM-DD-YYYY") & "'"
'        rptGeneralLedger.Formulas(4) = "BEGINNING = " & BEGINNING_BALANCE
'        rptGeneralLedger.Formulas(5) = "REPORTDATE = '" & Format(dtFrom, "MMM-DD-YYYY") & " to " & Format(dtTo, "MMM-DD-YYYY") & "'"
'        rptGeneralLedger.ReportTitle = "G E N E R A L  L E D G E R"
'        Dim rsProfile                                       As ADODB.Recordset
'        Set rsProfile = New ADODB.Recordset
'        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
'        If Not (rsProfile.EOF And rsProfile.BOF) Then
'
'            rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
'            rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
'            rptGeneralLedger.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
'
'            'ShowReport "AccountGeneralLedger", "Ledgers", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", "G E N E R A L  L E D G E R", "FROM: " & dtFrom & " TO: " & dtTO, True
'
'        End If
'        If grdAccountsLedger.TextMatrix(1, 2) = "BEGINNING BALANCE" And grdAccountsLedger.Rows = 2 Then
'            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger_Beg.Rpt", "{ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
'        Else
'            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
'            'PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= Cdate(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= Cdate(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
'        End If
'        LogAudit "V", "ACCOUNTS GENERAL LEDGER", txtCode
    End If
End Sub

Private Sub cmdShow_Click()
    FillGrids
End Sub

Private Sub FillGrid()
    Dim rsChartAccounts                                     As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccounts = New ADODB.Recordset
    gconDMIS.CommandTimeout = 500
    Set rsChartAccounts = gconDMIS.Execute("select AcctCode from AMIS_ChartAccount WHERE DESCRIPTION <> '-' order by AcctCode asc")
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
    PRB.Visible = False
    lblPRB.Visible = False
    initGrid
    initMemvars
    StoreMemVars
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
    JOURNALTYPE = ""
    Dim VARVOUCHERNO                                        As String
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
    ElseIf Left(grdAccountsLedger.Text, 3) = "CAJ" Then
        JOURNALTYPE = "CAJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "PAJ" Then
        JOURNALTYPE = "PAJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CLO" Then
        JOURNALTYPE = "CLO"
    ElseIf Left(grdAccountsLedger.Text, 3) = "DRJ" Then
        JOURNALTYPE = "DRJ"
    ElseIf Left(grdAccountsLedger.Text, 3) = "OPB" Then
        JOURNALTYPE = "OPB"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CCM" Then
        JOURNALTYPE = "CCM"
    ElseIf Left(grdAccountsLedger.Text, 3) = "CDM" Then
        JOURNALTYPE = "CDM"
    ElseIf Left(grdAccountsLedger.Text, 3) = "VCM" Then
        JOURNALTYPE = "VCM"
    ElseIf Left(grdAccountsLedger.Text, 3) = "VDM" Then
        JOURNALTYPE = "VDM"
    End If
    'JOURNALTYPE = Left(grdAccountsLedger.Text, 3)
    VARVOUCHERNO = Right(grdAccountsLedger.Text, 6)
    Screen.MousePointer = 11
    On Error Resume Next
    If JOURNALTYPE = "DRJ" Then
        If COMPANY_CODE = "HCA" Then
            Unload frmAMISJournalEntry_DRJ_2
            Call frmAMISJournalEntry_DRJ_2.LOADJOURNAL("DRJ")
            frmAMISJournalEntry_DRJ_2.Show
            Call frmAMISJournalEntry_DRJ_2.StoreSearch(VARVOUCHERNO)
        Else
            Unload frmAMISJournalEntry_DRJ
            Call frmAMISJournalEntry_DRJ.LOADJOURNAL("DRJ")
            frmAMISJournalEntry_DRJ.Show
            Call frmAMISJournalEntry_DRJ.StoreSearch(VARVOUCHERNO)
        End If
    ElseIf JOURNALTYPE = "GJ" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_GJ
        Call frmAMISJournalEntry_GJ.LOADJOURNAL("GJ")
        FormExistsShow frmAMISJournalEntry_GJ
        Call frmAMISJournalEntry_GJ.SearchVoucherNo(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "APJ" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_APJ
        Call frmAMISJournalEntry_APJ.LOADJOURNAL("APJ")
        FormExistsShow frmAMISJournalEntry_APJ
        Call frmAMISJournalEntry_APJ.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "CDJ" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_CDJ
        Call frmAMISJournalEntry_CDJ.LOADJOURNAL("CDJ")
        FormExistsShow frmAMISJournalEntry_CDJ
        Call frmAMISJournalEntry_CDJ.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "SJ" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_SJ
        Call frmAMISJournalEntry_SJ.LOADJOURNAL("SJ")
        FormExistsShow frmAMISJournalEntry_SJ
        Call frmAMISJournalEntry_SJ.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "CRJ" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_CRJ
        Call frmAMISJournalEntry_CRJ.LOADJOURNAL("CRJ")
        FormExistsShow frmAMISJournalEntry_CRJ
        Call frmAMISJournalEntry_CRJ.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "OPB" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_OPB
        Call frmAMISJournalEntry_OPB.LOADJOURNAL("OPB")
        FormExistsShow frmAMISJournalEntry_OPB
        Call frmAMISJournalEntry_OPB.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "CCM" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_CCM
        Call frmAMISJournalEntry_CCM.LOADJOURNAL("CCM")
        FormExistsShow frmAMISJournalEntry_CCM
        Call frmAMISJournalEntry_CCM.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "CDM" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_CDM
        Call frmAMISJournalEntry_CDM.LOADJOURNAL("CDM")
        FormExistsShow frmAMISJournalEntry_CDM
        Call frmAMISJournalEntry_CDM.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "VCM" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_VCM
        Call frmAMISJournalEntry_VCM.LOADJOURNAL("VCM")
        FormExistsShow frmAMISJournalEntry_VCM
        Call frmAMISJournalEntry_VCM.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "VDM" Then
        On Error Resume Next
        Unload frmAMISJournalEntry_VDM
        Call frmAMISJournalEntry_VDM.LOADJOURNAL("VDM")
        FormExistsShow frmAMISJournalEntry_VDM
        Call frmAMISJournalEntry_VDM.StoreSearch(VARVOUCHERNO)
        '    Else
        '        Unload frmAMISJournalEntry
        '        FormExistsShow frmAMISJournalEntry
        '        Call frmAMISJournalEntry.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "PAJ" Then
        On Error Resume Next
        Unload frmAMISJouirnalEntry_PAJE
        Call frmAMISJouirnalEntry_PAJE.LOADJOURNAL("PAJ")
        FormExistsShow frmAMISJouirnalEntry_PAJE
        Call frmAMISJouirnalEntry_PAJE.StoreSearch(VARVOUCHERNO)
    ElseIf JOURNALTYPE = "CAJ" Then
        On Error Resume Next
        Unload frmAMISJouirnalEntry_CAJE
        Call frmAMISJouirnalEntry_CAJE.LOADJOURNAL("CAJ")
        FormExistsShow frmAMISJouirnalEntry_CAJE
        Call frmAMISJouirnalEntry_CAJE.StoreSearch(VARVOUCHERNO)
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
        rsChartAccount.Bookmark = rsFind(rsChartAccount.Clone, "acctcode", lstAccounts.SelectedItem).Bookmark
    Else
        rsChartAccount.Bookmark = rsFind(rsChartAccount.Clone, "acctcode", lstAccounts.SelectedItem.SubItems(1)).Bookmark
    End If
    cleargrid grdAccountsLedger
    initGrid
    StoreMemVars
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
    lstAccounts.ColumnHeaders.Item(1).Width = 2200
End Sub

Private Sub optDescription_Click()
    If TextSearch = "" Then FillGrid2 Else FillSearchGrid2 (TextSearch.Text)
    On Error Resume Next
    TextSearch.SetFocus
    lstAccounts.ColumnHeaders.Item(1).Width = 432760
End Sub

Private Sub optPrintSelected_Click()
        rptGeneralLedger.Reset
        rptGeneralLedger.Formulas(3) = "BEG_DATE = '" & Format(dtFrom, "MMM-DD-YYYY") & "'"
        rptGeneralLedger.Formulas(4) = "BEGINNING = " & BEGINNING_BALANCE
        rptGeneralLedger.Formulas(5) = "REPORTDATE = '" & Format(dtFrom, "MMM-DD-YYYY") & " to " & Format(dtTo, "MMM-DD-YYYY") & "'"
        rptGeneralLedger.ReportTitle = "G E N E R A L  L E D G E R"
        Dim rsProfile                                       As ADODB.Recordset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then

            rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptGeneralLedger.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"

            'ShowReport "AccountGeneralLedger", "Ledgers", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTO) & "," & Month(dtTO) & "," & Day(dtTO) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", "G E N E R A L  L E D G E R", "FROM: " & dtFrom & " TO: " & dtTO, True

        End If
        If grdAccountsLedger.TextMatrix(1, 2) = "BEGINNING BALANCE" And grdAccountsLedger.Rows = 2 Then
            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger_Beg.Rpt", "{ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ")", DMIS_REPORT_Connection, 1
            'PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= Cdate(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= Cdate(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
        End If
        LogAudit "V", "ACCOUNTS GENERAL LEDGER", txtCode
'        picPrintAccount.Visible = False
End Sub

Private Sub optDetailed_KeyDown(KeyCode As Integer, Shift As Integer)
If COMPANY_CODE <> "DJM" Then Exit Sub
If optDetailed.Value = 0 Then Exit Sub
'JULIE  02062013: UPDATED DUE TO CORRECT THE BEGINNING BALANCE PER ACCOUNT
    Dim rsLOAD_APacc                                        As ADODB.Recordset
    Dim rsAllAcct                                           As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xVENDORCODE                                         As String
    Dim xVENDORCODE1                                        As String
    Dim xVENDORNAME                                         As String
    Dim currentacc                                          As String
    Dim prevacct                                            As String
    Dim xsumall                                             As String
    Dim csumaccount                                         As Double
    Dim nextbal                                             As Double
    Dim Credit                                              As Double
    Dim Debit                                               As Double
    Dim ycount                                              As Integer
    Dim shtcnt                                              As Integer
    Dim crrsht                                              As Integer
    Dim ictrsheet                                           As Integer


    PRB.Visible = True
    lblPRB.Visible = True
    lblCurrent(1).Visible = True
    lblOf(0).Visible = True
    lblMax(2).Visible = True
        
    Screen.MousePointer = 11
    xCounter = 10
    
    If Len(Dir(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")) <= 0 Then
        If EXTRACT_FILES(101, "AccountGeneralLedgerAllAccount.xlt") = False Then
            MsgBox "Please Put \Ledgers\AccountGeneralLedgerAllAccount.xlt on " & vbCrLf & CSMS_REPORT_PATH, vbInformation
            Exit Sub
        End If
    End If

        'remove harcode of acctcode after DJM
        'AND CA.ACCTCODE='21-06001-00'
        'AND CA.ACCTCODE='12-02002-00'
        Set rsAllAcct = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType not in ('CCM','CDM','VCM','VDM') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
        If Not (rsAllAcct.BOF And rsAllAcct.EOF) Then
                xCounter = 9
                shtcnt = 1

                Set xlApplication = CreateObject("Excel.Application")
                Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = COMPANY_TIN
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Cells(5, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(5, "A").Font.Bold = True
                
                'HEADER
                xlWorksheet.Cells(8, "A") = "DOC. DATE"
                xlWorksheet.Cells(8, "A").Font.Bold = True
                xlWorksheet.Cells(8, "A").ColumnWidth = "12"
                xlWorksheet.Cells(8, "B") = "REFERENCE"
                xlWorksheet.Cells(8, "B").Font.Bold = True
                xlWorksheet.Cells(8, "B").ColumnWidth = "16"
                xlWorksheet.Cells(8, "C") = "REF. INV."
                xlWorksheet.Cells(8, "C").Font.Bold = True
                xlWorksheet.Cells(8, "C").ColumnWidth = "12"
                xlWorksheet.Cells(8, "E") = "PARTICULARS/REMARKS"
                xlWorksheet.Cells(8, "E").ColumnWidth = "60"
                xlWorksheet.Cells(8, "E").Font.Bold = True
                xlWorksheet.Cells(8, "F") = "REFERENCE NAME"
                xlWorksheet.Cells(8, "F").Font.Bold = True
                xlWorksheet.Cells(8, "F").ColumnWidth = "45"
                xlWorksheet.Cells(8, "G") = "DEBIT"
                xlWorksheet.Cells(8, "G").Font.Bold = True
                xlWorksheet.Cells(8, "H") = "CREDIT"
                xlWorksheet.Cells(8, "H").Font.Bold = True
                xlWorksheet.Cells(8, "I") = "BALANCE"
                xlWorksheet.Cells(8, "I").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & 1
                crrsht = 10
                For ictrsheet = 2 To (crrsht)
                    xlWorkbook.ActiveSheet.Copy After:=xlWorkbook.Sheets(ictrsheet - 1)
                Next ictrsheet


         PRB.Max = rsAllAcct.RecordCount: lblMax(2).Caption = rsAllAcct.RecordCount
         PRB.Value = 0
'        Do While Not rsAllAcct.EOF
'        PRB.Value = PRB.Value + 1
'        lblPRB.Caption = "ACCT #: " & Null2String(rsAllAcct!ACCT_CODE)

        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
        Set rsLOAD_AP = New ADODB.Recordset
        If optPrintbyAccount = True Then
            rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,AP.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.VENDOR_NAME,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO,hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno  WHERE AP.STATUS='P'AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTo.Value & "'  order by ap.vendor_code asc", gconDMIS, adOpenKeyset
        Else
            'remove harcode of acctcode after DJM
            'and det.acct_code='21-06001-00'
            rsLOAD_AP.Open "SELECT distinct det.acct_code,CA.DESCRIPTION FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "'  GROUP BY DET.ACCT_CODE,CA.DESCRIPTION ORDER BY DET.ACCT_CODE", gconDMIS, adOpenKeyset
        End If
        ycount = 0
        If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
            xBALANCE = 0
            rsLOAD_AP.MoveFirst
            Do While Not rsLOAD_AP.EOF
                xAcctCode = Null2String(rsLOAD_AP!ACCT_CODE)
                
                'If xAcctCode = "12-01002-00" Then Stop
                'If xAcctCode = "12-02002-00" Then Stop
                'If xAcctCode = "21-02001-00" Then Stop
                
                Call FORWARDED_BALANCE
                
                'HSM added by kath 7.9.15
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                xlWorksheet.Cells(xCounter, "A").Font.Bold = True
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                xlWorksheet.Cells(xCounter, "B").Font.Bold = True
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
                
                If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                    xlWorksheet.Cells(xCounter, "F") = "0.00"
                    xlWorksheet.Cells(xCounter, "G") = "0.00"
                    xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(FWD_BALANCE)))
                Else
                    xlWorksheet.Cells(xCounter, "G") = "0.00"
                    xlWorksheet.Cells(xCounter, "H") = "0.00"
                    xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(FWD_BALANCE)))
                End If
                
                If COMPANY_CODE = "DJM" Then
                    'OLD DJM QUERY
                    'Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, (CASE WHEN HD.INVOICETYPE = 'CI' THEN HD.InvoiceNo WHEN HD.INVOICETYPE = 'SI' THEN 'SB' + '-' + HD.InvoiceNo WHEN HD.JTYPE = 'APJ' THEN HD.REFERENCENO ELSE HD.InvoiceType + '-' + HD.InvoiceNo END) AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit,HD.Remarks,  " & _
                    '                                    "Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode) " & _
                    '                                    "INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType not in ('CCM','CDM','VCM','VDM','DRJ') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")

                    'NEW QUERY FOR VIEWING OF FINANCING AS VENDOR NAME
                    '01-05-2016
                     Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDATE, HD.JTYPE + '-' + HD.VOUCHERNO AS VOUCHERNO, (CASE WHEN HD.INVOICETYPE = 'CI' THEN HD.INVOICENO WHEN HD.INVOICETYPE = 'SI' THEN 'SB' + '-' + HD.INVOICENO WHEN HD.JTYPE = 'APJ' THEN HD.REFERENCENO ELSE HD.INVOICETYPE + '-' + HD.INVOICENO END) AS INVOICENO,HD.INVOICENO AS RRNO,HD.CUSTOMERCODE, " & _
                                                        "(CASE WHEN HD.INVOICENO IS NULL THEN VENDOR.CODE " & _
                                                        "WHEN DET.ACCT_CODE = '21-02000-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02001-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02002-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02003-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02004-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02005-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_CODE FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) ELSE " & _
                                                        "VENDOR.Code END) AS CODE, " & _
                                                        "(CASE WHEN HD.INVOICENO IS NULL THEN VENDOR.NAMEOFVENDOR " & _
                                                        "WHEN DET.ACCT_CODE = '21-02000-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02001-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02002-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02003-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02004-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) " & _
                                                        "WHEN DET.ACCT_CODE = '21-02005-00' AND DET.JTYPE='APJ' THEN (SELECT VENDOR_NAME FROM AMIS_AP WHERE INVOICENO=HD.REFERENCENO) ELSE " & _
                                                        "VENDOR.NAMEOFVENDOR " & _
                                                        "END) AS NAMEOFVENDOR, " & _
                                                        "HD.STATUS,CUSTOMER.CUSTNAME,DET.ACCT_CODE, DET.DEBIT, DET.CREDIT,HD.REMARKS,  DET.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { OJ (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_VENDOR VENDOR ON HD.VENDORCODE = VENDOR.CODE) INNER JOIN  AMIS_JOURNAL_DET DET ON HD.JTYPE = DET.JTYPE AND HD.VOUCHERNO = DET.VOUCHERNO) LEFT OUTER JOIN ALL_CUSTMASTER_AMIS CUSTOMER ON HD.CUSTOMERCODE = CUSTOMER.CUSTCODE) INNER JOIN AMIS_CHARTACCOUNT CA ON DET.ACCT_CODE = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON DET.ENTITY = AE.COMPLET_CODE} WHERE HD.STATUS = 'P' AND HD.JTYPE NOT IN ('CCM','CDM','VCM','VDM','DRJ') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND " & _
                                                        "DET.ACCT_CODE = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY DET.ACCT_CODE ASC,HD.JDATE ASC,DET.ID ASC")

                Else
                    Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit,HD.Remarks,  " & _
                                                        "Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode) " & _
                                                        "INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType not in ('CCM','CDM','VCM','VDM') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
                End If
                
                If Not (rsLOAD_APacc.EOF And rsLOAD_APacc.BOF) Then
                    xBALANCE = FWD_BALANCE: TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
                    Do While Not rsLOAD_APacc.EOF
                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!Code)
                            xVENDORNAME = Null2String(rsLOAD_APacc!nameofvendor)
                        ElseIf Left(rsLOAD_APacc!VOUCHERNO, 3) = "DRJ" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        Else
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        End If
                        
                        If Len(rsLOAD_APacc!VOUCHERNO) = 10 Then
                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 3)
                        Else
                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 2)
                        End If
                        
                        xCounter = xCounter + 1
                            If xCounter > 32760 Then
                                shtcnt = shtcnt + 1
                                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                                xlWorksheet.Cells(1, "A").Font.Bold = True
                                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                                xlWorksheet.Cells(2, "A").Font.Bold = True
                                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                                xlWorksheet.Cells(3, "A").Font.Bold = True
                                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                                xlWorksheet.Cells(4, "A").Font.Bold = True

                                'header
                                xlWorksheet.Cells(8, "A") = "DOC. DATE"
                                xlWorksheet.Cells(8, "A").Font.Bold = True
                                xlWorksheet.Cells(8, "A").ColumnWidth = "12"
                                xlWorksheet.Cells(8, "B") = "REFERENCE"
                                xlWorksheet.Cells(8, "B").Font.Bold = True
                                xlWorksheet.Cells(8, "B").ColumnWidth = "16"
                                xlWorksheet.Cells(8, "C") = "REF. INV."
                                xlWorksheet.Cells(8, "C").Font.Bold = True
                                xlWorksheet.Cells(8, "C").ColumnWidth = "12"
                                xlWorksheet.Cells(8, "E") = "PARTICULARS/REMARKS"
                                xlWorksheet.Cells(8, "E").ColumnWidth = "60"
                                xlWorksheet.Cells(8, "E").Font.Bold = True
                                xlWorksheet.Cells(8, "F") = "REFERENCE NAME"
                                xlWorksheet.Cells(8, "F").Font.Bold = True
                                xlWorksheet.Cells(8, "F").ColumnWidth = "45"
                                xlWorksheet.Cells(8, "G") = "DEBIT"
                                xlWorksheet.Cells(8, "G").Font.Bold = True
                                xlWorksheet.Cells(8, "H") = "CREDIT"
                                xlWorksheet.Cells(8, "H").Font.Bold = True
                                xlWorksheet.Cells(8, "I") = "BALANCE"
                                xlWorksheet.Cells(8, "I").Font.Bold = True
                                xlWorksheet.Name = "Sheet-" & shtcnt
                                xCounter = 9
                            End If
                            
                        xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_APacc!JDATE))), "mm/dd/yyyy")
                        xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_APacc!VOUCHERNO)))
                        xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_APacc!INVOICENO)))
                        
                        'If xlWorksheet.Cells(xCounter, "C") = "OR000253" Then Stop
                        
                        If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                            xlWorksheet.Cells(xCounter, "D") = (Trim(Null2String(rsLOAD_APacc!remarks)))
                        Else
                            xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!remarks)))
                            xlWorksheet.Cells(xCounter, "E").WrapText = False
                        End If
                        
                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                            Else
                                xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                            End If
                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
                            Else
                                xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
                            End If
                       Else
                        xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                       End If

                        If NumericVal(rsLOAD_APacc!Credit) <> 0 Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "F") = (Trim("0.00"))
                                xlWorksheet.Cells(xCounter, "G") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
                            Else
                                xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
                            End If

                            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_APacc!Credit)), 2))
                            Credit = Credit + TOTAL_CREDIT
                                                        
                            'OLD
                            '01-08-16
                            'xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
                            If COMPANY_CODE = "DJM" Then
                                If (GetHeaders(xAcctCode) = "12" And GetTitleCode(xAcctCode) = "02") Or GetHeaders(xAcctCode) = "21" Or GetHeaders(xAcctCode) = "31" Or GetHeaders(xAcctCode) = "41" Or GetHeaders(xAcctCode) = "71" Then
                                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Credit)), 2))
                                Else
                                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
                                End If
                            Else
                                xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
                            End If
                            
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
                            Else
                                xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
                            End If
                            If COMPANY_CODE = "HNE" Then
                                xlWorksheet.Cells(xCounter, "I") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                                xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                            End If
                            
                        Else
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "F") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Debit))))
                                xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
                            Else
                                xlWorksheet.Cells(xCounter, "G") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Debit))))
                                xlWorksheet.Cells(xCounter, "H") = (Trim("0.00"))
                            End If
                            
                            TOTAL_DEBIT = Trim(ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_APacc!Debit)), 2)))
                            Debit = Debit + TOTAL_DEBIT
                            
                            'OLD
                            '01-08-16
                            'xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
                            If COMPANY_CODE = "DJM" Then
                                If (GetHeaders(xAcctCode) = "12" And GetTitleCode(xAcctCode) = "02") Or GetHeaders(xAcctCode) = "21" Or GetHeaders(xAcctCode) = "31" Or GetHeaders(xAcctCode) = "41" Or GetHeaders(xAcctCode) = "71" Then
                                    xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Debit)), 2))
                                Else
                                    xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
                                End If
                            Else
                                xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
                            End If
                            
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
                            Else
                                xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
                            End If
                            
                            If COMPANY_CODE = "HNE" Then
                                'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "I") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                                xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                            End If
                            
                        End If
                        rsLOAD_APacc.MoveNext
                        If rsLOAD_APacc.EOF = True Then
                            xCounter = xCounter + 1
                            
                            If xCounter > 32760 Then '3276A0 Then
                                shtcnt = shtcnt + 1
                                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                                xlWorksheet.Cells(1, "A").Font.Bold = True
                                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                                xlWorksheet.Cells(2, "A").Font.Bold = True
                                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                                xlWorksheet.Cells(3, "A").Font.Bold = True
                                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                                xlWorksheet.Cells(4, "A").Font.Bold = True
                                xlWorksheet.Name = "Sheet-" & shtcnt
                                xCounter = 9
                            End If
                            
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                xlWorksheet.Cells(xCounter, "F") = TOTAL_DEBIT
                                xlWorksheet.Cells(xCounter, "F").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "G") = TOTAL_CREDIT
                                xlWorksheet.Cells(xCounter, "G").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "H") = xBALANCE
                                xlWorksheet.Cells(xCounter, "H").Font.Bold = True
                            Else
                                xlWorksheet.Cells(xCounter, "G") = ToDoubleNumber(TOTAL_DEBIT)
                                xlWorksheet.Cells(xCounter, "G").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "H") = ToDoubleNumber(TOTAL_CREDIT)
                                xlWorksheet.Cells(xCounter, "H").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "I") = ToDoubleNumber(xBALANCE)
                                xlWorksheet.Cells(xCounter, "I").Font.Bold = True
                            End If
                            
                        End If
                        
                        If PRB.Value = PRB.Max Then
                            PRB.Enabled = False
                        Else
                            PRB.Value = PRB.Value + 1: lblCurrent(1).Caption = PRB.Value
                            lblPRB.Caption = "ACCT #: " & Null2String(rsLOAD_AP!ACCT_CODE)
                        End If
                    Loop
                            
                End If
                rsLOAD_AP.MoveNext
                ycount = 1
                DoEvents
            Loop
        End If

            xCounter = xCounter + 1
            If xCounter > 32760 Then
                shtcnt = shtcnt + 1
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & shtcnt
                xCounter = 9
            End If

            rsAllAcct.MoveNext

            xCounter = xCounter + 1
            If xCounter > 32760 Then
                shtcnt = shtcnt + 1
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & shtcnt
                xCounter = 10
            End If

        If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
            xlWorksheet.Cells(xCounter, "E") = "TOTAL"
            xlWorksheet.Cells(xCounter, "E").Font.Bold = True
            xlWorksheet.Cells(xCounter, "F") = ToDoubleNumber(Debit)
            xlWorksheet.Cells(xCounter, "F").Font.Bold = True
            xlWorksheet.Cells(xCounter, "G") = ToDoubleNumber(Credit)
            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
            xlWorksheet.Cells(xCounter, "H") = ToDoubleNumber(Debit - Credit)
            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
        Else
            xlWorksheet.Cells(xCounter, "F") = "TOTAL"
            xlWorksheet.Cells(xCounter, "F").Font.Bold = True
            xlWorksheet.Cells(xCounter, "G") = ToDoubleNumber(Debit)
            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
            xlWorksheet.Cells(xCounter, "H") = ToDoubleNumber(Credit)
            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
            xlWorksheet.Cells(xCounter, "I") = ToDoubleNumber(Debit - Credit)
            xlWorksheet.Cells(xCounter, "I").Font.Bold = True
        End If
        
        xlWorksheet.Columns("F:I").AutoFit
        xlApplication.Visible = True
        xlWorkbook.Sheets(1).Activate
        
        PRB.Visible = True
        lblPRB.Visible = True
        
        Set xlApplication = Nothing
        Set xlWorkbook = Nothing
        Set xlWorksheet = Nothing
        Set rsLOAD_AP = Nothing
        Set rsAllAcct = Nothing
        Screen.MousePointer = 0
    End If
        
        PRB.Visible = False
        lblPRB.Visible = False
        lblCurrent(1).Visible = False
        lblOf(0).Visible = False
        lblMax(2).Visible = False
        
        optPrintbyAccount.Value = 0
        optPrintAll.Value = 0
        picPrintopt.Visible = False
        Frame4.ZOrder 1
        
        '01202016
        'START OF AUTOMATICALLY SAVE XLS FILE TO CSV
        If MsgBox("Do you want to convert excel file to CSV format now?", vbYesNo, "Question") = vbYes Then
            Dim ws As Excel.Worksheet
            Dim path As String
            
            path = AMIS_REPORT_PATH & "LEDGERS\" & "ALL_ACCOUNT_GL_" & Date$ & "\"
            If Dir(path, vbDirectory) = "" Then
                MkDir (path)
            End If
            
            For Each ws In Excel.Worksheets
                If ActiveWorkbook.Worksheets(ws.Name).Cells(10, "A") = "" Then GoTo YeahNxt
                ws.Copy
                ActiveWorkbook.SaveAs FileName:=(path) & ws.Name & ".csv", FileFormat:=xlCSV, CreateBackup:=False
                ActiveWorkbook.Close False
YeahNxt:
            Next
            
            MsgBox "CSV files saved to:" & vbCrLf & path, vbInformation, "Information"
            Shell "explorer " & path
        End If

End Sub

Private Sub optPrintAll_KeyDown(KeyCode As Integer, Shift As Integer)

If COMPANY_CODE = "DJM" Then Frame4.ZOrder 0: Exit Sub
 'JULIE  02062013: UPDATED DUE TO CORRECT THE BEGINNING BALANCE PER ACCOUNT
    Dim rsLOAD_APacc                                        As ADODB.Recordset
    Dim rsAllAcct                                           As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xVENDORCODE                                         As String
    Dim xVENDORCODE1                                        As String
    Dim xVENDORNAME                                         As String
    Dim currentacc                                          As String
    Dim prevacct                                            As String
    Dim xsumall                                             As String
    Dim csumaccount                                         As Double
    Dim nextbal                                             As Double
    Dim Credit                                              As Double
    Dim Debit                                               As Double
    Dim ycount                                              As Integer
    Dim shtcnt                                              As Integer
    Dim crrsht                                              As Integer
    Dim ictrsheet                                           As Integer


    PRB.Visible = True
    lblPRB.Visible = True
    lblCurrent(1).Visible = True
    lblOf(0).Visible = True
    lblMax(2).Visible = True
        
    Screen.MousePointer = 11
    xCounter = 10
    
    If Len(Dir(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")) <= 0 Then
        If EXTRACT_FILES(101, "AccountGeneralLedgerAllAccount.xlt") = False Then
            MsgBox "Please Put \Ledgers\AccountGeneralLedgerAllAccount.xlt on " & vbCrLf & CSMS_REPORT_PATH, vbInformation
            Exit Sub
        End If
    End If


        Set rsAllAcct = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType not in ('CCM','CDM','VCM','VDM') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
        If Not (rsAllAcct.BOF And rsAllAcct.EOF) Then
                xCounter = 9
                shtcnt = 1

                Set xlApplication = CreateObject("Excel.Application")
                Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & 1
                crrsht = 10
                For ictrsheet = 2 To (crrsht)
                    xlWorkbook.ActiveSheet.Copy After:=xlWorkbook.Sheets(ictrsheet - 1)
                Next ictrsheet


         PRB.Max = rsAllAcct.RecordCount: lblMax(2).Caption = rsAllAcct.RecordCount
         PRB.Value = 0
'        Do While Not rsAllAcct.EOF
'        PRB.Value = PRB.Value + 1
'        lblPRB.Caption = "ACCT #: " & Null2String(rsAllAcct!ACCT_CODE)

        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
        Set rsLOAD_AP = New ADODB.Recordset
        If optPrintbyAccount = True Then
            rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,AP.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.VENDOR_NAME,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO,hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno  WHERE AP.STATUS='P'AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTo.Value & "'  order by ap.vendor_code asc", gconDMIS, adOpenKeyset
        Else
            rsLOAD_AP.Open "SELECT distinct det.acct_code,CA.DESCRIPTION FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' GROUP BY DET.ACCT_CODE,CA.DESCRIPTION ORDER BY DET.ACCT_CODE", gconDMIS, adOpenKeyset
        End If
        ycount = 0
        If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
            xBALANCE = 0
            rsLOAD_AP.MoveFirst
            Do While Not rsLOAD_AP.EOF
                xAcctCode = Null2String(rsLOAD_AP!ACCT_CODE)
                Call FORWARDED_BALANCE
        'HSM added by kath 7.9.15
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                xlWorksheet.Cells(xCounter, "A").Font.Bold = True
                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                xlWorksheet.Cells(xCounter, "B").Font.Bold = True
                xCounter = xCounter + 1
                xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
                If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                    xlWorksheet.Cells(xCounter, "F") = "0.00"
                    xlWorksheet.Cells(xCounter, "G") = "0.00"
                    xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(FWD_BALANCE)))
                Else
                    xlWorksheet.Cells(xCounter, "G") = "0.00"
                    xlWorksheet.Cells(xCounter, "H") = "0.00"
                    xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(FWD_BALANCE)))
                End If
                

                Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit,HD.Remarks,  " & _
                                                    "Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode) " & _
                                                    "INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType not in ('CCM','CDM','VCM','VDM') AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
                If Not (rsLOAD_APacc.EOF And rsLOAD_APacc.BOF) Then
                    xBALANCE = FWD_BALANCE: TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
                    Do While Not rsLOAD_APacc.EOF
                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!Code)
                            xVENDORNAME = Null2String(rsLOAD_APacc!nameofvendor)
                        ElseIf Left(rsLOAD_APacc!VOUCHERNO, 3) = "DRJ" Then
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        Else
                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
                        End If
                        If Len(rsLOAD_APacc!VOUCHERNO) = 10 Then
                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 3)
                        Else
                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 2)
                        End If
                        xCounter = xCounter + 1
                            If xCounter > 32760 Then
                                shtcnt = shtcnt + 1
                                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                                xlWorksheet.Cells(1, "A").Font.Bold = True
                                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                                xlWorksheet.Cells(2, "A").Font.Bold = True
                                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                                xlWorksheet.Cells(3, "A").Font.Bold = True
                                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                                xlWorksheet.Cells(4, "A").Font.Bold = True
                                xlWorksheet.Name = "Sheet-" & shtcnt
                                xCounter = 9
                            End If
                        xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_APacc!JDATE))), "mm/dd/yyyy")
                        xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_APacc!VOUCHERNO)))
                        xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_APacc!INVOICENO)))
                        If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                            xlWorksheet.Cells(xCounter, "D") = (Trim(Null2String(rsLOAD_APacc!remarks)))
                        Else
                            xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!remarks)))
                        End If
                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                                xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                            Else
                                xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                            End If
                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                                xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
                            Else
                                xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
                            End If
                       Else
                          xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
                        End If

                        If NumericVal(rsLOAD_APacc!Credit) <> 0 Then
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                            'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "F") = (Trim("0.00"))
                                xlWorksheet.Cells(xCounter, "G") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
                            Else
                                xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
                            End If

                            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_APacc!Credit)), 2))
                            Credit = Credit + TOTAL_CREDIT
                            xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                            'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
                            Else
                                xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
                            End If
                            If COMPANY_CODE = "HNE" Then
                                xlWorksheet.Cells(xCounter, "I") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                                xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                            End If
                            
                        Else
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                            'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!Debit)))
                                xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
                            Else
                                xlWorksheet.Cells(xCounter, "G") = (Trim(Null2String(rsLOAD_APacc!Debit)))
                                xlWorksheet.Cells(xCounter, "H") = (Trim("0.00"))
                            End If
                            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_APacc!Debit)), 2))
                            Debit = Debit + TOTAL_DEBIT
                            xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLB" Or COMPANY_CODE = "HBC" Or COMPANY_CODE = "HCR" Then
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
                            Else
                                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
                            End If
                            
                            If COMPANY_CODE = "HNE" Then
                            'added by kath 05.19.15
                                xlWorksheet.Cells(xCounter, "I") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
                                xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
                            End If
                            
                        End If
                        rsLOAD_APacc.MoveNext
                        If rsLOAD_APacc.EOF = True Then
                            xCounter = xCounter + 1
                            If xCounter > 32760 Then
                                shtcnt = shtcnt + 1
                                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                                xlWorksheet.Cells(1, "A").Font.Bold = True
                                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                                xlWorksheet.Cells(2, "A").Font.Bold = True
                                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                                xlWorksheet.Cells(3, "A").Font.Bold = True
                                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                                xlWorksheet.Cells(4, "A").Font.Bold = True
                                xlWorksheet.Name = "Sheet-" & shtcnt
                                xCounter = 9
                            End If
                            If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
                                xlWorksheet.Cells(xCounter, "F") = TOTAL_DEBIT
                                xlWorksheet.Cells(xCounter, "F").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "G") = TOTAL_CREDIT
                                xlWorksheet.Cells(xCounter, "G").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "H") = xBALANCE
                                xlWorksheet.Cells(xCounter, "H").Font.Bold = True
                            Else
                                xlWorksheet.Cells(xCounter, "G") = TOTAL_DEBIT
                                xlWorksheet.Cells(xCounter, "G").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "H") = TOTAL_CREDIT
                                xlWorksheet.Cells(xCounter, "H").Font.Bold = True
                                xlWorksheet.Cells(xCounter, "I") = xBALANCE
                                xlWorksheet.Cells(xCounter, "I").Font.Bold = True
                            End If
                        End If
                            If PRB.Value = PRB.Max Then
                            PRB.Enabled = False
                            Else
                            PRB.Value = PRB.Value + 1: lblCurrent(1).Caption = PRB.Value
                            lblPRB.Caption = "ACCT #: " & Null2String(rsLOAD_AP!ACCT_CODE)
                            End If
                    Loop
                            
                End If
                rsLOAD_AP.MoveNext
                ycount = 1
                DoEvents
            Loop
        End If

            xCounter = xCounter + 1
            If xCounter > 32760 Then
                shtcnt = shtcnt + 1
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & shtcnt
                xCounter = 9
            End If

            rsAllAcct.MoveNext
'        Loop

            xCounter = xCounter + 1
            If xCounter > 32760 Then
                shtcnt = shtcnt + 1
                Set xlWorksheet = xlWorkbook.Worksheets(shtcnt)
                xlWorksheet.Cells(1, "A") = COMPANY_NAME
                xlWorksheet.Cells(1, "A").Font.Bold = True
                xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
                xlWorksheet.Cells(2, "A").Font.Bold = True
                xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
                xlWorksheet.Cells(3, "A").Font.Bold = True
                xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
                xlWorksheet.Cells(4, "A").Font.Bold = True
                xlWorksheet.Name = "Sheet-" & shtcnt
                xCounter = 10
            End If

        If COMPANY_CODE = "HNE" Or COMPANY_CODE = "HSM" Then
            xlWorksheet.Cells(xCounter, "E") = "TOTAL"
            xlWorksheet.Cells(xCounter, "E").Font.Bold = True
            xlWorksheet.Cells(xCounter, "F") = Debit
            xlWorksheet.Cells(xCounter, "F").Font.Bold = True
            xlWorksheet.Cells(xCounter, "G") = Credit
            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
            xlWorksheet.Cells(xCounter, "H") = Debit - Credit
            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
        Else
            xlWorksheet.Cells(xCounter, "F") = "TOTAL"
            xlWorksheet.Cells(xCounter, "F").Font.Bold = True
            xlWorksheet.Cells(xCounter, "G") = Debit
            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
            xlWorksheet.Cells(xCounter, "H") = Credit
            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
            xlWorksheet.Cells(xCounter, "I") = Debit - Credit
            xlWorksheet.Cells(xCounter, "I").Font.Bold = True
        End If
        xlApplication.Visible = True
        xlWorkbook.Sheets(1).Activate
        
        PRB.Visible = True
        lblPRB.Visible = True
        
        Set xlApplication = Nothing
        Set xlWorkbook = Nothing
        Set xlWorksheet = Nothing
        Set rsLOAD_AP = Nothing
        Set rsAllAcct = Nothing
        Screen.MousePointer = 0
    End If
        
        PRB.Visible = False
        lblPRB.Visible = False
        lblCurrent(1).Visible = False
        lblOf(0).Visible = False
        lblMax(2).Visible = False
        
        optPrintbyAccount.Value = 0
        optPrintAll.Value = 0
        picPrintopt.Visible = False
        picPrintopt.ZOrder 1

''' JULIE  02062013: UPDATED DUE TO CORRECT THE BEGINNING BALANCE PER ACCOUNT
''    Dim xVOUCHERNO                                          As String
''    Dim xVENDORCODE                                         As String
''    Dim xVENDORCODE1                                        As String
''    Dim xVENDORNAME                                         As String
''    Dim csumaccount                                         As Double
''    Dim currentacc                                          As String
''    Dim prevacct                                            As String
''    Dim xsumall                                             As String
''    Dim ycount                                              As Integer
''    Dim rsLOAD_APacc                                        As ADODB.Recordset
''    Dim rsAllAcct                                           As ADODB.Recordset
''    Dim nextbal                                             As Double
''    Dim Credit                                              As Double
''    Dim Debit                                               As Double
''
'''PRB.Visible = True
'''lblPRB.Visible = True
''
''    Screen.MousePointer = 11
''    xCounter = 10
''    If Len(Dir(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")) <= 0 Then
''        If EXTRACT_FILES(101, "AccountGeneralLedgerAllAccount.xlt") = False Then
''            MsgBox "Please Put \Ledgers\AccountGeneralLedgerAllAccount.xlt on " & vbCrLf & CSMS_REPORT_PATH, vbInformation
''            Exit Sub
''        End If
''    End If
''
''    Set xlApplication = CreateObject("Excel.Application")
''    Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "\Ledgers\AccountGeneralLedgerAllAccount.xlt")
''    Set xlWorksheet = xlWorkbook.Worksheets(1)
''
''    If COMPANY_CODE <> "HSM" Then
''        xlWorksheet.Cells(1, "A") = COMPANY_NAME
''        xlWorksheet.Cells(1, "A").Font.Bold = True
''        xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
''        xlWorksheet.Cells(2, "A").Font.Bold = True
''        xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
''        xlWorksheet.Cells(3, "A").Font.Bold = True
''        xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
''        xlWorksheet.Cells(4, "A").Font.Bold = True
''
'''         PRB.Max = rsAllAcct.RecordCount
'''         PRB.Value = 0
'''         Do While Not rsAllAcct.EOF
'''         PRB.Value = PRB.Value + 1
'''         lblPRB.Caption = "ACCT #: " & Null2String(rsAllAcct!ACCT_CODE)
'''         Loop
''
''        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
''        Set rsLOAD_AP = New ADODB.Recordset
''        If optPrintbyAccount = True Then
''            rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,AP.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.VENDOR_NAME,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO,hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno  WHERE AP.STATUS='P'AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTo.Value & "'  order by ap.vendor_code asc", gconDMIS, adOpenKeyset
''        Else
''            rsLOAD_AP.Open "SELECT distinct det.acct_code,CA.DESCRIPTION FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' GROUP BY DET.ACCT_CODE,CA.DESCRIPTION", gconDMIS, adOpenKeyset
''        End If
''        ycount = 0
''        If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
''            xBALANCE = 0
''            'xlApplication.Visible = True
''            rsLOAD_AP.MoveFirst
''            Do While Not rsLOAD_AP.EOF
''                xAcctCode = Null2String(rsLOAD_AP!ACCT_CODE)
''                Call FORWARDED_BALANCE
''                xCounter = xCounter + 1
''                xlWorksheet.Cells(xCounter, "A") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
''                xlWorksheet.Cells(xCounter, "A").Font.Bold = True
''                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
''                xlWorksheet.Cells(xCounter, "B").Font.Bold = True
''                xCounter = xCounter + 1
''                xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
''                xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
''                xlWorksheet.Cells(xCounter, "F") = "0.00"
''                xlWorksheet.Cells(xCounter, "G") = "0.00"
''                xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(FWD_BALANCE)))
''
''                If COMPANY_CODE <> "DGI" Then
''                Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.REMARKS AS REMARKS, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
''                Else
''                Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.REMARKS AS REMARKS, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
''                End If
''                If Not (rsLOAD_APacc.EOF And rsLOAD_APacc.BOF) Then
''                    xBALANCE = FWD_BALANCE: TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
''                    Do While Not rsLOAD_APacc.EOF
''                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!Code)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!nameofvendor)
''                        ElseIf Left(rsLOAD_APacc!VOUCHERNO, 3) = "DRJ" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        Else
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        End If
''                        If Len(rsLOAD_APacc!VOUCHERNO) = 10 Then
''                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 3)
''                        Else
''                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 2)
''                        End If
''                        xCounter = xCounter + 1
''                        xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_APacc!JDATE))), "mm/dd/yyyy")
''                        xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_APacc!VOUCHERNO)))
''                        xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_APacc!INVOICENO)))
''                        xlWorksheet.Cells(xCounter, "D") = (Trim(Null2String(rsLOAD_APacc!remarks)))
''                        If COMPANY_CODE = "HNE" Then
''                            xlWorksheet.Cells(xCounter, "I") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
''                            xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_APacc!DESCRIPTION)))
''                        End If
''                        On Error Resume Next
''                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
''                            xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
''                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
''                            xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
''                       Else
''                          xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
''                       On Error Resume Next
''                        End If
''
''                        If NumericVal(rsLOAD_APacc!Credit) <> 0 Then
''                            xlWorksheet.Cells(xCounter, "F") = (Trim("0.00"))
''                            xlWorksheet.Cells(xCounter, "G") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
''
''                            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_APacc!Credit)), 2))
''                            Credit = Credit + TOTAL_CREDIT
''                            xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
''                            xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
''
''                        Else
''                            xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!Debit)))
''                            xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
''                            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_APacc!Debit)), 2))
''                            Debit = Debit + TOTAL_DEBIT
''                            xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
''                            xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(xBALANCE)))
''                        End If
''                        rsLOAD_APacc.MoveNext
''                        If rsLOAD_APacc.EOF = True Then
''                            xCounter = xCounter + 1
''                            xlWorksheet.Cells(xCounter, "F") = TOTAL_DEBIT
''                            xlWorksheet.Cells(xCounter, "F").Font.Bold = True
''                            xlWorksheet.Cells(xCounter, "G") = TOTAL_CREDIT
''                            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
''                            xlWorksheet.Cells(xCounter, "H") = xBALANCE
''                            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
''                        End If
''
''                    Loop
''                End If
''                rsLOAD_AP.MoveNext
''                ycount = 1
''                DoEvents
''            Loop
''        End If
''        xCounter = xCounter + 2
''        xlWorksheet.Cells(xCounter, "E") = "TOTAL"
''        xlWorksheet.Cells(xCounter, "E").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "F") = Debit
''        xlWorksheet.Cells(xCounter, "F").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "G") = Credit
''        xlWorksheet.Cells(xCounter, "G").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "H") = Debit - Credit
''        xlWorksheet.Cells(xCounter, "H").Font.Bold = True
''        xlApplication.Visible = True
''        Set xlApplication = Nothing
''        Set xlWorkbook = Nothing
''        Set xlWorksheet = Nothing
''        Set rsLOAD_AP = Nothing
''        Screen.MousePointer = 0
''
''    'SJR 7/1/14
''    'Remarks for details
''    'Else for HSM only
''    Else
''
''        xlWorksheet.Cells(1, "A") = COMPANY_NAME
''        xlWorksheet.Cells(1, "A").Font.Bold = True
''        xlWorksheet.Cells(2, "A") = COMPANY_ADDRESS
''        xlWorksheet.Cells(2, "A").Font.Bold = True
''        xlWorksheet.Cells(3, "A") = "From: " & Format(dtFrom.Value, "mm/dd/yyyy") & " To: " & Format(dtTo.Value, "mm/dd/yyyy")
''        xlWorksheet.Cells(3, "A").Font.Bold = True
''        xlWorksheet.Cells(4, "A") = "GENERAL LEDGER"
''        xlWorksheet.Cells(4, "A").Font.Bold = True
''
'''         PRB.Max = rsAllAcct.RecordCount
'''         PRB.Value = 0
'''         Do While Not rsAllAcct.EOF
'''         PRB.Value = PRB.Value + 1
'''         lblPRB.Caption = "ACCT #: " & Null2String(rsAllAcct!ACCT_CODE)
'''         Loop
''
''        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0: xBALANCE = 0
''        Set rsLOAD_AP = New ADODB.Recordset
''        If optPrintbyAccount = True Then
''            rsLOAD_AP.Open "SELECT ap.JDATE,ap.VOUCHERNO AS APVOUCHERNO,AP.INVOICENO AS INVOICE,ap.AMOUNT2PAY,ap.AMOUNTPAID,ap.INVOICENO,ap.INVOICETYPE,ap.VENDOR_CODE,AP.VENDOR_NAME,AP.ACCT_CODE,RIGHT(ap.VOUCHERNO,6) AS VOUCHERNO,hd.remarks FROM AMIS_AP AP INNER JOIN AMIS_CHARTACCOUNT AC ON AP.ACCT_CODE=AC.ACCTCODE inner join amis_journal_hd hd on ap.voucherno=hd.jtype+'-'+hd.voucherno  WHERE AP.STATUS='P'AND ap.Jdate >= '" & dtFrom.Value & "'and ap.Jdate <= '" & dtTo.Value & "'  order by ap.vendor_code asc", gconDMIS, adOpenKeyset
''        Else
''            rsLOAD_AP.Open "SELECT distinct det.acct_code,CA.DESCRIPTION FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' GROUP BY DET.ACCT_CODE,CA.DESCRIPTION", gconDMIS, adOpenKeyset
''        End If
''        ycount = 0
''        If Not rsLOAD_AP.EOF And Not rsLOAD_AP.BOF Then
''            xBALANCE = 0
''            'xlApplication.Visible = True
''            rsLOAD_AP.MoveFirst
''            Do While Not rsLOAD_AP.EOF
''                xAcctCode = Null2String(rsLOAD_AP!ACCT_CODE)
''                Call FORWARDED_BALANCE
''                xCounter = xCounter + 1
''                xlWorksheet.Cells(xCounter, "A") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
''                xlWorksheet.Cells(xCounter, "A").Font.Bold = True
''                xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_AP!DESCRIPTION)))
''                xlWorksheet.Cells(xCounter, "B").Font.Bold = True
''                xCounter = xCounter + 1
''                xlWorksheet.Cells(xCounter, "A") = Format(dtFrom.Value, "mm/dd/yyyy")
''                xlWorksheet.Cells(xCounter, "B") = "FWD BALANCE"
''                xlWorksheet.Cells(xCounter, "G") = "0.00"
''                xlWorksheet.Cells(xCounter, "H") = "0.00"
''                xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(FWD_BALANCE)))
''
''                If COMPANY_CODE <> "DGI" Then
''                Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.REMARKS AS REMARKS, HD.InvoiceType + '-' + HD.InvoiceNo AS INVOICENO, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME, det.adj_remarks " & _
''                                                    " FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
''                Else
''                Set rsLOAD_APacc = gconDMIS.Execute("SELECT HD.ENTITY_CLASS,HD.JDate, HD.JType + '-' + HD.VoucherNo AS VOUCHERNO, HD.REMARKS AS REMARKS, HD.CUSTOMERCODE,VENDOR.CODE, HD.Status,Customer.CustName,Vendor.NameofVendor,Det.Acct_Code, Det.Debit, Det.Credit, Det.ID,CA.ACCTCODE, CA.HEADERCODE, CA.DESCRIPTION,AE.ACCOUNTNAME FROM { oj (((( AMIS_JOURNAL_HD HD LEFT OUTER JOIN ALL_Vendor Vendor ON HD.VendorCode = Vendor.Code) INNER JOIN  AMIS_JOURNAL_DET Det ON HD.JType = Det.JType AND HD.VoucherNo = Det.VoucherNo) LEFT OUTER JOIN ALL_CustMaster_AMIS Customer ON HD.CustomerCode = Customer.CustCode)INNER JOIN AMIS_ChartAccount CA ON Det.Acct_Code = CA.ACCTCODE) LEFT OUTER JOIN ALL_ENTITY AE ON Det.Entity = AE.COMPLET_CODE} WHERE HD.Status = 'P' AND HD.JType <> 'CCM' AND HD.Jdate >= '" & dtFrom.Value & "' and HD.Jdate <= '" & dtTo.Value & "' AND det.acct_code = '" & Null2String(rsLOAD_AP!ACCT_CODE) & "' ORDER BY Det.Acct_Code ASC,HD.JDate ASC,Det.ID ASC")
''                End If
''                If Not (rsLOAD_APacc.EOF And rsLOAD_APacc.BOF) Then
''                    xBALANCE = FWD_BALANCE: TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
''                    Do While Not rsLOAD_APacc.EOF
''                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!Code)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!nameofvendor)
''                        ElseIf Left(rsLOAD_APacc!VOUCHERNO, 3) = "DRJ" Then
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        Else
''                            xVENDORCODE = Null2String(rsLOAD_APacc!CustomerCode)
''                            xVENDORNAME = Null2String(rsLOAD_APacc!CUSTNAME)
''                        End If
''                        If Len(rsLOAD_APacc!VOUCHERNO) = 10 Then
''                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 3)
''                        Else
''                            xVOUCHERNO = Left(Null2String(rsLOAD_APacc!VOUCHERNO), 2)
''                        End If
''                        xCounter = xCounter + 1
''                        xlWorksheet.Cells(xCounter, "A") = Format((Trim(Null2String(rsLOAD_APacc!JDATE))), "mm/dd/yyyy")
''                        xlWorksheet.Cells(xCounter, "B") = (Trim(Null2String(rsLOAD_APacc!VOUCHERNO)))
''                        xlWorksheet.Cells(xCounter, "C") = (Trim(Null2String(rsLOAD_APacc!INVOICENO)))
''                        xlWorksheet.Cells(xCounter, "D") = (Trim(Null2String(rsLOAD_APacc!remarks)))
''
''                        'SJR 7/1/14
''                        'Remarks for details
''                        xlWorksheet.Cells(xCounter, "E") = (Trim(Null2String(rsLOAD_APacc!ADJ_REMARKS)))
''                        If COMPANY_CODE = "HNE" Then
''                            xlWorksheet.Cells(xCounter, "J") = (Trim(Null2String(rsLOAD_AP!ACCT_CODE)))
''                            xlWorksheet.Cells(xCounter, "K") = (Trim(Null2String(rsLOAD_APacc!DESCRIPTION)))
''                        End If
''                        On Error Resume Next
''                        If (rsLOAD_APacc!ENTITY_CLASS) = "C" Then
''                            xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
''                        ElseIf (rsLOAD_APacc!ENTITY_CLASS) = "V" Then
''                            xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!nameofvendor)))
''                       Else
''                          xlWorksheet.Cells(xCounter, "F") = (Trim(Null2String(rsLOAD_APacc!CUSTNAME)))
''                       On Error Resume Next
''                        End If
''
''                        If NumericVal(rsLOAD_APacc!Credit) <> 0 Then
''                            xlWorksheet.Cells(xCounter, "G") = (Trim("0.00"))
''                            xlWorksheet.Cells(xCounter, "H") = (Trim(ToDoubleNumber(NumericVal(rsLOAD_APacc!Credit))))
''
''                            TOTAL_CREDIT = ToDoubleNumber(Round((TOTAL_CREDIT + NumericVal(rsLOAD_APacc!Credit)), 2))
''                            Credit = Credit + TOTAL_CREDIT
''                            xBALANCE = ToDoubleNumber(Round((xBALANCE - NumericVal(rsLOAD_APacc!Credit)), 2))
''                            xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
''
''                        Else
''                            xlWorksheet.Cells(xCounter, "G") = (Trim(Null2String(rsLOAD_APacc!Debit)))
''                            xlWorksheet.Cells(xCounter, "H") = (Trim("0.00"))
''                            TOTAL_DEBIT = ToDoubleNumber(Round((TOTAL_DEBIT + NumericVal(rsLOAD_APacc!Debit)), 2))
''                            Debit = Debit + TOTAL_DEBIT
''                            xBALANCE = ToDoubleNumber(Round((xBALANCE + NumericVal(rsLOAD_APacc!Debit)), 2))
''                            xlWorksheet.Cells(xCounter, "I") = (Trim(ToDoubleNumber(xBALANCE)))
''                        End If
''                        rsLOAD_APacc.MoveNext
''                        If rsLOAD_APacc.EOF = True Then
''                            xCounter = xCounter + 1
''                            xlWorksheet.Cells(xCounter, "G") = TOTAL_DEBIT
''                            xlWorksheet.Cells(xCounter, "G").Font.Bold = True
''                            xlWorksheet.Cells(xCounter, "H") = TOTAL_CREDIT
''                            xlWorksheet.Cells(xCounter, "H").Font.Bold = True
''                            xlWorksheet.Cells(xCounter, "I") = xBALANCE
''                            xlWorksheet.Cells(xCounter, "I").Font.Bold = True
''                        End If
''
''                    Loop
''                End If
''                rsLOAD_AP.MoveNext
''                ycount = 1
''                DoEvents
''            Loop
''        End If
''        xCounter = xCounter + 2
''        xlWorksheet.Cells(xCounter, "F") = "TOTAL"
''        xlWorksheet.Cells(xCounter, "F").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "G") = Debit
''        xlWorksheet.Cells(xCounter, "G").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "H") = Credit
''        xlWorksheet.Cells(xCounter, "H").Font.Bold = True
''        xlWorksheet.Cells(xCounter, "I") = Debit - Credit
''        xlWorksheet.Cells(xCounter, "I").Font.Bold = True
''        xlApplication.Visible = True
''        Set xlApplication = Nothing
''        Set xlWorkbook = Nothing
''        Set xlWorksheet = Nothing
''        Set rsLOAD_AP = Nothing
''        Screen.MousePointer = 0
''    End If
''    'SJR 7/1/14
''    'Remarks for details
''    'Else for HSM only
''
''        optPrintbyAccount.Value = 0
''        optPrintAll.Value = 0
''        picPrintopt.Visible = False
''        picPrintopt.ZOrder 1
    
End Sub

Private Sub optPrintbyAccount_KeyDown(KeyCode As Integer, Shift As Integer)
        rptGeneralLedger.Reset
        rptGeneralLedger.Formulas(3) = "BEG_DATE = '" & Format(dtFrom, "MMM-DD-YYYY") & "'"
        rptGeneralLedger.Formulas(4) = "BEGINNING = " & BEGINNING_BALANCE
        rptGeneralLedger.Formulas(5) = "REPORTDATE = '" & Format(dtFrom, "MMM-DD-YYYY") & " to " & Format(dtTo, "MMM-DD-YYYY") & "'"
        rptGeneralLedger.ReportTitle = "G E N E R A L  L E D G E R"
        Dim rsProfile                                       As ADODB.Recordset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then

            rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptGeneralLedger.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
        End If
        
        If grdAccountsLedger.TextMatrix(1, 2) = "BEGINNING BALANCE" And grdAccountsLedger.Rows = 2 Then
            PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger_Beg.Rpt", "{ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
        Else
            If COMPANY_CODE = "HNE" Then
                PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{AMIS_VW_GENERALLEDGER_FINANCINGCOMPANY.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {AMIS_VW_GENERALLEDGER_FINANCINGCOMPANY.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {AMIS_VW_GENERALLEDGER_FINANCINGCOMPANY.STATUS}='P' and {AMIS_VW_GENERALLEDGER_FINANCINGCOMPANY.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
            Else
                PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Ledgers\AccountGeneralLedger.Rpt", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {Journal_Hd.STATUS}='P' and {ChartAccount.AcctCode} = '" & txtCode.Text & "'", DMIS_REPORT_Connection, 1
            End If
        End If
        
        LogAudit "V", "ACCOUNTS GENERAL LEDGER", txtCode
        optPrintbyAccount.Value = 0
        optPrintAll.Value = 0
        picPrintopt.Visible = False
        picPrintopt.ZOrder 1
End Sub

Private Sub optSummary_KeyDown(KeyCode As Integer, Shift As Integer)
    If COMPANY_CODE <> "DJM" Then Exit Sub
    If optSummary.Value = 0 Then Exit Sub
    
        rptGeneralLedger.Reset
            rptGeneralLedger.ReportTitle = "GENERAL LEDGER SUMMARY"
            
            Dim rsProfile                                       As ADODB.Recordset
            Set rsProfile = New ADODB.Recordset
            Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
            If Not (rsProfile.EOF And rsProfile.BOF) Then
    
                rptGeneralLedger.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
                rptGeneralLedger.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
                rptGeneralLedger.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
    
                PrintSQLReport rptGeneralLedger, AMIS_REPORT_PATH & "Summary\AGLSummary.Rpt", "{Journal_Hd.Jdate} >= date(" & Year(dtFrom) & "," & Month(dtFrom) & "," & Day(dtFrom) & ") and {Journal_Hd.Jdate} <= date(" & Year(dtTo) & "," & Month(dtTo) & "," & Day(dtTo) & ") AND {Journal_Hd.STATUS}='P'", DMIS_REPORT_Connection, 1
                
            End If
            LogAudit "V", "ACCOUNTS GENERAL LEDGER SUMMARY", txtCode
            optPrintbyAccount.Value = 0
            optPrintAll.Value = 0
            picPrintopt.Visible = False
            picPrintopt.ZOrder 1
            
            Frame3.ZOrder 0

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
    Dim RS                                                  As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT REFNO FROM AMIS_JOURNAL_HD WHERE JTYPE ='" & XXX & "' AND VOUCHERNO = '" & YYY & "'")
    If Not (RS.EOF And RS.BOF) Then
        getRefNo = Null2String(RS!REFNO)
    End If
    Set RS = Nothing
End Function

Function OpeningBalance(xACCOUNTCODE As String) As Double
    Dim rsOpening                                           As ADODB.Recordset
    Set rsOpening = New ADODB.Recordset
    rsOpening.Open "SELECT DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE ACCT_CODE='" & xACCOUNTCODE & "' AND JTYPE='OPB'", gconDMIS, adOpenForwardOnly
    If Not rsOpening.EOF And Not rsOpening.BOF Then
        'OLD
        '01/08/2015
        'If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or GetAccountType(txtCode.Text) = "8" Then
        If GetAccountType(txtCode.Text) = "2" Or GetAccountType(txtCode.Text) = "3" Or GetAccountType(txtCode.Text) = "4" Or (GetAccountType(txtCode.Text) = "8" And COMPANY_CODE <> "DJM") Or (GetAccountType(txtCode.Text) = "7" And COMPANY_CODE = "DJM") Then
            OpeningBalance = rsOpening!Credit - rsOpening!Debit
        Else
            OpeningBalance = rsOpening!Debit - rsOpening!Credit
        End If
    End If
    Set rsOpening = Nothing
End Function

Sub FORWARDED_BALANCE()
    Set rsJournal_HDDet = New ADODB.Recordset
    Dim OUTBALANCE                                          As Double
    Dim cnt                                                 As Integer
    
    If COMPANY_CODE = "DJM" Then
        If (GetHeaders(xAcctCode) = "12" And GetTitleCode(xAcctCode) = "02") Or GetHeaders(xAcctCode) = "21" Or GetHeaders(xAcctCode) = "31" Or GetHeaders(xAcctCode) = "41" Or GetHeaders(xAcctCode) = "71" Then
            Set rsJournal_HDDet = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT, ROUND (SUM(CREDIT) - SUM(DEBIT),2) AS BALANCE from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & xAcctCode & "'")
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT)), 2)
                FWD_BALANCE = OUTBALANCE
            End If
        Else
            Set rsJournal_HDDet = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT, ROUND (SUM(DEBIT) - SUM(CREDIT),2) AS BALANCE from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & xAcctCode & "'")
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT)), 2)
                FWD_BALANCE = OUTBALANCE
            End If
        End If
    Else
        If GetAccountType(xAcctCode) = "2" Or GetAccountType(xAcctCode) = "3" Or GetAccountType(xAcctCode) = "4" Then
            Set rsJournal_HDDet = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT, ROUND (SUM(CREDIT) - SUM(DEBIT),2) AS BALANCE from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & xAcctCode & "'")
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT)), 2)
                FWD_BALANCE = OUTBALANCE
            End If
        Else
            Set rsJournal_HDDet = gconDMIS.Execute("select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT, ROUND (SUM(DEBIT) - SUM(CREDIT),2) AS BALANCE from AMIS_vw_vLEDGER where (JTYPE <> 'VPJ' and JTYPE <> 'COB' and JTYPE <>'BOB') AND Jdate < '" & dtFrom & "' and Acct_Code = '" & xAcctCode & "'")
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                OUTBALANCE = Round(OUTBALANCE + (N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT)), 2)
                FWD_BALANCE = OUTBALANCE
            End If
        End If
    End If
End Sub





