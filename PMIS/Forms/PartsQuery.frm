VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmPMISInquiry_Query 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Query"
   ClientHeight    =   9120
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   14625
   Begin XtremeReportControl.ReportControl grd_Detail 
      Height          =   3855
      Left            =   30
      TabIndex        =   40
      Top             =   4980
      Width           =   14565
      _Version        =   655364
      _ExtentX        =   25691
      _ExtentY        =   6800
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      AllowColumnResize=   0   'False
      AllowColumnSort =   0   'False
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin XtremeReportControl.ReportControl grd_Hdr 
      Height          =   3645
      Left            =   30
      TabIndex        =   30
      Top             =   840
      Width           =   14565
      _Version        =   655364
      _ExtentX        =   25691
      _ExtentY        =   6429
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   345
      Left            =   13440
      TabIndex        =   71
      Top             =   360
      Width           =   1035
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7485
      TabIndex        =   41
      Top             =   8850
      Width           =   7485
      Begin VB.Label lbl_Index 
         Caption         =   "Cancelled Transactions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   4500
         TabIndex        =   47
         Top             =   15
         Width           =   1785
      End
      Begin VB.Label lbl_Index 
         BackColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   4290
         TabIndex        =   46
         Top             =   15
         Width           =   195
      End
      Begin VB.Label lbl_Index 
         Caption         =   "Posted Transactions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2670
         TabIndex        =   45
         Top             =   15
         Width           =   1785
      End
      Begin VB.Label lbl_Index 
         BackColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   2310
         TabIndex        =   44
         Top             =   15
         Width           =   195
      End
      Begin VB.Label lbl_Index 
         Caption         =   "Un-Posted Transactions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   43
         Top             =   15
         Width           =   1785
      End
      Begin VB.Label lbl_Index 
         BackColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   42
         Top             =   15
         Width           =   195
      End
   End
   Begin VB.PictureBox PIC_BOTTOM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      DrawStyle       =   2  'Dot
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   30
      ScaleHeight     =   435
      ScaleWidth      =   14535
      TabIndex        =   31
      Top             =   4500
      Width           =   14565
      Begin VB.CommandButton cmd_Print 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   375
         Left            =   13440
         TabIndex        =   70
         Top             =   30
         Width           =   1095
      End
      Begin VB.PictureBox picPartsInquiry 
         BackColor       =   &H80000013&
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   30
         ScaleHeight     =   405
         ScaleWidth      =   7695
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   7695
         Begin VB.ComboBox cboLedger_StockOption 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   34
            Top             =   30
            Width           =   2145
         End
         Begin VB.CommandButton cmdLedger_BalanceLedger 
            Caption         =   "Balance Ledger"
            Enabled         =   0   'False
            Height          =   375
            Left            =   6150
            TabIndex        =   35
            Top             =   30
            Width           =   1485
         End
         Begin VB.CommandButton cmdLedger_ViewUnbalanceStock 
            Caption         =   "View Unbalanced Stocks"
            Height          =   375
            Left            =   4050
            TabIndex        =   37
            Top             =   30
            Width           =   2115
         End
         Begin VB.CommandButton cmdLedger_StockOption 
            Caption         =   "Ok"
            Height          =   375
            Left            =   3600
            TabIndex        =   36
            Top             =   30
            Width           =   465
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000013&
            Caption         =   "Options : (Show)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   33
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   10320
         MaxLength       =   35
         TabIndex        =   38
         Top             =   37
         Width           =   3075
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter Details: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9330
         TabIndex        =   39
         ToolTipText     =   "Search Transaction List by entering any Keyword (like RO-0000001, RIV-000001)"
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.PictureBox pic_Top_Ledger 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5D8BC&
      BorderStyle     =   0  'None
      FillColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   14385
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   14385
      Begin VB.OptionButton opt_Ledger_ByModel 
         Caption         =   "By Model Application"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   1605
      End
      Begin VB.OptionButton opt_Ledger_ByProdNo 
         Caption         =   "By Stock#"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Search Part by Product Number"
         Top             =   300
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox txt_Ledger_Search 
         BackColor       =   &H00FFFFFF&
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
         Height          =   375
         Left            =   9930
         MaxLength       =   35
         TabIndex        =   4
         Top             =   300
         Width           =   3345
      End
      Begin VB.ComboBox cboLedger_HARI_NONHARI 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":030A
         Left            =   60
         List            =   "PartsQuery.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   300
         Width           =   2325
      End
      Begin VB.OptionButton opt_Ledger_ByDescription 
         Caption         =   "By Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Search Parts by Part Description"
         Top             =   300
         Width           =   1605
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   19
         Left            =   8430
         TabIndex        =   69
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   18
         Left            =   2430
         TabIndex        =   68
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**Optional Classification"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   17
         Left            =   0
         TabIndex        =   67
         Top             =   60
         Width           =   2265
      End
   End
   Begin VB.PictureBox pic_Top_ISS 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      FillColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   14385
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   14385
      Begin VB.ComboBox cbo_ISS_Transtatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":030E
         Left            =   2370
         List            =   "PartsQuery.frx":031E
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton opt_ISS_PartNo 
         Caption         =   "Stock #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton opt_ISS_RIV 
         Caption         =   "RIV #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   975
      End
      Begin VB.ComboBox cbo_ISS_Type 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":0370
         Left            =   60
         List            =   "PartsQuery.frx":0372
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   300
         Width           =   2325
      End
      Begin VB.TextBox txt_ISS_Search 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   10290
         MaxLength       =   35
         TabIndex        =   17
         Top             =   300
         Width           =   2985
      End
      Begin VB.OptionButton opt_ISS_No 
         Caption         =   "Issuance #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Search Part by Product Number"
         Top             =   300
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton opt_ISS_Customer 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**Optional ISS Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   11
         Left            =   0
         TabIndex        =   60
         Top             =   60
         Width           =   1890
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   10
         Left            =   4500
         TabIndex        =   59
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   9
         Left            =   2340
         TabIndex        =   58
         Top             =   60
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   8
         Left            =   8790
         TabIndex        =   57
         Top             =   360
         Width           =   1485
      End
   End
   Begin VB.PictureBox pic_Top_DETAIL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      FillColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   14385
      TabIndex        =   25
      Top             =   60
      Visible         =   0   'False
      Width           =   14385
      Begin VB.ComboBox cbo_Tran_Transtatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":0374
         Left            =   2370
         List            =   "PartsQuery.frx":0384
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton OPT_TRAN_TRANNO 
         Caption         =   "By Tran#"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   1605
      End
      Begin VB.OptionButton OPT_TRAN_PARTNO 
         Caption         =   "By Stock#"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Search Part by Product Number"
         Top             =   300
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox TXT_TRAN_SEARCH 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   10740
         MaxLength       =   35
         TabIndex        =   28
         Top             =   300
         Width           =   2535
      End
      Begin VB.ComboBox cbo_Tran_TranType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":03D6
         Left            =   60
         List            =   "PartsQuery.frx":03D8
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**Optional ISS Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   15
         Left            =   0
         TabIndex        =   65
         Top             =   60
         Width           =   1890
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   14
         Left            =   4500
         TabIndex        =   64
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   13
         Left            =   2340
         TabIndex        =   63
         Top             =   60
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword (press enter):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   12
         Left            =   8010
         TabIndex        =   62
         Top             =   360
         Width           =   2670
      End
   End
   Begin VB.PictureBox pic_Top_RR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      FillColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   14385
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   14385
      Begin VB.ComboBox cboRR_Transtatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":03DA
         Left            =   2370
         List            =   "PartsQuery.frx":03EA
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton opt_MRR_ByPartNumber 
         Caption         =   "By Stock #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   1605
      End
      Begin VB.OptionButton opt_MRR_ByRR 
         Caption         =   "By RR#."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Search Part by Product Number"
         Top             =   300
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.TextBox txt_MRR_Search 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   9930
         MaxLength       =   35
         TabIndex        =   11
         Top             =   300
         Width           =   3345
      End
      Begin VB.ComboBox cbo_MRR_Suppliers 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "PartsQuery.frx":043C
         Left            =   60
         List            =   "PartsQuery.frx":043E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   2325
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   4
         Left            =   8430
         TabIndex        =   51
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   3
         Left            =   2370
         TabIndex        =   50
         Top             =   60
         Width           =   1740
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   2
         Left            =   4530
         TabIndex        =   49
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**Optional Vendor  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   60
         Width           =   1860
      End
   End
   Begin VB.PictureBox pic_Top_PO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      FillColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   14385
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   14385
      Begin VB.ComboBox cbo_PO_Transtatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PartsQuery.frx":0440
         Left            =   2370
         List            =   "PartsQuery.frx":0450
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   300
         Width           =   2115
      End
      Begin VB.OptionButton opt_PO_Partno 
         Caption         =   "By Stock #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6870
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Search Part by Model Description"
         Top             =   300
         Width           =   1185
      End
      Begin VB.OptionButton opt_PO_HARIPO 
         Caption         =   "By HARI PO#"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Search Parts by Part Description"
         Top             =   300
         Width           =   1185
      End
      Begin VB.ComboBox cbo_PO_Suppliers 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PartsQuery.frx":04A2
         Left            =   60
         List            =   "PartsQuery.frx":04A4
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   300
         Width           =   2325
      End
      Begin VB.TextBox txt_PO_Search 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   9960
         MaxLength       =   35
         TabIndex        =   24
         Top             =   300
         Width           =   3315
      End
      Begin VB.OptionButton opt_PO_PONo 
         Caption         =   "By PO #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Search Part by Product Number"
         Top             =   300
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Keyword :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   7
         Left            =   8460
         TabIndex        =   56
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Index           =   1
         Left            =   4500
         TabIndex        =   55
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "**Optional Vendor  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   6
         Left            =   30
         TabIndex        =   54
         Top             =   60
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Index           =   5
         Left            =   2370
         TabIndex        =   53
         Top             =   60
         Width           =   1740
      End
   End
   Begin wizProgBar.Prg Prg 
      Height          =   255
      Left            =   9165
      TabIndex        =   72
      Top             =   8850
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   450
      Picture         =   "PartsQuery.frx":04A6
      ForeColor       =   0
      BorderStyle     =   2
      BarPicture      =   "PartsQuery.frx":04C2
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
   Begin VB.Label lbl_Index 
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   8400
      TabIndex        =   73
      Top             =   8850
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000D&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   14565
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mn"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Stock Number"
      End
      Begin VB.Menu mnuOpenMaster 
         Caption         =   "Open Master File"
      End
      Begin VB.Menu mnuMac 
         Caption         =   "Apply Last Moving Average Cost"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPMISInquiry_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StockType                                          As String
Dim REC                                                As XtremeReportControl.ReportRecord
Dim RSHD                                               As ADODB.Recordset
Dim SEARCH_STRING                                      As String
Dim SEARCH_TRANTYPE                                    As String
Dim UNION_QUERY                                        As String

Public Sub SetTYPE(str_type As String)
    StockType = str_type
End Sub

Private Sub cbo_ISS_Transtatus_Click()
    ORDTRANSACTIONS_FILLGRID "BY STATUS"
End Sub

Private Sub cbo_MRR_Suppliers_Click()
    MRRTRANSACTIONS_FILLGRID "BY VENDOR"
End Sub

Private Sub cbo_PO_Suppliers_Click()
    If cbo_PO_Suppliers.ListIndex = -1 Or cbo_PO_Suppliers.ListIndex = 0 Then: Exit Sub
    POTRANSACTIONS_FILLGRID "BY VENDOR"

End Sub

Private Sub cbo_PO_Transtatus_Click()
    POTRANSACTIONS_FILLGRID "BY STATUS"
End Sub

Private Sub cbo_Tran_Transtatus_Click()
    TRANDETAILS_FILLGRID "BY STATUS"
End Sub

Private Sub cbo_Tran_TranType_Click()
    TRANDETAILS_FILLGRID "BY STATUS"
End Sub

Private Sub cboRR_Transtatus_Click()
    MRRTRANSACTIONS_FILLGRID "BY STATUS"

End Sub

Private Sub cmd_Print_Click()
    On Error GoTo ErrorCode:
    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter
    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If
    On Error Resume Next
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
    wsXL.Name = "PARTS QUERY"
    Prg.Max = grd_Detail.Rows.Count - 1
    Prg.Value = 0
    Screen.MousePointer = 11
    For intCol = 0 To grd_Detail.Columns.Count - 1
        wsXL.Cells(1, intCol + 1).Value = "" & CStr(grd_Detail.Columns(intCol).Caption) & "  "
    Next
    For intRow = 0 To grd_Detail.Rows.Count - 1
        For intCol = 0 To grd_Detail.Columns.Count - 1
            wsXL.Cells(intRow + 2, intCol + 1).Value = "" & CStr(grd_Detail.Rows(intRow).Record(intCol).Value) & "  "
        Next
        Prg.Value = Prg.Value + 1
        Prg.Text = Round((Prg.Value / Prg.Max) * 100, 0) & "%"
    Next
    For intCol = 1 To grd_Detail.Columns.Count
        wsXL.Columns(intCol).AutoFit
    Next
    wsXL.Range("A1", Right(wsXL.Columns(grd_Detail.Columns.Count).AddressLocal, 1) & grd_Detail.Rows.Count + 1).AutoFormat 2
    objXL.Visible = True
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    MsgBox err.Description
    err.Clear

End Sub

Private Sub cmdLedger_BalanceLedger_Click()
 
    If MsgBox("Are you Sure you want to balance the ledger", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    Dim SQL                                            As String
    
'    SQL = ""
'    SQL = " UPDATE PMIS_STOCKMAS SET LASTM_OH=LEDGERVIEW.LEDGERBALANCE"
'    SQL = SQL & " FROM "
'    SQL = SQL & " (  SELECT * FROM (SELECT STOCKNO, ONHAND MASTERFILE,("
'    SQL = SQL & "    SELECT ISNULL(SUM(TRANQTY),0) FROM ("
'    SQL = SQL & "    SELECT    TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='BEG' AND TYPE='" & StockType & "' AND STATUS IN('P','B') "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT  (TRANQTY)  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT (TRANQTY)    FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "'  AND STATUS IN('P','B')  "
'    SQL = SQL & "     ) T) AS LEDGERBALANCE FROM PMIS_STOCKMAS WHERE TYPE='" & StockType & "') C WHERE C.MASTERFILE<>C.LEDGERBALANCE"
'    SQL = SQL & "    ) LEDGERVIEW"
'    SQL = SQL & " INNER JOIN PMIS_STOCKMAS ON LEDGERVIEW.STOCKNO=PMIS_STOCKMAS.STOCKNO"
'    gconDMIS.Execute SQL
'
'    SQL = ""
'
'    SQL = " UPDATE PMIS_STOCKMAS SET ONHAND=LEDGERVIEW.LEDGERBALANCE"
'    SQL = SQL & " FROM "
'    SQL = SQL & " (  SELECT * FROM (SELECT STOCKNO, ONHAND MASTERFILE,("
'    SQL = SQL & "    SELECT ISNULL(SUM(TRANQTY),0) FROM ("
'    SQL = SQL & "    SELECT    TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='BEG' AND TYPE='" & StockType & "' AND STATUS IN('P','B') "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT  (TRANQTY)  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT (TRANQTY)   FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')  "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT (TRANQTY)    FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT (TRANQTY)    FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B') "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "'  AND STATUS IN('P','B')  "
'    SQL = SQL & "    UNION ALL"
'    SQL = SQL & "    SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & "     ) T) AS LEDGERBALANCE FROM PMIS_STOCKMAS WHERE TYPE='" & StockType & "') C WHERE C.MASTERFILE<>C.LEDGERBALANCE"
'    SQL = SQL & "    ) LEDGERVIEW"
'    SQL = SQL & " INNER JOIN PMIS_STOCKMAS ON LEDGERVIEW.STOCKNO=PMIS_STOCKMAS.STOCKNO"
'    gconDMIS.Execute SQL
    
    Dim cmd                                            As ADODB.Command
    
    On Error GoTo ERROR_BALANCE
    Set cmd = New ADODB.Command
    With cmd
        .NamedParameters = True
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PMIS_BALANCED_STOCK"
        .ActiveConnection = gconDMIS
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("@TYPE", adVarChar, adParamInput, 50, StockType)
    End With
    cmd.Execute

    Screen.MousePointer = 0
    STOCK_LEDGER_FILLGRID


    cmdLedger_BalanceLedger.Enabled = False
    MsgBox "Balancing Ledger vs Stockmaster is sucessfully completed", vbInformation
    
    Exit Sub
ERROR_BALANCE:
    MsgBox err.Number & " " & err.Description
    Screen.MousePointer = 0
    err.Clear
End Sub

Private Sub cmdLedger_StockOption_Click()
    If cboLedger_StockOption.ListIndex = -1 Then Exit Sub
    Select Case cboLedger_StockOption.ListIndex
        Case 0
            STOCK_LEDGER_FILLGRID "ALL"
        Case 1
            STOCK_LEDGER_FILLGRID "ACTIVE"
        Case 2
            STOCK_LEDGER_FILLGRID "INACTIVE"
        Case 3
            STOCK_LEDGER_FILLGRID "BEG"
        Case 4
            STOCK_LEDGER_FILLGRID "ZERO"
        Case 5
            STOCK_LEDGER_FILLGRID "ZEROACTIVE"
        Case 6
            STOCK_LEDGER_FILLGRID "NEGATIVE"
        Case 7
            STOCK_LEDGER_FILLGRID "INTRANS"
        Case 8
            STOCK_LEDGER_FILLGRID "NOTINTRANS"
        Case 9
            STOCK_LEDGER_FILLGRID "INACTIVEBUTGREATER"
        Case 10
            STOCK_LEDGER_FILLGRID "NEGATIVEMAC"
        '
    End Select

End Sub

Private Sub cmdLedger_ViewUnbalanceStock_Click()
    Screen.MousePointer = 11
    Dim RSHD                                           As ADODB.Recordset
    Dim SQL                                            As String
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim cmd                                            As ADODB.Command
    
    On Error GoTo ERROR_BALANCE
    Set cmd = New ADODB.Command
    With cmd
        .NamedParameters = True
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PMIS_VIEW_BALANCED_STOCK"
        .ActiveConnection = gconDMIS
        .CommandTimeout = 1000
        .Parameters.Append .CreateParameter("@STOCKTYPE", adVarChar, adParamInput, 50, StockType)
    End With

'
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
'    DoEvents

    Set RSHD = cmd.Execute
' UPDATED: JBF
' TO update time out expired error
' July 21 2010
'
'    SQL = " SELECT STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from pmis_stockmas WHERE STOCKNO IN "
'    SQL = SQL & " (SELECT STOCKNO FROM (SELECT STOCKNO, ONHAND MASTERFILE,"
'    SQL = SQL & " ("
'    SQL = SQL & " SELECT ISNULL(SUM(TRANQTY),0) FROM ("
'    SQL = SQL & " SELECT    TRANQTY  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='BEG' AND TYPE='" & StockType & "' AND STATUS IN('P','B') "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT  (TRANQTY)  FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')   "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT (TRANQTY)   FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='111111' AND STATUS IN('P','B')  "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT (TRANQTY)    FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT (TRANQTY)    FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B')    "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE='ADJ' AND TRANNO='000000' AND STATUS IN('P','B') "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & " UNION ALL"
'    SQL = SQL & " SELECT -1 *(TRANQTY) FROM PMIS_TDAYTRAN  WHERE STOCK_ORD=STOCKNO AND TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "' AND STATUS IN('P','B')  "
'    SQL = SQL & "  ) T) AS LEDGERBALANCE FROM PMIS_STOCKMAS WHERE TYPE='" & StockType & "') C WHERE C.MASTERFILE<>C.LEDGERBALANCE)"

'    DoEvents
'      SQL = "SELECT  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND,ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE FROM PMIS_STOCKMAS A INNER JOIN" & vbCrLf
'      SQL = SQL & "  (" & vbCrLf
'      SQL = SQL & "SELECT STOCK_ORD, SUM(ISNULL(TRANQTY,0)) AS TRANQTY FROM" & vbCrLf
'      SQL = SQL & "  (" & vbCrLf
'      SQL = SQL & "  SELECT  STOCK_ORD,  TRANQTY  FROM PMIS_DAYTRAN WHERE TRANTYPE='BEG' AND TYPE='" & StockType & "' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT  STOCK_ORD,(TRANQTY)  FROM PMIS_DAYTRAN WHERE  TRANTYPE='ADJ' AND TYPE='" & StockType & "' AND TRANNO='111111' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,(TRANQTY)   FROM PMIS_TDAYTRAN WHERE TRANTYPE='ADJ' AND TYPE='" & StockType & "' AND TRANNO='111111' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,(TRANQTY)    FROM PMIS_DAYTRAN WHERE  TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,(TRANQTY)    FROM PMIS_TDAYTRAN  WHERE TRANTYPE='RR' AND TYPE='" & StockType & "' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,-1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE TRANTYPE='ADJ' AND TYPE='" & StockType & "' AND TRANNO='000000' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,-1 *(TRANQTY) FROM PMIS_TDAYTRAN WHERE TRANTYPE='ADJ' AND TYPE='" & StockType & "' AND TRANNO='000000' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,-1 *(TRANQTY) FROM PMIS_DAYTRAN WHERE TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & "  Union All" & vbCrLf
'      SQL = SQL & "  SELECT STOCK_ORD,-1 *(TRANQTY) FROM PMIS_TDAYTRAN  WHERE TRANTYPE IN('CSH','CHG','RIV','DR') AND TYPE='" & StockType & "' AND STATUS IN('P','B')" & vbCrLf
'      SQL = SQL & ")T GROUP BY STOCK_ORD" & vbCrLf
'      SQL = SQL & "  ) B ON LTRIM(RTRIM(A.STOCKNO)) = LTRIM(RTRIM(B.STOCK_ORD)) WHERE A.TYPE = '" & StockType & "' AND A.ONHAND <> B.TRANQTY" & vbCrLf
'
'    DoEvents
'    Set RSHD = gconDMIS.Execute(SQL)
'    DoEvents




    While Not RSHD.EOF
        Set REC = grd_Hdr.Records.Add
        REC.AddItem (Trim(RSHD!STOCKNO))
        REC.AddItem (Trim(RSHD!STOCKDESC))
        REC.AddItem (Trim(RSHD!MODELCODE))
        REC.AddItem (Trim(RSHD!ONHAND))
        REC.AddItem (FormatNumber(RSHD!MAC))
        REC.AddItem (FormatNumber(RSHD!SRP))
        REC.AddItem ((RSHD!LASTM_OH))
        REC.AddItem (FormatNumber(RSHD!LASTM_MAC))
        REC.AddItem (Trim(RSHD!Location))
        REC.AddItem (Trim(RSHD!purchases))
        REC.AddItem (Trim(RSHD!RECEIPTS))
        REC.AddItem (Trim(RSHD!ISSUANCES))
        REC.AddItem (Trim(RSHD!tpoqty))
        REC.AddItem (Trim(RSHD!TRECQTY))
        REC.AddItem (Trim(RSHD!TISSQTY))
        REC.AddItem (Trim(RSHD!mad))
        REC.AddItem (Trim(RSHD!InvClass))
        REC.AddItem (Trim(RSHD!Active))
        grd_Hdr.Populate
        RSHD.MoveNext
        Set REC = Nothing
    Wend


    grd_Hdr.Populate
    Set RSHD = Nothing

    If grd_Hdr.Rows.Count >= 1 Then
        cmdLedger_BalanceLedger.Enabled = True
        MessagePop InfoFriend, "Information", "Complete"
    Else
        MsgBox "Ledger vs Stock Master file is Balanced", vbInformation
    End If
    Screen.MousePointer = 0
    
    Exit Sub
    
ERROR_BALANCE:
    MsgBox err.Number & " " & err.Description
    Screen.MousePointer = 0
    err.Clear
    
End Sub

Sub MRRTRANSACTIONS_FILLGRID(Optional ByVal SEARCH_METHOD As String = "")
    On Error GoTo ErrorCode
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    DoEvents


    SEARCH_STRING = " WHERE TYPE ='" & StockType & "'"
    Select Case cboRR_Transtatus.ListIndex
        Case 1                                        'P
            SEARCH_STRING = SEARCH_STRING & " AND STATUS IN('P','B')"
        Case 2                                        'U
            SEARCH_STRING = SEARCH_STRING & " AND (ISNULL(STATUS,'U')='U' OR STATUS='N' )"
        Case 3                                        'C
            SEARCH_STRING = SEARCH_STRING & " AND STATUS='C'"
    End Select

    If cbo_MRR_Suppliers.ListIndex <> -1 And cbo_MRR_Suppliers.ListIndex <> 0 Then
        SEARCH_STRING = SEARCH_STRING & " AND RECVD_CODE IN(SELECT CODE FROM ALL_VENDOR  WHERE ID=" & cbo_MRR_Suppliers.ItemData(cbo_MRR_Suppliers.ListIndex) & ") "
    End If

    If SEARCH_METHOD = "" Then
        Set RSHD = gconDMIS.Execute("SELECT RRNO,RRDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,CANCDATE,LISTED,ID FROM PMIS_RR_HD " & SEARCH_STRING & "  ORDER BY ID DESC")
    ElseIf SEARCH_METHOD = "BY STATUS" Then
        Set RSHD = gconDMIS.Execute("SELECT RRNO,RRDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,CANCDATE,LISTED,ID  FROM PMIS_vw_RR_Trans " & SEARCH_STRING & " ORDER BY ID DESC")
    ElseIf SEARCH_METHOD = "ALL" Then
        If LTrim(RTrim(txt_MRR_Search)) <> "" Then
            If opt_MRR_ByRR.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND RRNO LIKE " & N2Str2Null("%" & N2Str2Null(txt_MRR_Search) & "%")
            Else
                SEARCH_STRING = SEARCH_STRING & " AND RRNO IN (SELECT TRANNO FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD LIKE " & N2Str2Null(txt_MRR_Search & "%") & " AND TRANTYPE ='RR' AND TYPE='" & StockType & "')"
            End If
        End If
        Set RSHD = gconDMIS.Execute("SELECT TOP 50 RRNO,RRDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,CANCDATE,LISTED,ID  FROM PMIS_vw_RR_Trans " & SEARCH_STRING & " ORDER BY ID DESC")
    ElseIf SEARCH_METHOD = "BY VENDOR" Then
        Set RSHD = gconDMIS.Execute("SELECT RRNO,RRDATE,PONO,PODATE,RECVD_CODE,RECVD_FROM,DRNO,INVNO,CLASSCODE,TERMS,TTLRRAMT,DS1,DS_DESC1,DS_AMT1,NETRRAMT,STATUS,CANCDATE,LISTED,ID  FROM PMIS_vw_RR_Trans " & SEARCH_STRING & " ORDER BY ID DESC")
    End If


    If Not RSHD.EOF Or Not RSHD.BOF Then
        Do While Not RSHD.EOF
            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem Null2String(RSHD!RRNO)
                .AddItem Format(Null2String(RSHD!RRDATE), "mm/dd/yyyy")
                .AddItem Null2String(RSHD!PONO)
                .AddItem Null2String(RSHD!PODATE)
                .AddItem Null2String(RSHD!recvd_code)
                .AddItem Null2String(RSHD!recvd_from)
                .AddItem Null2String(RSHD!drno)
                .AddItem Null2String(RSHD!invno)
                .AddItem Null2String(RSHD!classcode)
                .AddItem Null2String(RSHD!TERMS)
                .AddItem FormatNumber(N2Str2Zero(RSHD!ttlrramt))
                .AddItem N2Str2IntZero(RSHD!ds1)
                .AddItem Null2String(RSHD!ds_desc1)
                .AddItem FormatNumber(N2Str2Zero(RSHD!DS_AMT1))
                .AddItem FormatNumber(N2Str2Zero(RSHD!netrramt))
                .AddItem Null2String(RSHD!Status)
                .AddItem Null2String(RSHD!CANCDATE)
                .AddItem Null2String(RSHD!LISTED)
                .AddItem RSHD!ID
            End With
            grd_Hdr.Populate
            RSHD.MoveNext
        Loop
        Set RSHD = Nothing
        Screen.MousePointer = 0
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub ORDTRANSACTIONS_FILLGRID(Optional ByVal SEARCH_METHOD As String = "")
    On Error GoTo ErrorCode
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    DoEvents

    SEARCH_STRING = " WHERE TYPE ='" & StockType & "'"


    Select Case cbo_ISS_Transtatus.ListIndex
        Case 1                                        'P
            SEARCH_STRING = SEARCH_STRING & " AND STATUS IN('P','B')"
        Case 2                                        'U
            SEARCH_STRING = SEARCH_STRING & " AND (ISNULL(STATUS,'U')='U' OR STATUS='N' )"
        Case 3                                        'C
            SEARCH_STRING = SEARCH_STRING & " AND STATUS='C'"
    End Select


    If cbo_ISS_Type.ListIndex <> -1 Then
        Select Case cbo_ISS_Type.ListIndex
            Case 1
                SEARCH_TRANTYPE = " AND TRANTYPE='CSH'"
            Case 2
                SEARCH_TRANTYPE = " AND TRANTYPE='CHG'"
            Case 3
                SEARCH_TRANTYPE = " AND TRANTYPE='DR'"
            Case 4
                SEARCH_TRANTYPE = " AND TRANTYPE='ADB'"
            Case 5
                SEARCH_TRANTYPE = " AND TRANTYPE='RIV'"
            Case Else
                SEARCH_TRANTYPE = " AND TRANTYPE in('ADB','CHG','DR','CSH','RIV') "
        End Select
    End If

    SEARCH_STRING = SEARCH_STRING & SEARCH_TRANTYPE

    If SEARCH_METHOD = "" Then
        Set RSHD = gconDMIS.Execute("Select * from PMIS_Ord_Hd  where TYPE ='" & StockType & "' AND TRANTYPE NOT IN('ARS','PRS','MRS') order by TRANTYPE ASC, TRANDATE DESC,tranno  DESC ")

    ElseIf SEARCH_METHOD = "BY STATUS" Then
        Select Case cbo_ISS_Transtatus.ListIndex
            Case 0, 1, -1
                UNION_QUERY = "SELECT TOP 200 * FROM PMIS_ORD_HD   " & SEARCH_STRING & vbCrLf & _
                              "SELECT TOP 200 * FROM PMIS_ORD_HIST   " & SEARCH_STRING & " ORDER BY TRANTYPE ASC, TRANDATE DESC,TRANNO  DESC "
                Set RSHD = gconDMIS.Execute(UNION_QUERY)
            Case Else
                UNION_QUERY = "SELECT  * FROM PMIS_ORD_HD   " & SEARCH_STRING & vbCrLf & _
                              "SELECT * FROM PMIS_ORD_HIST   " & SEARCH_STRING & " ORDER BY TRANTYPE ASC, TRANDATE DESC,TRANNO  DESC "
                Set RSHD = gconDMIS.Execute(UNION_QUERY)
        End Select


    ElseIf SEARCH_METHOD = "ALL" Then
        If LTrim(RTrim(txt_ISS_Search)) <> "" Then
            If opt_ISS_Customer.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND CUSTNAME LIKE " & N2Str2Null(LTrim(RTrim(txt_ISS_Search)) & "%")
            ElseIf opt_ISS_No.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND TRANNO LIKE " & N2Str2Null("%" & LTrim(RTrim(txt_ISS_Search)) & "%")
            ElseIf opt_ISS_RIV.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND RONO LIKE " & N2Str2Null("%" & LTrim(RTrim(txt_ISS_Search)) & "%")
            Else
                SEARCH_STRING = SEARCH_STRING & " AND TRANNO IN (SELECT TRANNO FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD LIKE " & N2Str2Null(LTrim(RTrim(txt_ISS_Search)) & "%") & SEARCH_TRANTYPE & " AND TYPE='" & StockType & "')"
            End If
        End If

        UNION_QUERY = "SELECT TOP 200 * FROM PMIS_ORD_HD   " & SEARCH_STRING & vbCrLf & _
                      "SELECT TOP 200 * FROM PMIS_ORD_HIST   " & SEARCH_STRING & " ORDER BY TRANTYPE ASC, TRANDATE DESC,TRANNO  DESC "
        Set RSHD = gconDMIS.Execute(UNION_QUERY)
    End If


    If Not RSHD.EOF And Not RSHD.BOF Then
        Screen.MousePointer = 11
        Do While Not RSHD.EOF

            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem Null2String(RSHD!TranType)
                .AddItem Null2String(RSHD!TRANNO)
                .AddItem Null2String(RSHD!trandate)
                .AddItem Null2String(RSHD!CUSTCODE)
                .AddItem Replace(Replace(Null2String(RSHD!CUSTNAME), Chr(13), " "), Chr(10), " ")
                .AddItem Null2String(RSHD!RONO)
                .AddItem Null2String(RSHD!smname)
                .AddItem Null2String(RSHD!TERMS)
                .AddItem FormatNumber(N2Str2Zero(RSHD!ttlinvamt))
                .AddItem N2Str2IntZero(RSHD!ds1)
                .AddItem FormatNumber(N2Str2Zero(RSHD!DS_AMT1))
                .AddItem FormatNumber(N2Str2Zero(RSHD!NETINVAMT))
                .AddItem Null2String(RSHD!Status)
                .AddItem RSHD!ID
            End With

            RSHD.MoveNext
            grd_Hdr.Populate
        Loop
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub STOCK_LEDGER_FILLGRID(Optional ByVal XCHOICE As String)
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    Screen.MousePointer = 11
    grd_Detail.Columns(6).FooterText = 0
    grd_Detail.Columns(7).FooterText = 0
    grd_Detail.Columns(8).FooterText = 0

    DoEvents

    Dim RSHD                                           As ADODB.Recordset
    Dim REC                                            As XtremeReportControl.ReportRecord
    Screen.MousePointer = 11
    If XCHOICE = "ALL" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "ACTIVE" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND ISNULL(ACTIVE,'N') ='Y' order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "INACTIVE" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND (ACTIVE = 'N' OR  ACTIVE IS NULL OR ACTIVE ='') order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "BEG" Then
        Set RSHD = gconDMIS.Execute("SELECT STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  FROM PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND STOCKNO IN (SELECT STOCK_ORD FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN ('BEG')) ORDER BY LTRIM(RTRIM(STOCKNO)) ASC")
    ElseIf XCHOICE = "ZERO" Then
        Set RSHD = gconDMIS.Execute("SELECT STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  FROM PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND ISNULL(ONHAND,0)=0 ORDER BY LTRIM(RTRIM(STOCKNO)) ASC")
    ElseIf XCHOICE = "ZEROACTIVE" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND  ISNULL(ONHAND,0)=0 AND ACTIVE='Y' order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "NEGATIVE" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND ISNULL(ONHAND,0)<0 order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "INTRANS" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND STOCKNO IN(SELECT STOCK_ORD FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('RIV','RR','ADJ','CSH','CHG','DR','BEG')) order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "NOTINTRANS" Then
        If StockType = "P" Then
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "'AND STOCKNO NOT IN(SELECT STOCK_ORD FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('PO','PRS','RIV','RR','ADJ','CSH','CHG','DR','BEG')) order by LTRIM(RTRIM(STOCKNO)) asc")
        ElseIf StockType = "A" Then
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "'AND STOCKNO NOT IN(SELECT STOCK_ORD FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('PO','ARS','RIV','RR','ADJ','CSH','CHG','DR','BEG')) order by LTRIM(RTRIM(STOCKNO)) asc")
        Else
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "'AND STOCKNO NOT IN(SELECT STOCK_ORD FROM PMIS_ALLDAYTRAN WHERE TRANTYPE IN('PO','MRS','RIV','RR','ADJ','CSH','CHG','DR','BEG')) order by LTRIM(RTRIM(STOCKNO)) asc")
        End If
    ElseIf XCHOICE = "INACTIVEBUTGREATER" Then
        Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' AND ISNULL(ACTIVE,'N') ='N'  AND ONHAND>0 order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "SRPLESSMAC" Then
        Set RSHD = gconDMIS.Execute("Select  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and SRP < MAC order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "MACEQUALTOZERO" Then
        Set RSHD = gconDMIS.Execute("Select  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and MAC <= 0 order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "WITHONHAND" Then
        Set RSHD = gconDMIS.Execute("Select  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and onhand > 0 order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "WITHNOSRP" Then
        Set RSHD = gconDMIS.Execute("Select  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and isnull(srp,0) = 0 order by LTRIM(RTRIM(STOCKNO)) asc")
    ElseIf XCHOICE = "NEGATIVEMAC" Then
        Set RSHD = gconDMIS.Execute("Select  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and MAC<0 order by LTRIM(RTRIM(STOCKNO)) asc")
    Else
        Set RSHD = gconDMIS.Execute("Select top 100  STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE   from PMIS_STOCKMAS WHERE TYPE='" & StockType & "' and ACTIVE = 'Y' order by LTRIM(RTRIM(STOCKNO)) asc")
    End If
    Screen.MousePointer = 11
    Do While Not RSHD.EOF
        DoEvents
        Screen.MousePointer = 11
        Set REC = grd_Hdr.Records.Add
        REC.AddItem (Trim(RSHD!STOCKNO))
        REC.AddItem (Trim(RSHD!STOCKDESC))
        REC.AddItem (Trim(RSHD!MODELCODE))
        REC.AddItem (Trim(RSHD!ONHAND))
        REC.AddItem (FormatNumber(RSHD!MAC))
        REC.AddItem (FormatNumber(RSHD!SRP))
        REC.AddItem ((RSHD!LASTM_OH))
        REC.AddItem (FormatNumber(RSHD!LASTM_MAC))
        REC.AddItem (Trim(RSHD!Location))
        REC.AddItem (Trim(RSHD!purchases))
        REC.AddItem (Trim(RSHD!RECEIPTS))
        REC.AddItem (Trim(RSHD!ISSUANCES))
        REC.AddItem (Trim(RSHD!tpoqty))
        REC.AddItem (Trim(RSHD!TRECQTY))
        REC.AddItem (Trim(RSHD!TISSQTY))
        REC.AddItem (Trim(RSHD!mad))
        REC.AddItem (Trim(RSHD!InvClass))
        REC.AddItem (Trim(RSHD!Active))
        RSHD.MoveNext
        'grd_Hdr.Populate
        Set REC = Nothing

    Loop
    Screen.MousePointer = 0
    grd_Hdr.Populate
    Screen.MousePointer = 0
    Set RSHD = Nothing
End Sub


Sub STOCK_LEDGER_FILLGRID_SEARCH(XXX As String)
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate

    grd_Detail.Columns(6).FooterText = 0
    grd_Detail.Columns(7).FooterText = 0
    grd_Detail.Columns(8).FooterText = 0
    DoEvents

    Dim RSHD                                           As ADODB.Recordset
    Set RSHD = New ADODB.Recordset

    XXX = Repleys(LTrim(RTrim(XXX)))


    If CSMS_PARTSQUERY = True Then
        If opt_Ledger_ByProdNo.Value = True Then
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKNO like '" & XXX & "%' and active='Y' order by STOCKNO asc")
        End If
        If opt_Ledger_ByDescription.Value = True Then
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKDESC like '" & XXX & "%' and active='Y' order by STOCKDESC asc")
        End If
        If opt_Ledger_ByModel.Value = True Then
            Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC, vehtype, modelcode, Onhand, srp from PMIS_STOCKMAS where  TYPE=" & N2Str2Null(StockType) & " AND modelcode like '" & XXX & "%' and active='Y' order by modelcode asc")
        End If
    Else
        If cboLedger_HARI_NONHARI.ListIndex = 1 Then
            If opt_Ledger_ByProdNo.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKNO like '" & XXX & "%'   and NON_HARI = 'N' order by STOCKNO asc")
            End If
            If opt_Ledger_ByDescription.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKDESC like '" & XXX & "%'   and NON_HARI = 'N' order by STOCKDESC asc")
            End If
            If opt_Ledger_ByModel.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND  modelcode like '%" & XXX & "%'  and NON_HARI = 'N' order by modelcode asc")
            End If
        ElseIf cboLedger_HARI_NONHARI.ListIndex = 2 = True Then
            If opt_Ledger_ByProdNo.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKNO like '" & XXX & "%'   and NON_HARI = 'Y' order by STOCKNO asc")
            End If
            If opt_Ledger_ByDescription.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND STOCKDESC like '" & XXX & "%'   and NON_HARI = 'Y' order by STOCKDESC asc")
            End If
            If opt_Ledger_ByModel.Value = True Then
                Set RSHD = gconDMIS.Execute("Select STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND  modelcode like '%" & XXX & "%'   and NON_HARI = 'Y' order by modelcode asc")
            End If
        Else
            If opt_Ledger_ByProdNo.Value = True Then
                Set RSHD = gconDMIS.Execute("Select TOP 200 STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where  TYPE=" & N2Str2Null(StockType) & " AND STOCKNO like '" & XXX & "%'   order by STOCKNO asc")
            End If
            If opt_Ledger_ByDescription.Value = True Then
                Set RSHD = gconDMIS.Execute("Select TOP 200 STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where TYPE=" & N2Str2Null(StockType) & " AND  STOCKDESC like '" & XXX & "%'  order by STOCKDESC asc")
            End If
            If opt_Ledger_ByModel.Value = True Then
                Set RSHD = gconDMIS.Execute("Select TOP 200 STOCKNO , STOCKDESC,  MODELCODE, ISNULL(ONHAND,0) ONHAND, ROUND(ISNULL(MAC,0),2) MAC, ISNULL(SRP,0) SRP, ISNULL(LASTM_OH,0)LASTM_OH, ROUND(ISNULL(LASTM_MAC,0),2) LASTM_MAC,LOCATION,ISNULL(PURCHASES,0)PURCHASES,ISNULL(RECEIPTS,0)RECEIPTS, ISNULL(ISSUANCES,0)ISSUANCES, ISNULL(TPOQTY,0)TPOQTY, ISNULL(TRECQTY,0)TRECQTY, ISNULL(TISSQTY,0)TISSQTY, ISNULL(MAD,0)MAD,ISNULL(INVCLASS,'')INVCLASS,ISNULL(ACTIVE,'N')ACTIVE  from PMIS_STOCKMAS where  TYPE=" & N2Str2Null(StockType) & " AND modelcode like '%" & XXX & "%' order by modelcode asc")
            End If
        End If
    End If
    Dim REC                                            As XtremeReportControl.ReportRecord
    grd_Hdr.Records.DeleteAll
    While Not RSHD.EOF
        Set REC = grd_Hdr.Records.Add
        REC.AddItem (Trim(RSHD!STOCKNO))
        REC.AddItem (Trim(RSHD!STOCKDESC))
        REC.AddItem (Trim(RSHD!MODELCODE))
        REC.AddItem (Trim(RSHD!ONHAND))
        REC.AddItem (FormatNumber(RSHD!MAC))
        REC.AddItem (FormatNumber(RSHD!SRP))

        REC.AddItem ((RSHD!LASTM_OH))
        REC.AddItem (FormatNumber(RSHD!LASTM_MAC))
        REC.AddItem (Trim(RSHD!Location))

        REC.AddItem (Trim(RSHD!purchases))
        REC.AddItem (Trim(RSHD!RECEIPTS))
        REC.AddItem (Trim(RSHD!ISSUANCES))

        REC.AddItem (Trim(RSHD!tpoqty))
        REC.AddItem (Trim(RSHD!TRECQTY))
        REC.AddItem (Trim(RSHD!TISSQTY))

        REC.AddItem (Trim(RSHD!mad))
        REC.AddItem (Trim(RSHD!InvClass))
        REC.AddItem (Trim(RSHD!Active))
        RSHD.MoveNext
        Set REC = Nothing
    Wend
    grd_Hdr.Populate
    Set RSHD = Nothing
End Sub

Sub POTRANSACTIONS_FILLGRID(Optional ByVal SEARCH_METHOD As String = "")
    On Error GoTo ErrorCode

    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    DoEvents

    SEARCH_STRING = " WHERE TYPE ='" & StockType & "' "

    Select Case cbo_PO_Transtatus.ListIndex
        Case 1                                        'P
            SEARCH_STRING = SEARCH_STRING & " AND STATUS IN('P','B')"
        Case 2                                        'U
            SEARCH_STRING = SEARCH_STRING & " AND (ISNULL(STATUS,'U')='U' OR STATUS='N' )"
        Case 3                                        'C
            SEARCH_STRING = SEARCH_STRING & " AND STATUS='C'"
    End Select

    If cbo_PO_Suppliers.ListIndex <> -1 And cbo_PO_Suppliers.ListIndex <> 0 Then
        SEARCH_STRING = SEARCH_STRING & " AND SupCode IN(SELECT CODE FROM ALL_VENDOR  WHERE ID=" & cbo_PO_Suppliers.ItemData(cbo_PO_Suppliers.ListIndex) & ") "
    End If

    If SEARCH_METHOD = "" Then
        Set RSHD = gconDMIS.Execute("SELECT * FROM PMIS_PO_HD  WHERE TYPE ='" & StockType & "' ORDER BY PONO ASC ")
    ElseIf SEARCH_METHOD = "BY STATUS" Then
        Set RSHD = gconDMIS.Execute("SELECT  * FROM PMIS_vw_PO_HISTORY  " & SEARCH_STRING & " ORDER BY PONO DESC ")
    ElseIf SEARCH_METHOD = "ALL" Then
        If LTrim(RTrim(txt_PO_Search)) <> "" Then
            If opt_PO_HARIPO.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND DON LIKE " & N2Str2Null(LTrim(RTrim(txt_PO_Search)) & "%")
            ElseIf opt_PO_PONo.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND PONO LIKE " & N2Str2Null("%" & LTrim(RTrim(txt_PO_Search)) & "%")
            Else
                SEARCH_STRING = SEARCH_STRING & " AND PONO IN (SELECT TRANNO FROM PMIS_ALLDAYTRAN WHERE STOCK_ORD LIKE " & N2Str2Null(LTrim(RTrim(txt_PO_Search)) & "%") & " AND TRANTYPE ='PO' AND TYPE=" & N2Str2Null(StockType)
            End If
        End If
        Set RSHD = gconDMIS.Execute("SELECT TOP 50 * FROM PMIS_vw_PO_HISTORY  " & SEARCH_STRING & " ORDER BY PONO DESC ")
    ElseIf SEARCH_METHOD = "BY VENDOR" Then
        Set RSHD = gconDMIS.Execute("SELECT  * FROM PMIS_vw_PO_HISTORY  " & SEARCH_STRING & " ORDER BY PONO DESC ")
    End If


    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate

    grd_Detail.Records.DeleteAll
    grd_Detail.Populate

    If Not RSHD.EOF Or Not RSHD.BOF Then
        Do While Not RSHD.EOF
            Set REC = grd_Hdr.Records.Add
            REC.AddItem (Null2String(RSHD!PONO))
            REC.AddItem (Null2String(RSHD!PODATE))
            REC.AddItem (Null2String(RSHD!ORDERTYPE))
            REC.AddItem (Null2String(RSHD!DON))
            REC.AddItem (Null2String(RSHD!SupCode))
            REC.AddItem (Null2String(RSHD!supname))
            REC.AddItem (Null2String(RSHD!DEALERCODE))
            REC.AddItem (Null2String(RSHD!Shipto))
            REC.AddItem FormatNumber(N2Str2Zero(RSHD!po_amount))
            REC.AddItem FormatNumber(N2Str2IntZero(RSHD!ds1))
            REC.AddItem (Null2String(RSHD!ds_desc1))
            REC.AddItem (N2Str2Zero(RSHD!DS_AMT1))
            REC.AddItem FormatNumber(N2Str2Zero(RSHD!netpoamt))
            REC.AddItem (Null2String(RSHD!Status))
            REC.AddItem (Null2String(RSHD!LISTED))
            '            REC.AddItem (rsHD!ID)
            RSHD.MoveNext
        Loop
    End If
    grd_Hdr.Populate
    Set RSHD = Nothing

    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Sub TRANDETAILS_FILLGRID(Optional ByVal SEARCH_METHOD As String = "")

    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    DoEvents
    Screen.MousePointer = 11
    Dim Qty                                            As Long
    Dim LINERATE                                       As Double


    SEARCH_STRING = " WHERE TYPE ='" & StockType & "'"

    Select Case cbo_Tran_Transtatus.ListIndex
        Case 1                                        'P
            SEARCH_STRING = SEARCH_STRING & " AND STATUS IN('P','B')"
        Case 2                                        'U
            SEARCH_STRING = SEARCH_STRING & " AND (ISNULL(STATUS,'U')='U' OR STATUS='N' )"
        Case 3                                        'C
            SEARCH_STRING = SEARCH_STRING & " AND STATUS='C'"
    End Select

    Select Case cbo_Tran_TranType.ListIndex

        Case 1                                        'CSH
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='CSH' "
        Case 2                                        'CHG
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='CHG' "
        Case 3                                        'DR
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='DR' "
        Case 4                                        'ADB
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='ADB' "
        Case 5                                        'RIV
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='RIV' "
        Case 6                                        'PRS
            If StockType = "P" Then
                SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='PRS' "
            ElseIf StockType = "M" Then
                SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='MRS' "
            Else
                SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='ARS' "
            End If

        Case 7                                        'PO
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='PO' "
        Case 8                                        'RR
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='RR' "
        Case 9                                        'ADJ
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='ADJ' "
        Case 10                                       'ADJ
            SEARCH_STRING = SEARCH_STRING & " AND TRANTYPE='BEG' "

    End Select



    If SEARCH_METHOD = "" Then
        Set RSHD = gconDMIS.Execute("SELECT top 50 * FROM PMIS_TDAYTRAN  WHERE TYPE ='" & StockType & "' ORDER BY TRANNO DESC ")
    ElseIf SEARCH_METHOD = "BY STATUS" Then
        Select Case cbo_Tran_Transtatus.ListIndex
            Case 0, -1, 1                             'P
                Set RSHD = gconDMIS.Execute("SELECT   * FROM PMIS_ALLDAYTRAN  " & SEARCH_STRING & " ORDER BY TRANNO DESC ")
            Case Else
                Set RSHD = gconDMIS.Execute("SELECT  * FROM PMIS_ALLDAYTRAN  " & SEARCH_STRING & " ORDER BY TRANNO DESC ")
        End Select
    ElseIf SEARCH_METHOD = "ALL" Then
        If LTrim(RTrim(TXT_TRAN_SEARCH)) <> "" Then
            If OPT_TRAN_PARTNO.Value = True Then
                SEARCH_STRING = SEARCH_STRING & " AND STOCK_ORD LIKE " & N2Str2Null(LTrim(RTrim(TXT_TRAN_SEARCH)) & "%")
            Else
                SEARCH_STRING = SEARCH_STRING & " AND TRANNO LIKE " & N2Str2Null("%" & LTrim(RTrim(TXT_TRAN_SEARCH)) & "%")
            End If
        End If
        Set RSHD = gconDMIS.Execute("SELECT TOP 200 * FROM PMIS_ALLDAYTRAN  " & SEARCH_STRING & " ORDER BY TRANNO DESC ")
    End If



    If Not RSHD.EOF Or Not RSHD.BOF Then
        Do While Not RSHD.EOF
            Set REC = grd_Hdr.Records.Add
            REC.AddItem (Null2String(RSHD!TranType))
            REC.AddItem (Null2String(RSHD!TRANNO))
            REC.AddItem (Null2String(RSHD!trandate))
            REC.AddItem (Null2String(RSHD!STOCK_ORD))
            REC.AddItem (SetSTOCKDESC(Null2String(RSHD!STOCK_ORD)))
            Qty = N2Str2IntZero(RSHD!TRANQTY)
            REC.AddItem Qty

            If Null2String(RSHD!TranType) = "BEG" Then
                LINERATE = N2Str2IntZero(RSHD!TRANUCOST)
            ElseIf Null2String(RSHD!TranType) = "ADJ" Then
                LINERATE = N2Str2IntZero(RSHD!TRANUCOST)
            ElseIf Null2String(RSHD!TranType) = "RR" Then
                LINERATE = N2Str2IntZero(RSHD!TRANUCOST)
            ElseIf Null2String(RSHD!TranType) = "PO" Then
                LINERATE = N2Str2IntZero(RSHD!TRANUCOST)
            Else
                LINERATE = N2Str2IntZero(RSHD!TRANUPRICE)
            End If

            REC.AddItem FormatNumber(LINERATE)
            REC.AddItem FormatNumber(LINERATE * Qty)
            REC.AddItem (Null2String(RSHD!Status))
            RSHD.MoveNext
        Loop
    End If
    grd_Hdr.Populate
    Screen.MousePointer = 0
    Set RSHD = Nothing
    Exit Sub
ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub


    '    On Error GoTo ERRORCODE
    '    Dim YzaCnt                                         As Integer
    '    YzaCnt = 0
    '    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
    '        Screen.MousePointer = 11
    '        RSTDAYTRAN.MoveFirst
    '        Do While Not RSTDAYTRAN.EOF
    '            YzaCnt = YzaCnt + 1
    '            grdQUERY.AddItem Null2String(RSTDAYTRAN!TranType) & Chr(9) & _
                 '                             Null2String(RSTDAYTRAN!TRANNO) & Chr(9) & _
                 '                             Format(Null2String(RSTDAYTRAN!TRANDATE), "mm/dd/yyyy") & Chr(9) & _
                 '                             Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
                 '                             Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                 '                             SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & Chr(9) & _
                 '                             N2Str2IntZero(RSTDAYTRAN!TRANQTY) & Chr(9) & _
                 '                             N2Str2Zero(RSTDAYTRAN!TRANUCOST)
    '            RSTDAYTRAN.MoveNext
    '            If YzaCnt = 1 Then grdQUERY.RemoveItem 1
    '            DoEvents
    '        Loop
    '        Screen.MousePointer = 0
    '    Else
    '        cleargrid grdQUERY
    '    End If
    '    Exit Sub
    '
    'ERRORCODE:
    '    ShowVBError
    '    Exit Sub
End Sub

Private Sub Command1_Click()
    If grd_Hdr.Rows.Count = 0 Then
        MsgBox "No Record(s) to Print", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorCode:
    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim intRow                                         As Integer    ' counter
    Dim intCol                                         As Integer    ' counter
    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If
    On Error Resume Next
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
    wsXL.Name = "PARTS QUERY"
    For intCol = 0 To grd_Hdr.Columns.Count - 1
        wsXL.Cells(1, intCol + 1).Value = "" & CStr(grd_Hdr.Columns(intCol).Caption) & "  "
    Next
    For intRow = 0 To grd_Hdr.Rows.Count - 1
        For intCol = 0 To grd_Hdr.Columns.Count - 1
            wsXL.Cells(intRow + 2, intCol + 1).Value = "" & CStr(grd_Hdr.Rows(intRow).Record(intCol).Value) & "  "
        Next
    Next
    For intCol = 1 To grd_Hdr.Columns.Count
        wsXL.Columns(intCol).AutoFit
    Next
    wsXL.Range("A1", Right(wsXL.Columns(grd_Hdr.Columns.Count).AddressLocal, 1) & grd_Hdr.Rows.Count + 1).AutoFormat 2
    objXL.Visible = True
    Exit Sub
ErrorCode:
    MsgBox err.Description
    err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS COMPUTERIZED STOCKCARDS)"
            Call frmALL_AuditInquiry.DisplayHistory("", "PARTS COMPUTERIZED STOCKCARDS", "PRINTING")
        
        Case vbKeyEscape
            'Unload Me
        Case vbKeyF3:
            On Error Resume Next
            txt_Ledger_Search.SetFocus
            
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1

    If PARTSQUERY = 1 Then
        pic_Top_Ledger.Visible = True
        picPartsInquiry.Visible = True
        Prg.Visible = True
        Prg.ZOrder 0
        lbl_Index(6).Visible = True
        STOCK_LEDGER_INITGRID
        If fromParts = False Then
            STOCK_LEDGER_FILLGRID
        End If
        Me.Caption = "PARTS MANAGEMENT INFORMATION SYSTEMS' QUERY: STOCK LEDGER"
    ElseIf PARTSQUERY = 3 Then
        pic_Top_PO.Visible = True
        POTRANSACTIONS_INITGRID
        POTRANSACTIONS_FILLGRID
        Me.Caption = "PARTS MANAGEMENT INFORMATION SYSTEMS' QUERY: PO TRANSACTIONS"
    ElseIf PARTSQUERY = 4 Then
        pic_Top_RR.Visible = True
        MRRTRANSACTIONS_INITGRID
        MRRTRANSACTIONS_FILLGRID
        Me.Caption = "PARTS MANAGEMENT INFORMATION SYSTEMS' QUERY: RR TRANSACTIONS"
    ElseIf PARTSQUERY = 5 Then
        pic_Top_ISS.Visible = True
        ORDTRANSACTIONS_INITGRID
        ORDTRANSACTIONS_FILLGRID
        Me.Caption = "PARTS MANAGEMENT INFORMATION SYSTEMS' QUERY: ISSUSANCES TRANSACTION"
    ElseIf PARTSQUERY = 7 Then
        Me.Caption = "PARTS MANAGEMENT INFORMATION SYSTEMS' QUERY: TRANSACTION DETAILS"
        PIC_BOTTOM.Visible = False
        pic_Top_DETAIL.Visible = True
        grd_Hdr.Height = 7995
        grd_Detail.Visible = False
        TRANDETAILS_INITGRID
        TRANDETAILS_FILLGRID
    ElseIf PARTSQUERY = 8 Then
    End If
End Sub





Private Sub grd_Detail_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If PARTSQUERY = 1 Then
        If Row.Record(13).Value = "C" Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Underline = True
        ElseIf Row.Record(13).Value = "N" Or Row.Record(13).Value = "U" Or Row.Record(13).Value = "" Then
            Metrics.ForeColor = vbBlue
        End If
    End If
End Sub

Private Sub grd_Detail_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim RS                               As ADODB.Recordset
    Dim sqlcommand                       As String
    Dim XTRANNO                          As String
    Dim xTRANTYPE                        As String
    
    xTRANTYPE = Trim(Left(grd_Detail.SelectedRows(0).Record(2).Value, 3))
    If FOR_CAS = True Then
        If xTRANTYPE = "RR" Or xTRANTYPE = "PO" Then
            XTRANNO = Right(grd_Detail.SelectedRows(0).Record(2).Value, 6)
        Else
            XTRANNO = Right(grd_Detail.SelectedRows(0).Record(2).Value, 7)
        End If
    Else
        If COMPANY_CODE = "DSSC" And xTRANTYPE = "CSH" Then
            XTRANNO = Right(grd_Detail.SelectedRows(0).Record(2).Value, 7)
        Else
            XTRANNO = Right(grd_Detail.SelectedRows(0).Record(2).Value, 6)
        End If
    End If
    
    
    If xTRANTYPE = "RR" Then
        sqlcommand = " Select 'hist' as kailan from pmis_rec_hist where rrno = '" & XTRANNO & "' and [type] = '" & StockType & "' " & _
                     " union all " & _
                     " Select 'current' as kailan from pmis_rr_hd where rrno = '" & XTRANNO & "' and [type] = '" & StockType & "' "
    Else
        sqlcommand = " Select 'hist' as kailan from pmis_ord_hist where tranno = '" & XTRANNO & "' and [TYPE] = '" & StockType & "' and trantype = '" & xTRANTYPE & "' " & _
                     " union all " & _
                     " Select 'current' as kailan from pmis_ord_hd where tranno = '" & XTRANNO & "' and [TYPE] = '" & StockType & "' and trantype = '" & xTRANTYPE & "' "
    End If
    
    Set RS = gconDMIS.Execute(sqlcommand)
    If Not (RS.EOF And RS.BOF) Then
        If StockType = "P" Then
            If Null2String(RS!kailan) = "hist" Then
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "PARTS TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
                    FormExistsShow frmPMISInquiry_ReceivingHist
                    frmPMISInquiry_ReceivingHist.textSearch.Text = XTRANNO
                    frmPMISInquiry_ReceivingHist.click
                Else
                    If xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                    Unload frmPMISInquiry_CustomerOrderHist
                    COUNTERTYPE = xTRANTYPE
                    frmPMISInquiry_CustomerOrderHist.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISInquiry_CustomerOrderHist
                    frmPMISInquiry_CustomerOrderHist.textSearch.Text = XTRANNO
                    frmPMISInquiry_CustomerOrderHist.click
                End If
            Else
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "PURCHASE RECEIVING STORING", "TRANSACTION") = False Then Exit Sub
                    FormExistsShow frmPMISTrans_Receiving2
                    frmPMISTrans_Receiving2.textSearch.Text = XTRANNO
                    frmPMISTrans_Receiving2.click
                Else
                    If xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "PARTS ISSUANCE SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "PARTS DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                    Unload frmPMISTrans_CustomerOrder
                    Unload frmPMISTrans_ADB_Issuances
                    COUNTERTYPE = xTRANTYPE
                    frmPMISTrans_CustomerOrder.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISTrans_CustomerOrder
                    frmPMISTrans_CustomerOrder.textSearch.Text = XTRANNO
                    frmPMISTrans_CustomerOrder.click
                End If
            End If
        ElseIf StockType = "A" Then
            If Null2String(RS!kailan) = "hist" Then
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "ACCESSORIES TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
                    FormExistsShow frmPMISAC_ReceivingHist
                    frmPMISAC_ReceivingHist.textSearch.Text = XTRANNO
                    frmPMISAC_ReceivingHist.click
                Else
                    If xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                    Unload frmPMISInquiry_CustomerOrderHist_AC
                    COUNTERTYPE = xTRANTYPE
                    frmPMISInquiry_CustomerOrderHist_AC.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISInquiry_CustomerOrderHist_AC
                    frmPMISInquiry_CustomerOrderHist_AC.textSearch.Text = XTRANNO
                    frmPMISInquiry_CustomerOrderHist_AC.click
                End If
            Else
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "ACCESSORIES RECEIVING", "TRANSACTION") = False Then Exit Sub
                    FormExistsShow frmPMISTrans_Receiving2_AC
                    frmPMISTrans_Receiving2_AC.textSearch.Text = XTRANNO
                    frmPMISTrans_Receiving2_AC.click
                Else
                    If xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "ACCESSORIES SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "ACCESSORIES DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                    Unload frmPMISTrans_CustomerOrder_AC
                    COUNTERTYPE = xTRANTYPE
                    frmPMISTrans_CustomerOrder_AC.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISTrans_CustomerOrder_AC
                    frmPMISTrans_CustomerOrder_AC.textSearch.Text = XTRANNO
                    frmPMISTrans_CustomerOrder_AC.click
                End If
            End If
        Else
            If Null2String(RS!kailan) = "hist" Then
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "MATERIALS TRANSACTION HISTORY RECEIVING AND STORING", "INQUIRY") = False Then Exit Sub
                    FormExistsShow frmPMISMat_ReceivingHist
                    frmPMISMat_ReceivingHist.textSearch.Text = XTRANNO
                    frmPMISMat_ReceivingHist.click
                Else
                    If xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CHARGE COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY CASH COUNTER ISSUANCE", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY REQUISTION ISSUANCE VOUCHER", "INQUIRY") = False Then Exit Sub
                    ElseIf xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "TRANSACTION HISTORY DR OUT ISSUANCE", "INQUIRY") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                    Unload frmPMISInquiry_CustomerOrderHist_MAT
                    COUNTERTYPE = xTRANTYPE
                    frmPMISInquiry_CustomerOrderHist_MAT.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISInquiry_CustomerOrderHist_MAT
                    frmPMISInquiry_CustomerOrderHist_MAT.textSearch.Text = XTRANNO
                    frmPMISInquiry_CustomerOrderHist_MAT.click
                End If
            Else
                If xTRANTYPE = "RR" Then
                    If Module_Access(LOGID, "MATERIALS RECEIVING", "TRANSACTION") = False Then Exit Sub
                    FormExistsShow frmPMISTrans_Receiving2_MAT
                    frmPMISTrans_Receiving2_MAT.textSearch.Text = XTRANNO
                    frmPMISTrans_Receiving2_MAT.click
                Else
                    If xTRANTYPE = "CHG" Then
                        If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "CSH" Then
                        If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "RIV" Then
                        If Module_Access(LOGID, "MATERIALS SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    ElseIf xTRANTYPE = "DR" Then
                        If Module_Access(LOGID, "MATERIALS DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
                    Else
                        Exit Sub
                    End If
                   
                    Unload frmPMISTrans_CustomerOrder_MAT
                    COUNTERTYPE = xTRANTYPE
                    frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = xTRANTYPE
                    FormExistsShow frmPMISTrans_CustomerOrder_MAT
                    frmPMISTrans_CustomerOrder_MAT.textSearch.Text = XTRANNO
                    frmPMISTrans_CustomerOrder_MAT.click
                End If
            End If
        End If
    End If

End Sub

Private Sub grd_Hdr_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If PARTSQUERY = 3 Then
        If Row.Record(13).Value = "C" Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Underline = True
        ElseIf Row.Record(13).Value = "N" Or Row.Record(13).Value = "U" Or Row.Record(13).Value = "" Then
            Metrics.ForeColor = vbBlue
        End If
    ElseIf PARTSQUERY = 1 Then

        If Row.Record(17).Value = "N" Then
            Metrics.ForeColor = vbRed

        Else
            Metrics.ForeColor = vbBlack
        End If

    ElseIf PARTSQUERY = 4 Then
        If Row.Record(15).Value = "C" Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Underline = True
        ElseIf Row.Record(15).Value = "N" Or Row.Record(15).Value = "U" Or Row.Record(15).Value = "" Then
            Metrics.ForeColor = vbBlue
        End If
    ElseIf PARTSQUERY = 5 Then
        If Row.Record(12).Value = "C" Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Underline = True
        ElseIf Row.Record(12).Value = "N" Or Row.Record(12).Value = "U" Or Row.Record(12).Value = "" Then
            Metrics.ForeColor = vbBlue
        End If
    ElseIf PARTSQUERY = 7 Then
        If Row.Record(8).Value = "C" Then
            Metrics.ForeColor = vbRed
            Metrics.Font.Underline = True
        ElseIf Row.Record(8).Value = "N" Or Row.Record(8).Value = "U" Or Row.Record(8).Value = "" Then
            Metrics.ForeColor = vbBlue
        End If

    End If
End Sub



Private Sub grd_Hdr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmd_Print.Enabled = False
        If PARTSQUERY = 1 Then
            SHOWLEDGER
        ElseIf PARTSQUERY = 3 Then
            SHOWPOTRANSACTION
        ElseIf PARTSQUERY = 4 Then
            SHOWRRTRANSACTION
        ElseIf PARTSQUERY = 5 Then
            SHOWORDTRANSACTION
        End If
        If grd_Detail.Rows.Count > 0 Then
            cmd_Print.Enabled = True
        End If
    End If
End Sub

Private Sub grd_Hdr_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    grd_Hdr_KeyDown 13, 1
End Sub


Sub STOCK_LEDGER_INITGRID()
    With cboLedger_StockOption
        .AddItem "All Stock"
        .AddItem "Active Stocks"
        .AddItem "In-active Stocks"
        .AddItem "Stock# With Begining Balance"
        .AddItem "Stock# With Zero On-hand"
        .AddItem "Stock# Active but Zero On-hand"
        .AddItem "Stock# with Negative On-hand"
        .AddItem "Stock# In Daily Transaction File"
        .AddItem "Stock# Not In Daily Transaction File"
        .AddItem "Stock# In-Active but On-hand>0"
        .AddItem "Stock# SRP less than mac"
        .AddItem "Stock# Mac is equal to zero"
        .AddItem "Stock# With Onhand"
        .AddItem "Stock# With no SRP"
        .AddItem "Stock# Negative Mac"
        .ListIndex = 0
        SetComboWidth cboLedger_StockOption, 300
    End With
    With cboLedger_HARI_NONHARI
        .AddItem "All STOCK TYPE"
        .AddItem "HARI PARTS"
        .AddItem "NON-HARI PARTS"
        .ListIndex = 0
    End With
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr
        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .Columns.Add 0, "Part #", 90, True: .Columns(0).Resizable = False
        .Columns.Add 1, "Description", 150, True
        .Columns.Add 2, "Model Code", 100, True
        .Columns.Add 3, "On-Hand", 60, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "MAC", 80, True: .Columns(4).Alignment = xtpAlignmentRight
        .Columns.Add 5, "SRP", 80, True: .Columns(5).Alignment = xtpAlignmentRight
        .Columns.Add 6, "Last OH", 80, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Last Mac", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Location", 100, True
        .Columns.Add 9, "Total PO", 80, True: .Columns(9).Alignment = xtpAlignmentCenter
        .Columns.Add 10, "Total RR", 80, True: .Columns(10).Alignment = xtpAlignmentCenter
        .Columns.Add 11, "Total ISS", 80, True: .Columns(11).Alignment = xtpAlignmentCenter
        .Columns.Add 12, "MTD PO", 80, True: .Columns(12).Alignment = xtpAlignmentCenter
        .Columns.Add 13, "MTD RR", 80, True: .Columns(13).Alignment = xtpAlignmentCenter
        .Columns.Add 14, "MTD ISS", 80, True: .Columns(14).Alignment = xtpAlignmentCenter
        .Columns.Add 15, "MAD", 80, True: .Columns(15).Alignment = xtpAlignmentCenter
        .Columns.Add 16, "RANK", 80, True: .Columns(16).Alignment = xtpAlignmentCenter
        .Columns.Add 17, "Active", 80, True: .Columns(17).Alignment = xtpAlignmentCenter
    End With
    With grd_Detail
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.NoItemsText = "No Record for Select Stock Number"
        .Columns.DeleteAll
        .Columns.Add 0, "Part Number", "80", False
        .Columns.Add 1, "Tran Date", "70", False
        .Columns.Add 2, "Tran No", "80", False
        .Columns.Add 3, "Supplier Name/Cust Name/Remarks", "250", False
        .Columns.Add 4, "RO Number", "70", False
        .Columns.Add 5, "Ref. No.", "70", True
        .Columns.Add 6, "Received", "60", False: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Issued", "60", False: .Columns(7).Alignment = xtpAlignmentCenter
        .Columns.Add 8, "Balance", "60", False: .Columns(8).Alignment = xtpAlignmentCenter
        .Columns.Add 9, "Unit Cost", "80", False: .Columns(9).Alignment = xtpAlignmentRight
        .Columns.Add 10, "MAC", "80", False: .Columns(10).Alignment = xtpAlignmentRight
        .Columns.Add 11, "EXT. MAC", "80", False: .Columns(11).Alignment = xtpAlignmentRight
        .Columns.Add 12, "SRP", "80", False: .Columns(12).Alignment = xtpAlignmentRight
        .Columns.Add 13, "Status", "80", False: .Columns(13).Alignment = xtpAlignmentCenter
        .Columns.Add 14, "User", "60", False: .Columns(14).Alignment = xtpAlignmentCenter

        .ShowFooter = True
        .Columns(0).DrawFooterDivider = False
        .Columns(1).DrawFooterDivider = False
        .Columns(2).DrawFooterDivider = False
        .Columns(3).DrawFooterDivider = False
        .Columns(4).DrawFooterDivider = False
        .Columns(9).DrawFooterDivider = False
        .Columns(10).DrawFooterDivider = False
        .Columns(11).DrawFooterDivider = False
        .Columns(12).DrawFooterDivider = False
        .Columns(13).DrawFooterDivider = False
        .Columns(14).DrawFooterDivider = False





        .Columns(5).FooterAlignment = xtpAlignmentCenter
        .Columns(5).FooterText = "Total(s):"
        .Columns(6).FooterAlignment = xtpAlignmentCenter
        .Columns(7).FooterAlignment = xtpAlignmentCenter
        .Columns(8).FooterAlignment = xtpAlignmentCenter

    End With
End Sub

Sub MRRTRANSACTIONS_INITGRID()
    '************************************************************************************************************************************************************
    Dim rsVENDOR                                       As ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("SELECT NAMEOFVENDOR,ID FROM ALL_VENDOR WHERE CODE IN (SELECT RECVD_CODE FROM PMIS_vw_RR_Trans) ORDER BY NAMEOFVENDOR")
    cbo_MRR_Suppliers.Clear
    While Not rsVENDOR.EOF
        cbo_MRR_Suppliers.AddItem UCase(Null2String(rsVENDOR!nameofvendor))
        cbo_MRR_Suppliers.ItemData(cbo_MRR_Suppliers.NewIndex) = Null2String(rsVENDOR!ID)
        rsVENDOR.MoveNext
    Wend
    cbo_MRR_Suppliers.AddItem "ALL VENDOR", 0
    'cbo_MRR_Suppliers.ListIndex = 0
    SetComboWidth cbo_MRR_Suppliers, 300
    '************************************************************************************************************************************************************
    SetComboWidth cboRR_Transtatus, 200
    '************************************************************************************************************************************************************
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr
        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .Columns.Add 0, "RR#", 50, True: .Columns(0).Resizable = False: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "RR Date", 70, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "PO#", 40, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "PO Date", 70, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Code", 50, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Supplier Name ", 180, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "DR#", 60, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Invoice#", 70, True: .Columns(7).Alignment = xtpAlignmentCenter
        .Columns.Add 8, "Class Code", 70, True: .Columns(8).Alignment = xtpAlignmentCenter
        .Columns.Add 9, " Terms", 50, True: .Columns(9).Alignment = xtpAlignmentCenter
        .Columns.Add 10, "Total Amt", 80, True: .Columns(10).Alignment = xtpAlignmentRight
        .Columns.Add 11, "Tax", 40, True: .Columns(11).Alignment = xtpAlignmentCenter
        .Columns.Add 12, "...", 40, True: .Columns(12).Alignment = xtpAlignmentCenter
        .Columns.Add 13, "VAT Amt.", 80, True: .Columns(13).Alignment = xtpAlignmentRight
        .Columns.Add 14, "Net Invoice", 80, True: .Columns(14).Alignment = xtpAlignmentRight
        .Columns.Add 15, "Status", 50, True: .Columns(15).Alignment = xtpAlignmentCenter
        .Columns.Add 16, "Date Cancelled", 70, True: .Columns(16).Alignment = xtpAlignmentCenter
    End With


    With grd_Detail
        .Columns.DeleteAll
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .Columns.Add 0, "Type", 70, True: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "Tran#", 60, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Item#", 60, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Parts Ordered", 100, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Parts Supplied", 100, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Description", 200, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "Quantity ", 80, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Unit Cost", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Line Amount", 80, True: .Columns(8).Alignment = xtpAlignmentRight
        .Columns(5).FooterAlignment = xtpAlignmentRight
        .Columns(5).FooterText = "Total:"
        .Columns(6).FooterAlignment = xtpAlignmentCenter
        .Columns(8).FooterAlignment = xtpAlignmentRight
        .ShowFooter = True
    End With
End Sub

Sub ORDTRANSACTIONS_INITGRID()

    '************************************************************************************************************************************************************
    cbo_ISS_Type.Clear
    cbo_ISS_Type.AddItem "ALL ISSUANCE TRANSACTIONS", 0
    cbo_ISS_Type.AddItem "CASH ISSUANCE"
    cbo_ISS_Type.AddItem "CHARGE ISSUANCE"
    cbo_ISS_Type.AddItem "DR-OUT ISSUANCE"
    cbo_ISS_Type.AddItem "ADVANCE BILL"
    cbo_ISS_Type.AddItem "SERVICE ISSUANCE"
    cbo_ISS_Type.ListIndex = 0
    SetComboWidth cbo_ISS_Type, 300
    '************************************************************************************************************************************************************
    SetComboWidth cbo_ISS_Transtatus, 200
    '************************************************************************************************************************************************************
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr

        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .Columns.Add 0, "Type", 50, True: .Columns(0).Resizable = False: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "Tran #", 70, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Tran Date", 70, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Cust. Code", 70, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Customer Name", 150, True: .Columns(4).Alignment = xtpAlignmentLeft
        .Columns.Add 5, "RO Number", 70, True: .Columns(5).Alignment = xtpAlignmentCenter
        .Columns.Add 6, "Salesman", 180, True: .Columns(6).Alignment = xtpAlignmentLeft
        .Columns.Add 7, " Terms", 50, True: .Columns(7).Alignment = xtpAlignmentCenter
        .Columns.Add 8, "Total Invoice", 80, True: .Columns(8).Alignment = xtpAlignmentRight
        .Columns.Add 9, "(%) Disc.", 60, True: .Columns(9).Alignment = xtpAlignmentCenter
        .Columns.Add 10, "Amount Disc.", 80, True: .Columns(10).Alignment = xtpAlignmentRight
        .Columns.Add 11, "Net Invoice", 80, True: .Columns(11).Alignment = xtpAlignmentRight
        .Columns.Add 12, "Status", 50, True: .Columns(12).Alignment = xtpAlignmentCenter

    End With
    With grd_Detail
        .Columns.DeleteAll
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .Columns.Add 0, "Tran Type", 70, True: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "Tran#", 60, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Item#", 60, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Parts Ordered", 100, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Parts Supplied", 100, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Description", 200, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "Quantity ", 80, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Unit Price", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Line Amount", 80, True: .Columns(8).Alignment = xtpAlignmentRight
        .Columns(5).FooterAlignment = xtpAlignmentRight
        .Columns(5).FooterText = "Total:"
        .Columns(6).FooterAlignment = xtpAlignmentCenter
        .Columns(8).FooterAlignment = xtpAlignmentRight
        .ShowFooter = True
    End With
End Sub

Sub POTRANSACTIONS_INITGRID()
    '************************************************************************************************************************************************************
    Dim rsVENDOR                                       As ADODB.Recordset
    Set rsVENDOR = gconDMIS.Execute("SELECT NAMEOFVENDOR,ID FROM ALL_VENDOR WHERE CODE IN (SELECT SUPCODE FROM PMIS_vw_PO_HISTORY) ORDER BY NAMEOFVENDOR")
    cbo_PO_Suppliers.Clear
    While Not rsVENDOR.EOF
        cbo_PO_Suppliers.AddItem UCase(Null2String(rsVENDOR!nameofvendor))
        cbo_PO_Suppliers.ItemData(cbo_PO_Suppliers.NewIndex) = Null2String(rsVENDOR!ID)
        rsVENDOR.MoveNext
    Wend
    cbo_PO_Suppliers.AddItem "ALL VENDOR", 0
    cbo_PO_Suppliers.ListIndex = 0
    SetComboWidth cbo_PO_Suppliers, 300
    '************************************************************************************************************************************************************
    SetComboWidth cbo_PO_Transtatus, 200
    '************************************************************************************************************************************************************
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr
        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .PaintManager.NoItemsText = "No Purchase order for selected Data"
        .Columns.Add 0, "PO#", 50, True: .Columns(0).Resizable = False: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "PO Date", 70, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Type", 40, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "HARI Order#", 80, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Code", 70, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Supplier Name ", 180, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "Dealer Code", 60, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Delivered To", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Total Amt.", 80, True: .Columns(8).Alignment = xtpAlignmentRight
        .Columns.Add 9, " Tax", 40, True: .Columns(9).Alignment = xtpAlignmentCenter
        .Columns.Add 10, "....", 40, True: .Columns(10).Alignment = xtpAlignmentCenter
        .Columns.Add 11, "VAT.", 60, True: .Columns(11).Alignment = xtpAlignmentRight
        .Columns.Add 12, "Net Amount", 80, True: .Columns(12).Alignment = xtpAlignmentRight
        .Columns.Add 13, "Status", 60, True: .Columns(13).Alignment = xtpAlignmentCenter

    End With

    With grd_Detail
        .Columns.DeleteAll
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .Columns.Add 0, "Type", 70, True: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "Tran#", 60, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Item#", 60, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Parts Ordered", 100, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Parts Supplied", 100, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Description", 200, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "Quantity ", 80, True: .Columns(6).Alignment = xtpAlignmentCenter
        .Columns.Add 7, "Unit Cost", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Line Amount", 80, True: .Columns(8).Alignment = xtpAlignmentRight
        .Columns(5).FooterAlignment = xtpAlignmentRight
        .Columns(5).FooterText = "Total:"
        .Columns(6).FooterAlignment = xtpAlignmentCenter
        .Columns(8).FooterAlignment = xtpAlignmentRight
        .ShowFooter = True
    End With

End Sub
Sub TRANDETAILS_INITGRID()
    SetComboWidth cbo_Tran_Transtatus, 200
    '************************************************************************************************************************************************************
    With cbo_Tran_TranType
        .Clear
        .AddItem "ALL ISSUANCE TRANSACTIONS", 0
        .AddItem "CASH ISSUANCE"
        .AddItem "CHARGE ISSUANCE"
        .AddItem "DR-OUT ISSUANCE"
        .AddItem "ADVANCE BILL"
        .AddItem "SERVICE ISSUANCE"
        .AddItem "STOCK REQUISTION"
        .AddItem "PO TRANSACTION"
        .AddItem "RR TRANSACTIONS"
        .AddItem "ADJUSTMENT TRANSACTIONS"
        .AddItem "BEGINING BALANCES"
    End With

    SetComboWidth cbo_Tran_TranType, 300
    '************************************************************************************************************************************************************
    SetComboWidth cbo_ISS_Transtatus, 200
    '************************************************************************************************************************************************************
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr

        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .Columns.Add 0, "Type", 50, True: .Columns(0).Resizable = False: .Columns(0).Alignment = xtpAlignmentCenter
        .Columns.Add 1, "Tran #", 70, True: .Columns(1).Alignment = xtpAlignmentCenter
        .Columns.Add 2, "Tran Date", 70, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Stock No", 80, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Stock Description", 180, True: .Columns(4).Alignment = xtpAlignmentLeft
        .Columns.Add 5, "Qty", 80, True: .Columns(5).Alignment = xtpAlignmentCenter
        .Columns.Add 6, "Rate/Unit Cost", 80, True: .Columns(6).Alignment = xtpAlignmentRight
        .Columns.Add 7, "Line Amount", 80, True: .Columns(7).Alignment = xtpAlignmentRight
        .Columns.Add 8, "Status", 50, True: .Columns(8).Alignment = xtpAlignmentCenter
    End With
End Sub

Private Sub grd_Hdr_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If PARTSQUERY = 1 Then
        PopupMenu mnuRightClick
    End If
End Sub

Private Sub mnuCopy_Click()
    Clipboard.SetText (grd_Hdr.SelectedRows(0).Record(0).Value)
End Sub

Private Sub mnuMac_Click()
    If grd_Detail.Records.Count = 0 Then Exit Sub
     If Module_Access(LOGID, "APPLY LAST MOVING AVERAGE COST", "SYSTEM") = False Then Exit Sub
    'If MsgBox("Are you Sure You Want to Apply Last Moving Average for this Item", vbYesNo) = vbNo Then: Exit Sub
    'If grd_Detail.Records(grd_Detail.Records.Count - 1).Item(0).Value = grd_Hdr.SelectedRows(0).Record(0).Value Then
    '    gconDMIS.Execute ("update pmis_stockmas set MAC=" & NumericVal(grd_Detail.Records(grd_Detail.Records.Count - 1).Item(10).Value) & " WHERE STOCKNO=" & N2Str2Null(grd_Detail.Records(grd_Detail.Records.Count - 1).Item(0).Value))
    'End If
    '




End Sub

Private Sub mnuOpenMaster_Click()
    If StockType = "P" Then
        If Module_Access(LOGID, "PARTS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
        frmMasterFile_Parts.SETSTOCKTYPE ("P")
        FormExistsShow frmMasterFile_Parts
        Call frmMasterFile_Parts.SearchStock(grd_Hdr.SelectedRows(0).Record(0).Value, StockType)
    ElseIf StockType = "A" Then
        If Module_Access(LOGID, "ACCESSORIES MASTER FILE", "DATA ENTRY") = False Then Exit Sub
        frmMasterFile_Accessories.SETSTOCKTYPE ("A")
        FormExistsShow frmMasterFile_Accessories
        Call frmMasterFile_Accessories.SearchStock(grd_Hdr.SelectedRows(0).Record(0).Value, StockType)
    Else
        If Module_Access(LOGID, "MATERIALS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
        frmMasterFile_Material.SETSTOCKTYPE ("M")
        FormExistsShow frmMasterFile_Material
        Call frmMasterFile_Material.SearchStock(grd_Hdr.SelectedRows(0).Record(0).Value, StockType)
    End If
End Sub

Private Sub opt_Ledger_ByDescription_Click()
    txt_Ledger_Search.SetFocus
End Sub

Private Sub opt_Ledger_ByModel_Click()
    txt_Ledger_Search.SetFocus
End Sub

Private Sub opt_Ledger_ByProdNo_Click()
    txt_Ledger_Search.SetFocus
End Sub

Private Sub opt_MRR_ByPartNumber_Click()
    txt_MRR_Search.SetFocus
End Sub

Private Sub opt_MRR_ByRR_Click()
    txt_MRR_Search.SetFocus
End Sub

Function SetSTOCKDESC(ppp As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("SELECT STOCKNO,STOCKDESC FROM PMIS_STOCKMAS WHERE STOCKNO = '" & ppp & "' AND TYPE='" & StockType & "'")
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function

Private Sub SHOWLEDGER()
    Dim rsRR_HD                                        As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Dim STOCKNUMBER                                    As String
    Dim MOVINGAVERAGECOST                              As Double
    Dim BALANS                                         As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim TOTAL_RR                                       As Long
    Dim TOTAL_ISS                                      As Long
    Dim REMARKSFIELD                                   As String
    Dim rsDet                                          As ADODB.Recordset
    
    Dim RR                                              As Integer
    Dim RIV                                             As Integer
    Dim ADJ_IN                                          As Integer
    Dim ADJ_OUT                                         As Integer
    
    grd_Detail.FilterText = ""
    STOCKNUMBER = grd_Hdr.SelectedRows(0).Record(0).Value
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate

    grd_Detail.Columns(6).FooterText = 0
    grd_Detail.Columns(7).FooterText = 0
    grd_Detail.Columns(8).FooterText = 0


    Screen.MousePointer = 11




    If STOCKNUMBER <> "" Then
        Command1.Enabled = False
        UNION_QUERY = "SELECT * FROM " & vbCrLf & _
                    " ( " & vbCrLf & _
                    "       SELECT 'HIST' AS TABLE_NAMEX,ID,ITEMNO,STOCK_ORD,TRANTYPE,TRANDATE,TRANNO,TRANQTY,TRANUCOST,MAC,STATUS,IN_OUT,TRANUPRICE,USERCODE  FROM PMIS_DAYTRAN WHERE TYPE ='" & StockType & "' AND TRANTYPE IN('RIV','RR','ADJ','CSH','CHG','DR','BEG') AND LTRIM(RTRIM(STOCK_ORD)) = " & N2Str2Null(STOCKNUMBER) & vbCrLf & _
                    "           UNION ALL" & vbCrLf & _
                    "       SELECT 'CURRENT' AS TABLE_NAMEX,ID,ITEMNO,STOCK_ORD,TRANTYPE,TRANDATE,TRANNO,TRANQTY,TRANUCOST,MAC,STATUS,IN_OUT,TRANUPRICE,USERCODE FROM PMIS_TDAYTRAN WHERE TYPE ='" & StockType & "' AND TRANTYPE IN('RIV','RR','ADJ','CSH','CHG','DR','BEG') AND LTRIM(RTRIM(STOCK_ORD)) = " & N2Str2Null(STOCKNUMBER) & _
                    " ) " & vbCrLf & _
                    " CTX  ORDER BY TRANDATE,id  ASC"

        Set RSTDAYTRAN = gconDMIS.Execute(UNION_QUERY)
        If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
            RSTDAYTRAN.MoveFirst
            Screen.MousePointer = 11
            Prg.Max = RSTDAYTRAN.RecordCount
            Prg.Value = 0
            Do While Not RSTDAYTRAN.EOF
                If Null2String(RSTDAYTRAN!TranType) = "BEG" Or Null2String(RSTDAYTRAN!TranType) = "IN" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "BEGINNING"
                        .AddItem ""
                        .AddItem ""
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem 0
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                End If
                If Null2String(RSTDAYTRAN!TranType) = "RR" Then
                    If Null2String(RSTDAYTRAN!TABLE_NAMEX) = "HIST" Then
                        Set rsRR_HD = gconDMIS.Execute("SELECT RRNO,RRDATE,RECVD_FROM,INVNO FROM PMIS_REC_HIST WHERE TYPE ='" & StockType & "' AND RRNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    Else
                        Set rsRR_HD = gconDMIS.Execute("SELECT RRNO,RRDATE,RECVD_FROM,INVNO FROM PMIS_RR_HD WHERE TYPE ='" & StockType & "' AND RRNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    End If


                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                            BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            'RR = RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        End If
                    'If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        Set REC = grd_Detail.Records.Add
                        With REC
                            .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                            .AddItem Format(Null2String(rsRR_HD!RRDATE), "mm/dd/yyyy")
                            .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(rsRR_HD!RRNO)
                            .AddItem Null2String(rsRR_HD!recvd_from)
                            .AddItem ""
                            .AddItem Null2String(rsRR_HD!invno)
                            .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            .AddItem 0
                            .AddItem BALANS
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                            .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                            .AddItem "0.00"
                            .AddItem Null2String(RSTDAYTRAN!Status)
                            .AddItem Null2String(RSTDAYTRAN!USERCODE)
                        End With
                        grd_Detail.Populate
                        grd_Detail.TopRowIndex = grd_Detail.Records.Count
                        MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                    End If
                End If
                If Null2String(RSTDAYTRAN!TranType) = "RIV" Or Null2String(RSTDAYTRAN!TranType) = "CSH" Or Null2String(RSTDAYTRAN!TranType) = "CHG" Or Null2String(RSTDAYTRAN!TranType) = "DR" Or Null2String(RSTDAYTRAN!TranType) = "OUT" Then
                    If Null2String(RSTDAYTRAN!TABLE_NAMEX) = "HIST" Then
                        Set RSORD_HD = gconDMIS.Execute("SELECT STATUS2,TRANTYPE,TRANNO,TRANDATE,CUSTNAME,RONO FROM PMIS_ORD_HIST WHERE TYPE ='" & StockType & "' AND TRANTYPE = " & N2Str2Null(RSTDAYTRAN!TranType) & " AND TRANNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    Else
                        Set RSORD_HD = gconDMIS.Execute("SELECT STATUS2,TRANTYPE,TRANNO,TRANDATE,CUSTNAME,RONO FROM PMIS_ORD_HD WHERE TYPE ='" & StockType & "' AND TRANTYPE = " & N2Str2Null(RSTDAYTRAN!TranType) & " AND TRANNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    End If

                   
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                            BALANS = BALANS - N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            TOTAL_ISS = TOTAL_ISS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            'RIV = RIV + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        End If
                        If Null2String(RSORD_HD!STATUS2) = "R" Then

                            Set rsDet = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & StockType & "' AND RONO='" & Null2String(RSORD_HD!RONO) & "' UNION  SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & StockType & "' AND RONO='" & Null2String(RSORD_HD!RONO) & "'")

                            If Not rsDet.EOF Or Not rsDet.BOF Then
                                REMARKSFIELD = "ADB#" & rsDet!TRANNO
                            End If


                        Else
                            REMARKSFIELD = ""
                        End If
                    'If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        Set REC = grd_Detail.Records.Add
                        With REC
                            .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                            .AddItem Format(Null2String(RSORD_HD!trandate), "mm/dd/yyyy")
                            .AddItem Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO)
                            .AddItem Replace(Replace(Null2String(RSORD_HD!CUSTNAME), Chr(13), " "), Chr(10), " ")
                            .AddItem Null2String(RSORD_HD!RONO)
                            .AddItem REMARKSFIELD
                            .AddItem 0
                            .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            .AddItem BALANS
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                            .AddItem ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS)
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                            .AddItem Null2String(RSTDAYTRAN!Status)
                            .AddItem Null2String(RSTDAYTRAN!USERCODE)
                        End With
                        grd_Detail.Populate
                        grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    End If
                End If

                If Null2String(RSTDAYTRAN!TranType) = "ADJ" And Null2String(RSTDAYTRAN!IN_OUT) = "O" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS - N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_ISS = TOTAL_ISS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        'ADJ_OUT = ADJ_OUT + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "ADJUSTMENTS OUT"
                        .AddItem ""
                        .AddItem ""
                        .AddItem 0
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                End If

                If Null2String(RSTDAYTRAN!TranType) = "ADJ" And Null2String(RSTDAYTRAN!IN_OUT) = "I" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        'ADJ_IN = ADJ_IN + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "ADJUSTMENTS IN"
                        .AddItem ""
                        .AddItem ""
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem 0
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                End If
                DoEvents
                Prg.Value = Prg.Value + 1
                Prg.Text = Round((Prg.Value / Prg.Max) * 100, 0) & "%"
                grd_Detail.Columns(6).FooterText = TOTAL_RR
                grd_Detail.Columns(7).FooterText = TOTAL_ISS
                grd_Detail.Columns(8).FooterText = BALANS
                RSTDAYTRAN.MoveNext
            Loop
        End If
        grd_Detail.Columns(6).FooterText = TOTAL_RR
        grd_Detail.Columns(7).FooterText = TOTAL_ISS
        grd_Detail.Columns(8).FooterText = BALANS
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Transaction on Selected Parts..."
        Screen.MousePointer = 0
        Exit Sub
    End If
    Command1.Enabled = True

    grd_Detail.Populate
    grd_Detail.TopRowIndex = grd_Detail.Records.Count
    Screen.MousePointer = 0

End Sub

Sub SHOWORDTRANSACTION()
    Dim ORDnumber                                      As String
    Dim TOTAL_ORD_QTY                                  As Long
    Dim TOTAL_ORD_AMT                                  As Double
    Dim LINE_AMOUNT                                    As Double
    Dim Qty                                            As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Dim ORDTYPE                                        As String
    grd_Detail.FilterText = ""
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    ORDnumber = grd_Hdr.SelectedRows(0).Record(1).Value
    ORDTYPE = grd_Hdr.SelectedRows(0).Record(0).Value
    If ORDnumber <> "" Then
        Set RSTDAYTRAN = New ADODB.Recordset
        RSTDAYTRAN.Open "SELECT TRANTYPE,TRANNO,ITEMNO,TRANDATE,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUPRICE FROM PMIS_ALLDAYTRAN WHERE TYPE ='" & StockType & "' AND TRANNO = '" & ORDnumber & "' AND TRANTYPE='" & ORDTYPE & "' ORDER BY TRANDATE,TRANTYPE,TRANNO,ITEMNO ASC", gconDMIS
        If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
            RSTDAYTRAN.MoveFirst
            Do While Not RSTDAYTRAN.EOF
                Qty = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                LINE_AMOUNT = N2Str2Zero(RSTDAYTRAN!TRANUPRICE) * Qty
                Set REC = grd_Detail.Records.Add
                With REC
                    .AddItem Null2String(RSTDAYTRAN!TranType)
                    .AddItem Null2String(RSTDAYTRAN!TRANNO)
                    .AddItem Format(Null2String(RSTDAYTRAN!itemno), "0000")
                    .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                    .AddItem Null2String(RSTDAYTRAN!STOCK_SUP)
                    .AddItem SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))
                    .AddItem Qty
                    .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                    .AddItem FormatNumber(LINE_AMOUNT)
                End With
                TOTAL_ORD_QTY = TOTAL_ORD_QTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                TOTAL_ORD_AMT = TOTAL_ORD_AMT + LINE_AMOUNT
                grd_Detail.Populate
                RSTDAYTRAN.MoveNext
            Loop
        End If
        If grd_Detail.Rows.Count > 0 Then
            grd_Detail.Columns(6).FooterText = TOTAL_ORD_QTY
            grd_Detail.Columns(8).FooterText = FormatNumber(TOTAL_ORD_AMT)
        End If
        Set REC = Nothing
        NEW_LogAudit "I", "TRANSACTION DETAILS", "", "", "Issuances", ORDnumber, "", ""
    Else
        MsgSpeechBox "No Entry on Issuance"
        Exit Sub
    End If

End Sub

Sub SHOWPOTRANSACTION()
    Dim PONUMBER                                       As String
    Dim TOTAL_PO_QTY                                   As Long
    Dim TOTAL_PO_AMT                                   As Double
    Dim LINE_AMOUNT                                    As Double
    Dim Qty                                            As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    grd_Detail.FilterText = ""
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate

    PONUMBER = grd_Hdr.SelectedRows(0).Record(0).Value
    Set RSTDAYTRAN = gconDMIS.Execute("select trantype,tranno,trandate,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost from PMIS_ALLdayTran where TYPE ='" & StockType & "' AND trantype = 'PO' and tranno = '" & PONUMBER & "' order by trandate,trantype,tranno,itemno asc")
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            Qty = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
            LINE_AMOUNT = N2Str2Zero(RSTDAYTRAN!TRANUCOST) * Qty
            Set REC = grd_Detail.Records.Add
            With REC
                .AddItem Null2String(RSTDAYTRAN!TranType)
                .AddItem Null2String(RSTDAYTRAN!TRANNO)
                .AddItem Format(Null2String(RSTDAYTRAN!itemno), "0000")
                .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                .AddItem Null2String(RSTDAYTRAN!STOCK_SUP)
                .AddItem SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))
                .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                .AddItem FormatNumber(LINE_AMOUNT)
            End With
            TOTAL_PO_QTY = TOTAL_PO_QTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
            TOTAL_PO_AMT = TOTAL_PO_AMT + LINE_AMOUNT
            grd_Detail.Populate
            RSTDAYTRAN.MoveNext
        Loop
        If grd_Detail.Rows.Count > 0 Then
            grd_Detail.Columns(6).FooterText = TOTAL_PO_QTY
            grd_Detail.Columns(8).FooterText = FormatNumber(TOTAL_PO_AMT)
        End If
        Set REC = Nothing
        NEW_LogAudit "I", "PO TRANSACTIONS", "", "", "Purchase Order", PONUMBER, "", ""
    Else
        'MsgSpeechBox "No Entry on PO"
        grd_Detail.PaintManager.NoItemsText = "No Line item(s) for the PO#:" & PONUMBER
        'Exit Sub
    End If
End Sub

Sub SHOWRRTRANSACTION()
    Dim RRNUMBER                                       As String
    Dim TOTAL_RR_QTY                                   As Long
    Dim TOTAL_RR_AMT                                   As Double
    Dim LINE_AMOUNT                                    As Double
    Dim Qty                                            As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    grd_Detail.FilterText = ""
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate
    RRNUMBER = grd_Hdr.SelectedRows(0).Record(0).Value
    If RRNUMBER <> "" Then
        Set RSTDAYTRAN = gconDMIS.Execute("SELECT TRANTYPE,TRANNO,TRANDATE,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TYPE ='" & StockType & "' AND TRANTYPE = 'RR' AND TRANNO = '" & RRNUMBER & "' ORDER BY TRANDATE,TRANTYPE,TRANNO,ITEMNO ASC")
        If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
            RSTDAYTRAN.MoveFirst
            Do While Not RSTDAYTRAN.EOF
                Qty = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                LINE_AMOUNT = N2Str2Zero(RSTDAYTRAN!TRANUCOST) * Qty
                Set REC = grd_Detail.Records.Add
                With REC
                    .AddItem Null2String(RSTDAYTRAN!TranType)
                    .AddItem Null2String(RSTDAYTRAN!TRANNO)
                    .AddItem Format(Null2String(RSTDAYTRAN!itemno), "0000")
                    .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                    .AddItem Null2String(RSTDAYTRAN!STOCK_SUP)
                    .AddItem SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))
                    .AddItem Qty
                    .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                    .AddItem FormatNumber(LINE_AMOUNT)
                End With
                TOTAL_RR_QTY = TOTAL_RR_QTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                TOTAL_RR_AMT = TOTAL_RR_AMT + LINE_AMOUNT
                grd_Detail.Populate
                RSTDAYTRAN.MoveNext
            Loop
            If grd_Detail.Rows.Count > 0 Then
                grd_Detail.Columns(6).FooterText = TOTAL_RR_QTY
                grd_Detail.Columns(8).FooterText = FormatNumber(TOTAL_RR_AMT)
            End If
            Set REC = Nothing
        End If
        NEW_LogAudit "I", "MRR TRANSACTIONS", "", "", "Receiving", RRNUMBER, "", ""
    Else
        MsgSpeechBox "No Entry on MRR"
        Exit Sub
    End If
End Sub

Private Sub Text1_Change()
    grd_Detail.FilterText = Text1
    grd_Detail.Populate
End Sub

Private Sub txt_ISS_Search_Change()
    ORDTRANSACTIONS_FILLGRID "ALL"
End Sub

Private Sub txt_Ledger_Search_Change()
    If Trim(txt_Ledger_Search.Text) <> "" Then
        STOCK_LEDGER_FILLGRID_SEARCH (txt_Ledger_Search.Text)
    Else
        STOCK_LEDGER_FILLGRID
    End If
End Sub

Private Sub txt_Ledger_Search_GotFocus()
    txt_Ledger_Search.BackColor = &HC0FFC0
End Sub

Private Sub txt_Ledger_Search_LostFocus()
    txt_Ledger_Search.BackColor = vbWhite
End Sub

Private Sub txt_MRR_Search_Change()
    MRRTRANSACTIONS_FILLGRID "ALL"
End Sub

Private Sub txt_PO_Search_Change()
    POTRANSACTIONS_FILLGRID "ALL"
End Sub

Private Sub TXT_TRAN_SEARCH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TRANDETAILS_FILLGRID "ALL"
    End If
End Sub

Public Sub frommaster_SHOWLEDGER(XXX As String)
    Dim rsRR_HD                                        As ADODB.Recordset
    Dim RSORD_HD                                       As ADODB.Recordset
    Dim RSTDAYTRAN                                     As ADODB.Recordset
    Dim STOCKNUMBER                                    As String
    Dim MOVINGAVERAGECOST                              As Double
    Dim BALANS                                         As Long
    Dim REC                                            As XtremeReportControl.ReportRecord
    Dim TOTAL_RR                                       As Long
    Dim TOTAL_ISS                                      As Long
    Dim REMARKSFIELD                                   As String
    Dim rsDet                                          As ADODB.Recordset
    
    Dim RR                                              As Integer
    Dim RIV                                             As Integer
    Dim ADJ_IN                                          As Integer
    Dim ADJ_OUT                                         As Integer
    
    grd_Detail.FilterText = ""
    STOCKNUMBER = XXX
    grd_Detail.Records.DeleteAll
    grd_Detail.Populate

    grd_Detail.Columns(6).FooterText = 0
    grd_Detail.Columns(7).FooterText = 0
    grd_Detail.Columns(8).FooterText = 0


    Screen.MousePointer = 11




    If STOCKNUMBER <> "" Then
        UNION_QUERY = "SELECT * FROM " & vbCrLf & _
                    " ( " & vbCrLf & _
                    "       SELECT 'HIST' AS TABLE_NAMEX,ID,ITEMNO,STOCK_ORD,TRANTYPE,TRANDATE,TRANNO,TRANQTY,TRANUCOST,MAC,STATUS,IN_OUT,TRANUPRICE,USERCODE  FROM PMIS_DAYTRAN WHERE TYPE ='" & StockType & "' AND TRANTYPE IN('RIV','RR','ADJ','CSH','CHG','DR','BEG') AND LTRIM(RTRIM(STOCK_ORD)) = " & N2Str2Null(STOCKNUMBER) & vbCrLf & _
                    "           UNION ALL" & vbCrLf & _
                    "       SELECT 'CURRENT' AS TABLE_NAMEX,ID,ITEMNO,STOCK_ORD,TRANTYPE,TRANDATE,TRANNO,TRANQTY,TRANUCOST,MAC,STATUS,IN_OUT,TRANUPRICE,USERCODE FROM PMIS_TDAYTRAN WHERE TYPE ='" & StockType & "' AND TRANTYPE IN('RIV','RR','ADJ','CSH','CHG','DR','BEG') AND LTRIM(RTRIM(STOCK_ORD)) = " & N2Str2Null(STOCKNUMBER) & _
                    " ) " & vbCrLf & _
                    " CTX  ORDER BY TRANDATE ,ID ASC"

        Set RSTDAYTRAN = gconDMIS.Execute(UNION_QUERY)
        If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
            RSTDAYTRAN.MoveFirst
            Screen.MousePointer = 11
            Prg.Max = RSTDAYTRAN.RecordCount
            Prg.Value = 0
            Do While Not RSTDAYTRAN.EOF
                If Null2String(RSTDAYTRAN!TranType) = "BEG" Or Null2String(RSTDAYTRAN!TranType) = "IN" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "BEGINNING"
                        .AddItem ""
                        .AddItem ""
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem 0
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                End If
                If Null2String(RSTDAYTRAN!TranType) = "RR" Then
                    If Null2String(RSTDAYTRAN!TABLE_NAMEX) = "HIST" Then
                        Set rsRR_HD = gconDMIS.Execute("SELECT RRNO,RRDATE,RECVD_FROM,INVNO FROM PMIS_REC_HIST WHERE TYPE ='" & StockType & "' AND RRNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    Else
                        Set rsRR_HD = gconDMIS.Execute("SELECT RRNO,RRDATE,RECVD_FROM,INVNO FROM PMIS_RR_HD WHERE TYPE ='" & StockType & "' AND RRNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    End If


                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                            BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            'RR = RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        End If
                    'If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        Set REC = grd_Detail.Records.Add
                        With REC
                            .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                            .AddItem Format(Null2String(rsRR_HD!RRDATE), "mm/dd/yyyy")
                            .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(rsRR_HD!RRNO)
                            .AddItem Null2String(rsRR_HD!recvd_from)
                            .AddItem ""
                            .AddItem Null2String(rsRR_HD!invno)
                            .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            .AddItem 0
                            .AddItem BALANS
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                            .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                            .AddItem "0.00"
                            .AddItem Null2String(RSTDAYTRAN!Status)
                            .AddItem Null2String(RSTDAYTRAN!USERCODE)
                        End With
                        grd_Detail.Populate
                        grd_Detail.TopRowIndex = grd_Detail.Records.Count
                        MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                    End If
                End If
                If Null2String(RSTDAYTRAN!TranType) = "RIV" Or Null2String(RSTDAYTRAN!TranType) = "CSH" Or Null2String(RSTDAYTRAN!TranType) = "CHG" Or Null2String(RSTDAYTRAN!TranType) = "DR" Or Null2String(RSTDAYTRAN!TranType) = "OUT" Then
                    If Null2String(RSTDAYTRAN!TABLE_NAMEX) = "HIST" Then
                        Set RSORD_HD = gconDMIS.Execute("SELECT STATUS2,TRANTYPE,TRANNO,TRANDATE,CUSTNAME,RONO FROM PMIS_ORD_HIST WHERE TYPE ='" & StockType & "' AND TRANTYPE = " & N2Str2Null(RSTDAYTRAN!TranType) & " AND TRANNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    Else
                        Set RSORD_HD = gconDMIS.Execute("SELECT STATUS2,TRANTYPE,TRANNO,TRANDATE,CUSTNAME,RONO FROM PMIS_ORD_HD WHERE TYPE ='" & StockType & "' AND TRANTYPE = " & N2Str2Null(RSTDAYTRAN!TranType) & " AND TRANNO = " & N2Str2Null(RSTDAYTRAN!TRANNO))
                    End If

                   
                    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                            BALANS = BALANS - N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            TOTAL_ISS = TOTAL_ISS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            'RIV = RIV + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        End If
                        If Null2String(RSORD_HD!STATUS2) = "R" Then

                            Set rsDet = gconDMIS.Execute("SELECT TRANNO FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND TYPE='" & StockType & "' AND RONO='" & Null2String(RSORD_HD!RONO) & "' UNION  SELECT TRANNO FROM PMIS_ORD_HIST WHERE TRANTYPE='ADB' AND TYPE='" & StockType & "' AND RONO='" & Null2String(RSORD_HD!RONO) & "'")

                            If Not rsDet.EOF Or Not rsDet.BOF Then
                                REMARKSFIELD = "ADB#" & rsDet!TRANNO
                            End If


                        Else
                            REMARKSFIELD = ""
                        End If
                    'If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
                        Set REC = grd_Detail.Records.Add
                        With REC
                            .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                            .AddItem Format(Null2String(RSORD_HD!trandate), "mm/dd/yyyy")
                            .AddItem Null2String(RSORD_HD!TranType) & " #" & Null2String(RSORD_HD!TRANNO)
                            .AddItem Replace(Replace(Null2String(RSORD_HD!CUSTNAME), Chr(13), " "), Chr(10), " ")
                            .AddItem Null2String(RSORD_HD!RONO)
                            .AddItem REMARKSFIELD
                            .AddItem 0
                            .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                            .AddItem BALANS
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                            .AddItem ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS)
                            .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                            .AddItem Null2String(RSTDAYTRAN!Status)
                            .AddItem Null2String(RSTDAYTRAN!USERCODE)
                        End With
                        grd_Detail.Populate
                        grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    End If
                End If

                If Null2String(RSTDAYTRAN!TranType) = "ADJ" And Null2String(RSTDAYTRAN!IN_OUT) = "O" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS - N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_ISS = TOTAL_ISS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        'ADJ_OUT = ADJ_OUT + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "ADJUSTMENTS OUT"
                        .AddItem ""
                        .AddItem ""
                        .AddItem 0
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                End If

                If Null2String(RSTDAYTRAN!TranType) = "ADJ" And Null2String(RSTDAYTRAN!IN_OUT) = "I" Then
                    If (Null2String(RSTDAYTRAN!Status) = "P" Or Null2String(RSTDAYTRAN!Status) = "B") And (Null2String(RSTDAYTRAN!Status) <> "C" Or Null2String(RSTDAYTRAN!Status) <> "N") Then
                        BALANS = BALANS + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        TOTAL_RR = TOTAL_RR + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        'ADJ_IN = ADJ_IN + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    End If
                    Set REC = grd_Detail.Records.Add
                    With REC
                        .AddItem Null2String(RSTDAYTRAN!STOCK_ORD)
                        .AddItem Format(Null2String(RSTDAYTRAN!trandate), "mm/dd/yyyy")
                        .AddItem Null2String(RSTDAYTRAN!TranType) & " #" & Null2String(RSTDAYTRAN!TRANNO)
                        .AddItem "ADJUSTMENTS IN"
                        .AddItem ""
                        .AddItem ""
                        .AddItem N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        .AddItem 0
                        .AddItem BALANS
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
                        .AddItem FormatNumber(N2Str2Zero(RSTDAYTRAN!MAC))
                        .AddItem Format(N2Str2Zero(RSTDAYTRAN!MAC) * BALANS, MAXIMUM_DIGIT)
                        .AddItem "0.00"
                        .AddItem Null2String(RSTDAYTRAN!Status)
                        .AddItem Null2String(RSTDAYTRAN!USERCODE)
                    End With
                    grd_Detail.Populate
                    grd_Detail.TopRowIndex = grd_Detail.Records.Count
                    MOVINGAVERAGECOST = N2Str2Zero(RSTDAYTRAN!MAC)
                End If
                DoEvents
                Prg.Value = Prg.Value + 1
                grd_Detail.Columns(6).FooterText = TOTAL_RR
                grd_Detail.Columns(7).FooterText = TOTAL_ISS
                grd_Detail.Columns(8).FooterText = BALANS
                RSTDAYTRAN.MoveNext
            Loop
        End If
        grd_Detail.Columns(6).FooterText = TOTAL_RR
        grd_Detail.Columns(7).FooterText = TOTAL_ISS
        grd_Detail.Columns(8).FooterText = BALANS
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "No Transaction on Selected Parts..."
        Screen.MousePointer = 0
        Exit Sub
    End If


    grd_Detail.Populate
    grd_Detail.TopRowIndex = grd_Detail.Records.Count
    Screen.MousePointer = 0

End Sub


