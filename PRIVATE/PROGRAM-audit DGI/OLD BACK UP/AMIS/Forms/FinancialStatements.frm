VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAMISFinancialStatements 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Financial Statements"
   ClientHeight    =   20145
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10080
   ForeColor       =   &H00FFFFFF&
   Icon            =   "FinancialStatements.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   20145
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   19200
      Left            =   120
      ScaleHeight     =   19200
      ScaleWidth      =   9405
      TabIndex        =   5
      Top             =   720
      Width           =   9405
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   19200
         Left            =   0
         ScaleHeight     =   19200
         ScaleWidth      =   9405
         TabIndex        =   6
         Top             =   0
         Width           =   9405
         Begin SHDocVwCtl.WebBrowser browCashFlow 
            Height          =   1785
            Left            =   360
            TabIndex        =   26
            Top             =   20000
            Width           =   1125
            ExtentX         =   1984
            ExtentY         =   3149
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin SHDocVwCtl.WebBrowser browCashFlowOLD 
            Height          =   1065
            Left            =   8370
            TabIndex        =   25
            Top             =   20000
            Width           =   885
            ExtentX         =   1561
            ExtentY         =   1879
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   "http:///"
         End
         Begin Crystal.CrystalReport rptAMISIncomeStatement 
            Left            =   90
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Financial Statements"
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowGroupTree=   -1  'True
            WindowAllowDrillDown=   -1  'True
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
         End
         Begin VB.Label labOwnersEquity 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Statement of Owners Equity"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   510
            MouseIcon       =   "FinancialStatements.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   3720
            Width           =   4035
         End
         Begin VB.Label labCashFlow 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Statement of Cash Flow"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4890
            MouseIcon       =   "FinancialStatements.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   3720
            Width           =   4035
         End
         Begin VB.Label labIncomeStatementByProductCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statement by Product - Cumulative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   210
            MouseIcon       =   "FinancialStatements.frx":091E
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   11910
            Width           =   4395
         End
         Begin VB.Label labIncomeStatementByProductCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statement by Product - Current"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   4860
            MouseIcon       =   "FinancialStatements.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   11910
            Width           =   4395
         End
         Begin VB.Label labScheduleOAdministrativeExpensesCumulative 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Administrative Expenses"
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
            Height          =   795
            Left            =   2250
            MouseIcon       =   "FinancialStatements.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   13965
            Width           =   5055
         End
         Begin VB.Label labScheduleOfAdminAndSellingExpensesCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Admin and Selling Expenses - Current"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   5280
            MouseIcon       =   "FinancialStatements.frx":123C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   16800
            Width           =   3375
         End
         Begin VB.Label labScheduleOfAdminAndSellingExpensesCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Admin and Selling Expenses - Cummulative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   840
            MouseIcon       =   "FinancialStatements.frx":1546
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   16800
            Width           =   3375
         End
         Begin VB.Image Image22 
            Height          =   720
            Left            =   8190
            MouseIcon       =   "FinancialStatements.frx":1850
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":19A2
            Top             =   8850
            Width           =   585
         End
         Begin VB.Image Image21 
            Height          =   660
            Left            =   5010
            MouseIcon       =   "FinancialStatements.frx":3064
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":31B6
            Top             =   8820
            Width           =   1095
         End
         Begin VB.Image Image20 
            Height          =   615
            Left            =   1500
            MouseIcon       =   "FinancialStatements.frx":57C8
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":591A
            Top             =   8880
            Width           =   1755
         End
         Begin VB.Image Image19 
            Height          =   720
            Left            =   8070
            MouseIcon       =   "FinancialStatements.frx":91BC
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":930E
            Top             =   4830
            Width           =   585
         End
         Begin VB.Image Image18 
            Height          =   660
            Left            =   5010
            MouseIcon       =   "FinancialStatements.frx":A9D0
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":AB22
            Top             =   4860
            Width           =   1095
         End
         Begin VB.Image Image17 
            Height          =   615
            Left            =   1410
            MouseIcon       =   "FinancialStatements.frx":D134
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":D286
            Top             =   4950
            Width           =   1755
         End
         Begin VB.Label labIncomeStatements 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statements"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4620
            MouseIcon       =   "FinancialStatements.frx":10B28
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Label labBalanceSheets 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Balance Sheets"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   1800
            MouseIcon       =   "FinancialStatements.frx":10E32
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   1560
            Width           =   3135
         End
         Begin VB.Image Image16 
            Height          =   1380
            Left            =   4260
            MouseIcon       =   "FinancialStatements.frx":1113C
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1128E
            ToolTipText     =   "View Schedule of Accounts"
            Top             =   17190
            Width           =   930
         End
         Begin VB.Image Image15 
            Height          =   2415
            Left            =   5910
            MouseIcon       =   "FinancialStatements.frx":11D23
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":11E75
            ToolTipText     =   "View Schedule of Admin and Selling Expenses - Current"
            Top             =   14580
            Width           =   2250
         End
         Begin VB.Image Image13 
            Height          =   2415
            Left            =   3735
            MouseIcon       =   "FinancialStatements.frx":14252
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":143A4
            ToolTipText     =   "View Schedule of Administrative Expenses"
            Top             =   12300
            Width           =   2250
         End
         Begin VB.Image Image12 
            Height          =   1305
            Left            =   6690
            MouseIcon       =   "FinancialStatements.frx":1602E
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":16180
            ToolTipText     =   "View Schedule of Selling Expenses - Service Section"
            Top             =   8130
            Width           =   1020
         End
         Begin VB.Image Image11 
            Height          =   1305
            Left            =   3750
            MouseIcon       =   "FinancialStatements.frx":16C8D
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":16DDF
            ToolTipText     =   "View Schedule of Selling Expenses - Parts Section"
            Top             =   8160
            Width           =   1020
         End
         Begin VB.Image Image10 
            Height          =   1305
            Left            =   480
            MouseIcon       =   "FinancialStatements.frx":178EC
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":17A3E
            ToolTipText     =   "View Schedule of Selling Expenses - Sales Section"
            Top             =   8160
            Width           =   1020
         End
         Begin VB.Image Image9 
            Height          =   1230
            Left            =   6360
            MouseIcon       =   "FinancialStatements.frx":1854B
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1869D
            ToolTipText     =   "View Income Statement Service - Cumulative"
            Top             =   4320
            Width           =   1290
         End
         Begin VB.Image Image8 
            Height          =   1305
            Left            =   5490
            MouseIcon       =   "FinancialStatements.frx":19184
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":192D6
            ToolTipText     =   "View Schedule of Selling Expenses - Current"
            Top             =   5910
            Width           =   1245
         End
         Begin VB.Image Image7 
            Height          =   1305
            Left            =   2670
            MouseIcon       =   "FinancialStatements.frx":19E9C
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":19FEE
            ToolTipText     =   "View Schedule of Selling Expenses - Cumulative"
            Top             =   5910
            Width           =   1245
         End
         Begin VB.Image Image6 
            Height          =   1230
            Left            =   3450
            MouseIcon       =   "FinancialStatements.frx":1ABB4
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1AD06
            ToolTipText     =   "View Income Statement Parts - Cumulative"
            Top             =   4320
            Width           =   1290
         End
         Begin VB.Image Image5 
            Height          =   2415
            Left            =   5910
            MouseIcon       =   "FinancialStatements.frx":1B7ED
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1B93F
            ToolTipText     =   "View Income Statement by Product - Current"
            Top             =   9690
            Width           =   2250
         End
         Begin VB.Image Image3 
            Height          =   1230
            Left            =   60
            MouseIcon       =   "FinancialStatements.frx":1D5F7
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1D749
            ToolTipText     =   "View Income Statement Vehicles - Cumulative"
            Top             =   4320
            Width           =   1290
         End
         Begin VB.Image Image1 
            Height          =   1230
            Left            =   2490
            MouseIcon       =   "FinancialStatements.frx":1E230
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":1E382
            ToolTipText     =   "View Balance Sheets"
            Top             =   210
            Width           =   1380
         End
         Begin VB.Label labScheduleOfAccounts 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Accounts"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   2640
            MouseIcon       =   "FinancialStatements.frx":1EF33
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   18750
            Width           =   4155
         End
         Begin VB.Label labScheduleOfSellingExpensesServiceSection 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Selling Expenses - Service Section"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   585
            Left            =   7620
            MouseIcon       =   "FinancialStatements.frx":1F23D
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   8190
            Width           =   1815
         End
         Begin VB.Label labScheduleOfSellingExpensesPartsSection 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Selling Expenses - Parts Section"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   585
            Left            =   4740
            MouseIcon       =   "FinancialStatements.frx":1F547
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   8220
            Width           =   1815
         End
         Begin VB.Label labScheduleOfSellingExpensesSalesSection 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Selling Expenses - Sales Section"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   585
            Left            =   1470
            MouseIcon       =   "FinancialStatements.frx":1F851
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   8220
            Width           =   1665
         End
         Begin VB.Label labScheduleOfSellingExpenseCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Selling Expenses - Cumulative"
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
            Height          =   675
            Left            =   1650
            MouseIcon       =   "FinancialStatements.frx":1FB5B
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   7260
            Width           =   3465
         End
         Begin VB.Label labIncomeStatementServiceCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statement Service - Cumulative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   585
            Left            =   7740
            MouseIcon       =   "FinancialStatements.frx":1FE65
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   4320
            Width           =   1515
         End
         Begin VB.Label labIncomeStatementPartsCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statement Parts - Cumulative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   4860
            MouseIcon       =   "FinancialStatements.frx":2016F
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   4320
            Width           =   1515
         End
         Begin VB.Label labIncomeStatementVehiclesCumulative 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Income Statement Vehicles - Cumulative"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   1560
            MouseIcon       =   "FinancialStatements.frx":20479
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   4320
            Width           =   1515
         End
         Begin VB.Image Image4 
            Height          =   2415
            Left            =   1380
            MouseIcon       =   "FinancialStatements.frx":20783
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":208D5
            ToolTipText     =   "View Income Statement by Product - Cumulative"
            Top             =   9690
            Width           =   2250
         End
         Begin VB.Image Image14 
            Height          =   2415
            Left            =   1380
            MouseIcon       =   "FinancialStatements.frx":2251F
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":22671
            ToolTipText     =   "View Schedule of Admin and Selling Expenses - Cummulative"
            Top             =   14580
            Width           =   2250
         End
         Begin VB.Label labScheduleOfSellingExpenseCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Selling Expenses - Current"
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
            Height          =   645
            Left            =   4650
            MouseIcon       =   "FinancialStatements.frx":24A4E
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   7260
            Width           =   3285
         End
         Begin VB.Image Image24 
            Height          =   1770
            Left            =   1410
            MouseIcon       =   "FinancialStatements.frx":24D58
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":24EAA
            Stretch         =   -1  'True
            ToolTipText     =   "View Balance Sheets"
            Top             =   2040
            Width           =   2160
         End
         Begin VB.Image Image23 
            Height          =   1800
            Left            =   6000
            MouseIcon       =   "FinancialStatements.frx":3DFEA
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":3E13C
            ToolTipText     =   "View Income Statements"
            Top             =   2010
            Width           =   1785
         End
         Begin VB.Image Image2 
            Height          =   1230
            Left            =   5670
            MouseIcon       =   "FinancialStatements.frx":3FF44
            MousePointer    =   99  'Custom
            Picture         =   "FinancialStatements.frx":40096
            ToolTipText     =   "View Income Statements"
            Top             =   210
            Width           =   990
         End
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   585
      Left            =   2130
      TabIndex        =   1
      Top             =   60
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48431105
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   585
      Left            =   6150
      TabIndex        =   2
      Top             =   60
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   48431105
      CurrentDate     =   38216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   450
      TabIndex        =   4
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5010
      TabIndex        =   3
      Top             =   90
      Width           =   945
   End
   Begin MSForms.ScrollBar ScrollBar1 
      Height          =   9000
      Left            =   9585
      TabIndex        =   0
      Top             =   0
      Width           =   435
      Size            =   "767;15875"
      Max             =   11000
      SmallChange     =   500
      LargeChange     =   500
      Delay           =   0
   End
End
Attribute VB_Name = "frmAMISFinancialStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset
Dim rsJournal_Det                                 As ADODB.Recordset
Dim V_NetIncomeOrLoss                             As Double
Dim V_ProvisionForBonus                           As Double
Dim V_ProvisionForTax                             As Double
Dim V_Total_Current_Asset                         As Double
Dim V_Net_Propert_Equipment                       As Double
Dim V_Other_Assets                                As Double
Dim V_TaxCredit                                   As Double
Dim Prev_V_NetIncomeOrLoss                        As Double

Dim VEHICLES_SECTION                              As String
Dim PARTS_SECTION                                 As String
Dim SERVICE_SECTION                               As String
Dim ADMIN_SECTION                                 As String

Dim CF_Depreciation_And_Depletion                 As Double

Dim CF_Provision_For_Income_Tax                   As Double
Dim CF_Allowance_For_Doubtful_Accounts            As Double
Dim CF_Interest_Income                            As Double
Dim CF_Gain_On_Disposal_Of_Equipment              As Double
Dim CF_Accounts_Receivable_Trade                  As Double
Dim CF_Accounts_Receivable_Non_Trade              As Double
Dim CF_Inventories                                As Double
Dim CF_Prepaid_Expenses                           As Double
Dim CF_Tax_Credits                                As Double
Dim CF_Other_Assets                               As Double

Dim CF_Accounts_Payable_Trade                     As Double
Dim CF_Accounts_Payable_Non_Trade                 As Double
Dim CF_Accrued_Expenses                           As Double
Dim CF_Remittances_Payable                        As Double
Dim CF_Taxes_Liabilities                          As Double
Dim CF_Other_Payables                             As Double

Dim CF_Interest_Income_Received                   As Double

Dim CF_Acquiositions_Of_Propert_And_Equipment     As Double

Dim CF_Advances_From_StockHolders_And_Invenstors  As Double

Dim CF_Cash_At_Beginning                          As Double

'For Owners Equity
Dim OE_CapitalStock, OE_Prev_CapitalStock         As Double
Attribute OE_Prev_CapitalStock.VB_VarUserMemId = 1073938467
Dim OE_BalanceBeg, OE_Prev_BalanceBeg             As Double
Attribute OE_BalanceBeg.VB_VarUserMemId = 1073938473
Attribute OE_Prev_BalanceBeg.VB_VarUserMemId = 1073938473
Dim OE_ClearingTB, OE_Prev_ClearingTB             As Double
Attribute OE_ClearingTB.VB_VarUserMemId = 1073938475
Attribute OE_Prev_ClearingTB.VB_VarUserMemId = 1073938475
Dim OE_NetIncome, OE_Prev_NetIncome               As Double
Attribute OE_NetIncome.VB_VarUserMemId = 1073938477
Attribute OE_Prev_NetIncome.VB_VarUserMemId = 1073938477

Function CheckDate() As Boolean
    If Year(dtpFrom) <> Year(dtpTo) Then
        MsgBox "Invalid Date Range!", vbExclamation + vbCritical, "Error"
        CheckDate = False
    Else
        CheckDate = True
    End If
End Function

Function CashBeginningOfPeriod() As Double
    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Function
    End If
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Cash from AMIS_Journal_Det where left(Acct_Code,5) = '11-01' and Status = 'P' and (jdate < '" & CDate(dtpFrom) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        CashBeginningOfPeriod = N2Str2Zero(rsJournal_HD!Cash)
    End If
End Function

Sub ShowBalanceSheetReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME = 'AMIS'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.Reset
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"


        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.Formulas(3) = "NetIncomeOrLoss = " & V_NetIncomeOrLoss
        CrystalRpt.Formulas(4) = "Prev_NetIncomeOrLoss = " & Prev_V_NetIncomeOrLoss
        CrystalRpt.Formulas(11) = "CurrentMonthYear = DateSerial (" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Sub ShowOwnersEquityReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset

    Dim CrystalRpt                                As Crystal.CrystalReport
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME = 'AMIS'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.Reset
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"


        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.Formulas(11) = "CurrentMonthYear = DateSerial (" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"

        CrystalRpt.Formulas(3) = "Prev_CapitalStock = " & OE_Prev_CapitalStock
        CrystalRpt.Formulas(4) = "Prev_BalanceBeg = " & OE_Prev_BalanceBeg
        CrystalRpt.Formulas(5) = "Prev_ClearingTB = " & OE_Prev_ClearingTB
        CrystalRpt.Formulas(6) = "Prev_Net_Income = " & OE_Prev_NetIncome

        CrystalRpt.Formulas(7) = "CapitalStock = " & OE_CapitalStock
        CrystalRpt.Formulas(8) = "BalanceBeg = " & OE_BalanceBeg
        CrystalRpt.Formulas(9) = "ClearingTB = " & OE_ClearingTB
        CrystalRpt.Formulas(10) = "CF_Net_Income = " & OE_NetIncome


        CrystalRpt.ReportTitle = "Statement of Owners Equity": CrystalRpt.WindowTitle = "Statement of Owners Equity"
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Sub ShowCashFlowReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset

    Dim CrystalRpt                                As Crystal.CrystalReport
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE WHERE MODULENAME = 'AMIS'")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.Reset
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"


        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"

        CrystalRpt.Formulas(3) = "CF_Net_Income = " & V_NetIncomeOrLoss

        CrystalRpt.Formulas(21) = "CF_Depreciation_And_Depletion = " & CF_Depreciation_And_Depletion

        CrystalRpt.Formulas(22) = "CF_Provision_For_Income_Tax = " & CF_Provision_For_Income_Tax
        CrystalRpt.Formulas(23) = "CF_Allowance_For_Doubtful_Accounts = " & CF_Allowance_For_Doubtful_Accounts
        CrystalRpt.Formulas(24) = "CF_Interest_Income = " & CF_Interest_Income
        CrystalRpt.Formulas(25) = "CF_Gain_On_Disposal_Of_Equipment = " & CF_Gain_On_Disposal_Of_Equipment
        CrystalRpt.Formulas(26) = "CF_Accounts_Receivable_Trade = " & CF_Accounts_Receivable_Trade
        CrystalRpt.Formulas(27) = "CF_Accounts_Receivable_Non_Trade = " & CF_Accounts_Receivable_Non_Trade
        CrystalRpt.Formulas(28) = "CF_Inventories = " & CF_Inventories
        CrystalRpt.Formulas(29) = "CF_Prepaid_Expenses = " & CF_Prepaid_Expenses
        CrystalRpt.Formulas(30) = "CF_Tax_Credits = " & CF_Tax_Credits
        CrystalRpt.Formulas(31) = "CF_Other_Assets = " & CF_Other_Assets

        CrystalRpt.Formulas(32) = "CF_Accounts_Payable_Trade = " & CF_Accounts_Payable_Trade
        CrystalRpt.Formulas(33) = "CF_Accounts_Payable_Non_Trade = " & CF_Accounts_Payable_Non_Trade
        CrystalRpt.Formulas(34) = "CF_Accrued_Expenses = " & CF_Accrued_Expenses
        CrystalRpt.Formulas(35) = "CF_Remittances_Payable = " & CF_Remittances_Payable
        CrystalRpt.Formulas(36) = "CF_Taxes_Liabilities = " & CF_Taxes_Liabilities
        CrystalRpt.Formulas(37) = "CF_Other_Payables = " & CF_Other_Payables

        CrystalRpt.Formulas(38) = "CF_Interest_Income_Received = " & CF_Interest_Income_Received

        CrystalRpt.Formulas(39) = "CF_Acquisitions_Of_Property_And_Equipment = " & CF_Acquiositions_Of_Propert_And_Equipment

        CrystalRpt.Formulas(41) = "CF_Advances_From_StockHolders_And_Invenstors = " & CF_Advances_From_StockHolders_And_Invenstors

        CrystalRpt.Formulas(42) = "CF_Cash_At_Beginning = " & CF_Cash_At_Beginning

        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

Sub PrintCashFlow()

    Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
    Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

    Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
    Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double

    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Sub
    End If
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (Jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome
    Else
        ShowNoRecord
        Exit Sub
    End If

    Dim PreviousMonthFrom                         As String
    Dim PreviousMonthTo                           As String

    PreviousMonthFrom = DateSerial(Year(dtpFrom), Month(dtpFrom) - 1, Day(dtpFrom))
    PreviousMonthTo = DateSerial(Year(dtpTo), Month(dtpTo) - 1, Day(dtpTo))

    Open App.Path & "\CashFlow.html" For Output As #1
    Screen.MousePointer = 11
    Print #1, "<HTML>"
    Print #1, "<HEAD>"
    Print #1, "<TITLE>STATEMENT OF CASH FLOW</TITLE>"
    Print #1, "</HEAD>"
    Print #1, "<BODY BGCOLOR=FFFFFF>"
    Print #1, "<TABLE ALIGN=BLEEDLEFT WIDTH=100% CELLSPACING=0 CELLPADDING=0 BORDER=0>"
    Print #1, "<TR>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=1.399%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD><TD WIDTH=2.239%></TD>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=1.119%></TD><TD WIDTH=3.778%></TD><TD WIDTH=0.341%></TD><TD WIDTH=0.638%></TD>"
    Print #1, "<TD WIDTH=1.399%></TD><TD WIDTH=0.980%></TD><TD WIDTH=0.140%></TD><TD WIDTH=3.210%></TD><TD WIDTH=1.889%></TD>"
    Print #1, "<TD WIDTH=4.277%></TD><TD WIDTH=13.189%></TD><TD WIDTH=0.140%></TD><TD WIDTH=4.058%></TD><TD WIDTH=0.525%></TD>"
    Print #1, "<TD WIDTH=4.898%></TD><TD WIDTH=0.840%></TD><TD WIDTH=2.379%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.280%></TD><TD WIDTH=0.140%></TD><TD WIDTH=1.399%></TD><TD WIDTH=2.230%></TD><TD WIDTH=7.985%></TD>"
    Print #1, "<TD WIDTH=0.420%></TD><TD WIDTH=1.119%></TD><TD WIDTH=0.140%></TD><TD WIDTH=1.119%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=6.218%></TD><TD WIDTH=0.499%></TD><TD WIDTH=4.898%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD>"
    Print #1, "<TD WIDTH=2.099%></TD><TD WIDTH=0.061%></TD><TD WIDTH=0.638%></TD><TD WIDTH=1.399%></TD><TD WIDTH=5.317%></TD>"
    Print #1, "<TD WIDTH=2.239%></TD><TD WIDTH=0.315%></TD><TD WIDTH=0.385%></TD><TD WIDTH=2.230%></TD><TD WIDTH=0.009%></TD>"
    Print #1, "<TD WIDTH=5.798%></TD><TD WIDTH=0.533%></TD><TD WIDTH=0.805%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.280%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.778%></TD><TD WIDTH=0.009%></TD>"
    Print #1, "</TR>"
    Print #1, "</TABLE>"
    Print #1, "<TABLE ALIGN=BLEEDLEFT WIDTH=100% CELLSPACING=0 CELLPADDING=0 BORDER=0>"
    Print #1, "<TR>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=1.399%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD><TD WIDTH=2.239%></TD>"
    Print #1, "<TD WIDTH=0.140%></TD><TD WIDTH=1.119%></TD><TD WIDTH=3.778%></TD><TD WIDTH=0.341%></TD><TD WIDTH=0.638%></TD>"
    Print #1, "<TD WIDTH=1.399%></TD><TD WIDTH=0.980%></TD><TD WIDTH=0.140%></TD><TD WIDTH=3.210%></TD><TD WIDTH=1.889%></TD>"
    Print #1, "<TD WIDTH=4.277%></TD><TD WIDTH=13.189%></TD><TD WIDTH=0.140%></TD><TD WIDTH=4.058%></TD><TD WIDTH=0.525%></TD>"
    Print #1, "<TD WIDTH=4.898%></TD><TD WIDTH=0.840%></TD><TD WIDTH=2.379%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.280%></TD><TD WIDTH=0.140%></TD><TD WIDTH=1.399%></TD><TD WIDTH=2.230%></TD><TD WIDTH=7.985%></TD>"
    Print #1, "<TD WIDTH=0.420%></TD><TD WIDTH=1.119%></TD><TD WIDTH=0.140%></TD><TD WIDTH=1.119%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=6.218%></TD><TD WIDTH=0.499%></TD><TD WIDTH=4.898%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.420%></TD>"
    Print #1, "<TD WIDTH=2.099%></TD><TD WIDTH=0.061%></TD><TD WIDTH=0.638%></TD><TD WIDTH=1.399%></TD><TD WIDTH=5.317%></TD>"
    Print #1, "<TD WIDTH=2.239%></TD><TD WIDTH=0.315%></TD><TD WIDTH=0.385%></TD><TD WIDTH=2.230%></TD><TD WIDTH=0.009%></TD>"
    Print #1, "<TD WIDTH=5.798%></TD><TD WIDTH=0.533%></TD><TD WIDTH=0.805%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.140%></TD>"
    Print #1, "<TD WIDTH=0.280%></TD><TD WIDTH=0.140%></TD><TD WIDTH=0.778%></TD><TD WIDTH=0.009%></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=4> </TD>"
    Print #1, "<TD ALIGN=CENTER COLSPAN=55><FONT SIZE=4 FACE=Times New Roman><B>" & COMPANY_NAME & "</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=4> </TD>"
    Print #1, "<TD ALIGN=CENTER COLSPAN=55><FONT SIZE=4 FACE=Times New Roman><B>STATEMENT OF CASH FLOW</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=4> </TD>"
    Print #1, "<TD ALIGN=CENTER COLSPAN=55><FONT SIZE=3 FACE=Times New Roman>AS OF: " & Format(dtpTo, "Long Date") & "</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=52><FONT SIZE=3 FACE=Times New Roman><B>CASH FLOW FROM OPERATING ACTIVITIES</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=30><FONT SIZE=3 FACE=Times New Roman><B>NET INCOME</B></FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=5 NOWRAP><FONT SIZE=2 FACE=Times New Roman><B>" & Format(Round(V_NetIncomeOrLoss, 2), "###,###,###,##0.00") & "</B></FONT></TD>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=35><FONT SIZE=3 FACE=Times New Roman>Adjustments to reconcile net income to net cash</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=30><FONT SIZE=3 FACE=Times New Roman>provided by operating activities:</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Dim rsChartAccount                            As ADODB.Recordset

    Dim GrandTotalOperatingAmount, BeginningTotalOperatingAmount, TotalOperatingAmount As Double
    Dim GrandTotalInvestingAmount, BeginningTotalInvestingAmount, TotalInvestingAmount As Double
    Dim GrandTotalFinancingAmount, BeginningTotalFinancingAmount, TotalFinancingAmount As Double

    Dim rsOperating                               As ADODB.Recordset
    Dim rsInvesting                               As ADODB.Recordset
    Dim rsFinancing                               As ADODB.Recordset

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=55><FONT SIZE=3 FACE=Times New Roman>Subtract Income Accounts that do not represent cash flows:</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    GrandTotalOperatingAmount = 0

    Set rsOperating = New ADODB.Recordset
    Set rsOperating = gconDMIS.Execute("Select * from AMIS_TitleCode Where CashFlowCode = 'OA' AND (HeaderCode = '4' or HeaderCode = '5' or HeaderCode = '6' or HeaderCode = '8') Order By Code Asc")
    If Not rsOperating.EOF And Not rsOperating.BOF Then
        rsOperating.MoveFirst:
        Do While Not rsOperating.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles = " & N2Str2Null(rsOperating!Code) & " Order By AcctCode")
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                rsChartAccount.MoveFirst: BeginningTotalOperatingAmount = 0: TotalOperatingAmount = 0
                Do While Not rsChartAccount.EOF
                    Set rsJournal_Det = New ADODB.Recordset
                    If Null2String(rsOperating!HeaderCode) = "2" Or Null2String(rsOperating!HeaderCode) = "3" Or Null2String(rsOperating!HeaderCode) = "4" Or Null2String(rsOperating!HeaderCode) = "8" Then
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    Else
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    End If
                    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
                        TotalOperatingAmount = TotalOperatingAmount + N2Str2Zero(rsJournal_Det!TOTAL_AMOUNT)
                    End If
                    rsChartAccount.MoveNext
                Loop
                BeginningTotalOperatingAmount = 0
                If TotalOperatingAmount - BeginningTotalOperatingAmount <> 0 Then
                    Print #1, "<TR VALIGN=TOP>"
                    Print #1, "<TD COLSPAN=12> </TD>"
                    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>" & Null2String(rsOperating!Description) & "</FONT></TD>"
                    Print #1, "<TD COLSPAN=2> </TD>"
                    Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(TotalOperatingAmount - BeginningTotalOperatingAmount, "###,###,###,##0.00") & "</FONT></TD>"
                    Print #1, "</TR>"
                End If
                GrandTotalOperatingAmount = GrandTotalOperatingAmount + (TotalOperatingAmount - BeginningTotalOperatingAmount)
            End If
            rsOperating.MoveNext
        Loop
    End If
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=55><FONT SIZE=3 FACE=Times New Roman>Add expense accounts that do not represent cash flows:</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Set rsOperating = New ADODB.Recordset
    Set rsOperating = gconDMIS.Execute("Select * from AMIS_TitleCode Where CashFlowCode = 'OA' AND (HeaderCode = '7' or HeaderCode = '9') Order By Code Asc")
    If Not rsOperating.EOF And Not rsOperating.BOF Then
        rsOperating.MoveFirst:
        Do While Not rsOperating.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles = " & N2Str2Null(rsOperating!Code) & " Order By AcctCode")
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                rsChartAccount.MoveFirst: BeginningTotalOperatingAmount = 0: TotalOperatingAmount = 0
                Do While Not rsChartAccount.EOF
                    Set rsJournal_Det = New ADODB.Recordset
                    Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
                        TotalOperatingAmount = TotalOperatingAmount + N2Str2Zero(rsJournal_Det!TOTAL_AMOUNT)
                    End If
                    rsChartAccount.MoveNext
                Loop
                If TotalOperatingAmount <> 0 Then
                    Print #1, "<TR VALIGN=TOP>"
                    Print #1, "<TD COLSPAN=12> </TD>"
                    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>" & Null2String(rsOperating!Description) & "</FONT></TD>"
                    Print #1, "<TD COLSPAN=2> </TD>"
                    Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(TotalOperatingAmount - BeginningTotalOperatingAmount, "###,###,###,##0.00") & "</FONT></TD>"
                    Print #1, "</TR>"
                End If
                GrandTotalOperatingAmount = GrandTotalOperatingAmount + (TotalOperatingAmount)
            End If
            rsOperating.MoveNext
        Loop
    End If

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Operating income before working capital changes</FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=5 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(V_NetIncomeOrLoss + GrandTotalOperatingAmount, "###,###,###,##0.00") & "</FONT></TD>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Dim IncreaseArrayDescription(1 To 50)         As String
    Dim IncreaseArrayValue(1 To 100)              As Double
    Dim DecreaseArrayDescription(100)             As String
    Dim DecreaseArrayValue(100)                   As Double

    Dim IncreaseArrayCount                        As Integer
    Dim DecreaseArrayCount                        As Integer

    TotalOperatingAmount = 0

    Set rsOperating = New ADODB.Recordset
    Set rsOperating = gconDMIS.Execute("Select * from AMIS_TitleCode Where CashFlowCode = 'OA' AND (HeaderCode = '1' or HeaderCode = '2' or HeaderCode = '3') Order By Code Asc")
    If Not rsOperating.EOF And Not rsOperating.BOF Then
        rsOperating.MoveFirst: IncreaseArrayCount = 0: DecreaseArrayCount = 0
        Do While Not rsOperating.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles = " & N2Str2Null(rsOperating!Code) & " Order By AcctCode")
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                rsChartAccount.MoveFirst: BeginningTotalOperatingAmount = 0: TotalOperatingAmount = 0
                Do While Not rsChartAccount.EOF
                    Set rsJournal_Det = New ADODB.Recordset
                    If Null2String(rsOperating!HeaderCode) = "2" Or Null2String(rsOperating!HeaderCode) = "3" Or Null2String(rsOperating!HeaderCode) = "4" Or Null2String(rsOperating!HeaderCode) = "8" Then
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    Else
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    End If
                    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
                        TotalOperatingAmount = TotalOperatingAmount + N2Str2Zero(rsJournal_Det!TOTAL_AMOUNT)
                    End If
                    rsChartAccount.MoveNext
                Loop
                BeginningTotalOperatingAmount = 0
                If TotalOperatingAmount = 0 Then GoTo Sunod
                'If TotalOperatingAmount > 0 Then
                If Null2String(rsOperating!HeaderCode) = "1" Then
                    GrandTotalOperatingAmount = GrandTotalOperatingAmount - (TotalOperatingAmount - BeginningTotalOperatingAmount)
                    DecreaseArrayCount = DecreaseArrayCount + 1
                    DecreaseArrayDescription(DecreaseArrayCount) = Null2String(rsOperating!Description)
                    DecreaseArrayValue(DecreaseArrayCount) = Round(TotalOperatingAmount - BeginningTotalOperatingAmount, 2)
                Else
                    GrandTotalOperatingAmount = GrandTotalOperatingAmount + (BeginningTotalOperatingAmount - TotalOperatingAmount)
                    IncreaseArrayCount = IncreaseArrayCount + 1
                    IncreaseArrayDescription(IncreaseArrayCount) = Null2String(rsOperating!Description)
                    IncreaseArrayValue(IncreaseArrayCount) = Abs(Round(TotalOperatingAmount - BeginningTotalOperatingAmount, 2))
                End If
                'Else
                '    If Null2String(rsOperating!HeaderCode) = "1" Then
                '        GrandTotalOperatingAmount = GrandTotalOperatingAmount + (TotalOperatingAmount - BeginningTotalOperatingAmount)
                '        DecreaseArrayCount = DecreaseArrayCount + 1
                '        DecreaseArrayDescription(DecreaseArrayCount) = Null2String(rsOperating!Description)
                '        DecreaseArrayValue(DecreaseArrayCount) = 0 - Round(TotalOperatingAmount - BeginningTotalOperatingAmount, 2)
                '    Else
                '        GrandTotalOperatingAmount = GrandTotalOperatingAmount - (TotalOperatingAmount - BeginningTotalOperatingAmount)
                '        DecreaseArrayCount = DecreaseArrayCount + 1
                '        DecreaseArrayDescription(DecreaseArrayCount) = Null2String(rsOperating!Description)
                '        DecreaseArrayValue(DecreaseArrayCount) = 0 - Round(TotalOperatingAmount - BeginningTotalOperatingAmount, 2)
                '    End If
                'End If
            End If
Sunod:
            rsOperating.MoveNext
        Loop
    End If
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=55><FONT SIZE=3 FACE=Times New Roman>Increase in:</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Dim KIM                                       As Integer
    For KIM = 1 To IncreaseArrayCount
        Print #1, "<TR VALIGN=TOP>"
        Print #1, "<TD COLSPAN=12> </TD>"
        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>" & IncreaseArrayDescription(KIM) & "</FONT></TD>"
        Print #1, "<TD COLSPAN=2> </TD>"
        Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(IncreaseArrayValue(KIM), "###,###,###,##0.00") & "</FONT></TD>"
        Print #1, "</TR>"
    Next

    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=55><FONT SIZE=3 FACE=Times New Roman>Decrease in:</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    For KIM = 1 To DecreaseArrayCount
        Print #1, "<TR VALIGN=TOP>"
        Print #1, "<TD COLSPAN=12> </TD>"
        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>" & DecreaseArrayDescription(KIM) & "</FONT></TD>"
        Print #1, "<TD COLSPAN=2> </TD>"
        Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(DecreaseArrayValue(KIM), "###,###,###,##0.00") & "</FONT></TD>"
        Print #1, "</TR>"
    Next

    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>NET CASH PROVIDED BY (USED IN) OPERATING ACTIVITIES</FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=5 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(V_NetIncomeOrLoss + GrandTotalOperatingAmount, "###,###,###,##0.00") & "</FONT></TD>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=52><FONT SIZE=3 FACE=Times New Roman><B>CASH FLOW FROM INVESTING ACTIVITIES</B></FONT></TD>"
    Print #1, "</TR>"

    Set rsInvesting = New ADODB.Recordset
    Set rsInvesting = gconDMIS.Execute("Select * from AMIS_TitleCode Where CashFlowCode = 'IA' Order By Code Asc")
    If Not rsInvesting.EOF And Not rsInvesting.BOF Then
        rsInvesting.MoveFirst: GrandTotalInvestingAmount = 0
        Do While Not rsInvesting.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles = " & N2Str2Null(rsInvesting!Code) & " Order By AcctCode")
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                rsChartAccount.MoveFirst: BeginningTotalInvestingAmount = 0: TotalInvestingAmount = 0
                Do While Not rsChartAccount.EOF
                    Set rsJournal_Det = New ADODB.Recordset
                    If Null2String(rsInvesting!HeaderCode) = "2" Or Null2String(rsInvesting!HeaderCode) = "3" Or Null2String(rsInvesting!HeaderCode) = "4" Or Null2String(rsInvesting!HeaderCode) = "8" Then
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    Else
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    End If
                    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
                        TotalInvestingAmount = TotalInvestingAmount + N2Str2Zero(rsJournal_Det!TOTAL_AMOUNT)
                    End If
                    rsChartAccount.MoveNext
                Loop
                If TotalInvestingAmount - BeginningTotalInvestingAmount = 0 Then GoTo Sunod1
                Print #1, "<TR VALIGN=TOP>"
                Print #1, "<TD COLSPAN=12> </TD>"
                BeginningTotalInvestingAmount = 0
                If TotalInvestingAmount - BeginningTotalInvestingAmount > 0 Then
                    If Null2String(rsInvesting!HeaderCode) = "1" Then
                        GrandTotalInvestingAmount = GrandTotalInvestingAmount - (TotalInvestingAmount - BeginningTotalInvestingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Increase in: " & Null2String(rsInvesting!Description) & "</FONT></TD>"
                    Else
                        GrandTotalInvestingAmount = GrandTotalInvestingAmount + (TotalInvestingAmount - BeginningTotalInvestingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Increase in: " & Null2String(rsInvesting!Description) & "</FONT></TD>"
                    End If
                Else
                    If Null2String(rsInvesting!HeaderCode) = "1" Then
                        GrandTotalInvestingAmount = GrandTotalInvestingAmount + (TotalInvestingAmount - BeginningTotalInvestingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Decrease in: " & Null2String(rsInvesting!Description) & "</FONT></TD>"
                    Else
                        GrandTotalInvestingAmount = GrandTotalInvestingAmount - (TotalInvestingAmount - BeginningTotalInvestingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Decrease in: " & Null2String(rsInvesting!Description) & "</FONT></TD>"
                    End If
                End If
                Print #1, "<TD COLSPAN=2> </TD>"
                Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(TotalInvestingAmount - BeginningTotalInvestingAmount, "###,###,###,##0.00") & "</FONT></TD>"
                Print #1, "</TR>"
            End If
Sunod1:
            rsInvesting.MoveNext
        Loop
    End If

    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>NET CASH PROVIDED BY (USED IN) INVESTING ACTIVITIES</FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=5 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(V_NetIncomeOrLoss + GrandTotalOperatingAmount + GrandTotalInvestingAmount, "###,###,###,##0.00") & "</FONT></TD>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"


    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=52><FONT SIZE=3 FACE=Times New Roman><B>CASH FLOW FROM FINANCING ACTIVITIES</B></FONT></TD>"
    Print #1, "</TR>"

    Set rsFinancing = New ADODB.Recordset
    Set rsFinancing = gconDMIS.Execute("Select * from AMIS_TitleCode Where CashFlowCode = 'FA' Order By Code Asc")
    If Not rsFinancing.EOF And Not rsFinancing.BOF Then
        rsFinancing.MoveFirst: GrandTotalFinancingAmount = 0
        Do While Not rsFinancing.EOF
            Set rsChartAccount = New ADODB.Recordset
            Set rsChartAccount = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles = " & N2Str2Null(rsFinancing!Code) & " Order By AcctCode")
            If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
                rsChartAccount.MoveFirst: BeginningTotalFinancingAmount = 0: TotalFinancingAmount = 0
                Do While Not rsChartAccount.EOF
                    Set rsJournal_Det = New ADODB.Recordset
                    If Null2String(rsFinancing!HeaderCode) = "2" Or Null2String(rsFinancing!HeaderCode) = "3" Or Null2String(rsFinancing!HeaderCode) = "4" Or Null2String(rsFinancing!HeaderCode) = "8" Then
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    Else
                        Set rsJournal_Det = gconDMIS.Execute("Select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as TOTAL_AMOUNT from AMIS_Journal_Det where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.Jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & CDate(dtpTo) & "') AND AMIS_Journal_Det.Acct_Code=" & N2Str2Null(rsChartAccount!ACCTCODE))
                    End If
                    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
                        TotalFinancingAmount = TotalFinancingAmount + N2Str2Zero(rsJournal_Det!TOTAL_AMOUNT)
                    End If
                    rsChartAccount.MoveNext
                Loop
                BeginningTotalFinancingAmount = 0
                If TotalFinancingAmount - BeginningTotalFinancingAmount = 0 Then GoTo Sunod2
                Print #1, "<TR VALIGN=TOP>"
                Print #1, "<TD COLSPAN=12> </TD>"
                If TotalFinancingAmount - BeginningTotalFinancingAmount > 0 Then
                    If Null2String(rsFinancing!HeaderCode) = "1" Then
                        GrandTotalFinancingAmount = GrandTotalFinancingAmount - (TotalFinancingAmount - BeginningTotalFinancingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Increase in: " & Null2String(rsFinancing!Description) & "</FONT></TD>"
                    Else
                        GrandTotalFinancingAmount = GrandTotalFinancingAmount + (TotalFinancingAmount - BeginningTotalFinancingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Increase in: " & Null2String(rsFinancing!Description) & "</FONT></TD>"
                    End If
                Else
                    If Null2String(rsFinancing!HeaderCode) = "1" Then
                        GrandTotalFinancingAmount = GrandTotalFinancingAmount + (TotalFinancingAmount - BeginningTotalFinancingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Decrease in: " & Null2String(rsFinancing!Description) & "</FONT></TD>"
                    Else
                        GrandTotalFinancingAmount = GrandTotalFinancingAmount - (TotalFinancingAmount - BeginningTotalFinancingAmount)
                        Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>Decrease in: " & Null2String(rsFinancing!Description) & "</FONT></TD>"
                    End If
                End If
                Print #1, "<TD COLSPAN=2> </TD>"
                Print #1, "<TD ALIGN=RIGHT COLSPAN=6 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(TotalFinancingAmount - BeginningTotalFinancingAmount, "###,###,###,##0.00") & "</FONT></TD>"
                Print #1, "</TR>"
            End If
Sunod2:
            rsFinancing.MoveNext
        Loop
    End If

    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=9> </TD>"
    Print #1, "<TD COLSPAN=27><FONT SIZE=3 FACE=Times New Roman>NET CASH PROVIDED BY (USED IN) FINANCING ACTIVITIES</FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=5 NOWRAP><FONT SIZE=2 FACE=Times New Roman>" & Format(V_NetIncomeOrLoss + GrandTotalOperatingAmount + GrandTotalInvestingAmount + GrandTotalFinancingAmount, "###,###,###,##0.00") & "</FONT></TD>"
    Print #1, "</TR><TR VALIGN=TOP><TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD></TR>"

    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=30><FONT SIZE=3 FACE=Times New Roman><B>NET INCREASE (DECREASE) IN CASH</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=25><FONT SIZE=3 FACE=Times New Roman><B>AND CASH EQUIVALENTS</B></FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=9 NOWRAP><FONT SIZE=2 FACE=Times New Roman><B>" & Format(GrandTotalOperatingAmount + GrandTotalInvestingAmount + GrandTotalFinancingAmount, "###,###,###,##0.00") & "</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=30><FONT SIZE=3 FACE=Times New Roman><B>CASH AND CASH EQUIVALENTS</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=25><FONT SIZE=3 FACE=Times New Roman><B>AT BEGINNING OF PERIOD</B></FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=9 NOWRAP><FONT SIZE=2 FACE=Times New Roman><B>" & Format(CashBeginningOfPeriod(), "###,###,###,##0.00") & "</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=6> </TD>"
    Print #1, "<TD COLSPAN=30><FONT SIZE=3 FACE=Times New Roman><B>CASH AND CASH EQUIVALENTS</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=11> </TD>"
    Print #1, "<TD COLSPAN=25><FONT SIZE=3 FACE=Times New Roman><B>AT END OF PERIOD</B></FONT></TD>"
    Print #1, "<TD COLSPAN=15> </TD>"
    Print #1, "<TD ALIGN=RIGHT COLSPAN=9 NOWRAP><FONT SIZE=2 FACE=Times New Roman><B>" & Format(CashBeginningOfPeriod + (V_NetIncomeOrLoss + GrandTotalOperatingAmount + GrandTotalInvestingAmount + GrandTotalFinancingAmount), "###,###,###,##0.00") & "</B></FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "<TR VALIGN=TOP>"
    Print #1, "<TD COLSPAN=64><FONT SIZE=1>&nbsp;</FONT></TD>"
    Print #1, "</TR>"
    Print #1, "</TABLE>"
    Print #1, "</BODY></HTML>"
    Close #1
    DoEvents
    On Error GoTo ErrorCode
    Open App.Path & "\CashFlow.html" For Input As #1
    If EOF(1) Then
        Screen.MousePointer = 0
        MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
    Else
        Close #1
        browCashFlow.Navigate App.Path & "\CashFlow.html": DoEvents
        browCashFlow.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DONTPROMPTUSER
        Screen.MousePointer = 0
    End If
    Exit Sub

ErrorCode:
    Resume Next
End Sub

Sub GenerateLastIncomeStatement()
    Dim Last_dtpFrom
    Dim Last_dtpTo
    Dim Prev_dtpFrom, Prev_dtpTo                  As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
        Last_dtpFrom = CDate(Month(dtpFrom) & "/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        If Month(dtpTo) = 2 And Day(dtpTo) = 29 Then
            Last_dtpTo = CDate(Month(dtpTo) & "/" & 28 & "/" & Year(dtpTo) - 1)
        Else
            Last_dtpTo = CDate(Month(dtpTo) & "/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
        End If
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))

    End If

    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where STATUS = 'P' AND (jdate >= '" & Last_dtpFrom & "' AND jdate <= '" & Last_dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then

        rptAMISIncomeStatement.Reset

        '        '=========================================
        '        '================LAST CURRENT ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(31) = "LastCurent_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(32) = "LastCurent_Charge_GrosSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(33) = "lastcurrent_cash_salesdiscountandReturn = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(34) = "lastcurrent_charge_SalesDiscountandReturn = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(35) = "lastCurrent_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(36) = "lastCurrent_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "')" & _
                                             " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '" & ADMIN_SECTION & "')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(37) = "LastCurrent_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '" & ADMIN_SECTION & "'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(38) = "lastCurrent_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(39) = "LastCurrent_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(40) = "LastCurrent_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "MM/DD/YYYY") & "#"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatements.rpt", "{Journal_Hd.jtype} = 'CLO' AND {Journal_Hd.jdate} >= date(" & Year(Last_dtpFrom) & "," & Month(Last_dtpFrom) & "," & Day(Last_dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(Last_dtpTo) & "," & Month(Last_dtpTo) & "," & Day(Last_dtpTo) & ") and year({Journal_Hd.jdate}) = " & Year(Last_dtpTo), DMIS_REPORT_Connection, 1
    Else
        'ShowNoRecord
    End If
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError

End Sub

Sub GenerateLastScheduleOfSellingExpenseSales()


    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    Dim Last_dtpTo
    Dim Last_dtpFrom

    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
        Last_dtpFrom = CDate(Month(dtpFrom) & "/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
        Last_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
        Last_dtpTo = CDate(Month(dtpTo) & "/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    End If


    Dim LastCurrent_NetSales                      As Double
    Dim LastCURRENT_GROSSSALES                    As Double
    Dim LastCURRENT_DISCOUNTSRETURNS              As Double

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    LastCurrent_NetSales = LastCURRENT_GROSSSALES - LastCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
    End If
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(33) = "LastCurrent_NetSales= " & NumericVal(LastCurrent_NetSales)
    rptAMISIncomeStatement.Formulas(31) = "Last_ToJDate = #" & Format(Last_dtpTo, "Short Date") & "#"
    rptAMISIncomeStatement.Formulas(32) = "Last_fromJDate = #" & Format(Last_dtpFrom, "Short Date") & "#"

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Sub GenerateLastScheduleOfSellingExpensePartSection()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    Dim Last_dtpTo
    Dim Last_dtpFrom

    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
        Last_dtpFrom = CDate(Month(dtpFrom) & "/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)

    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
        Last_dtpTo = CDate(Month(dtpTo) & "/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    End If


    Dim LastCURRENT_GROSSSALES                    As Double
    Dim LastCURRENT_DISCOUNTSRETURNS              As Double
    Dim LastCurrent_NetSales                      As Double

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    LastCurrent_NetSales = LastCURRENT_GROSSSALES - LastCURRENT_DISCOUNTSRETURNS

    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
    End If

    rptAMISIncomeStatement.Formulas(32) = "LastCurrent_NetSales = " & NumericVal(LastCurrent_NetSales)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    rptAMISIncomeStatement.Formulas(31) = "Last_ToJDate = #" & Format(Last_dtpTo, "Short Date") & "#"


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Sub GenerateScheduleOfSellingExpenseServiceSection()
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    Dim Last_dtpTo
    Dim Last_dtpFrom

    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
        Last_dtpFrom = CDate(Month(dtpFrom) & "/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
        Last_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
        Last_dtpTo = CDate(Month(dtpTo) & "/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)

    End If

    Dim vCUMULATIVE_NETSALES                      As Double
    Dim vCURRENT_NETSALES                         As Double
    Dim LastCurrent_NetSales                      As Double
    Dim LastCURRENT_GROSSSALES                    As Double
    Dim LastCURRENT_DISCOUNTSRETURNS              As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        LastCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    LastCurrent_NetSales = LastCURRENT_GROSSSALES - LastCURRENT_DISCOUNTSRETURNS
    Dim rsProfile                                 As New ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")

    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(2) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(3) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
    End If
    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    rptAMISIncomeStatement.Formulas(32) = "LastCurrent_NetSales= " & NumericVal(LastCurrent_NetSales)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    'rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    rptAMISIncomeStatement.Formulas(31) = "last_ToJDate = #" & Format(Last_dtpTo, "Short Date") & "#"
End Sub

Sub GenerateIncomeStatementVehicleCumulative()
    Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT         As String
    Dim Last_dtpTo
    Dim Last_dtpFrom
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
        Last_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
        Last_dtpFrom = CDate(Month(dtpFrom) & "/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)

    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
        Last_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
        Last_dtpTo = CDate(Month(dtpTo) & "/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    End If
    rptAMISIncomeStatement.Formulas(43) = "last_ToJDate = #" & Format(Last_dtpTo, "Short Date") & "#"
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where (jdate >= '" & Last_dtpFrom & "' AND jdate <= '" & Last_dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - VEHICLE"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - VEHICLE"
        End If
        PRODUCT = "'" & VEHICLES_SECTION & "'"
        '=========================================
        '================ Last CURRENT ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(33) = "LastCurrent_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(34) = "LastCurrent_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(35) = "LastCurrent_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(36) = "LastCurrent_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(37) = "LastCurrent_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(38) = "LastCurrent_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(39) = "LastCurrent_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(40) = "LastCurrent_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(41) = "LastCurrent_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Last_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Last_dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(42) = "LastCurrent_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        'rptAMISIncomeStatement.Formulas(30) = "ToJDate = '" & Format(dtpTo, "Short Date") & "'"
        'PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementVehiclesCumulative.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "Audit Inquiry (FinancialStatement)"
        Call frmALL_AuditInquiry.DisplayHistory("", "FinancialStatement", "PRINTING")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Height = 9400
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Me.Top = Me.Top + 200
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    Screen.MousePointer = 0

    If COMPANY_CODE = "HAI" Then
        VEHICLES_SECTION = "10"
        PARTS_SECTION = "30"
        SERVICE_SECTION = "20"
        ADMIN_SECTION = "40"

        CASH_SALES = "'41'"
        CHARGE_SALES = "'41'"
        CASH_DISCOUNT = "'51'"
        CHARGE_DISCOUNT = "'52'"
        CASH_COSTOFSALES = "'61'"
        CHARGE_COSTOFSALES = "'62'"
        OPERATIONAL_EXPENSE = "'71'"
        ADMIN_EXPENSE = "'40'"
        OTHER_INCOME = "'81'"
        OTHER_EXPENSE = "'91'"
        CURRENT_ASSET = "'11'"

        TAX_CREDITS = "1107"
        PROPERTY_EQUIPMENT = "1201"
        ACCUMULATED_DEPRECIATION = "1202"
        OTHER_ASSET = "1204"
        
    ElseIf COMPANY_CODE = "HCC" Then
        VEHICLES_SECTION = "10"
        PARTS_SECTION = "20"
        SERVICE_SECTION = "30"
        ADMIN_SECTION = "40"

        CASH_SALES = "'41'"
        CHARGE_SALES = "'41'"
        CASH_DISCOUNT = "'51'"
        CHARGE_DISCOUNT = "'52'"
        CASH_COSTOFSALES = "'61'"
        CHARGE_COSTOFSALES = "'62'"
        OPERATIONAL_EXPENSE = "'71'"
        ADMIN_EXPENSE = "'40'"
        OTHER_INCOME = "'81'"
        OTHER_EXPENSE = "'91'"
        CURRENT_ASSET = "'11'"

        TAX_CREDITS = "1107"
        PROPERTY_EQUIPMENT = "1201"
        ACCUMULATED_DEPRECIATION = "1202"
        OTHER_ASSET = "1204"
    Else
        VEHICLES_SECTION = "10"
        PARTS_SECTION = "30"
        SERVICE_SECTION = "20"
        ADMIN_SECTION = "40"

        CASH_SALES = "'41'"
        CHARGE_SALES = "'41'"
        CASH_DISCOUNT = "'51'"
        CHARGE_DISCOUNT = "'52'"
        CASH_COSTOFSALES = "'61'"
        CHARGE_COSTOFSALES = "'62'"
        OPERATIONAL_EXPENSE = "'71'"
        ADMIN_EXPENSE = "'40'"
        OTHER_INCOME = "'81'"
        OTHER_EXPENSE = "'91'"
        CURRENT_ASSET = "'11'"

        TAX_CREDITS = "1107"
        PROPERTY_EQUIPMENT = "1201"
        ACCUMULATED_DEPRECIATION = "1202"
        OTHER_ASSET = "1204"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAMISRange = Nothing
End Sub

Private Sub dtpFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub dtpTo_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Image1_Click()
    On Error GoTo ErrorCode:
    labBalanceSheets_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Image10_Click()
    labScheduleOfSellingExpensesSalesSection_Click
End Sub

Private Sub Image11_Click()
    labScheduleOfSellingExpensesPartsSection_Click
End Sub

Private Sub Image12_Click()
    labScheduleOfSellingExpensesServiceSection_Click

End Sub

Private Sub Image13_Click()
    labScheduleOAdministrativeExpensesCumulative_Click
End Sub

Private Sub Image14_Click()
    labScheduleOfAdminAndSellingExpensesCumulative_Click
End Sub

Private Sub Image15_Click()
    labScheduleOfAdminAndSellingExpensesCurrent_Click
End Sub

Private Sub Image16_Click()
    labScheduleOfAccounts_Click
End Sub

Private Sub Image17_Click()
    labIncomeStatementVehiclesCumulative_Click
End Sub

Private Sub Image18_Click()
    labIncomeStatementPartsCumulative_Click
End Sub

Private Sub Image19_Click()
    labIncomeStatementServiceCumulative_Click
End Sub

Private Sub Image2_Click()
    On Error GoTo ErrorCode:

    labIncomeStatements_Click
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub Image20_Click()
    labScheduleOfSellingExpensesSalesSection_Click
End Sub

Private Sub Image21_Click()
    labScheduleOfSellingExpensesPartsSection_Click
End Sub

Private Sub Image22_Click()
    labScheduleOfSellingExpensesServiceSection_Click

End Sub

Private Sub Image23_Click()
    labCashFlow_Click
End Sub

Private Sub Image24_Click()
    labOwnersEquity_Click
End Sub

Private Sub Image3_Click()
    On Error GoTo ErrorCode:
    labIncomeStatementVehiclesCumulative_Click
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub Image4_Click()
    labIncomeStatementByProductCumulative_Click
End Sub

Private Sub Image5_Click()
    labIncomeStatementByProductCurrent_Click
End Sub

Private Sub Image6_Click()
    On Error GoTo ErrorCode:

    labIncomeStatementPartsCumulative_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Image7_Click()
    On Error GoTo ErrorCode:

    labScheduleOfSellingExpenseCumulative_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Image8_Click()
    labScheduleOfSellingExpenseCurrent_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Image9_Click()
    On Error GoTo ErrorCode:

    labIncomeStatementServiceCumulative_Click
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labBalanceSheets_Click()
'FOR HAI


    Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
    Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

    Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
    Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double

    Dim Prev_V_GrossSales, Prev_V_SalesDiscountsAndReturns, Prev_V_CostOfSales As Double
    Dim Prev_V_LessSellingExpense, Prev_V_LessAdminExpense, Prev_V_LessOtherExpense, Prev_V_AddOtherIncome As Double

    Dim Prev_Cummulative_Cash_GrossSales, Prev_Cummulative_Charge_GrossSales, Prev_Cummulative_Cash_SalesDiscountsAndReturns, Prev_Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Prev_Cummulative_Cash_CostOfSales, Prev_Cummulative_Charge_CostOfSales, Prev_Cummulative_LessSellingExpense, Prev_Cummulative_LessAdminExpense As Double
    Dim Prev_Cummulative_LessOtherExpense, Prev_Cummulative_AddOtherIncome As Double

    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Sub
    End If

    'NetIncomeOrLoss Previous Period
    Dim PreviousMonthTo                           As String

    'PreviousMonthFrom = DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1
    PreviousMonthTo = DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (jdate <= '" & CDate(PreviousMonthTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Prev_V_GrossSales = Prev_Cummulative_Cash_GrossSales + Prev_Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Prev_V_SalesDiscountsAndReturns = Prev_Cummulative_Cash_SalesDiscountsAndReturns + Prev_Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Prev_V_CostOfSales = Prev_Cummulative_Cash_CostOfSales + Prev_Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Prev_V_LessSellingExpense = Prev_Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Prev_V_LessAdminExpense = Prev_Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Prev_V_LessOtherExpense = Prev_Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        Prev_V_AddOtherIncome = Prev_Cummulative_AddOtherIncome

        Prev_V_NetIncomeOrLoss = (((Prev_V_GrossSales - Prev_V_SalesDiscountsAndReturns) - Prev_V_CostOfSales) - (Prev_V_LessSellingExpense + Prev_V_LessAdminExpense)) - Prev_V_LessOtherExpense + Prev_V_AddOtherIncome
    End If

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "'" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome

        ShowBalanceSheetReport "BalanceSheet", "FinancialStatement\FinancialStatements\", "({Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", "BALANCE SHEETS", "AS OF: " & Format(dtpTo, "long date"), True
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Call NEW_LogAudit("G", "FinancialStatement", "BALANCE SHEET", "", "", "BALANCE SHEET" & "-" & dtpTo, "", "")
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labCashFlow_Click()

    Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
    Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

    Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
    Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double

    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Sub
    End If

    Dim PreviousMonthFrom                         As String
    Dim PreviousMonthTo                           As String

    PreviousMonthFrom = firstDay(DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1)
    PreviousMonthTo = lastDay(PreviousMonthFrom)

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (jdate >= '" & CDate(dtpFrom) & "' AND jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & CDate(dtpFrom) & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome


        'For CashFlow
        Dim CF_Depreciation_And_Depletion_Current As Double

        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles='7115'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Depreciation_And_Depletion_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If

        CF_Depreciation_And_Depletion = CF_Depreciation_And_Depletion_Current

        Dim CF_Provision_For_Income_Tax_Current   As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles='9103'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Provision_For_Income_Tax_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        CF_Provision_For_Income_Tax = CF_Provision_For_Income_Tax_Current

        Dim CF_Allowance_For_Doubtful_Accounts_Current As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles='1104'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Allowance_For_Doubtful_Accounts_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        CF_Allowance_For_Doubtful_Accounts = CF_Allowance_For_Doubtful_Accounts_Current

        Dim CF_Interest_Income_Current            As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles='1104'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Interest_Income_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        CF_Interest_Income = CF_Interest_Income_Current

        '        Dim CF_Gain_On_Disposal_Of_Equipment_Current, CF_Gain_On_Disposal_Of_Equipment_Previous As Double
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Titles='1104'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Gain_On_Disposal_Of_Equipment_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        '        End If
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1104'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Gain_On_Disposal_Of_Equipment_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        '        End If
        '        CF_Gain_On_Disposal_Of_Equipment = CF_Gain_On_Disposal_Of_Equipment_Current - CF_Gain_On_Disposal_Of_Equipment_Previous

        Dim CF_Accounts_Receivable_Trade_Current, CF_Accounts_Receivable_Trade_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1102'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Receivable_Trade_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1102'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Receivable_Trade_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Accounts_Receivable_Trade = CF_Accounts_Receivable_Trade_Previous - CF_Accounts_Receivable_Trade_Current

        Dim CF_Accounts_Receivable_Non_Trade_Current, CF_Accounts_Receivable_Non_Trade_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1103'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Receivable_Non_Trade_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1103'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Receivable_Non_Trade_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Accounts_Receivable_Non_Trade = CF_Accounts_Receivable_Non_Trade_Previous - CF_Accounts_Receivable_Non_Trade_Current

        Dim CF_Inventories_Current, CF_Inventories_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1105'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Inventories_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1105'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Inventories_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Inventories = CF_Inventories_Previous - CF_Inventories_Current

        Dim CF_Prepaid_Expenses_Current, CF_Prepaid_Expenses_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1106'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Prepaid_Expenses_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1106'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Prepaid_Expenses_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Prepaid_Expenses = CF_Prepaid_Expenses_Previous - CF_Prepaid_Expenses_Current

        Dim CF_Tax_Credits_Current, CF_Tax_Credits_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1107'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Tax_Credits_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1107'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Tax_Credits_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Tax_Credits = CF_Tax_Credits_Previous - CF_Tax_Credits_Current

        Dim CF_Other_Assets_Current, CF_Other_Assets_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1204'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Other_Assets_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1204'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Other_Assets_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Other_Assets = CF_Other_Assets_Previous - CF_Other_Assets_Current



        Dim CF_Accounts_Payable_Trade_Current, CF_Accounts_Payable_Trade_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2101'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Payable_Trade_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2101'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Payable_Trade_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Accounts_Payable_Trade = CF_Accounts_Payable_Trade_Previous - CF_Accounts_Payable_Trade_Current

        Dim CF_Accounts_Payable_Non_Trade_Current, CF_Accounts_Payable_Non_Trade_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2102'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Payable_Non_Trade_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2102'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accounts_Payable_Non_Trade_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Accounts_Payable_Non_Trade = CF_Accounts_Payable_Non_Trade_Previous - CF_Accounts_Payable_Non_Trade_Current

        Dim CF_Accrued_Expenses_Current, CF_Accrued_Expenses_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2103'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accrued_Expenses_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2103'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Accrued_Expenses_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Accrued_Expenses = CF_Accrued_Expenses_Previous - CF_Accrued_Expenses_Current

        Dim CF_Remittances_Payable_Current, CF_Remittances_Payable_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2104'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Remittances_Payable_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2104'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Remittances_Payable_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Remittances_Payable = CF_Remittances_Payable_Previous - CF_Remittances_Payable_Current

        Dim CF_Taxes_Liabilities_Current, CF_Taxes_Liabilities_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2105'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Taxes_Liabilities_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2105'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Taxes_Liabilities_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Taxes_Liabilities = CF_Taxes_Liabilities_Previous - CF_Taxes_Liabilities_Current

        Dim CF_Other_Payables_Current, CF_Other_Payables_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Titles='2106' OR AMIS_ChartAccount.Titles='2107')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Other_Payables_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND (AMIS_ChartAccount.Titles='2106' OR AMIS_ChartAccount.Titles='2107')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Other_Payables_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Other_Payables = CF_Other_Payables_Previous - CF_Other_Payables_Current

        '        Dim CF_Loans_Payable_Current, CF_Loans_Payable_Previous As Double
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='2101'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Loans_Payable_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        '        End If
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='2101'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Loans_Payable_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        '        End If
        '        CF_Loans_Payable = CF_Loans_Payable_Current - CF_Loans_Payable_Previous

        '        Dim CF_Interest_Income_Received_Current, CF_Interest_Income_Received_Previous As Double
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='8104'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Interest_Income_Received_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        '        End If
        '        Set rsJournal_Det = New ADODB.Recordset
        '        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='8104'")
        '        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        '            CF_Interest_Income_Received_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        '        End If
        '        CF_Interest_Income_Received = CF_Interest_Income_Received_Current - CF_Interest_Income_Received_Previous

        Dim CF_Acquiositions_Of_Propert_And_Equipment_Current, CF_Acquiositions_Of_Propert_And_Equipment_Previous As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Titles='1201'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Acquiositions_Of_Propert_And_Equipment_Current = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1201'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Acquiositions_Of_Propert_And_Equipment_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Acquiositions_Of_Propert_And_Equipment = CF_Acquiositions_Of_Propert_And_Equipment_Previous - CF_Acquiositions_Of_Propert_And_Equipment_Current

        Dim CF_Cash_At_Beginning_Previous         As Double
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.Titles='1101'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            CF_Cash_At_Beginning_Previous = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        CF_Cash_At_Beginning = CF_Cash_At_Beginning_Previous

        ShowCashFlowReport "StatementOfCashFlow", "FinancialStatement\FinancialStatements\", "({Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", "STATEMENT OF CASH FLOW", "AS OF: " & Format(dtpTo, "long date"), True
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "StatementOfCashFlow", "", "", "Statement Of CashFlow" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labIncomeStatementByProductCumulative_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim PRODUCT                                   As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If

    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where Status = 'P' and jdate <= '" & dtpTo & "' and year(jdate) = " & Year(dtpTo))
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        rptAMISIncomeStatement.Reset
        Dim rsProfile                             As ADODB.Recordset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - BY PRODUCT - CUMMULATIVE"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - BY PRODUCT - CUMMULATIVE"
        End If
        PRODUCT = "'" & VEHICLES_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.5
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='81') AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        PRODUCT = "'" & PARTS_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.2
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='81') AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        PRODUCT = "'" & SERVICE_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.3
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementByProductCumulative.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "Income Statement By Product Cumulative", "", "", "Income Statement By Product Cumulative" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labIncomeStatementByProductCurrent_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim PRODUCT                                   As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where STATUS = 'P' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - BY PRODUCT - CURRENT"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - BY PRODUCT - CURRENT"
        End If
        PRODUCT = "'" & VEHICLES_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.5
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        PRODUCT = "'" & PARTS_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.2
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        PRODUCT = "'" & SERVICE_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.3
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementByProductCurrent.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "Income Statement By Product Current", "", "", "Income Statement By Product Current" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labIncomeStatementPartsCumulative_Click()
    If CheckDate = False Then Exit Sub
    Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT         As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where STATUS = 'P' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - PARTS & ACCESSORIES / MOTOR OIL & LUBRICANTS"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - PARTS & ACCESSORIES / MOTOR OIL & LUBRICANTS"
        End If
        PRODUCT = "'" & PARTS_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.5
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.5
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense) * 0.5
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = '" & Format(dtpTo, "Short Date") & "'"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementPartsCumulative.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
        Set rsJournal_Det = Nothing

    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "Income Statement Vehicles Cumulative", "", "", "Income Statement Vehicles Cumulative" & "-" & dtpTo, "", "")
End Sub

Private Sub labIncomeStatements_Click()


    Dim Prev_dtpFrom, Prev_dtpTo                  As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))

    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))

    End If

    Dim Prev_V_GrossSales, Prev_V_SalesDiscountsAndReturns, Prev_V_CostOfSales As Double
    Dim Prev_V_LessSellingExpense, Prev_V_LessAdminExpense, Prev_V_LessOtherExpense, Prev_V_AddOtherIncome As Double

    Dim Prev_Cummulative_Cash_GrossSales, Prev_Cummulative_Charge_GrossSales, Prev_Cummulative_Cash_SalesDiscountsAndReturns, Prev_Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Prev_Cummulative_Cash_CostOfSales, Prev_Cummulative_Charge_CostOfSales, Prev_Cummulative_LessSellingExpense, Prev_Cummulative_LessAdminExpense As Double
    Dim Prev_Cummulative_LessOtherExpense, Prev_Cummulative_AddOtherIncome As Double

    'NetIncomeOrLoss Previous Period
    Dim PreviousMonthTo                           As String

    'PreviousMonthFrom = DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1
    PreviousMonthTo = DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (jdate <= '" & CDate(PreviousMonthTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Prev_V_GrossSales = Prev_Cummulative_Cash_GrossSales + Prev_Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Prev_V_SalesDiscountsAndReturns = Prev_Cummulative_Cash_SalesDiscountsAndReturns + Prev_Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Prev_V_CostOfSales = Prev_Cummulative_Cash_CostOfSales + Prev_Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Prev_V_LessSellingExpense = Prev_Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Prev_V_LessAdminExpense = Prev_Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Prev_V_LessOtherExpense = Prev_Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & CDate(PreviousMonthTo) & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Prev_Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        Prev_V_AddOtherIncome = Prev_Cummulative_AddOtherIncome

        Prev_V_NetIncomeOrLoss = (((Prev_V_GrossSales - Prev_V_SalesDiscountsAndReturns) - Prev_V_CostOfSales) - (Prev_V_LessSellingExpense + Prev_V_LessAdminExpense)) - Prev_V_LessOtherExpense + Prev_V_AddOtherIncome
    End If

    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where STATUS = 'P' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")

    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(90) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(91) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENTS"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENTS"
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode <> '" & ADMIN_SECTION & "')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If

        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '" & ADMIN_SECTION & "'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND Year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ CURRENT ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "')" & _
                                             " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '" & ADMIN_SECTION & "')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '" & ADMIN_SECTION & "'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ PREVIOUS ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='41'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='51'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='52'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='61' OR AMIS_ChartAccount.Headers='63')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='62'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "')" & _
                                             " AND ((AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.DepartmentCode <> '" & ADMIN_SECTION & "')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND (AMIS_ChartAccount.Headers='71' OR AMIS_ChartAccount.Headers='72') AND AMIS_ChartAccount.DepartmentCode = '" & ADMIN_SECTION & "'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='91' AND (AMIS_ChartAccount.AcctCode <> '91-03000-00' AND AMIS_ChartAccount.AcctCode <> '91-04000-00')")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_Journal_Det.Jtype <> 'CLO' AND AMIS_ChartAccount.Headers='81'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "MM/DD/YYYY") & "#"
        rptAMISIncomeStatement.Formulas(31) = "Previous_MonthYearGrossSales = " & Prev_V_GrossSales
        rptAMISIncomeStatement.Formulas(32) = "Previous_MonthYearSalesDiscountsAndReturns = " & Prev_V_SalesDiscountsAndReturns
        rptAMISIncomeStatement.Formulas(33) = "Previous_MonthYearCostOfSales = " & Prev_V_CostOfSales
        rptAMISIncomeStatement.Formulas(34) = "Previous_MonthYearLessSellingExpense = " & Prev_V_LessSellingExpense
        rptAMISIncomeStatement.Formulas(35) = "Previous_MonthYearLessAdminExpense = " & Prev_V_LessAdminExpense
        rptAMISIncomeStatement.Formulas(36) = "Previous_MonthYearLessOtherExpense = " & Prev_V_LessOtherExpense
        rptAMISIncomeStatement.Formulas(37) = "Previous_MonthYearAddOtherIncome = " & Prev_V_AddOtherIncome
        rptAMISIncomeStatement.Formulas(38) = "CurrentMonthYear = DateSerial (" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatements.rpt", "{Journal_Hd.jtype} = 'CLO' AND {Journal_Hd.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Hd.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and year({Journal_Hd.jdate}) = " & Year(dtpTo), DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord

    End If
    Call NEW_LogAudit("G", "FinancialStatement", "INCOME STATEMENT", "", "", "INCOME STATEMENT" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labIncomeStatementServiceCumulative_Click()
    If CheckDate = False Then Exit Sub
    Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT         As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where year(jdate) = " & Year(dtpTo) & " and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - SERVICE"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - SERVICE"
        End If
        PRODUCT = "'" & SERVICE_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND (AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ PREVIOUS ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND (AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = '" & Format(dtpTo, "Short Date") & "'"
        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementserviceCumulative.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "Income Statementservice Cumulative", "", "", "Income Statementservice Cumulative" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing

End Sub

Private Sub labIncomeStatementVehiclesCumulative_Click()
    If CheckDate = False Then Exit Sub
    Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT         As String
    If dtpFrom > dtpTo Then
        MsgSpeechBox "Error In From and To date"
        Exit Sub
    End If
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    GenerateIncomeStatementVehicleCumulative
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("select * from AMIS_Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then
        Dim rsProfile                             As ADODB.Recordset
        rptAMISIncomeStatement.Reset
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
        If Not (rsProfile.EOF And rsProfile.BOF) Then
            rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
            rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
            rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - VEHICLE"
            rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - VEHICLE"
        End If

        PRODUCT = "'" & VEHICLES_SECTION & "'"
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ CURRENT ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        '================ PREVIOUS ================
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='41' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='42' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='51' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='52' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='61' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='62' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='71' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='72' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='91' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "') AND AMIS_ChartAccount.Headers='81' AND AMIS_ChartAccount.DepartmentCode = " & PRODUCT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        '=========================================
        rptAMISIncomeStatement.Formulas(30) = "ToJDate = '" & Format(dtpTo, "Short Date") & "'"

        PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementVehiclesCumulative.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "Income Statement Vehicles Cumulative", "", "", "Income Statement Vehicles Cumulative" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing

End Sub

Private Sub labOwnersEquity_Click()

    Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
    Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

    Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
    Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
    Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double

    If IsDate(dtpTo) = False Then
        MsgSpeechBox "Error In Date"
        Exit Sub
    End If

    Dim PreviousMonthFrom                         As String
    Dim PreviousMonthTo                           As String

    PreviousMonthFrom = firstDay(DateSerial(Year(dtpFrom), Month(dtpFrom), Day(dtpFrom)) - 1)
    PreviousMonthTo = lastDay(PreviousMonthFrom)

    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_Det where Status = 'P' and (jdate <= '" & CDate(dtpTo) & "')", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_HD.EOF And Not rsJournal_HD.EOF Then

        'Previous
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome
        OE_Prev_NetIncome = V_NetIncomeOrLoss


        'Current

        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Cash_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CASH_SALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_GrossSales = N2Str2Zero(rsJournal_Det!Cash_GrossSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Charge_GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers='42'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_GrossSales = N2Str2Zero(rsJournal_Det!Charge_GrossSales)
        End If
        V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CASH_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CHARGE_DISCOUNT)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
        End If
        V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND (AMIS_ChartAccount.Headers='63' OR AMIS_ChartAccount.Headers=" & CASH_COSTOFSALES & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Cash_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as CostOfSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & CHARGE_COSTOFSALES)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_Charge_CostOfSales = N2Str2Zero(rsJournal_Det!CostOfSales)
        End If
        V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessSellingExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode" & _
                                             " where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "')" & _
                                             " AND (AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode <> " & ADMIN_EXPENSE & ")")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessSellingExpense = N2Str2Zero(rsJournal_Det!LessSellingExpense)
        End If
        V_LessSellingExpense = Cummulative_LessSellingExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessAdminExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OPERATIONAL_EXPENSE & " AND AMIS_ChartAccount.DepartmentCode = " & ADMIN_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessAdminExpense = N2Str2Zero(rsJournal_Det!LessAdminExpense)
        End If
        V_LessAdminExpense = Cummulative_LessAdminExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as LessOtherExpense from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OTHER_EXPENSE)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_LessOtherExpense = N2Str2Zero(rsJournal_Det!LessOtherExpense)
        End If
        V_LessOtherExpense = Cummulative_LessOtherExpense
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as AddOtherIncome from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.Headers=" & OTHER_INCOME)
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            Cummulative_AddOtherIncome = N2Str2Zero(rsJournal_Det!AddOtherIncome)
        End If
        V_AddOtherIncome = Cummulative_AddOtherIncome

        V_NetIncomeOrLoss = (((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense + V_AddOtherIncome
        OE_NetIncome = V_NetIncomeOrLoss


        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.AcctCode='31-00001-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_Prev_CapitalStock = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.AcctCode='31-00001-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_CapitalStock = N2Str2Zero(rsJournal_Det!Value_Current)
        End If

        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate <= '" & PreviousMonthTo & "' AND AMIS_ChartAccount.AcctCode='31-00002-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_Prev_BalanceBeg = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Previous from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate < '" & dtpFrom & "' AND AMIS_ChartAccount.AcctCode='31-00002-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_BalanceBeg = N2Str2Zero(rsJournal_Det!Value_Previous)
        End If


        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate < '" & dtpFrom & "') AND AMIS_ChartAccount.AcctCode='31-00200-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_Prev_ClearingTB = N2Str2Zero(rsJournal_Det!Value_Current)
        End If
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as Value_Current from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND (AMIS_Journal_Det.jdate <= '" & dtpTo & "') AND AMIS_ChartAccount.AcctCode='31-00200-00'")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
            OE_ClearingTB = N2Str2Zero(rsJournal_Det!Value_Current)
        End If


        ShowOwnersEquityReport "StatementOwnersEquity", "FinancialStatement\FinancialStatements\", "({Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", "STATEMENT OF CASH FLOW", "AS OF: " & Format(dtpTo, "long date"), True
        Screen.MousePointer = 0
    Else
        ShowNoRecord
    End If
    Call NEW_LogAudit("G", "FinancialStatement", "StatementOwnersEquity", "", "", "Statement of Owners Equity" & "-" & dtpTo, "", "")
    Set rsJournal_Det = Nothing
    Set rsJournal_HD = Nothing
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOAdministrativeExpensesCumulative_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Dim vCUMULATIVE_GROSSSALES, vCUMULATIVE_DISCOUNTSRETURNS, vCUMULATIVE_NETSALES As Double
    Dim vCURRENT_GROSSSALES, vCURRENT_DISCOUNTSRETURNS, vCURRENT_NETSALES As Double
    Dim vPREVIOUS_GROSSSALES, vPREVIOUS_DISCOUNTSRETURNS, vPREVIOUS_NETSALES As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
    End If
    Set rsJournal_Det = Nothing

    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdministrativeExpensesCumulative.rpt", "{Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Administrative Expenses Cumulative", "", "", "Schedule Of Administrative Expenses Cumulative" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfAccounts_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder                              As String
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    rptAMISIncomeStatement.WindowShowSearchBtn = True
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(1) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(2) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ACCOUNTS"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ACCOUNTS"
        rptAMISIncomeStatement.Formulas(3) = "ReportDate = '" & "As of: " & Format(dtpTo, "long date") & "'"
    End If
    ReportFolder = "FinancialStatement\FinancialStatements\"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAccounts.rpt", "{Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "ScheduleOfAccounts", "", "", "ScheduleOfAccounts" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:17
Private Sub labScheduleOfAdminAndSellingExpensesCumulative_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder                              As String
    Dim vSALES_CUMULATIVE_NETSALES, vSALES_CUMULATIVE_GROSSSALES, vSALES_CUMULATIVE_DISCOUNTSRETURNS As Double
    Dim vPARTS_CUMULATIVE_NETSALES, vPARTS_CUMULATIVE_GROSSSALES, vPARTS_CUMULATIVE_DISCOUNTSRETURNS As Double
    Dim vSERVICE_CUMULATIVE_NETSALES, vSERVICE_CUMULATIVE_GROSSSALES, vSERVICE_CUMULATIVE_DISCOUNTSRETURNS As Double
    Dim vADMIN_CUMULATIVE_NETSALES, vADMIN_CUMULATIVE_GROSSSALES, vADMIN_CUMULATIVE_DISCOUNTSRETURNS As Double
    Dim SALES_PRODUCT                             As String
    Dim SERVICE_PRODUCT                           As String
    Dim PARTS_PRODUCT                             As String
    Dim ADMIN_CODE                                As String

    SALES_PRODUCT = "'" & VEHICLES_SECTION & "'"
    SERVICE_PRODUCT = "'" & SERVICE_SECTION & "'"
    PARTS_PRODUCT = "'" & PARTS_SECTION & "'"
    ADMIN_CODE = "'" & ADMIN_SECTION & "'"

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & SALES_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & SALES_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & PARTS_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & PARTS_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & SERVICE_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & SERVICE_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vADMIN_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vADMIN_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vSALES_CUMULATIVE_NETSALES = vSALES_CUMULATIVE_GROSSSALES - vSALES_CUMULATIVE_DISCOUNTSRETURNS
    vPARTS_CUMULATIVE_NETSALES = vPARTS_CUMULATIVE_GROSSSALES - vPARTS_CUMULATIVE_DISCOUNTSRETURNS
    vSERVICE_CUMULATIVE_NETSALES = vSERVICE_CUMULATIVE_GROSSSALES - vSERVICE_CUMULATIVE_DISCOUNTSRETURNS
    vADMIN_CUMULATIVE_NETSALES = vADMIN_CUMULATIVE_GROSSSALES - vADMIN_CUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing

    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CUMULATIVE"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CUMULATIVE"
    End If
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdminAndSellingExpensesCummulative.rpt", "year({Journal_Det.jdate}) = " & Year(dtpTo) & " and {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Admin And SellingExpenses Cummulative", "", "", "Schedule Of Admin And SellingExpenses Cummulative" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfAdminAndSellingExpensesCurrent_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder                              As String
    Dim vSALES_CURRENT_NETSALES, vSALES_CURRENT_GROSSSALES, vSALES_CURRENT_DISCOUNTSRETURNS As Double
    Dim vPARTS_CURRENT_NETSALES, vPARTS_CURRENT_GROSSSALES, vPARTS_CURRENT_DISCOUNTSRETURNS As Double
    Dim vSERVICE_CURRENT_NETSALES, vSERVICE_CURRENT_GROSSSALES, vSERVICE_CURRENT_DISCOUNTSRETURNS As Double
    Dim vADMIN_CURRENT_NETSALES, vADMIN_CURRENT_GROSSSALES, vADMIN_CURRENT_DISCOUNTSRETURNS As Double
    Dim SALES_PRODUCT                             As String
    Dim SERVICE_PRODUCT                           As String
    Dim PARTS_PRODUCT                             As String
    Dim ADMIN_CODE                                As String

    SALES_PRODUCT = "'" & VEHICLES_SECTION & "'"
    SERVICE_PRODUCT = "'" & SERVICE_SECTION & "'"
    PARTS_PRODUCT = "'" & PARTS_SECTION & "'"
    ADMIN_CODE = "'" & ADMIN_SECTION & "'"

    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & SALES_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_CURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & SALES_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & PARTS_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_CURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & PARTS_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = " & SERVICE_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_CURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = " & SERVICE_PRODUCT)
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vADMIN_CURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52')")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vADMIN_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vSALES_CURRENT_NETSALES = vSALES_CURRENT_GROSSSALES - vSALES_CURRENT_DISCOUNTSRETURNS
    vPARTS_CURRENT_NETSALES = vPARTS_CURRENT_GROSSSALES - vPARTS_CURRENT_DISCOUNTSRETURNS
    vSERVICE_CURRENT_NETSALES = vSERVICE_CURRENT_GROSSSALES - vSERVICE_CURRENT_DISCOUNTSRETURNS
    vADMIN_CURRENT_NETSALES = vADMIN_CURRENT_GROSSSALES - vADMIN_CURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CURRENT"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CURRENT"
    End If
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdminAndSellingExpensesCurrent.rpt", "{Journal_Det.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Admin And Selling Expenses Current", "", "", "Schedule Of Admin And Selling Expenses Current" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfSellingExpenseCumulative_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder                              As String
    Dim vSALES_NETSALES, vSALES_GROSSSALES, vSALES_DISCOUNTSRETURNS As Double
    Dim vPARTS_NETSALES, vPARTS_GROSSSALES, vPARTS_DISCOUNTSRETURNS As Double
    Dim vSERVICE_NETSALES, vSERVICE_GROSSSALES, vSERVICE_DISCOUNTSRETURNS As Double
    Dim vTOTAL_NETSALES                           As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vSALES_NETSALES = vSALES_GROSSSALES - vSALES_DISCOUNTSRETURNS
    vPARTS_NETSALES = vPARTS_GROSSSALES - vPARTS_DISCOUNTSRETURNS
    vSERVICE_NETSALES = vSERVICE_GROSSSALES - vSERVICE_DISCOUNTSRETURNS
    vTOTAL_NETSALES = vSALES_NETSALES + vPARTS_NETSALES + vSERVICE_NETSALES
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - CUMULATIVE"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - CUMULATIVE"
    End If
    ReportFolder = "FinancialStatement\FinancialStatements\"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensesCumulative.rpt", "year({Journal_Det.jdate}) = " & Year(dtpTo) & " and {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Selling Expenses Cumulative", "", "", "Schedule Of Selling Expenses Cumulative" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfSellingExpenseCurrent_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder                              As String
    Dim vSALES_NETSALES, vSALES_GROSSSALES, vSALES_DISCOUNTSRETURNS As Double
    Dim vPARTS_NETSALES, vPARTS_GROSSSALES, vPARTS_DISCOUNTSRETURNS As Double
    Dim vSERVICE_NETSALES, vSERVICE_GROSSSALES, vSERVICE_DISCOUNTSRETURNS As Double
    Dim vTOTAL_NETSALES                           As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSALES_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPARTS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vSERVICE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vSALES_NETSALES = vSALES_GROSSSALES - vSALES_DISCOUNTSRETURNS
    vPARTS_NETSALES = vPARTS_GROSSSALES - vPARTS_DISCOUNTSRETURNS
    vSERVICE_NETSALES = vSERVICE_GROSSSALES - vSERVICE_DISCOUNTSRETURNS
    vTOTAL_NETSALES = vSALES_NETSALES + vPARTS_NETSALES + vSERVICE_NETSALES
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - CURRENT"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - CURRENT"
    End If

    ReportFolder = "FinancialStatement\FinancialStatements\"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensesCurrent.rpt", "{Journal_Det.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    Call NEW_LogAudit("G", "FinancialStatement", "ScheduleOfSellingExpensesCurrent", "", "", "Schedule Of Selling Expenses Current" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfSellingExpensesPartsSection_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Dim vCUMULATIVE_GROSSSALES, vCUMULATIVE_DISCOUNTSRETURNS, vCUMULATIVE_NETSALES As Double
    Dim vCURRENT_GROSSSALES, vCURRENT_DISCOUNTSRETURNS, vCURRENT_NETSALES As Double
    Dim vPREVIOUS_GROSSSALES, vPREVIOUS_DISCOUNTSRETURNS, vPREVIOUS_NETSALES As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & PARTS_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    'GenerateLastScheduleOfSellingExpensePartSection
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
    End If
    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensePartsSection.rpt", "{ChartAccount.DepartmentCode} = '" & PARTS_SECTION & "' AND {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Selling Expense Parts Section", "", "", "Schedule Of Selling Expense Parts Section" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfSellingExpensesSalesSection_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String


    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)

    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))

    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)

    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))

    End If



    Dim vCUMULATIVE_GROSSSALES, vCUMULATIVE_DISCOUNTSRETURNS, vCUMULATIVE_NETSALES As Double
    Dim vCURRENT_GROSSSALES, vCURRENT_DISCOUNTSRETURNS, vCURRENT_NETSALES As Double
    Dim vPREVIOUS_GROSSSALES, vPREVIOUS_DISCOUNTSRETURNS, vPREVIOUS_NETSALES As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & VEHICLES_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        'rptAMISIncomeStatement.Formulas(5) = "ToJdate = '" & CDate(dtpTo) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
    End If
    'GenerateLastScheduleOfSellingExpenseSales
    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpenseSalesSection.rpt", "{ChartAccount.DepartmentCode} = '" & VEHICLES_SECTION & "' AND {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Selling Expense Sales Section", "", "", "Schedule Of Selling ExpenseSales Section" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub labScheduleOfSellingExpensesServiceSection_Click()
    On Error GoTo ErrorCode:

    If CheckDate = False Then Exit Sub
    Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo    As String
    If Month(dtpFrom) = 1 Then
        Prev_dtpFrom = CDate("12/" & Day(dtpFrom) & "/" & Year(dtpFrom) - 1)
    Else
        Prev_dtpFrom = CDate(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom))
    End If
    If Month(dtpTo) = 1 Then
        Prev_dtpTo = CDate("12/" & Day(dtpTo) & "/" & Year(dtpTo) - 1)
    Else
        Prev_dtpTo = lastDay(Format(Month(dtpFrom) - 1 & "/" & Day(dtpFrom) & "/" & Year(dtpFrom), "short date"))
    End If
    Dim vCUMULATIVE_GROSSSALES, vCUMULATIVE_DISCOUNTSRETURNS, vCUMULATIVE_NETSALES As Double
    Dim vCURRENT_GROSSSALES, vCURRENT_DISCOUNTSRETURNS, vCURRENT_NETSALES As Double
    Dim vPREVIOUS_GROSSSALES, vPREVIOUS_DISCOUNTSRETURNS, vPREVIOUS_NETSALES As Double

    Dim rsJournal_Det                             As ADODB.Recordset
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND year(AMIS_Journal_Det.jdate) = " & Year(dtpTo) & " and AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Credit) - SUM(AMIS_Journal_Det.Debit) as GrossSales from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='41' OR AMIS_ChartAccount.Headers='42') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_GROSSSALES = N2Str2Zero(rsJournal_Det!GrossSales)
    End If
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select SUM(AMIS_Journal_Det.Debit) - SUM(AMIS_Journal_Det.Credit) as SalesDiscountsAndReturns from AMIS_Journal_Det inner Join AMIS_ChartAccount on AMIS_Journal_Det.Acct_Code = AMIS_ChartAccount.AcctCode where AMIS_Journal_Det.Status = 'P' AND AMIS_Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND AMIS_Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (AMIS_ChartAccount.Headers='51' OR AMIS_ChartAccount.Headers='52') AND AMIS_ChartAccount.DepartmentCode = '" & SERVICE_SECTION & "'")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.EOF Then
        vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJournal_Det!SalesDiscountsAndReturns)
    End If
    vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
    Set rsJournal_Det = Nothing
    Dim rsProfile                                 As ADODB.Recordset
    rptAMISIncomeStatement.Reset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    'GenerateScheduleOfSellingExpenseServiceSection
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        rptAMISIncomeStatement.Formulas(2) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        rptAMISIncomeStatement.Formulas(3) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
        rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
    End If
    rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
    rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
    ReportFolder = "FinancialStatement\FinancialStatements\"
    rptAMISIncomeStatement.Formulas(30) = "ToJDate = #" & Format(dtpTo, "Short Date") & "#"
    PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpenseServiceSection.rpt", "{ChartAccount.DepartmentCode} = '" & SERVICE_SECTION & "' AND year({Journal_Det.jdate}) = " & Year(dtpTo) & " and {Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", DMIS_REPORT_Connection, 1
    Call NEW_LogAudit("G", "FinancialStatement", "Schedule Of Selling Expense Service Section", "", "", "Schedule Of Selling Expense Service Section" & "-" & dtpTo, "", "")
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub ScrollBar1_Change()
    Picture2.Top = 0 - ScrollBar1.Value
End Sub

Public Sub ShowISReport(ReportName As Variant, ReportFolder As Variant, filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
    If CheckDate = False Then Exit Sub
    Screen.MousePointer = 11
    Dim rsProfile                                 As ADODB.Recordset
    Dim CrystalRpt                                As Crystal.CrystalReport
    frmMain.rptMain.Reset
    Set CrystalRpt = frmMain.rptMain
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_PROFILE")
    If Not (rsProfile.EOF And rsProfile.BOF) Then
        CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rtp"
        CrystalRpt.Formulas(0) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
        CrystalRpt.Formulas(1) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
        If WithDate = True Then CrystalRpt.Formulas(2) = "ReportDate = '" & REPORT_DATE & "'"
        CrystalRpt.Formulas(3) = "NetIncomeOrLoss = " & V_NetIncomeOrLoss
        CrystalRpt.Formulas(4) = "NetIncomeOrLoss2 = " & V_NetIncomeOrLoss
        CrystalRpt.Formulas(5) = "ProvisionForBonus = " & V_ProvisionForBonus
        CrystalRpt.Formulas(6) = "ProvisionForTax = " & V_ProvisionForTax
        CrystalRpt.Formulas(7) = "TOTAL_CURRENT_ASSET = " & V_Total_Current_Asset
        CrystalRpt.Formulas(8) = "NET_PROPERTY_EQUIPMENT = " & V_Net_Propert_Equipment
        CrystalRpt.Formulas(9) = "OTHER_ASSETS = " & V_Other_Assets
        CrystalRpt.Formulas(10) = "Tax_Credits = " & V_TaxCredit
        CrystalRpt.Formulas(11) = "CurrentMonthYear = date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")"
        CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
        PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", filter, DMIS_REPORT_Connection, 1
        CrystalRpt.PageZoom 89
    End If
    Screen.MousePointer = 0
End Sub

