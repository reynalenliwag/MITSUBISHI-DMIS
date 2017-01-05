VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAMISHYUNDAIFinancialStatements 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "HYUNDAI - Financial Statements"
   ClientHeight    =   18135
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10080
   ForeColor       =   &H00FFFFFF&
   Icon            =   "HYUNDAIFinancialStatements.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   18135
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   18000
      Left            =   120
      ScaleHeight     =   18000
      ScaleWidth      =   9405
      TabIndex        =   5
      Top             =   2730
      Width           =   9405
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   17400
         Left            =   0
         ScaleHeight     =   17400
         ScaleWidth      =   9405
         TabIndex        =   6
         Top             =   0
         Width           =   9405
         Begin Crystal.CrystalReport rptAMISIncomeStatement 
            Left            =   90
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowTitle     =   "Income Statements"
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":030A
            MousePointer    =   99  'Custom
            TabIndex        =   11
            Top             =   9270
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":0614
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   9270
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
            Left            =   2220
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":091E
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   11280
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":0C28
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   13830
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":0F32
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   13830
            Width           =   3375
         End
         Begin VB.Image Image22 
            Height          =   720
            Left            =   8190
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":123C
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":138E
            Top             =   6510
            Width           =   585
         End
         Begin VB.Image Image21 
            Height          =   660
            Left            =   5010
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":2A50
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":2BA2
            Top             =   6480
            Width           =   1095
         End
         Begin VB.Image Image20 
            Height          =   615
            Left            =   1500
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":51B4
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":5306
            Top             =   6540
            Width           =   1755
         End
         Begin VB.Image Image19 
            Height          =   720
            Left            =   8070
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":8BA8
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":8CFA
            Top             =   2520
            Width           =   585
         End
         Begin VB.Image Image18 
            Height          =   660
            Left            =   5010
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":A3BC
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":A50E
            Top             =   2550
            Width           =   1095
         End
         Begin VB.Image Image17 
            Height          =   615
            Left            =   1410
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":CB20
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":CC72
            Top             =   2610
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
            Left            =   4650
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":10514
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   1440
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
            Left            =   1830
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1081E
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   1410
            Width           =   3135
         End
         Begin VB.Image Image16 
            Height          =   1380
            Left            =   4260
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":10B28
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":10C7A
            Top             =   14010
            Width           =   930
         End
         Begin VB.Image Image15 
            Height          =   2415
            Left            =   5910
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1170F
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":11861
            Top             =   11610
            Width           =   2250
         End
         Begin VB.Image Image13 
            Height          =   2415
            Left            =   3660
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":13C3E
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":13D90
            Top             =   9510
            Width           =   2250
         End
         Begin VB.Image Image12 
            Height          =   1305
            Left            =   6690
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":15A1A
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":15B6C
            Top             =   5790
            Width           =   1020
         End
         Begin VB.Image Image11 
            Height          =   1305
            Left            =   3750
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":16679
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":167CB
            Top             =   5820
            Width           =   1020
         End
         Begin VB.Image Image10 
            Height          =   1305
            Left            =   480
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":172D8
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1742A
            Top             =   5820
            Width           =   1020
         End
         Begin VB.Image Image9 
            Height          =   1230
            Left            =   6360
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":17F37
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":18089
            Top             =   2010
            Width           =   1290
         End
         Begin VB.Image Image8 
            Height          =   1305
            Left            =   5490
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":18B70
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":18CC2
            Top             =   3600
            Width           =   1245
         End
         Begin VB.Image Image7 
            Height          =   1305
            Left            =   2670
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":19888
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":199DA
            Top             =   3600
            Width           =   1245
         End
         Begin VB.Image Image6 
            Height          =   1230
            Left            =   3450
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1A5A0
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1A6F2
            Top             =   2010
            Width           =   1290
         End
         Begin VB.Image Image5 
            Height          =   2415
            Left            =   5910
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1B1D9
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1B32B
            Top             =   7380
            Width           =   2250
         End
         Begin VB.Image Image3 
            Height          =   1230
            Left            =   60
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1CFE3
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1D135
            Top             =   2010
            Width           =   1290
         End
         Begin VB.Image Image2 
            Height          =   1230
            Left            =   5700
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1DC1C
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1DD6E
            Top             =   60
            Width           =   990
         End
         Begin VB.Image Image1 
            Height          =   1230
            Left            =   2520
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1E742
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":1E894
            Top             =   60
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1F445
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   15420
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1F74F
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   5850
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1FA59
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   5880
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":1FD63
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   5880
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
            Left            =   1680
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":2006D
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   4950
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":20377
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   2010
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":20681
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   2010
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
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":2098B
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   2010
            Width           =   1515
         End
         Begin VB.Image Image4 
            Height          =   2415
            Left            =   1380
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":20C95
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":20DE7
            Top             =   7380
            Width           =   2250
         End
         Begin VB.Image Image14 
            Height          =   2415
            Left            =   1380
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":22A31
            MousePointer    =   99  'Custom
            Picture         =   "HYUNDAIFinancialStatements.frx":22B83
            Top             =   11610
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
            Left            =   4680
            MouseIcon       =   "HYUNDAIFinancialStatements.frx":24F60
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   4950
            Width           =   3285
         End
      End
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   585
      Left            =   2130
      TabIndex        =   1
      Top             =   2070
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
      Format          =   22740993
      CurrentDate     =   38216
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   585
      Left            =   6150
      TabIndex        =   2
      Top             =   2070
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
      Format          =   22740993
      CurrentDate     =   38216
   End
   Begin VB.Image Image23 
      Height          =   2160
      Left            =   -120
      Picture         =   "HYUNDAIFinancialStatements.frx":2526A
      Top             =   -150
      Width           =   2910
   End
   Begin MSForms.ScrollBar ScrollBar1 
      Height          =   6525
      Left            =   9600
      TabIndex        =   0
      Top             =   60
      Width           =   435
      Size            =   "767;11509"
      Max             =   10500
      SmallChange     =   500
      LargeChange     =   500
      Delay           =   0
   End
   Begin VB.Image Image24 
      Height          =   1950
      Left            =   2250
      Picture         =   "HYUNDAIFinancialStatements.frx":288A0
      Top             =   60
      Width           =   7560
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
      Top             =   2100
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
      Top             =   2100
      Width           =   945
   End
End
Attribute VB_Name = "frmAMISHYUNDAIFinancialStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJOURNAL_HD As ADODB.Recordset
Dim rsJOURNAL_DET As ADODB.Recordset
Dim V_NetIncomeOrLoss As Double
Dim V_ProvisionForBonus As Double
Dim V_ProvisionForTax As Double
Dim V_Total_Current_Asset As Double
Dim V_Net_Propert_Equipment As Double
Dim V_Other_Assets As Double
Dim V_Propert_Equipment As Double
Dim V_AccumDepreciation As Double
Dim V_TaxCredit As Double
Dim DEALER_TYPE As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
Me.Height = 6960
CenterMe frmMain, Me, 1
dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
dtpTo = LOGDATE
DEALER_TYPE = "'2'"
Screen.MousePointer = 0
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

Function CheckDate() As Boolean
If Year(dtpFrom) <> Year(dtpTo) Then
   MsgBox "Invalid Date Range!", vbExclamation + vbCritical, "Error"
   CheckDate = False
Else
   CheckDate = True
End If
End Function

Private Sub Image1_Click()
labBalanceSheets_Click
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
labIncomeStatements_Click
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

Private Sub Image3_Click()
labIncomeStatementVehiclesCumulative_Click
End Sub

Private Sub Image4_Click()
labIncomeStatementByProductCumulative_Click
End Sub

Private Sub Image5_Click()
labIncomeStatementByProductCurrent_Click
End Sub

Private Sub Image6_Click()
labIncomeStatementPartsCumulative_Click
End Sub

Private Sub Image7_Click()
labScheduleOfSellingExpenseCumulative_Click
End Sub

Private Sub Image8_Click()
labScheduleOfSellingExpenseCurrent_Click
End Sub

Private Sub Image9_Click()
labIncomeStatementServiceCumulative_Click
End Sub

Private Sub labBalanceSheets_Click()
If CheckDate = False Then Exit Sub
Dim DateString As String
Dim V_GrossSales, V_SalesDiscountsAndReturns, V_CostOfSales As Double
Dim V_LessSellingExpense, V_LessAdminExpense, V_LessOtherExpense, V_AddOtherIncome As Double

Dim Cummulative_Cash_GrossSales, Cummulative_Charge_GrossSales, Cummulative_Cash_SalesDiscountsAndReturns, Cummulative_Charge_SalesDiscountsAndReturns As Double
Dim Cummulative_Cash_CostOfSales, Cummulative_Charge_CostOfSales, Cummulative_LessSellingExpense, Cummulative_LessAdminExpense As Double
Dim Cummulative_LessOtherExpense, Cummulative_AddOtherIncome As Double
If IsDate(dtpTo) = False Then
   MsgSpeechBox "Error In Date"
   Exit Sub
End If
Set rsJOURNAL_HD = New ADODB.Recordset
    rsJOURNAL_HD.Open "select * from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND (jdate <= '" & CDate(dtpTo) & "')", gconAmis, adOpenForwardOnly, adLockReadOnly
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   '================ CUMMULATIVE ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Cash_GrossSales = N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Charge_GrossSales = N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   V_GrossSales = Cummulative_Cash_GrossSales + Cummulative_Charge_GrossSales
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Cash_SalesDiscountsAndReturns = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Charge_SalesDiscountsAndReturns = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   V_SalesDiscountsAndReturns = Cummulative_Cash_SalesDiscountsAndReturns + Cummulative_Charge_SalesDiscountsAndReturns
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND (Headers='61' OR Headers='63')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Cash_CostOfSales = N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_Charge_CostOfSales = N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   V_CostOfSales = Cummulative_Cash_CostOfSales + Cummulative_Charge_CostOfSales
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "'" & _
                       " AND (Headers=" & OPERATIONAL_EXPENSE & " AND DepartmentCode <> " & ADMIN_EXPENSE & ")")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_LessSellingExpense = N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   V_LessSellingExpense = Cummulative_LessSellingExpense
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & OPERATIONAL_EXPENSE & " AND DepartmentCode = " & ADMIN_EXPENSE)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_LessAdminExpense = N2Str2Zero(rsJOURNAL_DET!LessAdminExpense)
   End If
   V_LessAdminExpense = Cummulative_LessAdminExpense
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & OTHER_EXPENSE)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_LessOtherExpense = N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   V_LessOtherExpense = Cummulative_LessOtherExpense
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jdate <= '" & dtpTo & "' AND Headers=" & OTHER_INCOME)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      Cummulative_AddOtherIncome = N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   V_AddOtherIncome = Cummulative_AddOtherIncome
   V_NetIncomeOrLoss = ((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)
   If V_NetIncomeOrLoss > 0 Then
      V_NetIncomeOrLoss = Round((((((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense))) - V_LessOtherExpense) + V_AddOtherIncome, 2)
      If V_NetIncomeOrLoss > 0 Then V_ProvisionForBonus = (V_NetIncomeOrLoss * 0.2)
      If V_NetIncomeOrLoss > 0 Then V_ProvisionForTax = ((V_NetIncomeOrLoss - V_ProvisionForBonus) * 0.32)
      If V_NetIncomeOrLoss > 0 Then V_NetIncomeOrLoss = V_NetIncomeOrLoss - (V_ProvisionForBonus + V_ProvisionForTax)
   Else
      V_NetIncomeOrLoss = Round(((((V_GrossSales - V_SalesDiscountsAndReturns) - V_CostOfSales) - (V_LessSellingExpense + V_LessAdminExpense)) - V_LessOtherExpense) + Abs(V_AddOtherIncome), 2)
      If V_NetIncomeOrLoss > 0 Then V_ProvisionForBonus = (V_NetIncomeOrLoss * 0.2)
      If V_NetIncomeOrLoss > 0 Then V_ProvisionForTax = ((V_NetIncomeOrLoss - V_ProvisionForBonus) * 0.32)
      If V_NetIncomeOrLoss > 0 Then V_NetIncomeOrLoss = V_NetIncomeOrLoss - (V_ProvisionForBonus + V_ProvisionForTax)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as Total_Current_Asset from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND (jdate <= '" & dtpTo & "') AND Headers=" & CURRENT_ASSET)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      V_Total_Current_Asset = N2Str2Zero(rsJOURNAL_DET!Total_Current_Asset)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as TaxCredit from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND (jdate <= '" & dtpTo & "') AND Titles=" & TAX_CREDITS)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      V_TaxCredit = N2Str2Zero(rsJOURNAL_DET!TaxCredit)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as Property_Equipment from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND (jdate <= '" & dtpTo & "') AND Titles=" & PROPERTY_EQUIPMENT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      V_Propert_Equipment = N2Str2Zero(rsJOURNAL_DET!PROPERTY_EQUIPMENT)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as AccumDepreciation from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND (jdate <= '" & dtpTo & "') AND Titles=" & ACCUMULATED_DEPRECIATION)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      V_AccumDepreciation = N2Str2Zero(rsJOURNAL_DET!AccumDepreciation)
   End If
   V_Net_Propert_Equipment = V_Propert_Equipment + V_AccumDepreciation
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as Other_Assets from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND (jdate <= '" & dtpTo & "') AND Titles=" & OTHER_ASSET)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      V_Other_Assets = N2Str2Zero(rsJOURNAL_DET!Other_Assets)
   End If
   ShowBalanceSheetReport "BalanceSheet", "FinancialStatement\", "{Journal_Det.SubtitleCode} = " & DEALER_TYPE & " AND ({Journal_Det.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & "))", "BALANCE SHEETS", "AS OF: " & Format(dtpTo, "long date"), True
   Screen.MousePointer = 0
Else
   ShowNoRecord
End If
End Sub

Public Sub ShowISReport(ReportName As Variant, ReportFolder As Variant, Filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
If CheckDate = False Then Exit Sub
Screen.MousePointer = 11
Dim rsProfile As ADODB.Recordset
Dim CrystalRpt As Crystal.CrystalReport
frmMain.rptMain.Reset
Set CrystalRpt = frmMain.rptMain
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
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
   CrystalRpt.Formulas(11) = "CurrentMonthYear = '" & Format(dtpTo, "MM/DD/YYYY") & "'"
   CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
   PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", Filter, AMIS_REPORT_Connection, 1
   CrystalRpt.PageZoom 89
End If
Screen.MousePointer = 0
End Sub

Private Sub labIncomeStatementByProductCumulative_Click()
If CheckDate = False Then Exit Sub
Dim PRODUCT As String
If dtpFrom > dtpTo Then
   MsgSpeechBox "Error In From and To date"
   Exit Sub
End If
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where jdate <= '" & dtpTo & "' and year(jdate) = " & Year(dtpTo))
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   rptAMISIncomeStatement.Reset
   Dim rsProfile As ADODB.Recordset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - BY PRODUCT - CUMMULATIVE"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - BY PRODUCT - CUMMULATIVE"
   End If
   '================ VEHICLES ================
   PRODUCT = "'10'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from Status = 'P' AND year(jdate) = " & Year(dtpTo) & " and jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.5
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ PARTS ================
   PRODUCT = "'30'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " and jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.2
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ SERVICE ================
   PRODUCT = "'20'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (((Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT & ") OR (AcctCode = '72-02200-20'))")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND AcctCode <> '72-02200-20' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.3
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND year(jdate) = " & Year(dtpTo) & " AND jtype <> 'CLO' and jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementByProductCumulative.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " AND {Journal_Det.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labIncomeStatementByProductCurrent_Click()
If CheckDate = False Then Exit Sub
Dim PRODUCT As String
If dtpFrom > dtpTo Then
   MsgSpeechBox "Error In From and To date"
   Exit Sub
End If
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   Dim rsProfile As ADODB.Recordset
   rptAMISIncomeStatement.Reset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - BY PRODUCT - CURRENT"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - BY PRODUCT - CURRENT"
   End If
   '================ VEHICLES ================
   PRODUCT = "'10'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
                                                                                              
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.5
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=======================================
   '================ PARTS ================
   PRODUCT = "'30'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.2
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ SERVICE ================
   PRODUCT = "'20'"
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "') AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (((Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT & ") OR (AcctCode = '72-02200-20'))")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND AcctCode <> '72-02200-20' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.3
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementByProductCurrent.rpt", "{Journal_Det.jtype} <> 'CLO' AND {Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND Journal_Det.jtype <> 'CLO' AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labIncomeStatementPartsCumulative_Click()
If CheckDate = False Then Exit Sub
Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT As String
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
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   Dim rsProfile As ADODB.Recordset
   rptAMISIncomeStatement.Reset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - PARTS"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - PARTS"
   End If
   PRODUCT = "'30'"
   '================ CUMMULATIVE ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   'used for service section
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20') AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='71' AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.2
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ CURRENT ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "') AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='71' AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.2
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ PREVIOUS ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.2
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementPartsCumulative.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labIncomeStatements_Click()
If CheckDate = False Then Exit Sub
Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT As String
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
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   Dim rsProfile As ADODB.Recordset
   rptAMISIncomeStatement.Reset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENTS"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENTS"
   End If
   '================ CUMMULATIVE ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='41'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='42'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='51'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='52'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='61' OR Headers='63')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='62'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND ((Headers='71' OR Headers='72') AND DepartmentCode <> '40')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND ChartAccount.Headers='91'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='81'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ CURRENT ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='41'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='42'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='51'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='52'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='61' OR Headers='63')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='62'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND ((Headers='71' OR Headers='72') AND DepartmentCode <> '40')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='81'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ PREVIOUS ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='41'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='42'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='51'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='52'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='61' OR Headers='63')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='62'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "')" & _
                       " AND ((Headers='71' OR Headers='72') AND DepartmentCode <> '40')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='91' AND (AcctCode <> '91-03000-00' AND AcctCode <> '91-04000-00')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='81'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\IncomeStatements.rpt", "{Journal_HD.jtype} = 'CLO' AND {Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ") and year({Journal_HD.jdate}) = " & Year(dtpTo), AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labIncomeStatementServiceCumulative_Click()
If CheckDate = False Then Exit Sub
Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT As String
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
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where year(jdate) = " & Year(dtpTo) & " and (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   Dim rsProfile As ADODB.Recordset
   rptAMISIncomeStatement.Reset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - SERVICE"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - SERVICE"
   End If
   PRODUCT = "'20'"
   '================ CUMMULATIVE ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jdate <= '" & dtpTo & "' AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
   'used for service section
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND ((Headers='61' OR Headers='63') OR AcctCode='72-02200-20') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.3
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ CURRENT ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "') AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND ((Headers='61' OR Headers='63') OR AcctCode='72-02200-20') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='71' AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.3
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ PREVIOUS ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers='51' AND ChartAccount.DepartmentCode = " & PRODUCT)
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND ((Headers='61' OR Headers='63') OR AcctCode='72-02200-20') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='71' OR Headers='72') AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.3
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & Prev_dtpFrom & "' AND jdate <= '" & Prev_dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementserviceCumulative.rpt", "{Journal_Det.jtype} <> 'CLO' AND {Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labIncomeStatementVehiclesCumulative_Click()
If CheckDate = False Then Exit Sub
Dim Prev_dtpFrom, Prev_dtpTo, PRODUCT As String
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
Set rsJOURNAL_HD = New ADODB.Recordset
Set rsJOURNAL_HD = gconAmis.Execute("select * from Journal_HD where (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')")
If Not rsJOURNAL_HD.EOF And Not rsJOURNAL_HD.EOF Then
   Dim rsProfile As ADODB.Recordset
   rptAMISIncomeStatement.Reset
   Set rsProfile = New ADODB.Recordset
   Set rsProfile = gconAmis.Execute("Select * from Profile")
   If Not (rsProfile.EOF And rsProfile.BOF) Then
      rptAMISIncomeStatement.Formulas(30) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
      rptAMISIncomeStatement.Formulas(31) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
      rptAMISIncomeStatement.ReportTitle = "INCOME STATEMENT - VEHICLE AND ACCESSORIES"
      rptAMISIncomeStatement.WindowTitle = "INCOME STATEMENT - VEHICLE AND ACCESSORIES"
   End If
   PRODUCT = "'10'"
   '================ CUMMULATIVE ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(0) = "Cummulative_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(1) = "Cummulative_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(2) = "Cummulative_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(3) = "Cummulative_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   'used for service section
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20') AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(4) = "Cummulative_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(5) = "Cummulative_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "'" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(6) = "Cummulative_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='71' AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(7) = "Cummulative_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.5
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(8) = "Cummulative_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND jdate <= '" & dtpTo & "' AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(9) = "Cummulative_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ CURRENT ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "') AND Headers=" & CASH_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(10) = "Current_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_SALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(11) = "Current_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CASH_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(12) = "Current_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_DISCOUNT & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(13) = "Current_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='61' OR Headers='63') AND DepartmentCode = " & PRODUCT)
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "') AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(14) = "Current_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers=" & CHARGE_COSTOFSALES & " AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(15) = "Current_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "')" & _
                       " AND Headers='71' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(16) = "Current_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='71' AND DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(17) = "Current_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.5
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND Headers='91' AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(18) = "Current_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from vCOA_Journal_Det where SubtitleCode = " & DEALER_TYPE & " AND Status = 'P' AND jtype <> 'CLO' AND (jdate >= '" & dtpFrom & "' AND jdate <= '" & dtpTo & "') AND (Headers='81' OR Headers='82' OR Headers='83') AND DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(19) = "Current_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   '================ PREVIOUS ================
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Cash_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers=" & CASH_SALES & " AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(20) = "Previous_Cash_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Cash_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as Charge_GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers=" & CHARGE_SALES & " AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(21) = "Previous_Charge_GrossSales = " & N2Str2Zero(rsJOURNAL_DET!Charge_GrossSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers=" & CASH_DISCOUNT & " AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(22) = "Previous_Cash_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers=" & CHARGE_DISCOUNT & " AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(23) = "Previous_Charge_SalesDiscountsAndReturns = " & N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND (ChartAccount.Headers='61' OR ChartAccount.Headers='63') AND ChartAccount.DepartmentCode = " & PRODUCT)
   'Set rsJournal_Det = gconAMIS.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND (ChartAccount.Headers='51' OR ChartAccount.AcctCode='72-02200-20')")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(24) = "Previous_Cash_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as CostOfSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers=" & CHARGE_COSTOFSALES & " AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(25) = "Previous_Charge_CostOfSales = " & N2Str2Zero(rsJOURNAL_DET!CostOfSales)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessSellingExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode" & _
                       " where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "')" & _
                       " AND ChartAccount.Headers='71' AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(26) = "Previous_LessSellingExpense = " & N2Str2Zero(rsJOURNAL_DET!LessSellingExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessAdminExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers='71' AND ChartAccount.DepartmentCode = '40'")
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(27) = "Previous_LessAdminExpense = " & N2Str2Zero(rsJOURNAL_DET!LessAdminExpense) * 0.5
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as LessOtherExpense from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND ChartAccount.Headers='91' AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(28) = "Previous_LessOtherExpense = " & N2Str2Zero(rsJOURNAL_DET!LessOtherExpense)
   End If
   Set rsJOURNAL_DET = New ADODB.Recordset
   Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as AddOtherIncome from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND (Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "') AND (ChartAccount.Headers='81' OR ChartAccount.Headers='82' OR ChartAccount.Headers='83') AND ChartAccount.DepartmentCode = " & PRODUCT)
   If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
      rptAMISIncomeStatement.Formulas(29) = "Previous_AddOtherIncome = " & N2Str2Zero(rsJOURNAL_DET!AddOtherIncome)
   End If
   '=========================================
   PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & "FinancialStatement\FinancialStatements\IncomeStatementVehiclesCumulative.rpt", "{Journal_HD.jtype} <> 'CLO' AND {Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Else
   ShowNoRecord
End If
End Sub

Private Sub labScheduleOAdministrativeExpensesCumulative_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo As String
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
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers= " & CHARGE_SALES & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.Formulas(5) = "ToJDate = Date(" & Year(CDate(dtpTo)) & "," & Month(CDate(dtpTo)) & "," & Day(CDate(dtpTo)) & ")"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMINISTRATIVE EXPENSES"
End If
rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdministrativeExpensesCumulative.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfAccounts_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder As String
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(1) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(2) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ACCOUNTS"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ACCOUNTS"
   rptAMISIncomeStatement.Formulas(3) = "ReportDate = '" & "As of: " & Format(dtpTo, "long date") & "'"
End If
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAccounts.rpt", "{Journal_Det.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfAdminAndSellingExpensesCumulative_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder As String
Dim vSALES_CUMULATIVE_NETSALES, vSALES_CUMULATIVE_GROSSSALES, vSALES_CUMULATIVE_DISCOUNTSRETURNS As Double
Dim vPARTS_CUMULATIVE_NETSALES, vPARTS_CUMULATIVE_GROSSSALES, vPARTS_CUMULATIVE_DISCOUNTSRETURNS As Double
Dim vSERVICE_CUMULATIVE_NETSALES, vSERVICE_CUMULATIVE_GROSSSALES, vSERVICE_CUMULATIVE_DISCOUNTSRETURNS As Double
Dim vADMIN_CUMULATIVE_NETSALES, vADMIN_CUMULATIVE_GROSSSALES, vADMIN_CUMULATIVE_DISCOUNTSRETURNS As Double

Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vADMIN_CUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vADMIN_CUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vSALES_CUMULATIVE_NETSALES = vSALES_CUMULATIVE_GROSSSALES - vSALES_CUMULATIVE_DISCOUNTSRETURNS
vPARTS_CUMULATIVE_NETSALES = vPARTS_CUMULATIVE_GROSSSALES - vPARTS_CUMULATIVE_DISCOUNTSRETURNS
vSERVICE_CUMULATIVE_NETSALES = vSERVICE_CUMULATIVE_GROSSSALES - vSERVICE_CUMULATIVE_DISCOUNTSRETURNS
vADMIN_CUMULATIVE_NETSALES = vADMIN_CUMULATIVE_GROSSSALES - vADMIN_CUMULATIVE_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CUMULATIVE"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CUMULATIVE"
End If
rptAMISIncomeStatement.Formulas(0) = "ADMIN_CUMULATIVE_NETSALES = " & NumericVal(vADMIN_CUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "SALES_CUMULATIVE_NETSALES = " & NumericVal(vSALES_CUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "PARTS_CUMULATIVE_NETSALES = " & NumericVal(vPARTS_CUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(3) = "SERVICE_CUMULATIVE_NETSALES = " & NumericVal(vSERVICE_CUMULATIVE_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdminAndSellingExpensesCummulative.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " AND {Journal_Det.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfAdminAndSellingExpensesCurrent_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder As String
Dim vSALES_CURRENT_NETSALES, vSALES_CURRENT_GROSSSALES, vSALES_CURRENT_DISCOUNTSRETURNS As Double
Dim vPARTS_CURRENT_NETSALES, vPARTS_CURRENT_GROSSSALES, vPARTS_CURRENT_DISCOUNTSRETURNS As Double
Dim vSERVICE_CURRENT_NETSALES, vSERVICE_CURRENT_GROSSSALES, vSERVICE_CURRENT_DISCOUNTSRETURNS As Double
Dim vADMIN_CURRENT_NETSALES, vADMIN_CURRENT_GROSSSALES, vADMIN_CURRENT_DISCOUNTSRETURNS As Double

Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_CURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_CURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_CURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vADMIN_CURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ")")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vADMIN_CURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vSALES_CURRENT_NETSALES = vSALES_CURRENT_GROSSSALES - vSALES_CURRENT_DISCOUNTSRETURNS
vPARTS_CURRENT_NETSALES = vPARTS_CURRENT_GROSSSALES - vPARTS_CURRENT_DISCOUNTSRETURNS
vSERVICE_CURRENT_NETSALES = vSERVICE_CURRENT_GROSSSALES - vSERVICE_CURRENT_DISCOUNTSRETURNS
vADMIN_CURRENT_NETSALES = vADMIN_CURRENT_GROSSSALES - vADMIN_CURRENT_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CURRENT"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF ADMIN & SELLING EXPENSES - CURRENT"
End If
rptAMISIncomeStatement.Formulas(0) = "ADMIN_CURRENT_NETSALES = " & NumericVal(vADMIN_CURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "SALES_CURRENT_NETSALES = " & NumericVal(vSALES_CURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "PARTS_CURRENT_NETSALES = " & NumericVal(vPARTS_CURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(3) = "SERVICE_CURRENT_NETSALES = " & NumericVal(vSERVICE_CURRENT_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfAdminAndSellingExpensesCurrent.rpt", "{Journal_Det.jtype} <> 'CLO' and {Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfSellingExpenseCumulative_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder As String
Dim vSALES_NETSALES, vSALES_GROSSSALES, vSALES_DISCOUNTSRETURNS As Double
Dim vPARTS_NETSALES, vPARTS_GROSSSALES, vPARTS_DISCOUNTSRETURNS As Double
Dim vSERVICE_NETSALES, vSERVICE_GROSSSALES, vSERVICE_DISCOUNTSRETURNS As Double
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.Jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vSALES_NETSALES = vSALES_GROSSSALES - vSALES_DISCOUNTSRETURNS
vPARTS_NETSALES = vPARTS_GROSSSALES - vPARTS_DISCOUNTSRETURNS
vSERVICE_NETSALES = vSERVICE_GROSSSALES - vSERVICE_DISCOUNTSRETURNS
vTOTAL_NETSALES = vSALES_NETSALES + vPARTS_NETSALES + vSERVICE_NETSALES
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - CUMULATIVE"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - CUMULATIVE"
End If
rptAMISIncomeStatement.Formulas(0) = "SALES_NETSALES = " & NumericVal(vSALES_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "PARTS_NETSALES = " & NumericVal(vPARTS_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "SERVICE_NETSALES = " & NumericVal(vSERVICE_NETSALES)
rptAMISIncomeStatement.Formulas(3) = "TOTAL_NETSALES = " & NumericVal(vTOTAL_NETSALES)

ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensesCumulative.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " AND {Journal_HD.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub labScheduleOfSellingExpenseCurrent_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder As String
Dim vSALES_NETSALES, vSALES_GROSSSALES, vSALES_DISCOUNTSRETURNS As Double
Dim vPARTS_NETSALES, vPARTS_GROSSSALES, vPARTS_DISCOUNTSRETURNS As Double
Dim vSERVICE_NETSALES, vSERVICE_GROSSSALES, vSERVICE_DISCOUNTSRETURNS As Double
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSALES_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPARTS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.Jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vSERVICE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vSALES_NETSALES = vSALES_GROSSSALES - vSALES_DISCOUNTSRETURNS
vPARTS_NETSALES = vPARTS_GROSSSALES - vPARTS_DISCOUNTSRETURNS
vSERVICE_NETSALES = vSERVICE_GROSSSALES - vSERVICE_DISCOUNTSRETURNS
vTOTAL_NETSALES = vSALES_NETSALES + vPARTS_NETSALES + vSERVICE_NETSALES
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(4) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(5) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - CURRENT"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - CURRENT"
End If
rptAMISIncomeStatement.Formulas(0) = "SALES_NETSALES = " & NumericVal(vSALES_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "PARTS_NETSALES = " & NumericVal(vPARTS_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "SERVICE_NETSALES = " & NumericVal(vSERVICE_NETSALES)
rptAMISIncomeStatement.Formulas(3) = "TOTAL_NETSALES = " & NumericVal(vTOTAL_NETSALES)

ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensesCurrent.rpt", "{Journal_HD.jdate} >= date(" & Year(dtpFrom) & "," & Month(dtpFrom) & "," & Day(dtpFrom) & ") AND {Journal_HD.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub labScheduleOfSellingExpensesPartsSection_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo As String
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
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '30'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.Formulas(5) = "ToJDate = '" & CDate(dtpTo) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - PARTS SECTION"
End If
rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpensePartsSection.rpt", "{Journal_Det.jtype} <> 'CLO' {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfSellingExpensesSalesSection_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo As String
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
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '10'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(3) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(4) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.Formulas(5) = "ToJdate = '" & CDate(dtpTo) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SALES SECTION"
End If
rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
rptAMISIncomeStatement.Formulas(2) = "PREVIOUS_NETSALES = " & NumericVal(vPREVIOUS_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpenseSalesSection.rpt", "{Journal_Det.jtype} <> 'CLO' {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub labScheduleOfSellingExpensesServiceSection_Click()
If CheckDate = False Then Exit Sub
Dim ReportFolder, Prev_dtpFrom, Prev_dtpTo As String
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
Dim vTOTAL_NETSALES  As Double

Dim rsJOURNAL_DET As ADODB.Recordset
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND year(Journal_Det.jdate) = " & Year(dtpTo) & " AND Journal_Det.jtype <> 'CLO' and Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCUMULATIVE_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCUMULATIVE_NETSALES = vCUMULATIVE_GROSSSALES - vCUMULATIVE_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & dtpFrom & "' AND Journal_Det.jdate <= '" & dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vCURRENT_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vCURRENT_NETSALES = vCURRENT_GROSSSALES - vCURRENT_DISCOUNTSRETURNS
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Credit) - SUM(Journal_Det.Debit) as GrossSales from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_SALES & " OR ChartAccount.Headers=" & CHARGE_SALES & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_GROSSSALES = N2Str2Zero(rsJOURNAL_DET!GrossSales)
End If
Set rsJOURNAL_DET = New ADODB.Recordset
Set rsJOURNAL_DET = gconAmis.Execute("select SUM(Journal_Det.Debit) - SUM(Journal_Det.Credit) as SalesDiscountsAndReturns from Journal_Det inner Join ChartAccount on Journal_Det.Acct_Code = ChartAccount.AcctCode where Journal_Det.Status = 'P' AND Journal_Det.jtype <> 'CLO' AND Journal_Det.jdate >= '" & Prev_dtpFrom & "' AND Journal_Det.jdate <= '" & Prev_dtpTo & "' AND (ChartAccount.Headers=" & CASH_DISCOUNT & " OR ChartAccount.Headers=" & CHARGE_DISCOUNT & ") AND ChartAccount.DepartmentCode = '20'")
If Not rsJOURNAL_DET.EOF And Not rsJOURNAL_DET.EOF Then
   vPREVIOUS_DISCOUNTSRETURNS = N2Str2Zero(rsJOURNAL_DET!SalesDiscountsAndReturns)
End If
vPREVIOUS_NETSALES = vPREVIOUS_GROSSSALES - vPREVIOUS_DISCOUNTSRETURNS
Dim rsProfile As ADODB.Recordset
rptAMISIncomeStatement.Reset
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   rptAMISIncomeStatement.Formulas(2) = "CompanyName = '" & Null2String(rsProfile!CompanyName) & "'"
   rptAMISIncomeStatement.Formulas(3) = "CompanyAddress = '" & Null2String(rsProfile!Companyaddress) & "'"
   rptAMISIncomeStatement.Formulas(4) = "ToJDate = '" & CDate(dtpTo) & "'"
   rptAMISIncomeStatement.ReportTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
   rptAMISIncomeStatement.WindowTitle = "SCHEDULE OF SELLING EXPENSES - SERVICE SECTION"
End If
rptAMISIncomeStatement.Formulas(0) = "CUMULATIVE_NETSALES = " & NumericVal(vCUMULATIVE_NETSALES)
rptAMISIncomeStatement.Formulas(1) = "CURRENT_NETSALES = " & NumericVal(vCURRENT_NETSALES)
ReportFolder = "FinancialStatement\FinancialStatements\"
PrintSQLReport rptAMISIncomeStatement, AMIS_REPORT_PATH & ReportFolder & "ScheduleOfSellingExpenseServiceSection.rpt", "year({Journal_HD.jdate}) = " & Year(dtpTo) & " AND {Journal_Det.jtype} <> 'CLO' and {Journal_HD.jdate} <= date(" & Year(dtpTo) & "," & Month(dtpTo) & "," & Day(dtpTo) & ")", AMIS_REPORT_Connection, 1
End Sub

Private Sub ScrollBar1_Change()
Picture2.Top = 0 - ScrollBar1.Value
End Sub

Sub ShowBalanceSheetReport(ReportName As Variant, ReportFolder As Variant, Filter As Variant, ReportHeading As String, REPORT_DATE As String, WithDate As Boolean)
Screen.MousePointer = 11
Dim rsProfile As ADODB.Recordset
Dim CrystalRpt As Crystal.CrystalReport
Set CrystalRpt = frmMain.rptMain
Set rsProfile = New ADODB.Recordset
Set rsProfile = gconAmis.Execute("Select * from Profile")
If Not (rsProfile.EOF And rsProfile.BOF) Then
   CrystalRpt.Reset
   CrystalRpt.ReportFileName = AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt"
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
   CrystalRpt.Formulas(11) = "CurrentMonthYear = '" & Format(dtpTo, "MM/DD/YYYY") & "'"
   CrystalRpt.ReportTitle = ReportHeading: CrystalRpt.WindowTitle = ReportHeading
   PrintSQLReport CrystalRpt, AMIS_REPORT_PATH & ReportFolder & "\" & ReportName & ".rpt", Filter, AMIS_REPORT_Connection, 1
   CrystalRpt.PageZoom 89
End If
Screen.MousePointer = 0
End Sub

