VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PMIS Main Menu"
   ClientHeight    =   6690
   ClientLeft      =   5700
   ClientTop       =   660
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "PMISMainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10965
   Begin XtremeSuiteControls.TabControl SS_MAIN 
      Height          =   6675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      _Version        =   655364
      _ExtentX        =   19288
      _ExtentY        =   11774
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Tables && Files"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Inquiry"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage4"
      Item(4).Caption =   "Other Setups"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "TabControlPage5"
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   6060
         Left            =   -69970
         TabIndex        =   1
         Top             =   585
         Visible         =   0   'False
         Width           =   10875
         _Version        =   655364
         _ExtentX        =   19182
         _ExtentY        =   10689
         _StockProps     =   0
         Begin VB.CommandButton cmdOther_Reminders 
            Height          =   645
            Left            =   390
            MouseIcon       =   "PMISMainMenu.frx":01CA
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":031C
            Style           =   1  'Graphical
            TabIndex        =   2
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton cmdOther_ComapnyProfile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   390
            MouseIcon       =   "PMISMainMenu.frx":0B97
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":0CE9
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "1405"
            ToolTipText     =   "Company Profile"
            Top             =   1260
            Width           =   720
         End
         Begin VB.CommandButton cmdOther_Password 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   390
            MouseIcon       =   "PMISMainMenu.frx":16E0
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":1832
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   2160
            Width           =   720
         End
         Begin VB.CommandButton cmdOther_PRICELISTCONVERSIONTOOL 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   390
            MouseIcon       =   "PMISMainMenu.frx":2156
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":22A8
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   3060
            Width           =   720
         End
         Begin VB.CommandButton cmdOther_MacTool 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   390
            MouseIcon       =   "PMISMainMenu.frx":251B
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":266D
            Style           =   1  'Graphical
            TabIndex        =   10
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   3960
            Width           =   720
         End
         Begin VB.Label Label98 
            BackStyle       =   0  'Transparent
            Caption         =   "REMINDERS"
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
            Height          =   375
            Left            =   1260
            TabIndex        =   3
            Top             =   570
            Width           =   2490
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD MAINTENANCE"
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
            Left            =   1260
            TabIndex        =   7
            Top             =   2355
            Width           =   6015
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY PROFILE"
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
            Height          =   345
            Left            =   1245
            TabIndex        =   5
            Top             =   1440
            Width           =   3195
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE LIST CONVERSION TOOL"
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
            Left            =   1260
            TabIndex        =   9
            Top             =   3255
            Width           =   6015
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "MAC TOOL"
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
            Left            =   1260
            TabIndex        =   11
            Top             =   4155
            Width           =   6015
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   6060
         Left            =   -69970
         TabIndex        =   207
         Top             =   585
         Visible         =   0   'False
         Width           =   10875
         _Version        =   655364
         _ExtentX        =   19182
         _ExtentY        =   10689
         _StockProps     =   0
         Begin VB.CommandButton cmdReport_Forcasting 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5130
            MouseIcon       =   "PMISMainMenu.frx":28E0
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":2A32
            Style           =   1  'Graphical
            TabIndex        =   226
            Tag             =   "1388"
            ToolTipText     =   "Forecasting Reports"
            Top             =   3615
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_PartsRundown 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5130
            MouseIcon       =   "PMISMainMenu.frx":3135
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":3287
            Style           =   1  'Graphical
            TabIndex        =   230
            Tag             =   "1700"
            ToolTipText     =   "Parts Rundown Reports"
            Top             =   4470
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_StockStatusReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5130
            MouseIcon       =   "PMISMainMenu.frx":3A5D
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":3BAF
            Style           =   1  'Graphical
            TabIndex        =   218
            Tag             =   "1385"
            ToolTipText     =   "Stock Status Report"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_RankingReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5145
            MouseIcon       =   "PMISMainMenu.frx":4314
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":4466
            Style           =   1  'Graphical
            TabIndex        =   222
            Tag             =   "1387"
            ToolTipText     =   "Ranking Reports"
            Top             =   2745
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_IssuanceOfTheMonth 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":4B5B
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":4CAD
            Style           =   1  'Graphical
            TabIndex        =   228
            Tag             =   "1384"
            ToolTipText     =   "Issuances for the Month"
            Top             =   4470
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_TransListingIssuance 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":5389
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":54DB
            Style           =   1  'Graphical
            TabIndex        =   220
            Tag             =   "1382"
            ToolTipText     =   "Transaction Listing Issuance Report"
            Top             =   2745
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_RecieptForTheMonth 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5145
            MouseIcon       =   "PMISMainMenu.frx":5BDB
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":5D2D
            Style           =   1  'Graphical
            TabIndex        =   214
            Tag             =   "1383"
            ToolTipText     =   "Receipts for the Month"
            Top             =   1035
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_TransListingReceipt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   360
            MouseIcon       =   "PMISMainMenu.frx":6435
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":6587
            Style           =   1  'Graphical
            TabIndex        =   216
            Tag             =   "1381"
            ToolTipText     =   "Transaction Listing Receipts Report"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_PISReportWorkinProgress 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":6BCD
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":6D1F
            Style           =   1  'Graphical
            TabIndex        =   212
            Tag             =   "1380"
            ToolTipText     =   "PIS Report for Work-In-Progress"
            Top             =   1035
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_DailySalesReport 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":73F4
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":7546
            Style           =   1  'Graphical
            TabIndex        =   208
            Tag             =   "1379"
            ToolTipText     =   "Daily Sales Report"
            Top             =   180
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_TranHist_PO 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":7CBD
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":7E0F
            Style           =   1  'Graphical
            TabIndex        =   224
            Tag             =   "1382"
            ToolTipText     =   "Transaction Listing Purchase Report"
            Top             =   3615
            Width           =   720
         End
         Begin VB.CommandButton cmdReport_PurchaseForTheMonth 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5145
            MouseIcon       =   "PMISMainMenu.frx":7FB5
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":8107
            Style           =   1  'Graphical
            TabIndex        =   210
            Tag             =   "1383"
            ToolTipText     =   "Purchase for the Month"
            Top             =   180
            Width           =   720
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS RUNDOWN REPORTS"
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
            Height          =   285
            Left            =   5985
            TabIndex        =   231
            Top             =   4650
            Width           =   2565
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "FORECASTING REPORTS"
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
            Height          =   195
            Left            =   5985
            TabIndex        =   227
            Top             =   3840
            Width           =   2835
         End
         Begin VB.Label Label52 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "RANKING REPORTS"
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
            Height          =   330
            Left            =   5985
            TabIndex        =   223
            Top             =   2902
            Width           =   2625
         End
         Begin VB.Label Label58 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "DAILY SALES REPORT"
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
            Height          =   285
            Left            =   1215
            TabIndex        =   209
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label Label59 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "PIS REPORT FOR WORK-IN PROGRESS"
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
            Height          =   330
            Left            =   1215
            TabIndex        =   213
            Top             =   1192
            Width           =   3480
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "RECEIPTS FOR THE MONTH"
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
            Height          =   345
            Left            =   5985
            TabIndex        =   215
            Top             =   1185
            Width           =   2565
         End
         Begin VB.Label Label61 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION LISTING RECEIPTS REPORT "
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
            Height          =   345
            Left            =   1200
            TabIndex        =   217
            Top             =   2040
            Width           =   3600
         End
         Begin VB.Label Label62 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION LISTING ISSUANCE REPORT "
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
            Height          =   270
            Left            =   1200
            TabIndex        =   221
            Top             =   2932
            Width           =   3645
         End
         Begin VB.Label Label63 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ISSUANCES FOR THE MONTH"
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
            Left            =   1215
            TabIndex        =   229
            Top             =   4635
            Width           =   2550
         End
         Begin VB.Label Label65 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "STOCK STATUS REPORT"
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
            Height          =   285
            Left            =   6000
            TabIndex        =   219
            Top             =   2070
            Width           =   2115
         End
         Begin VB.Label Label101 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION LISTING PURCHASE REPORT"
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
            Height          =   345
            Left            =   1215
            TabIndex        =   225
            Top             =   3765
            Width           =   3825
         End
         Begin VB.Label Label102 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "PURCHASE FOR THE MONTH"
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
            Height          =   345
            Left            =   5985
            TabIndex        =   211
            Top             =   330
            Width           =   2565
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   6060
         Left            =   -69970
         TabIndex        =   12
         Top             =   585
         Visible         =   0   'False
         Width           =   10875
         _Version        =   655364
         _ExtentX        =   19182
         _ExtentY        =   10689
         _StockProps     =   0
         Begin VB.CommandButton cmdInquiry_PartsAvalibity 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   360
            MouseIcon       =   "PMISMainMenu.frx":8937
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":8A89
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "1364"
            ToolTipText     =   "Parts Availability Inquiry"
            Top             =   180
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_Counter 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":9141
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":9293
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "1365"
            ToolTipText     =   "Counter Inquiry"
            Top             =   1035
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_CheckInventoryBalance 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":9864
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":99B6
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "1373"
            ToolTipText     =   "Check Inventory Balances"
            Top             =   1080
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_InventoryRankingInquiry 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":9EBF
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":A011
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "1374"
            ToolTipText     =   "Inventory Ranking Inquiry"
            Top             =   1935
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_DealerDNPComparision 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":A68B
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":A7DD
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "1376"
            ToolTipText     =   "Dealer/Distributor DNP Comparison"
            Top             =   3600
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_DealerSRPComparision 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":AE2B
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":AF7D
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "1377"
            ToolTipText     =   "Dealer/Distributor SRP Comparison"
            Top             =   4455
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_DealerSRPDNP 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":B617
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":B769
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "1375"
            ToolTipText     =   "Dealer SRP/DNP Listing"
            Top             =   2790
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_TransactionDetails 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":BEF1
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":C043
            Style           =   1  'Graphical
            TabIndex        =   39
            Tag             =   "1372"
            ToolTipText     =   "Transaction Details"
            Top             =   5325
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_BrowseErrorFiles 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":C6B1
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":C803
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "1378"
            ToolTipText     =   "Browse Error Files"
            Top             =   5310
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_IssuanceTransaction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":CDBB
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":CF0D
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "1371"
            ToolTipText     =   "Issuances Transactions"
            Top             =   4470
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_POTransaction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":D583
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":D6D5
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "1369"
            ToolTipText     =   "PO Transactions"
            Top             =   2745
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_MRRTransaction 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":DE32
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":DF84
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "1370"
            ToolTipText     =   "MRR Transactions"
            Top             =   3615
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_ComputeriedParts 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":E6CB
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":E81D
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "1366"
            ToolTipText     =   "Parts Computerized Stock Cards"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdInquiry_DelaerPartInquiry 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5340
            MouseIcon       =   "PMISMainMenu.frx":EFCB
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":F11D
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "1469"
            ToolTipText     =   "Dealer Part Inquiry"
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "BROWSE ERROR FILES"
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
            Height          =   330
            Left            =   6210
            TabIndex        =   40
            Top             =   5460
            Width           =   3225
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "PO TRANSACTIONS"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   27
            Top             =   2955
            Width           =   3225
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "MRR TRANSACTIONS"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   31
            Top             =   3810
            Width           =   3225
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "ISSUANCES TRANSACTIONS"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   33
            Top             =   4665
            Width           =   3765
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSACTION DETAILS"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   37
            Top             =   5505
            Width           =   3225
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "CHECK INVENTORY BALANCES"
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
            Height          =   330
            Left            =   6195
            TabIndex        =   20
            Top             =   1245
            Width           =   4260
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "INVENTORY RANKING INQUIRY"
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
            Height          =   330
            Left            =   6195
            TabIndex        =   24
            Top             =   2070
            Width           =   4080
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER SRP / DNP LISTING"
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
            Height          =   330
            Left            =   6195
            TabIndex        =   26
            Top             =   2940
            Width           =   4260
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER / DISTRIBUTOR DNP COMPARISON"
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
            Height          =   735
            Left            =   6195
            TabIndex        =   32
            Top             =   3630
            Width           =   3480
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER / DISTRIBUTOR SRP COMPARISON"
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
            Height          =   555
            Left            =   6195
            TabIndex        =   36
            Top             =   4485
            Width           =   3570
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS COMPUTERIZED STOCK CARDS"
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
            Height          =   555
            Left            =   1155
            TabIndex        =   22
            Top             =   2010
            Width           =   2865
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "COUNTER INQUIRY"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   18
            Top             =   1245
            Width           =   3135
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS AVAILABILITY INQUIRY"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   14
            Top             =   450
            Width           =   3990
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "DEALER PART INQUIRY"
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
            Left            =   6210
            TabIndex        =   16
            Top             =   450
            Width           =   3975
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   6060
         Left            =   30
         TabIndex        =   41
         Top             =   585
         Width           =   10875
         _Version        =   655364
         _ExtentX        =   19182
         _ExtentY        =   10689
         _StockProps     =   0
         Begin XtremeSuiteControls.TabControl SS_PARTS 
            Height          =   6075
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   10875
            _Version        =   655364
            _ExtentX        =   19182
            _ExtentY        =   10716
            _StockProps     =   64
            Appearance      =   2
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            PaintManager.FixedTabWidth=   300
            PaintManager.MinTabWidth=   120
            ItemCount       =   3
            Item(0).Caption =   "Parts"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "TabControlPage6"
            Item(1).Caption =   "Accessories"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "TabControlPage7"
            Item(2).Caption =   "Materials"
            Item(2).ControlCount=   1
            Item(2).Control(0)=   "TabControlPage8"
            Begin XtremeSuiteControls.TabControlPage TabControlPage8 
               Height          =   5460
               Left            =   -69970
               TabIndex        =   124
               Top             =   585
               Visible         =   0   'False
               Width           =   10815
               _Version        =   655364
               _ExtentX        =   19076
               _ExtentY        =   9631
               _StockProps     =   0
               Begin VB.CommandButton Command2 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   7350
                  MouseIcon       =   "PMISMainMenu.frx":F78B
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":F8DD
                  Style           =   1  'Graphical
                  TabIndex        =   234
                  Tag             =   "1307"
                  ToolTipText     =   "Advance Bill Data Entry"
                  Top             =   210
                  Width           =   720
               End
               Begin VB.CommandButton cmdAdvanceBill_Materials 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":10146
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":10298
                  Style           =   1  'Graphical
                  TabIndex        =   133
                  Tag             =   "1307"
                  ToolTipText     =   "Advance Bill Data Entry"
                  Top             =   2466
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_DROut 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":10B01
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":10C53
                  Style           =   1  'Graphical
                  TabIndex        =   134
                  Tag             =   "1306"
                  ToolTipText     =   "DR Out Issuance"
                  Top             =   3188
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_PO 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":114CF
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":11621
                  Style           =   1  'Graphical
                  TabIndex        =   136
                  Tag             =   "1318"
                  ToolTipText     =   "Materials Purchase Order"
                  Top             =   3910
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_RR 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":11E91
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":11FE3
                  Style           =   1  'Graphical
                  TabIndex        =   138
                  Tag             =   "1319"
                  ToolTipText     =   "Materials Receiving and Storing"
                  Top             =   4635
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_Adjustment 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   4425
                  MouseIcon       =   "PMISMainMenu.frx":12773
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":128C5
                  Style           =   1  'Graphical
                  TabIndex        =   125
                  Tag             =   "1308"
                  ToolTipText     =   "Materials Adjustment"
                  Top             =   210
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_ServiceIssuance 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":12FD9
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1312B
                  Style           =   1  'Graphical
                  TabIndex        =   131
                  Tag             =   "1305"
                  ToolTipText     =   "Materials Requisition Issuance"
                  Top             =   1744
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_OverTheCounter 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":13965
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":13AB7
                  Style           =   1  'Graphical
                  TabIndex        =   129
                  Tag             =   "1303"
                  ToolTipText     =   "Materials Issuance(Over the Counter)"
                  Top             =   1022
                  Width           =   720
               End
               Begin VB.CommandButton cmdMat_Requistion 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":141CA
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1431C
                  Style           =   1  'Graphical
                  TabIndex        =   126
                  Tag             =   "1467"
                  ToolTipText     =   "Materials Requisition Slip"
                  Top             =   300
                  Width           =   720
               End
               Begin XtremeSuiteControls.TabControl ss_Mat 
                  Height          =   4365
                  Left            =   4425
                  TabIndex        =   141
                  Top             =   990
                  Width           =   6375
                  _Version        =   655364
                  _ExtentX        =   11245
                  _ExtentY        =   7699
                  _StockProps     =   64
                  Appearance      =   2
                  Color           =   4
                  PaintManager.BoldSelected=   -1  'True
                  PaintManager.HotTracking=   -1  'True
                  PaintManager.ShowIcons=   -1  'True
                  PaintManager.LargeIcons=   -1  'True
                  PaintManager.MinTabWidth=   75
                  ItemCount       =   3
                  Item(0).Caption =   "Files"
                  Item(0).ControlCount=   1
                  Item(0).Control(0)=   "TabControlPage12"
                  Item(1).Caption =   "Inquiries"
                  Item(1).ControlCount=   1
                  Item(1).Control(0)=   "TabControlPage13"
                  Item(2).Caption =   "Reports"
                  Item(2).ControlCount=   1
                  Item(2).Control(0)=   "TabControlPage14"
                  Begin XtremeSuiteControls.TabControlPage TabControlPage12 
                     Height          =   3750
                     Left            =   30
                     TabIndex        =   142
                     Top             =   585
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdMatFiles_Master 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":14B14
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":14C66
                        Style           =   1  'Graphical
                        TabIndex        =   143
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Master File"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatFiles_Series 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":153D8
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1552A
                        Style           =   1  'Graphical
                        TabIndex        =   145
                        Tag             =   "1295"
                        ToolTipText     =   "Series No. Master File"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatFiles_PhysicalInvDatabase 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":15BDE
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":15D30
                        Style           =   1  'Graphical
                        TabIndex        =   147
                        Tag             =   "1295"
                        ToolTipText     =   "Create Physical Inventory Database"
                        Top             =   2025
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatFiles_PhysicalMenu 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":16446
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":16598
                        Style           =   1  'Graphical
                        TabIndex        =   149
                        Tag             =   "1295"
                        ToolTipText     =   "Physical Count Inventory Menu"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.Label Label80 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS MASTER FILE"
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
                        Height          =   330
                        Left            =   1080
                        TabIndex        =   144
                        Top             =   630
                        Width           =   4245
                     End
                     Begin VB.Label Label79 
                        BackStyle       =   0  'Transparent
                        Caption         =   "SERIES NO. MASTER FILE"
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
                        Height          =   330
                        Left            =   1080
                        TabIndex        =   146
                        Top             =   1425
                        Width           =   4245
                     End
                     Begin VB.Label Label57 
                        BackStyle       =   0  'Transparent
                        Caption         =   "CREATE PHYSICAL INVENTORY DATABASE"
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
                        Height          =   645
                        Left            =   1080
                        TabIndex        =   148
                        Top             =   2100
                        Width           =   4725
                     End
                     Begin VB.Label Label56 
                        BackStyle       =   0  'Transparent
                        Caption         =   "PHYSICAL COUNT INVENTORY MENU"
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
                        Height          =   570
                        Left            =   1080
                        TabIndex        =   150
                        Top             =   3000
                        Width           =   4605
                     End
                  End
                  Begin XtremeSuiteControls.TabControlPage TabControlPage13 
                     Height          =   3750
                     Left            =   -69970
                     TabIndex        =   160
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdMat_Inquiry_MRRTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":16D0E
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":16E60
                        Style           =   1  'Graphical
                        TabIndex        =   166
                        Tag             =   "1370"
                        ToolTipText     =   "MRR Transactions"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_POTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":175A7
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":176F9
                        Style           =   1  'Graphical
                        TabIndex        =   163
                        Tag             =   "1369"
                        ToolTipText     =   "PO Transactions"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_IssuanceTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":17E56
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":17FA8
                        Style           =   1  'Graphical
                        TabIndex        =   171
                        Tag             =   "1371"
                        ToolTipText     =   "Issuances Transactions"
                        Top             =   2040
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_TransactionDetail 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":1861E
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":18770
                        Style           =   1  'Graphical
                        TabIndex        =   175
                        Tag             =   "1372"
                        ToolTipText     =   "Transaction Details"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_CheckRunning 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":18DDE
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":18F30
                        Style           =   1  'Graphical
                        TabIndex        =   173
                        Tag             =   "1295"
                        ToolTipText     =   "Check Inventory Balances"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_CounterInquiry 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1952E
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":19680
                        Style           =   1  'Graphical
                        TabIndex        =   161
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Counter Inquiry"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_Location 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":19D34
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":19E86
                        Style           =   1  'Graphical
                        TabIndex        =   165
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Location File"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMat_Inquiry_Ledger 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1A514
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1A666
                        Style           =   1  'Graphical
                        TabIndex        =   169
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Ledger File"
                        Top             =   2040
                        Width           =   720
                     End
                     Begin VB.Label Label45 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS TRANSACTION DETAILS"
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
                        Height          =   450
                        Left            =   3840
                        TabIndex        =   176
                        Top             =   2910
                        Width           =   2025
                     End
                     Begin VB.Label Label43 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS ISSUANCES TRANSACTIONS"
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
                        Height          =   450
                        Left            =   3840
                        TabIndex        =   172
                        Top             =   2130
                        Width           =   2055
                     End
                     Begin VB.Label Label40 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS MRR TRANSACTIONS"
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
                        Height          =   450
                        Left            =   3840
                        TabIndex        =   168
                        Top             =   1350
                        Width           =   1905
                     End
                     Begin VB.Label Label34 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS PO TRANSACTIONS"
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
                        Height          =   480
                        Left            =   3840
                        TabIndex        =   164
                        Top             =   555
                        Width           =   1875
                     End
                     Begin VB.Label Label97 
                        BackStyle       =   0  'Transparent
                        Caption         =   "CHECK INVENTORY BALANCES"
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
                        Height          =   450
                        Left            =   1095
                        TabIndex        =   174
                        Top             =   2910
                        Width           =   1605
                     End
                     Begin VB.Label Label91 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS COUNTER INQUIRY"
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
                        Height          =   450
                        Left            =   1095
                        TabIndex        =   162
                        Top             =   555
                        Width           =   1665
                     End
                     Begin VB.Label Label90 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS LOCATION FILE"
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
                        Height          =   450
                        Left            =   1095
                        TabIndex        =   167
                        Top             =   1350
                        Width           =   1905
                     End
                     Begin VB.Label Label89 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MATERIALS LEDGER FILE"
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
                        Height          =   450
                        Left            =   1095
                        TabIndex        =   170
                        Top             =   2130
                        Width           =   1635
                     End
                  End
                  Begin XtremeSuiteControls.TabControlPage TabControlPage14 
                     Height          =   3750
                     Left            =   -69970
                     TabIndex        =   151
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdMatReports_DailySales 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1ADA6
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1AEF8
                        Style           =   1  'Graphical
                        TabIndex        =   152
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Daily Sales Report"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatReports_TransListing 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1B677
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1B7C9
                        Style           =   1  'Graphical
                        TabIndex        =   154
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Transaction Listings"
                        Top             =   1230
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatReports_MonthlyReports 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1BEB3
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1C005
                        Style           =   1  'Graphical
                        TabIndex        =   156
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Monthly Reports"
                        Top             =   1995
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdMatReports_MonthEndReport 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":1C6D6
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":1C828
                        Style           =   1  'Graphical
                        TabIndex        =   158
                        Tag             =   "1295"
                        ToolTipText     =   "Materials Month-End Inventory Reports"
                        Top             =   2775
                        Width           =   720
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "DAILY SALES REPORT"
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
                        Height          =   330
                        Index           =   30
                        Left            =   1065
                        TabIndex        =   153
                        Top             =   690
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "TRANSACTION LISTINGS"
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
                        Height          =   330
                        Index           =   31
                        Left            =   1065
                        TabIndex        =   155
                        Top             =   1440
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MONTHLY REPORTS"
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
                        Height          =   330
                        Index           =   32
                        Left            =   1065
                        TabIndex        =   157
                        Top             =   2220
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MONTH-END INVENTORY REPORTS"
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
                        Height          =   570
                        Index           =   33
                        Left            =   1065
                        TabIndex        =   159
                        Top             =   2970
                        Width           =   4605
                     End
                  End
               End
               Begin VB.PictureBox picMATSelect 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   2055
                  Left            =   3480
                  ScaleHeight     =   2025
                  ScaleWidth      =   2955
                  TabIndex        =   177
                  Top             =   1530
                  Width           =   2985
                  Begin VB.CommandButton cmdMATCancel 
                     Caption         =   "Cancel"
                     Height          =   405
                     Left            =   1410
                     TabIndex        =   182
                     Top             =   1500
                     Width           =   1035
                  End
                  Begin VB.OptionButton optMatIssuances 
                     Caption         =   "Issuances"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   180
                     Top             =   750
                     Width           =   2385
                  End
                  Begin VB.CommandButton cmdMATOk 
                     Caption         =   "Ok"
                     Height          =   405
                     Left            =   390
                     TabIndex        =   183
                     Top             =   1500
                     Width           =   1035
                  End
                  Begin VB.OptionButton optMatReceipts 
                     Caption         =   "Receiving"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   179
                     Top             =   420
                     Value           =   -1  'True
                     Width           =   2385
                  End
                  Begin VB.OptionButton optPO_Mat 
                     Caption         =   "Purchase Order"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   181
                     Top             =   1080
                     Width           =   2385
                  End
                  Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
                     Height          =   345
                     Left            =   0
                     TabIndex        =   178
                     Top             =   0
                     Width           =   2955
                     _Version        =   655364
                     _ExtentX        =   5212
                     _ExtentY        =   609
                     _StockProps     =   14
                     Caption         =   "Select Report Type"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "SERVICE ISSUANCES AGAINST ADVANCE BILL"
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
                  Height          =   525
                  Index           =   43
                  Left            =   8190
                  TabIndex        =   235
                  Top             =   330
                  Width           =   2505
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ADVANCE BILL DATA ENTRY"
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
                  Height          =   345
                  Index           =   41
                  Left            =   1245
                  TabIndex        =   140
                  Top             =   2610
                  Width           =   3015
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "DR OUT ISSUANCE"
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
                  Height          =   735
                  Index           =   38
                  Left            =   1245
                  TabIndex        =   135
                  Top             =   3360
                  Width           =   2655
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS INVENTORY ADJUSTMENT"
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
                  Height          =   675
                  Index           =   34
                  Left            =   5250
                  TabIndex        =   127
                  Top             =   195
                  Width           =   1845
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS ISSUANCE (SERVICE ISSUANCE)"
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
                  Height          =   735
                  Index           =   37
                  Left            =   1245
                  TabIndex        =   132
                  Top             =   1785
                  Width           =   2655
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS ISSUANCE (OVER THE COUNTER)"
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
                  Height          =   555
                  Index           =   36
                  Left            =   1245
                  TabIndex        =   130
                  Top             =   1065
                  Width           =   2325
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS PURCHASE ORDER"
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
                  Height          =   450
                  Index           =   39
                  Left            =   1245
                  TabIndex        =   137
                  Top             =   4065
                  Width           =   2775
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS RECEIVING AND STORING"
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
                  Height          =   435
                  Index           =   40
                  Left            =   1245
                  TabIndex        =   139
                  Top             =   4710
                  Width           =   2370
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "MATERIALS REQUISITION SLIP"
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
                  Height          =   210
                  Index           =   35
                  Left            =   1245
                  TabIndex        =   128
                  Top             =   480
                  Width           =   2940
               End
            End
            Begin XtremeSuiteControls.TabControlPage TabControlPage7 
               Height          =   5460
               Left            =   -69970
               TabIndex        =   43
               Top             =   585
               Visible         =   0   'False
               Width           =   10815
               _Version        =   655364
               _ExtentX        =   19076
               _ExtentY        =   9631
               _StockProps     =   0
               Begin VB.CommandButton Command3 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   7440
                  MouseIcon       =   "PMISMainMenu.frx":1CF5D
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1D0AF
                  Style           =   1  'Graphical
                  TabIndex        =   236
                  Tag             =   "1307"
                  ToolTipText     =   "Advance Bill Data Entry"
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_ISS_CHG 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":1D918
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1DA6A
                  Style           =   1  'Graphical
                  TabIndex        =   53
                  Tag             =   "1303"
                  ToolTipText     =   "Accessories Issuance(Over the Counter)"
                  Top             =   1744
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_ISS_DR 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":1E157
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1E2A9
                  Style           =   1  'Graphical
                  TabIndex        =   55
                  Tag             =   "1306"
                  ToolTipText     =   "DR Out Issuance"
                  Top             =   3188
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_Requisition 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":1EB25
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1EC77
                  Style           =   1  'Graphical
                  TabIndex        =   46
                  Tag             =   "1467"
                  ToolTipText     =   "Accessories Requisition Slip"
                  Top             =   300
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_ISS_CSH 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":1F508
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1F65A
                  Style           =   1  'Graphical
                  TabIndex        =   47
                  Tag             =   "1303"
                  ToolTipText     =   "Accessories Issuance(Over the Counter)"
                  Top             =   1022
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_ISS_RIV 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":1FD47
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":1FE99
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  Tag             =   "1305"
                  ToolTipText     =   "Accessories Requisition Issuance"
                  Top             =   2466
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_Adjustment 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   4425
                  MouseIcon       =   "PMISMainMenu.frx":206DF
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":20831
                  Style           =   1  'Graphical
                  TabIndex        =   45
                  Tag             =   "1308"
                  ToolTipText     =   "Accessories Adjustment"
                  Top             =   255
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_RR 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":210A1
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":211F3
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  Tag             =   "1319"
                  ToolTipText     =   "Accessories Receiving and Storing"
                  Top             =   4635
                  Width           =   720
               End
               Begin VB.CommandButton cmdAcc_PO 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":21958
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":21AAA
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  Tag             =   "1318"
                  ToolTipText     =   "Accessories Purchase Order"
                  Top             =   3910
                  Width           =   720
               End
               Begin XtremeSuiteControls.TabControl ss_AC 
                  Height          =   4365
                  Left            =   4425
                  TabIndex        =   60
                  Top             =   990
                  Width           =   6375
                  _Version        =   655364
                  _ExtentX        =   11245
                  _ExtentY        =   7699
                  _StockProps     =   64
                  Appearance      =   2
                  Color           =   4
                  PaintManager.BoldSelected=   -1  'True
                  PaintManager.HotTracking=   -1  'True
                  PaintManager.ShowIcons=   -1  'True
                  PaintManager.LargeIcons=   -1  'True
                  PaintManager.MinTabWidth=   75
                  ItemCount       =   3
                  Item(0).Caption =   "Files"
                  Item(0).ControlCount=   1
                  Item(0).Control(0)=   "TabControlPage9"
                  Item(1).Caption =   "Inquiries"
                  Item(1).ControlCount=   1
                  Item(1).Control(0)=   "TabControlPage10"
                  Item(2).Caption =   "Reports"
                  Item(2).ControlCount=   1
                  Item(2).Control(0)=   "TabControlPage11"
                  Begin XtremeSuiteControls.TabControlPage TabControlPage11 
                     Height          =   3750
                     Left            =   -69970
                     TabIndex        =   61
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdAcc_ReportDailySales 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":222E4
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":22436
                        Style           =   1  'Graphical
                        TabIndex        =   62
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Daily Sales Report"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_ReportTransListing 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":22BB5
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":22D07
                        Style           =   1  'Graphical
                        TabIndex        =   64
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Transaction Listings"
                        Top             =   1230
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_ReportMonthlyReport 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":233F1
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":23543
                        Style           =   1  'Graphical
                        TabIndex        =   66
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Monthly Reports"
                        Top             =   1995
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_ReportMonthEndReport 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":23C14
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":23D66
                        Style           =   1  'Graphical
                        TabIndex        =   68
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Month-End Inventory Reports"
                        Top             =   2775
                        Width           =   720
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "DAILY SALES REPORT"
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
                        Height          =   330
                        Index           =   26
                        Left            =   1065
                        TabIndex        =   63
                        Top             =   690
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "TRANSACTION LISTINGS"
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
                        Height          =   330
                        Index           =   27
                        Left            =   1065
                        TabIndex        =   65
                        Top             =   1440
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MONTHLY REPORTS"
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
                        Height          =   330
                        Index           =   28
                        Left            =   1065
                        TabIndex        =   67
                        Top             =   2220
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "MONTH-END INVENTORY REPORTS"
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
                        Height          =   570
                        Index           =   29
                        Left            =   1065
                        TabIndex        =   69
                        Top             =   2970
                        Width           =   4605
                     End
                  End
                  Begin XtremeSuiteControls.TabControlPage TabControlPage10 
                     Height          =   3750
                     Left            =   -69970
                     TabIndex        =   70
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdAcc_Inq_Counter 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":2449B
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":245ED
                        Style           =   1  'Graphical
                        TabIndex        =   71
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Counter Inquiry"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_Location 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":24CA1
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":24DF3
                        Style           =   1  'Graphical
                        TabIndex        =   75
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Location File"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_Ledger 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":25481
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":255D3
                        Style           =   1  'Graphical
                        TabIndex        =   79
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Ledger File"
                        Top             =   2040
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_CheckInvBal 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":25D13
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":25E65
                        Style           =   1  'Graphical
                        TabIndex        =   83
                        Tag             =   "1295"
                        ToolTipText     =   "Check Inventory Balances"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_MRRTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":26463
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":265B5
                        Style           =   1  'Graphical
                        TabIndex        =   76
                        Tag             =   "1370"
                        ToolTipText     =   "MRR Transactions"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_POTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":26CFC
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":26E4E
                        Style           =   1  'Graphical
                        TabIndex        =   73
                        Tag             =   "1369"
                        ToolTipText     =   "PO Transactions"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_IssuancesTransaction 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":275AB
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":276FD
                        Style           =   1  'Graphical
                        TabIndex        =   80
                        Tag             =   "1371"
                        ToolTipText     =   "Issuances Transactions"
                        Top             =   2040
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Inq_TransactionDetail 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   2985
                        MouseIcon       =   "PMISMainMenu.frx":27D73
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":27EC5
                        Style           =   1  'Graphical
                        TabIndex        =   84
                        Tag             =   "1372"
                        ToolTipText     =   "Transaction Details"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES COUNTER INQUIRY"
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
                        Height          =   480
                        Index           =   22
                        Left            =   1095
                        TabIndex        =   72
                        Top             =   555
                        Width           =   1875
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES LOCATION FILE"
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
                        Height          =   450
                        Index           =   23
                        Left            =   1095
                        TabIndex        =   77
                        Top             =   1350
                        Width           =   1905
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES LEDGER FILE"
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
                        Height          =   450
                        Index           =   24
                        Left            =   1095
                        TabIndex        =   81
                        Top             =   2130
                        Width           =   1935
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "CHECK INVENTORY BALANCES"
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
                        Height          =   450
                        Index           =   25
                        Left            =   1095
                        TabIndex        =   86
                        Top             =   2910
                        Width           =   1965
                     End
                     Begin VB.Label Label19 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES TRANSACTION DETAILS"
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
                        Height          =   450
                        Left            =   3840
                        TabIndex        =   85
                        Top             =   2895
                        Width           =   2385
                     End
                     Begin VB.Label Label20 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES ISSUANCES TRANSACTIONS"
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
                        Height          =   450
                        Left            =   3840
                        TabIndex        =   82
                        Top             =   2130
                        Width           =   2355
                     End
                     Begin VB.Label Label22 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES MRR TRANSACTIONS"
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
                        Height          =   450
                        Left            =   3825
                        TabIndex        =   78
                        Top             =   1350
                        Width           =   2325
                     End
                     Begin VB.Label Label33 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES PO TRANSACTIONS"
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
                        Height          =   480
                        Left            =   3840
                        TabIndex        =   74
                        Top             =   555
                        Width           =   1875
                     End
                  End
                  Begin XtremeSuiteControls.TabControlPage TabControlPage9 
                     Height          =   3750
                     Left            =   30
                     TabIndex        =   87
                     Top             =   585
                     Width           =   6315
                     _Version        =   655364
                     _ExtentX        =   11139
                     _ExtentY        =   6615
                     _StockProps     =   0
                     Begin VB.CommandButton cmdAcc_Files_AcMasterFile 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":28533
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":28685
                        Style           =   1  'Graphical
                        TabIndex        =   88
                        Tag             =   "1295"
                        ToolTipText     =   "Accessories Master File"
                        Top             =   480
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Files_Series 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":28E00
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":28F52
                        Style           =   1  'Graphical
                        TabIndex        =   90
                        Tag             =   "1295"
                        ToolTipText     =   "Series No. Master File"
                        Top             =   1260
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Files_CreatePhysicalInventory 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":29606
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":29758
                        Style           =   1  'Graphical
                        TabIndex        =   92
                        Tag             =   "1295"
                        ToolTipText     =   "Create Physical Inventory Database"
                        Top             =   2025
                        Width           =   720
                     End
                     Begin VB.CommandButton cmdAcc_Files_PhysicalCountInventoryMenu 
                        BeginProperty Font 
                           Name            =   "Verdana"
                           Size            =   14.25
                           Charset         =   0
                           Weight          =   400
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   645
                        Left            =   270
                        MouseIcon       =   "PMISMainMenu.frx":29E6E
                        MousePointer    =   99  'Custom
                        Picture         =   "PMISMainMenu.frx":29FC0
                        Style           =   1  'Graphical
                        TabIndex        =   94
                        Tag             =   "1295"
                        ToolTipText     =   "Physical Count Inventory Menu"
                        Top             =   2805
                        Width           =   720
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "ACCESSORIES MASTER FILE"
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
                        Height          =   330
                        Index           =   18
                        Left            =   1080
                        TabIndex        =   89
                        Top             =   630
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "SERIES NO. MASTER FILE"
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
                        Height          =   330
                        Index           =   19
                        Left            =   1080
                        TabIndex        =   91
                        Top             =   1425
                        Width           =   4245
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "CREATE PHYSICAL INVENTORY DATABASE"
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
                        Height          =   615
                        Index           =   20
                        Left            =   1080
                        TabIndex        =   93
                        Top             =   2100
                        Width           =   4740
                     End
                     Begin VB.Label Label38 
                        BackStyle       =   0  'Transparent
                        Caption         =   "PHYSICAL COUNT INVENTORY MENU"
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
                        Height          =   570
                        Index           =   21
                        Left            =   1080
                        TabIndex        =   95
                        Top             =   3000
                        Width           =   4605
                     End
                  End
               End
               Begin VB.PictureBox picACSelect 
                  Appearance      =   0  'Flat
                  ForeColor       =   &H80000008&
                  Height          =   2055
                  Left            =   3480
                  ScaleHeight     =   2025
                  ScaleWidth      =   2955
                  TabIndex        =   96
                  Top             =   1530
                  Width           =   2985
                  Begin VB.OptionButton optPO_AC 
                     Caption         =   "Purchase Order"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   100
                     Top             =   1080
                     Width           =   2385
                  End
                  Begin VB.OptionButton optACReceipt 
                     Caption         =   "Receiving"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   98
                     Top             =   420
                     Value           =   -1  'True
                     Width           =   2385
                  End
                  Begin VB.OptionButton optACIssuances 
                     Caption         =   "Issuances"
                     Height          =   285
                     Left            =   300
                     TabIndex        =   99
                     Top             =   750
                     Width           =   2385
                  End
                  Begin VB.CommandButton cmdCANCELACSelect 
                     Caption         =   "Cancel"
                     Height          =   405
                     Left            =   1410
                     TabIndex        =   101
                     Top             =   1500
                     Width           =   1035
                  End
                  Begin VB.CommandButton cmdACokSelect 
                     Caption         =   "Ok"
                     Height          =   405
                     Left            =   390
                     TabIndex        =   102
                     Top             =   1500
                     Width           =   1035
                  End
                  Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
                     Height          =   345
                     Left            =   0
                     TabIndex        =   97
                     Top             =   0
                     Width           =   2955
                     _Version        =   655364
                     _ExtentX        =   5212
                     _ExtentY        =   609
                     _StockProps     =   14
                     Caption         =   "Select Report Type"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                  End
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "SERVICE ISSUANCES AGAINST ADVANCE BILL"
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
                  Height          =   525
                  Index           =   44
                  Left            =   8310
                  TabIndex        =   237
                  Top             =   330
                  Visible         =   0   'False
                  Width           =   2505
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES ISSUANCE (OVER THE COUNTER - CHG)"
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
                  Height          =   555
                  Index           =   12
                  Left            =   1245
                  TabIndex        =   49
                  Top             =   1800
                  Width           =   2475
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "DR OUT ISSUANCE"
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
                  Height          =   225
                  Index           =   14
                  Left            =   1245
                  TabIndex        =   51
                  Top             =   3405
                  Width           =   2655
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES REQUISITION SLIP"
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
                  Height          =   300
                  Index           =   10
                  Left            =   1245
                  TabIndex        =   44
                  Top             =   465
                  Width           =   2955
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES RECEIVING AND STORING"
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
                  Height          =   435
                  Index           =   16
                  Left            =   1245
                  TabIndex        =   59
                  Top             =   4725
                  Width           =   2370
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES PURCHASE ORDER"
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
                  Height          =   420
                  Index           =   15
                  Left            =   1245
                  TabIndex        =   57
                  Top             =   3990
                  Width           =   2775
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES ISSUANCE (OVER THE COUNTER - CSH)"
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
                  Height          =   555
                  Index           =   11
                  Left            =   1245
                  TabIndex        =   48
                  Top             =   1065
                  Width           =   2415
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES ISSUANCE (SERVICE ISSUANCE)"
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
                  Height          =   735
                  Index           =   13
                  Left            =   1245
                  TabIndex        =   50
                  Top             =   2505
                  Width           =   2655
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ACCESSORIES INVENTORY ADJUSTMENT"
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
                  Height          =   465
                  Index           =   17
                  Left            =   5190
                  TabIndex        =   52
                  Top             =   345
                  Width           =   2205
               End
            End
            Begin XtremeSuiteControls.TabControlPage TabControlPage6 
               Height          =   5460
               Left            =   30
               TabIndex        =   103
               Top             =   585
               Width           =   10815
               _Version        =   655364
               _ExtentX        =   19076
               _ExtentY        =   9631
               _StockProps     =   0
               Begin VB.CommandButton Command1 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   450
                  MouseIcon       =   "PMISMainMenu.frx":2A736
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2A888
                  Style           =   1  'Graphical
                  TabIndex        =   232
                  Tag             =   "1307"
                  ToolTipText     =   "Advance Bill Data Entry"
                  Top             =   4470
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsPO 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   5160
                  MouseIcon       =   "PMISMainMenu.frx":2B0F1
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2B243
                  Style           =   1  'Graphical
                  TabIndex        =   108
                  Tag             =   "1318"
                  ToolTipText     =   "Purchase Order Data Entry"
                  Top             =   1245
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsReceiving 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   5160
                  MouseIcon       =   "PMISMainMenu.frx":2BA23
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2BB75
                  Style           =   1  'Graphical
                  TabIndex        =   111
                  Tag             =   "1319"
                  ToolTipText     =   "Purchase Receiving and Storing"
                  Top             =   2220
                  Width           =   735
               End
               Begin VB.CommandButton cmdMain_PartsInventoryAdjustment 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   5160
                  MouseIcon       =   "PMISMainMenu.frx":2C3E5
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2C537
                  Style           =   1  'Graphical
                  TabIndex        =   118
                  Tag             =   "1308"
                  ToolTipText     =   "Inventory Adjustment"
                  Top             =   4200
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsAdvanceBilling 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   5160
                  MouseIcon       =   "PMISMainMenu.frx":2CDAB
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2CEFD
                  Style           =   1  'Graphical
                  TabIndex        =   104
                  Tag             =   "1307"
                  ToolTipText     =   "Advance Bill Data Entry"
                  Top             =   270
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PQIR 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   5160
                  MouseIcon       =   "PMISMainMenu.frx":2D766
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2D8B8
                  Style           =   1  'Graphical
                  TabIndex        =   115
                  Tag             =   "1470"
                  ToolTipText     =   "Parts Quality Information Report"
                  Top             =   3225
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsIssuanceOC_CHG 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":2DF26
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2E078
                  Style           =   1  'Graphical
                  TabIndex        =   112
                  Tag             =   "1303"
                  ToolTipText     =   "Parts Issuance(Over the Counter)"
                  Top             =   1968
                  Width           =   735
               End
               Begin VB.CommandButton cmdMain_PartsDROut 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":2E7E5
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2E937
                  Style           =   1  'Graphical
                  TabIndex        =   117
                  Tag             =   "1306"
                  ToolTipText     =   "DR Out Issuance"
                  Top             =   3636
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsIssuanceServiceIssuance 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":2F1B3
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2F305
                  Style           =   1  'Graphical
                  TabIndex        =   114
                  Tag             =   "1305"
                  ToolTipText     =   "Parts Issuance(Service Issuance)"
                  Top             =   2802
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsIssuanceOC_CSH 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":2FA2B
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":2FB7D
                  Style           =   1  'Graphical
                  TabIndex        =   109
                  Tag             =   "1303"
                  ToolTipText     =   "Parts Issuance(Over the Counter)"
                  Top             =   1134
                  Width           =   720
               End
               Begin VB.CommandButton cmdMain_PartsRequisition 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   645
                  Left            =   405
                  MouseIcon       =   "PMISMainMenu.frx":302EA
                  MousePointer    =   99  'Custom
                  Picture         =   "PMISMainMenu.frx":3043C
                  Style           =   1  'Graphical
                  TabIndex        =   105
                  Tag             =   "1467"
                  ToolTipText     =   "Parts Requisition Slip"
                  Top             =   300
                  Width           =   720
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "SERVICE ISSUANCES AGAINST ADVANCE BILL"
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
                  Height          =   525
                  Index           =   42
                  Left            =   1290
                  TabIndex        =   233
                  Top             =   4590
                  Width           =   3105
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS INVENTORY ADJUSTMENT"
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
                  Index           =   5
                  Left            =   6075
                  TabIndex        =   119
                  Top             =   4410
                  Width           =   3735
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ADVANCE BILL DATA ENTRY"
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
                  Height          =   345
                  Index           =   1
                  Left            =   6060
                  TabIndex        =   107
                  Top             =   480
                  Width           =   3015
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PURCHASE ORDER DATA ENTRY"
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
                  Height          =   300
                  Index           =   2
                  Left            =   6045
                  TabIndex        =   110
                  Top             =   1425
                  Width           =   3810
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PURCHASE RECEIVING AND STORING"
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
                  Height          =   255
                  Index           =   3
                  Left            =   6060
                  TabIndex        =   113
                  Top             =   2415
                  Width           =   4335
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS QUALITY INFORMATION REPORT"
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
                  Index           =   4
                  Left            =   6060
                  TabIndex        =   116
                  Top             =   3405
                  Width           =   3945
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS ISSUANCE        (OVER THE COUNTER - CHG)"
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
                  Height          =   555
                  Index           =   7
                  Left            =   1245
                  TabIndex        =   121
                  Top             =   2025
                  Width           =   2385
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "DR OUT ISSUANCE"
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
                  Index           =   9
                  Left            =   1245
                  TabIndex        =   123
                  Top             =   3825
                  Width           =   1725
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS ISSUANCE        (SERVICE ISSUANCE)"
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
                  Height          =   525
                  Index           =   8
                  Left            =   1245
                  TabIndex        =   122
                  Top             =   2895
                  Width           =   1965
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS ISSUANCE        (OVER THE COUNTER - CSH)"
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
                  Height          =   555
                  Index           =   6
                  Left            =   1245
                  TabIndex        =   120
                  Top             =   1200
                  Width           =   2385
               End
               Begin VB.Label Label38 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PARTS REQUISITION SLIP"
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
                  Height          =   330
                  Index           =   0
                  Left            =   1245
                  TabIndex        =   106
                  Top             =   540
                  Width           =   2670
               End
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   6060
         Left            =   -69970
         TabIndex        =   184
         Top             =   585
         Visible         =   0   'False
         Width           =   10875
         _Version        =   655364
         _ExtentX        =   19182
         _ExtentY        =   10689
         _StockProps     =   0
         Begin VB.CommandButton cmdTable_CustomerMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":30AB2
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":30C04
            Style           =   1  'Graphical
            TabIndex        =   185
            Tag             =   "1289"
            ToolTipText     =   "Customer Master File"
            Top             =   180
            Width           =   720
         End
         Begin VB.CommandButton cmdTable_SupplierMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":3126B
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":313BD
            Style           =   1  'Graphical
            TabIndex        =   193
            Tag             =   "1291"
            ToolTipText     =   "Supplier Master File"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdTable_HariMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":31AD3
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":31C25
            Style           =   1  'Graphical
            TabIndex        =   189
            Tag             =   "1290"
            ToolTipText     =   "HARI Parts Master File"
            Top             =   1035
            Width           =   720
         End
         Begin VB.CommandButton cmdTable_PartsMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":32335
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":32487
            Style           =   1  'Graphical
            TabIndex        =   197
            Tag             =   "1294"
            ToolTipText     =   "Parts Master File"
            Top             =   2745
            Width           =   720
         End
         Begin VB.CommandButton cmdTable_CounterMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":32BE3
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":32D35
            Style           =   1  'Graphical
            TabIndex        =   205
            Tag             =   "1297"
            ToolTipText     =   "Counter Master File"
            Top             =   4470
            Width           =   720
         End
         Begin VB.CommandButton cmdTable_SalesManMasterFile 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   375
            MouseIcon       =   "PMISMainMenu.frx":33487
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":335D9
            Style           =   1  'Graphical
            TabIndex        =   201
            Tag             =   "1296"
            ToolTipText     =   "Salesman Master File"
            Top             =   3615
            Width           =   720
         End
         Begin VB.CommandButton cmdFile_BackOrder 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5070
            MouseIcon       =   "PMISMainMenu.frx":33CFF
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":33E51
            Style           =   1  'Graphical
            TabIndex        =   187
            Tag             =   "1319"
            ToolTipText     =   "Create Purchase Order to Mobis"
            Top             =   375
            Width           =   720
         End
         Begin VB.CommandButton cmdFile_UpdateLocation 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5070
            MouseIcon       =   "PMISMainMenu.frx":344F1
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":34643
            Style           =   1  'Graphical
            TabIndex        =   191
            Tag             =   "1287"
            ToolTipText     =   "Update Location"
            Top             =   1155
            Width           =   720
         End
         Begin VB.CommandButton cmdFile_PhysicalMenu 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5055
            MouseIcon       =   "PMISMainMenu.frx":34C7A
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":34DCC
            Style           =   1  'Graphical
            TabIndex        =   195
            Tag             =   "1299"
            ToolTipText     =   "Physical Inventory Menu"
            Top             =   1995
            Width           =   720
         End
         Begin VB.CommandButton cmdFile_CreateInventoryDatabase 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5040
            MouseIcon       =   "PMISMainMenu.frx":354C2
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":35614
            Style           =   1  'Graphical
            TabIndex        =   199
            Tag             =   "1301"
            ToolTipText     =   "Create Inventory Database"
            Top             =   2850
            Width           =   720
         End
         Begin VB.CommandButton cmdFile_Location 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5040
            MouseIcon       =   "PMISMainMenu.frx":35C93
            MousePointer    =   99  'Custom
            Picture         =   "PMISMainMenu.frx":35DE5
            Style           =   1  'Graphical
            TabIndex        =   203
            Tag             =   "1302"
            ToolTipText     =   "Location"
            Top             =   3705
            Width           =   720
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER MASTER FILE"
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
            Height          =   330
            Left            =   1185
            TabIndex        =   186
            Top             =   420
            Width           =   3225
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "COUNTER MASTER FILE"
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
            Height          =   330
            Left            =   1140
            TabIndex        =   206
            Top             =   4680
            Width           =   3225
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "SALESMAN MASTER FILE"
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
            Height          =   330
            Left            =   1155
            TabIndex        =   202
            Top             =   3825
            Width           =   3225
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS MASTER FILE"
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
            Height          =   330
            Left            =   1185
            TabIndex        =   198
            Top             =   3000
            Width           =   3225
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "SUPPLIER MASTER FILE"
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
            Height          =   330
            Left            =   1185
            TabIndex        =   194
            Top             =   2145
            Width           =   3225
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "HARI PARTS MASTER FILE"
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
            Height          =   330
            Left            =   1185
            TabIndex        =   190
            Top             =   1290
            Width           =   3225
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "BACK-ORDER ALLOCATION"
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
            Height          =   330
            Left            =   5940
            TabIndex        =   188
            Top             =   555
            Width           =   5295
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTS LOCATION FILE"
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
            Height          =   330
            Left            =   5895
            TabIndex        =   204
            Top             =   3885
            Width           =   3225
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PHYSICAL INVENTORY MENU"
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
            Height          =   225
            Left            =   5895
            TabIndex        =   196
            Top             =   2175
            Width           =   2415
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CREATE INVENTORY DATABASE"
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
            Height          =   225
            Left            =   5895
            TabIndex        =   200
            Top             =   3030
            Width           =   2670
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "UPDATE LOCATION"
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
            Height          =   330
            Left            =   5925
            TabIndex        =   192
            Top             =   1350
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPORT_TYPE                                        As String



Private Sub cmdAcc_Adjustment_Click()
    If Module_Access(LOGID, "ACCESSORIES INVENTORY ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
    frmPMISTrans_InventoryAdjustment_Accessories.SetStockType ("A")
    FormExistsShow frmPMISTrans_InventoryAdjustment_Accessories
End Sub

Private Sub cmdAcc_Files_AcMasterFile_Click()
    If Module_Access(LOGID, "ACCESSORIES MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Accessories.SetStockType ("A")
    FormExistsShow frmMasterFile_Accessories
End Sub

Private Sub cmdAcc_Files_CreatePhysicalInventory_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    C_TYPE = "A": DESC_TYPE = "ACCESSORIES"
    FormExistsShow frmPMIS_Physical_CreateINVDATA
End Sub

Private Sub cmdAcc_Files_PhysicalCountInventoryMenu_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMIS_Physical_INVMenu_New
    C_TYPE = "A": DESC_TYPE = "ACCESSORIES"
    FormExistsShow frmPMIS_Physical_INVMenu_New

End Sub

Private Sub cmdAcc_Files_Series_Click()
    If Module_Access(LOGID, "ACCESSORIES COUNTER", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Counter_Accessories.SetStockType ("A")
    FormExistsShow frmMasterFile_Counter_Accessories
End Sub

Private Sub cmdAcc_Inq_CheckInvBal_Click()
    If Module_Access(LOGID, "ACCESSORIES CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
    frmPMISInquiry_CheckPrevBal_Accessories.SetStockType ("A")
    FormExistsShow frmPMISInquiry_CheckPrevBal_Accessories
End Sub

Private Sub cmdAcc_Inq_Counter_Click()
    If Module_Access(LOGID, "ACCESSORIES COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
    frmPMIS_CounterInquiry_Accessories.SETSTOCK_TYPE ("A")
    FormExistsShow frmPMIS_CounterInquiry_Accessories
End Sub

Private Sub cmdAcc_Inq_IssuancesTransaction_Click()

    If Module_Access(LOGID, "ACCESSORIES TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 5
    frmPMISInquiry_Query.SetTYPE ("A")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdAcc_Inq_Ledger_Click()
    If Module_Access(LOGID, "ACCESSORIES LEDGER FILE", "INQUIRY") = False Then Exit Sub
    Unload frmPMISInquiry_Query
    PARTSQUERY = 1
    frmPMISInquiry_Query.SetTYPE ("A")
    FormExistsShow frmPMISInquiry_Query

End Sub

Private Sub cmdAcc_Inq_Location_Click()
    If Module_Access(LOGID, "ACCESSORIES LOCATION", "REPORTS") = False Then Exit Sub
    frmPMISReports_Location_Accessories.SETSTOCK_TYPE ("A")
    FormExistsShow frmPMISReports_Location_Accessories
End Sub

Private Sub cmdAcc_Inq_MRRTransaction_Click()

    If Module_Access(LOGID, "ACCESSORIES MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 4
    frmPMISInquiry_Query.SetTYPE ("A")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdAcc_Inq_POTransaction_Click()
    If Module_Access(LOGID, "ACCESSORIES PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 3
    frmPMISInquiry_Query.SetTYPE ("A")
    FormExistsShow frmPMISInquiry_Query

End Sub

Private Sub cmdAcc_Inq_TransactionDetail_Click()
    If Module_Access(LOGID, "ACCESSORIES TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 7
    frmPMISInquiry_Query.SetTYPE ("A")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdAcc_ISS_CHG_Click()
    On Error Resume Next
    If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
    Unload frmPMISTrans_CustomerOrder_AC
    COUNTERTYPE = "CHG"
    frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CHG"
    FormExistsShow frmPMISTrans_CustomerOrder_AC
End Sub

Private Sub cmdAcc_ISS_CSH_Click()
    On Error Resume Next
    If Module_Access(LOGID, "ACCESSORIES ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
    Unload frmPMISTrans_CustomerOrder_AC
    COUNTERTYPE = "CSH"
    frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "CSH"
    FormExistsShow frmPMISTrans_CustomerOrder_AC
End Sub

Private Sub cmdAcc_ISS_DR_Click()
    If Module_Access(LOGID, "ACCESSORIES DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_AC
    COUNTERTYPE = "DR"
    frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "DR"
    FormExistsShow frmPMISTrans_CustomerOrder_AC
End Sub

Private Sub cmdAcc_ISS_Riv_Click()
    If Module_Access(LOGID, "ACCESSORIES SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_AC
    COUNTERTYPE = "RIV"
    frmPMISTrans_CustomerOrder_AC.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_CustomerOrder_AC
End Sub

Private Sub cmdAcc_PO_Click()
    If Module_Access(LOGID, "ACCESSORIES PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    FormExistsShow frmPMISAC_Purchase
End Sub

Private Sub cmdAcc_ReportDailySales_Click()
    If Module_Access(LOGID, "ACCESSORIES DAILY SALES", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_DailySales_AC
End Sub

Private Sub cmdAcc_ReportMonthEndReport_Click()
    If Module_Access(LOGID, "MONTH END REPORT", "REPORTS") = False Then Exit Sub

End Sub

Private Sub cmdAcc_ReportMonthlyReport_Click()
    If Module_Access(LOGID, "ACCESSORIES MONTHLY REPORT", "REPORTS") = False Then Exit Sub
    REPORT_TYPE = "ACC_MONTHREPORTS"
    picACSelect.Visible = True
End Sub

Private Sub cmdAcc_ReportTransListing_Click()
    If Module_Access(LOGID, "ACCESSORIES TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
    REPORT_TYPE = "ACC_TRANLIST"
    picACSelect.Visible = True
End Sub

Private Sub cmdAcc_Requisition_Click()
    On Error Resume Next
    If Module_Access(LOGID, "ACCESSORIES REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
    Unload frmPMISAC_ARISForms
    WAREHOUSETYPE = "ARS"
    frmPMISAC_ARISForms.txtTranType.Text = "ARS"
    FormExistsShow frmPMISAC_ARISForms
End Sub

Private Sub cmdAcc_RR_Click()
    If Module_Access(LOGID, "ACCESSORIES RECEIVING", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISAC_Receiving
End Sub

Private Sub cmdACokSelect_Click()
    If REPORT_TYPE = "ACC_TRANLIST" Then
        If optACReceipt.Value = True Then frmPMISReports_RCRange_AC.Show
        If optACIssuances.Value = True Then frmPMISReports_ISRange_AC.Show
        If optPO_AC.Value = True Then frmPMISReports_PORange_AC.Show
    End If
    If REPORT_TYPE = "ACC_MONTHREPORTS" Then
        If optACReceipt.Value = True Then frmPMISReports_Receipts_AC.Show
        If optACIssuances.Value = True Then frmPMISReports_Issuances_AC.Show
        If optPO_AC.Value = True Then frmPMISReports_Purchase_For_The_Month_AC.Show

    End If
    picACSelect.Visible = False
End Sub

Private Sub cmdAdvanceBill_Materials_Click()
    If Module_Access(LOGID, "MATERIALS ADVANCE BILL DATA ENTRY", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_MAT
    COUNTERTYPE = "ADB"
    frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "ADB"
    FormExistsShow frmPMISTrans_CustomerOrder_MAT
End Sub

Private Sub cmdCANCELACSelect_Click()
    picACSelect.Visible = False
End Sub

Private Sub cmdFile_BackOrder_Click()
    If Module_Access(LOGID, "BACK-ORDER ALLOCATION", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISTrans_POConfirmationProcess
End Sub

Private Sub cmdFile_CreateInventoryDatabase_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    Unload frmPMIS_Physical_INVMenu_New
    C_TYPE = "P": DESC_TYPE = "PART"
    FormExistsShow frmPMIS_Physical_CreateINVDATA
End Sub

Private Sub cmdFile_Location_Click()
    If Module_Access(LOGID, "LOCATION", "REPORTS") = False Then Exit Sub
    frmPMISReports_Location_Parts.SETSTOCK_TYPE ("P")
    FormExistsShow frmPMISReports_Location_Parts
End Sub

Private Sub cmdFile_PhysicalMenu_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    On Error Resume Next
    C_TYPE = "P": DESC_TYPE = "PART"
    FormExistsShow frmPMIS_Physical_INVMenu_New
End Sub

Private Sub cmdFile_SystemConfig_Click()

End Sub

Private Sub cmdFile_UpdateLocation_Click()
    If Module_Access(LOGID, "UPDATE LOCATION", "SYSTEM") = False Then Exit Sub
    FormExistsShow frmPMISUpdateLocation
End Sub

Private Sub cmdInquiry_BrowseErrorFiles_Click()
    If Module_Access(LOGID, "BROWSE ERROR FILES", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_ErrorQuery
End Sub

Private Sub cmdInquiry_CheckInventoryBalance_Click()
    If Module_Access(LOGID, "PARTS CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
    frmPMISInquiry_CheckPrevBal_Parts.SetStockType ("P")
    FormExistsShow frmPMISInquiry_CheckPrevBal_Parts
End Sub

Private Sub cmdInquiry_ComputeriedParts_Click()
    If Module_Access(LOGID, "PARTS COMPUTERIZED STOCKCARDS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 1
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdInquiry_Counter_Click()
    If Module_Access(LOGID, "COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
    frmPMIS_CounterInquiry_Parts.SETSTOCK_TYPE ("P")
    FormExistsShow frmPMIS_CounterInquiry_Parts
End Sub

Private Sub cmdInquiry_DealerDNPComparision_Click()
    If Module_Access(LOGID, "DEALER DISTRIBUTOR DNP COMPARISON", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_PartsDNPComparison
End Sub

Private Sub cmdInquiry_DealerSRPComparision_Click()
    If Module_Access(LOGID, "DEALER DISTRIBUTOR SRP COMPARISON", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_PartsSRPComparison
End Sub

Private Sub cmdInquiry_DealerSRPDNP_Click()
    If Module_Access(LOGID, "DEALER SRP DNP LISTING", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_PartsSRPComparison
End Sub

Private Sub cmdInquiry_DelaerPartInquiry_Click()
    If Module_Access(LOGID, "DEALER PARTS INQUIRY", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISTrans_DealerPartInquiry
End Sub

Private Sub cmdInquiry_InventoryRankingInquiry_Click()
    If Module_Access(LOGID, "INVENTORY RANKING INQUIRY", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_RankingInquiry
End Sub

Private Sub cmdInquiry_IssuanceTransaction_Click()
    If Module_Access(LOGID, "TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 5
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdInquiry_MRRTransaction_Click()
    If Module_Access(LOGID, "MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 4
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdInquiry_PartsAvalibity_Click()
    If Module_Access(LOGID, "PARTS AVAILABILITY", "INQUIRY") = False Then Exit Sub
    FormExistsShow frmPMISInquiry_PartsInquiry
End Sub

Private Sub cmdInquiry_POTransaction_Click()
    If Module_Access(LOGID, "PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 3
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdInquiry_TransactionDetails_Click()
    If Module_Access(LOGID, "TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 7
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdMain_PartsAdvanceBilling_Click()
    If Module_Access(LOGID, "PARTS ADVANCE BILL DATA ENTRY", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    COUNTERTYPE = "ADB"
    frmPMISTrans_CustomerOrder.txtTranType.Text = "ADB"
    FormExistsShow frmPMISTrans_CustomerOrder
 End Sub

Private Sub cmdMain_PartsDROut_Click()
    If Module_Access(LOGID, "PARTS DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    COUNTERTYPE = "DR"
    frmPMISTrans_CustomerOrder.txtTranType.Text = "DR"
    FormExistsShow frmPMISTrans_CustomerOrder
End Sub

Private Sub cmdMain_PartsInventoryAdjustment_Click()
    If Module_Access(LOGID, "PARTS INVENTORY ADJUSTMENT", "TRANSACTION") = False Then Exit Sub
    frmPMISTrans_InventoryAdjustment_Parts.SetStockType ("P")
    FormExistsShow frmPMISTrans_InventoryAdjustment_Parts
     
    
End Sub

Private Sub cmdMain_PartsIssuanceOC_CHG_Click()
    If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CHARGE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    COUNTERTYPE = "CHG"
    frmPMISTrans_CustomerOrder.txtTranType.Text = "CHG"
    FormExistsShow frmPMISTrans_CustomerOrder
End Sub

Private Sub cmdMain_PartsIssuanceOC_CSH_Click()
    If Module_Access(LOGID, "PARTS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_ADB_Issuances
    Unload frmPMISTrans_CustomerOrder
    COUNTERTYPE = "CSH"
    frmPMISTrans_CustomerOrder.txtTranType.Text = "CSH"
    FormExistsShow frmPMISTrans_CustomerOrder

End Sub

Private Sub cmdMain_PartsIssuanceServiceIssuance_Click()
    If Module_Access(LOGID, "PARTS ISSUANCE SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    COUNTERTYPE = "RIV"
    frmPMISTrans_CustomerOrder.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_CustomerOrder
End Sub

Private Sub cmdMain_PartsPO_Click()
    If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISTrans_Purchase
End Sub

Private Sub cmdMain_PartsReceiving_Click()
    If Module_Access(LOGID, "PURCHASE RECEIVING STORING", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISTrans_Receiving2
End Sub

Private Sub cmdMain_PartsRequisition_Click()
    If Module_Access(LOGID, "PARTS REQUISTION SLIP", "TRANSACTION") = False Then Exit Sub
    WAREHOUSETYPE = "PRS"
    frmPMISTrans_PrisForms.txtTranType.Text = "PRS"
    FormExistsShow frmPMISTrans_PrisForms
End Sub

Private Sub cmdMain_PQIR_Click()
    If Module_Access(LOGID, "PARTS QUALITY INFORMATION REPORT", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmPMISTrans_PQIR
End Sub

Private Sub cmdMat_Adjustment_Click()
    If Module_Access(LOGID, "MATERIALS INVENTORY ADJUSTMENT", "DATA ENTRY") = False Then Exit Sub
        frmPMISTrans_InventoryAdjustment_Materials.SetStockType ("M")
    FormExistsShow frmPMISTrans_InventoryAdjustment_Materials
End Sub

Private Sub cmdMat_DROut_Click()
    If Module_Access(LOGID, "MATERIALS DR OUT ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_MAT
    COUNTERTYPE = "DR"
    frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "DR"
    FormExistsShow frmPMISTrans_CustomerOrder_MAT
End Sub

Private Sub cmdMat_Inquiry_CheckRunning_Click()
    If Module_Access(LOGID, "MATERIALS CHECK PREVIOUS BALANCE", "PROCESSING") = False Then Exit Sub
    frmPMISInquiry_CheckPrevBal_Materials.SetStockType ("M")
    FormExistsShow frmPMISInquiry_CheckPrevBal_Materials
End Sub

Private Sub cmdMat_Inquiry_CounterInquiry_Click()
    If Module_Access(LOGID, "MATERIALS COUNTER INQUIRY", "INQUIRY") = False Then Exit Sub
    frmPMIS_CounterInquiry_Materials.SETSTOCK_TYPE ("M")
    FormExistsShow frmPMIS_CounterInquiry_Materials
End Sub

Private Sub cmdMat_Inquiry_IssuanceTransaction_Click()
    If Module_Access(LOGID, "MATERIAL TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 5
    frmPMISInquiry_Query.SetTYPE ("M")
    FormExistsShow frmPMISInquiry_Query

End Sub

Private Sub cmdMat_Inquiry_Ledger_Click()
    If Module_Access(LOGID, "MATERIALS QUERY", "INQUIRY") = False Then Exit Sub
    Unload frmPMISInquiry_Query
    PARTSQUERY = 1
    frmPMISInquiry_Query.SetTYPE ("M")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdMat_Inquiry_Location_Click()
    If Module_Access(LOGID, "MATERIALS LOCATION", "REPORTS") = False Then Exit Sub
    frmPMISReports_Location_Materials.SETSTOCK_TYPE ("M")
    FormExistsShow frmPMISReports_Location_Materials
End Sub

Private Sub cmdMat_Inquiry_MRRTransaction_Click()
    If Module_Access(LOGID, "MATERIAL MRR TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 4
    frmPMISInquiry_Query.SetTYPE ("M")
    FormExistsShow frmPMISInquiry_Query

End Sub

Private Sub cmdMat_Inquiry_POTransaction_Click()
    If Module_Access(LOGID, "MATERIAL PO TRANSACTIONS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 3
    frmPMISInquiry_Query.SetTYPE ("P")
    FormExistsShow frmPMISInquiry_Query

End Sub

Private Sub cmdMat_Inquiry_TransactionDetail_Click()
    If Module_Access(LOGID, "MATERIAL TRANSACTION DETAILS", "INQUIRY") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISInquiry_Query
    PARTSQUERY = 7
    frmPMISInquiry_Query.SetTYPE ("M")
    FormExistsShow frmPMISInquiry_Query
End Sub

Private Sub cmdMat_OverTheCounter_Click()
    On Error Resume Next
    If Module_Access(LOGID, "MATERIALS ISSUANCE COUNTER CASH", "TRANSACTION") = False Then Exit Sub
    Unload frmPMISTrans_CustomerOrder_MAT
    COUNTERTYPE = "CSH"
    frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "CSH"
    FormExistsShow frmPMISTrans_CustomerOrder_MAT
End Sub

Private Sub cmdMat_PO_Click()
    If Module_Access(LOGID, "MATERIALS PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISMAT_Purchase
End Sub

Private Sub cmdMat_Requistion_Click()
    On Error Resume Next
    If Module_Access(LOGID, "MATERIALS REQUISITION SLIP", "TRANSACTION") = False Then Exit Sub
    Unload frmPMISMAT_MRISForms
    WAREHOUSETYPE = "MRS"
    frmPMISMAT_MRISForms.txtTranType.Text = "MRS"
    FormExistsShow frmPMISMAT_MRISForms
End Sub

Private Sub cmdMat_RR_Click()
    If Module_Access(LOGID, "MATERIALS RECEIVING", "TRANSACTION") = False Then Exit Sub
    FormExistsShow frmPMISMAT_Receiving
End Sub

Private Sub cmdMat_ServiceIssuance_Click()
    If Module_Access(LOGID, "MATERIALS SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_MAT
    COUNTERTYPE = "RIV"
    frmPMISTrans_CustomerOrder_MAT.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_CustomerOrder_MAT
End Sub

Private Sub cmdMATCancel_Click()
    picMATSelect.Visible = False
End Sub

Private Sub cmdMatFiles_Master_Click()
    If Module_Access(LOGID, "MATERIALS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Material.SetStockType ("M")
    FormExistsShow frmMasterFile_Material
End Sub

Private Sub cmdMatFiles_PhysicalInvDatabase_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    C_TYPE = "M": DESC_TYPE = "MATERIAL"
    FormExistsShow frmPMIS_Physical_CreateINVDATA
End Sub

Private Sub cmdMatFiles_PhysicalMenu_Click()
    If Module_Access(LOGID, "PHYSICAL COUNT", "SYSTEM") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMIS_Physical_INVMenu_New
    C_TYPE = "M": DESC_TYPE = "MATERIAL"
    FormExistsShow frmPMIS_Physical_INVMenu_New
End Sub

Private Sub cmdMatFiles_Series_Click()
    If Module_Access(LOGID, "MATERIALS COUNTER", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Counter_Materials.SetStockType ("M")
    FormExistsShow frmMasterFile_Counter_Materials
End Sub

Private Sub cmdMATOk_Click()
    If REPORT_TYPE = "MAT_TRANLIST" Then
        If optMatReceipts.Value = True Then frmPMISReports_RCRange_MAT.Show
        If optMatIssuances.Value = True Then frmPMISReports_ISRange_MAT.Show
        If optPO_Mat.Value = True Then frmPMISReports_PORange_MAT.Show
    End If
    If REPORT_TYPE = "MAT_MONTHREPORTS" Then
        If optMatReceipts.Value = True Then frmPMISReports_Receipts_MAT.Show
        If optMatIssuances.Value = True Then frmPMISReports_Issuances_MAT.Show
        If optPO_Mat.Value = True Then frmPMISReports_Purchase_For_The_Month_MAT.Show
    End If
    picMATSelect.Visible = False
End Sub

Private Sub cmdMatReports_DailySales_Click()
    If Module_Access(LOGID, "MATERIALS DAILY SALES", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_DailySales_MAT
End Sub

Private Sub cmdMatReports_MonthEndReport_Click()
    If Module_Access(LOGID, "MONTH END REPORT", "REPORTS") = False Then Exit Sub
End Sub

Private Sub cmdMatReports_MonthlyReports_Click()
    If Module_Access(LOGID, "MATERIALS MONTHLY REPORT", "REPORTS") = False Then Exit Sub
    REPORT_TYPE = "MAT_MONTHREPORTS"
    picMATSelect.Visible = True
End Sub

Private Sub cmdMatReports_TransListing_Click()
    If Module_Access(LOGID, "MATERIALS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
    REPORT_TYPE = "MAT_TRANLIST"
    picMATSelect.Visible = True
End Sub

Private Sub cmdOther_ComapnyProfile_Click()
    'frmAllTOOLS.Show
    If Module_Access(LOGID, "COMPANY PROFILE", "DATA ENTRY") = False Then Exit Sub
    frmPMISProfile.Show
End Sub

Private Sub cmdOther_MacTool_Click()
    'If Module_Access(LOGID, "MAC TOOL", "SYSTEM") = False Then Exit Sub
'    frmPmisMacTool.Show
frmMACTool.Show
End Sub

Private Sub cmdOther_Password_Click()
    FormExistsShow frmAccMaintenance
End Sub

Private Sub cmdOther_PRICELISTCONVERSIONTOOL_Click()
    'If Module_Access(LOGID, "PARTS CONVERSION TOOL", "PROCESSING") = False Then Exit Sub
    FormExistsShow frmPMIS_Tools_ExcelAcess
End Sub

Private Sub cmdOther_Reminders_Click()
    FormExistsShow frmSMIS_Log_Reminder
End Sub

Private Sub cmdReport_DailySalesReport_Click()
    If Module_Access(LOGID, "DAILY SALES REPORT", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_DailySales
End Sub

Private Sub cmdReport_Forcasting_Click()
    If Module_Access(LOGID, "FORCASTING REPORT", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_PrintForeCasting
End Sub

Private Sub cmdReport_IssuanceOfTheMonth_Click()
    If Module_Access(LOGID, "PARTS MONTHLY REPORT", "REPORTS") = False Then Exit Sub
    ISSREPTYPE = ""
    FormExistsShow frmPMISReports_Issuances
End Sub

Private Sub cmdReport_PartsRundown_Click()
    If Module_Access(LOGID, "PARTS RUNDOWN REPORT", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_PrintPartsRunDown
End Sub

Private Sub cmdReport_PISReportWorkinProgress_Click()
    If Module_Access(LOGID, "RIV FOR WORKINPROGRESS", "REPORTS") = False Then Exit Sub
    ISSREPTYPE = "RIV_INPROCESS"
    FormExistsShow frmPMISReports_Issuances
End Sub

Private Sub cmdReport_PurchaseForTheMonth_Click()
    If Module_Access(LOGID, "PURCHASE FOR THE MONTH", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_Purchase_For_The_Month
End Sub

Private Sub cmdReport_RankingReport_Click()
    If Module_Access(LOGID, "RANKING REPORT", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_PrintRankfle
End Sub

Private Sub cmdReport_RecieptForTheMonth_Click()
    If Module_Access(LOGID, "RECEIPT FOR THE MONTH", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_Receipts
End Sub

Private Sub cmdReport_StockStatusReport_Click()
    If Module_Access(LOGID, "STOCK STATUS REPORT", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_PrintStockStat
End Sub

Private Sub cmdReport_TranHist_PO_Click()
    If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_Parts_PORange
End Sub

Private Sub cmdReport_TransListingIssuance_Click()
    If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_ISRange
End Sub

Private Sub cmdReport_TransListingReceipt_Click()
    If Module_Access(LOGID, "PARTS TRANSACTION LISTING", "REPORTS") = False Then Exit Sub
    FormExistsShow frmPMISReports_RCRange
End Sub

Private Sub cmdTable_CounterMasterFile_Click()
    If Module_Access(LOGID, "PARTS COUNTER", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Counter_Parts.SetStockType ("P")
    FormExistsShow frmMasterFile_Counter_Parts

End Sub

Private Sub cmdTable_CustomerMasterFile_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAllCustomer
End Sub

Private Sub cmdTable_HariMasterFile_Click()
    If Module_Access(LOGID, "MASTER HARIPARTS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmPMISMaster_DNPPEntry
End Sub

Private Sub cmdTable_PartsMasterFile_Click()
    If Module_Access(LOGID, "PARTS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmMasterFile_Parts.SetStockType ("P")
    FormExistsShow frmMasterFile_Parts
End Sub

Private Sub cmdTable_SalesManMasterFile_Click()
    If Module_Access(LOGID, "SALESMAN MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmPMISMaster_SalesMan
End Sub

Private Sub cmdTable_SupplierMasterFile_Click()
    If Module_Access(LOGID, "VENDORS", "DATA ENTRY") = False Then Exit Sub
    FormExistsShow frmAMISMASTERFILEVendor
End Sub

Private Sub Command1_Click()
  
  
  If Module_Access(LOGID, "PARTS ISSUANCE SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    Unload frmPMISTrans_ADB_Issuances_Mat
    
    COUNTERTYPE = "RIV"
    frmPMISTrans_ADB_Issuances.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_ADB_Issuances
     
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "MATERIALS SERVICE ISSUANCE", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    Unload frmPMISTrans_CustomerOrder_MAT
    Unload frmPMISTrans_ADB_Issuances_Mat
    COUNTERTYPE = "RIV"
    frmPMISTrans_ADB_Issuances_Mat.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_ADB_Issuances_Mat
End Sub

Private Sub Command3_Click()
    Unload frmPMISTrans_CustomerOrder
    Unload frmPMISTrans_ADB_Issuances
    Unload frmPMISTrans_ADB_Issuances_Acc
    COUNTERTYPE = "RIV"
    frmPMISTrans_ADB_Issuances.txtTranType.Text = "RIV"
    FormExistsShow frmPMISTrans_ADB_Issuances_Acc
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SS_MAIN.SelectedItem = 0
    ss_AC.SelectedItem = 0
    ss_Mat.SelectedItem = 0
    SS_PARTS.SelectedItem = 0
    picACSelect.Visible = False
    picACSelect.ZOrder 0
    picMATSelect.Visible = False
    picMATSelect.ZOrder 0
End Sub

  
