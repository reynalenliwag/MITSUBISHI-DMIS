VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmPMIS_Physical_INVMenu_New 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Physical Inventory Menu"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2985
   ClipControls    =   0   'False
   FillColor       =   &H8000000D&
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
   Icon            =   "New_InvMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   2985
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2985
      _Version        =   655364
      _ExtentX        =   5265
      _ExtentY        =   9763
      _StockProps     =   64
      Appearance      =   8
      Color           =   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   4
      Item(0).Caption =   "Data Entry"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "cmdAddTagNumbers"
      Item(0).Control(1)=   "cmdAdEditPhysCount"
      Item(0).Control(2)=   "Label38(0)"
      Item(0).Control(3)=   "Label38(1)"
      Item(1).Caption =   "Inquiry"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "cmdTagMasterList"
      Item(1).Control(1)=   "cmdDisplayTagMasterFilebyPartNumber"
      Item(1).Control(2)=   "cmdDisplayLedgerFile"
      Item(1).Control(3)=   "Label38(2)"
      Item(1).Control(4)=   "Label38(3)"
      Item(1).Control(5)=   "Label38(4)"
      Item(2).Caption =   "Processing"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "cmdCheckCutOffBalances"
      Item(2).Control(1)=   "cmdGenerateLedgerFile"
      Item(2).Control(2)=   "cmdCreateCutOffMasterFile"
      Item(2).Control(3)=   "cmdConsolidatePhysicalCount"
      Item(2).Control(4)=   "Label38(5)"
      Item(2).Control(5)=   "Label38(6)"
      Item(2).Control(6)=   "Label38(7)"
      Item(2).Control(7)=   "Label38(8)"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   10
      Item(3).Control(0)=   "Picture4"
      Item(3).Control(1)=   "Picture5"
      Item(3).Control(2)=   "cmdUnacctTagNo"
      Item(3).Control(3)=   "cmdUnaccPartNo"
      Item(3).Control(4)=   "cmdVarianceReport"
      Item(3).Control(5)=   "cmdInventoryJustification"
      Item(3).Control(6)=   "Label38(9)"
      Item(3).Control(7)=   "Label38(10)"
      Item(3).Control(8)=   "Label38(11)"
      Item(3).Control(9)=   "Label38(12)"
      Begin VB.CommandButton cmdInventoryJustification 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":101C
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Inventory Justification"
         Top             =   3030
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdVarianceReport 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":1393
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":14E5
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "View Variance Report"
         Top             =   2250
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdUnaccPartNo 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":16C2
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":1814
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Unaccounted Part Number"
         Top             =   3810
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdUnacctTagNo 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":1AC9
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":1C1B
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Unaccounted Tag Number"
         Top             =   4590
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdConsolidatePhysicalCount 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":228B
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":23DD
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Consolidate Physical Count"
         Top             =   2580
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCreateCutOffMasterFile 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":2A55
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":2BA7
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Create Cut-Off Master File"
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdGenerateLedgerFile 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":2EE3
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         Picture         =   "New_InvMenu.frx":3035
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Generate Ledger File"
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCheckCutOffBalances 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":32D4
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":3426
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Check Cut-Off Balances"
         Top             =   4170
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDisplayLedgerFile 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":3B0E
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":3C60
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdDisplayTagMasterFilebyPartNumber 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":3E41
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":3F93
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Display Tag Master File by Part Number"
         Top             =   2010
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdTagMasterList 
         Height          =   720
         Left            =   -69850
         MouseIcon       =   "New_InvMenu.frx":41EE
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":4340
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Tag Master List"
         Top             =   2790
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdAdEditPhysCount 
         Height          =   720
         Left            =   90
         MouseIcon       =   "New_InvMenu.frx":46D6
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":4828
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Add/Edit Physical Count Ticket"
         Top             =   1560
         Width           =   735
      End
      Begin VB.CommandButton cmdAddTagNumbers 
         Height          =   720
         Left            =   90
         MouseIcon       =   "New_InvMenu.frx":4DAE
         MousePointer    =   99  'Custom
         Picture         =   "New_InvMenu.frx":4F00
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Data Entry of Tag Numbers"
         Top             =   750
         Width           =   735
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6705
         Left            =   -1.40000e5
         ScaleHeight     =   6705
         ScaleWidth      =   11070
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   11070
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
            Left            =   5175
            MouseIcon       =   "New_InvMenu.frx":5243
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":5395
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "1388"
            ToolTipText     =   "Forecasting Reports"
            Top             =   4020
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
            Left            =   5175
            MouseIcon       =   "New_InvMenu.frx":5A98
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":5BEA
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "1700"
            ToolTipText     =   "Parts Rundown Reports"
            Top             =   4920
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
            Left            =   5175
            MouseIcon       =   "New_InvMenu.frx":63C0
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":6512
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "1385"
            ToolTipText     =   "Stock Status Report"
            Top             =   2205
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
            Left            =   5190
            MouseIcon       =   "New_InvMenu.frx":6C77
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":6DC9
            Style           =   1  'Graphical
            TabIndex        =   21
            Tag             =   "1387"
            ToolTipText     =   "Ranking Reports"
            Top             =   3105
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
            Left            =   420
            MouseIcon       =   "New_InvMenu.frx":74BE
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":7610
            Style           =   1  'Graphical
            TabIndex        =   20
            Tag             =   "1384"
            ToolTipText     =   "Issuances for the Month"
            Top             =   4920
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
            Left            =   420
            MouseIcon       =   "New_InvMenu.frx":7CEC
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":7E3E
            Style           =   1  'Graphical
            TabIndex        =   19
            Tag             =   "1382"
            ToolTipText     =   "Transaction Listing Issuance Report"
            Top             =   3105
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
            Left            =   5190
            MouseIcon       =   "New_InvMenu.frx":853E
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":8690
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "1383"
            ToolTipText     =   "Receipts for the Month"
            Top             =   1245
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
            Left            =   405
            MouseIcon       =   "New_InvMenu.frx":8D98
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":8EEA
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "1381"
            ToolTipText     =   "Transaction Listing Receipts Report"
            Top             =   2205
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
            Left            =   420
            MouseIcon       =   "New_InvMenu.frx":9530
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":9682
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "1380"
            ToolTipText     =   "PIS Report for Work-In-Progress"
            Top             =   1245
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
            Left            =   420
            MouseIcon       =   "New_InvMenu.frx":9D57
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":9EA9
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "1379"
            ToolTipText     =   "Daily Sales Report"
            Top             =   300
            Width           =   720
         End
         Begin VB.CommandButton cmdTranHist_PO 
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
            Left            =   420
            MouseIcon       =   "New_InvMenu.frx":A620
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":A772
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "1382"
            ToolTipText     =   "Transaction Listing Purchase Report"
            Top             =   4020
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
            Left            =   5190
            MouseIcon       =   "New_InvMenu.frx":A918
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":AA6A
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "1383"
            ToolTipText     =   "Purchase for the Month"
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label30 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6030
            TabIndex        =   36
            Top             =   5175
            Width           =   2565
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6030
            TabIndex        =   35
            Top             =   4200
            Width           =   2835
         End
         Begin VB.Label Label52 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6030
            TabIndex        =   34
            Top             =   3315
            Width           =   2625
         End
         Begin VB.Label Label58 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1290
            TabIndex        =   33
            Top             =   510
            Width           =   2325
         End
         Begin VB.Label Label59 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1290
            TabIndex        =   32
            Top             =   1455
            Width           =   3480
         End
         Begin VB.Label Label60 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6030
            TabIndex        =   31
            Top             =   1485
            Width           =   2565
         End
         Begin VB.Label Label61 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1275
            TabIndex        =   30
            Top             =   2415
            Width           =   3600
         End
         Begin VB.Label Label62 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1275
            TabIndex        =   29
            Top             =   3330
            Width           =   3645
         End
         Begin VB.Label Label63 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1290
            TabIndex        =   28
            Top             =   5115
            Width           =   2550
         End
         Begin VB.Label Label65 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6045
            TabIndex        =   27
            Top             =   2400
            Width           =   2115
         End
         Begin VB.Label Label101 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1290
            TabIndex        =   26
            Top             =   4245
            Width           =   3825
         End
         Begin VB.Label Label102 
            BackColor       =   &H00FFFFFF&
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
            Left            =   6030
            TabIndex        =   25
            Top             =   525
            Width           =   2565
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6795
         Left            =   -1.40000e5
         ScaleHeight     =   6795
         ScaleWidth      =   11115
         TabIndex        =   1
         Top             =   585
         Visible         =   0   'False
         Width           =   11115
         Begin VB.CommandButton cmdOther_Reminders 
            Height          =   645
            Left            =   480
            MouseIcon       =   "New_InvMenu.frx":B29A
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":B3EC
            Style           =   1  'Graphical
            TabIndex        =   6
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   660
            Width           =   720
         End
         Begin VB.CommandButton cmdMainTainComapnyProfile 
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
            Left            =   480
            MouseIcon       =   "New_InvMenu.frx":BC67
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":BDB9
            Style           =   1  'Graphical
            TabIndex        =   5
            Tag             =   "1405"
            ToolTipText     =   "Company Profile"
            Top             =   1580
            Width           =   720
         End
         Begin VB.CommandButton cmdPassword 
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
            Left            =   480
            MouseIcon       =   "New_InvMenu.frx":C7B0
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":C902
            Style           =   1  'Graphical
            TabIndex        =   4
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   2500
            Width           =   720
         End
         Begin VB.CommandButton cmdPRICELISTCONVERSIONTOOL 
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
            Left            =   480
            MouseIcon       =   "New_InvMenu.frx":D226
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":D378
            Style           =   1  'Graphical
            TabIndex        =   3
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   3420
            Width           =   720
         End
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
            Left            =   480
            MouseIcon       =   "New_InvMenu.frx":D5EB
            MousePointer    =   99  'Custom
            Picture         =   "New_InvMenu.frx":D73D
            Style           =   1  'Graphical
            TabIndex        =   2
            Tag             =   "1407"
            ToolTipText     =   "Password Maintenance"
            Top             =   4260
            Width           =   720
         End
         Begin VB.Label Label98 
            BackStyle       =   0  'Transparent
            Caption         =   "REMINDERS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1350
            TabIndex        =   11
            Top             =   870
            Width           =   2490
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "PASSWORD MAINTENANCE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1350
            TabIndex        =   10
            Top             =   2745
            Width           =   6015
         End
         Begin VB.Label label 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY PROFILE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1335
            TabIndex        =   9
            Top             =   1770
            Width           =   3195
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PRICE LIST CONVERSION TOOL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1350
            TabIndex        =   8
            Top             =   3645
            Width           =   6015
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "MAC TOOL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1350
            TabIndex        =   7
            Top             =   4485
            Width           =   6015
         End
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Unaccounted Tag No."
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
         Index           =   12
         Left            =   -69010
         TabIndex        =   62
         ToolTipText     =   "Unacct. Tag No."
         Top             =   4770
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Unaccounted Part No."
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
         Index           =   11
         Left            =   -69010
         TabIndex        =   61
         ToolTipText     =   "&Unacct. Part No."
         Top             =   3930
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Inventory Justification"
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
         Index           =   10
         Left            =   -69010
         TabIndex        =   60
         ToolTipText     =   "&Inventory Justification"
         Top             =   3210
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Variance Report"
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
         Index           =   9
         Left            =   -69010
         TabIndex        =   59
         ToolTipText     =   "Variance Report"
         Top             =   2370
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Cut-Off &Balances"
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
         Height          =   495
         Index           =   8
         Left            =   -69040
         TabIndex        =   58
         ToolTipText     =   "Check Cut-Off &Balances"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Generate Ledger File"
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
         Height          =   495
         Index           =   7
         Left            =   -69040
         TabIndex        =   57
         ToolTipText     =   "&Generate Ledger File"
         Top             =   3570
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidate &Physical Count"
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
         Height          =   495
         Index           =   6
         Left            =   -69010
         TabIndex        =   56
         ToolTipText     =   "Consolidate &Physical Count"
         Top             =   2700
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Create Cut-Off Master File"
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
         Height          =   495
         Index           =   5
         Left            =   -69010
         TabIndex        =   55
         ToolTipText     =   "&Create Cut-Off Master File"
         Top             =   1920
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Tag &Master List"
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
         Height          =   495
         Index           =   4
         Left            =   -69040
         TabIndex        =   54
         ToolTipText     =   "Tag &Master List"
         Top             =   3030
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Display &Tag Master File by Part Number"
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
         Height          =   495
         Index           =   3
         Left            =   -69040
         TabIndex        =   53
         ToolTipText     =   "Display Ledger File"
         Top             =   2130
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Display Ledger File"
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
         Index           =   2
         Left            =   -69040
         TabIndex        =   52
         ToolTipText     =   "Display Ledger File"
         Top             =   1410
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Add/Edit Physical Count Ticket"
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
         Height          =   495
         Index           =   1
         Left            =   900
         TabIndex        =   51
         Top             =   1650
         Width           =   1785
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "&Data Entry of Tag Numbers"
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
         Height          =   495
         Index           =   0
         Left            =   900
         TabIndex        =   50
         Top             =   900
         Width           =   1785
      End
   End
   Begin Crystal.CrystalReport rptInventory 
      Left            =   3450
      Top             =   6990
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin MSComDlg.CommonDialog cmdDialogINV 
      Left            =   3375
      Top             =   5925
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPMIS_Physical_INVMenu_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose : To create USER DSN through VB code
'By      : Manish Kumar Pandey
'Put Following declaration in the form you want to create the dsn from.....
Option Explicit
Private Const REG_DWORD = 4&
Private Const REG_SZ = 1                              'Constant for a string variable type.
Private Const HKEY_CURRENT_USER = &H80000001

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias _
                                      "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
                                                       phkResult As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
                                       "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
                                                         ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
                                                                                                                      cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" _
                                     (ByVal hKey As Long) As Long
Dim FILNAME                                            As String
Dim INVENTORY_REPORT_CONNECTION                        As String
 
Sub SETODBC()
    Dim DataSourceName                                 As String
    Dim Description                                    As String
    Dim DriverPath                                     As String
    Dim DriverId                                       As Long
    Dim DriverName                                     As String
    Dim User                                           As String
    Dim PWD                                            As String

    Dim lResult                                        As Long
    Dim hKeyHandle                                     As Long
    Dim hKeyHandSub                                    As Long
    Dim DBQ                                            As String

    'Specify the DSN parameters.

    DataSourceName = "INVENTORY"
    DBQ = FILNAME
    Description = "PHYSICAL COUNT INVENTORY DATABASE"
    DriverPath = "E:\windows\System32\odbcjt32.dll"
    PWD = ""
    DriverId = 19
    User = "admin"
    DriverName = "Microsoft Access Driver (*.mdb)"

    'Create the new DSN key.

    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName, hKeyHandle)

    'Set the values of the new DSN key.

    lResult = RegSetValueEx(hKeyHandle, "DBQ", 0&, REG_SZ, _
                            ByVal DBQ, Len(DBQ))
    lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
                            ByVal Description, Len(Description))
    lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
                            ByVal DriverPath, Len(DriverPath))
    lResult = RegSetValueEx(hKeyHandle, "DriverID", 0&, REG_DWORD, _
                            25, 4)
    lResult = RegSetValueEx(hKeyHandle, "FIL", 0&, REG_SZ, _
                            ByVal "MS Access", 9)
    lResult = RegSetValueEx(hKeyHandle, "PWD", 0&, REG_SZ, _
                            ByVal PWD, Len(PWD))
    lResult = RegSetValueEx(hKeyHandle, "SafeTransactions", 0&, REG_DWORD, _
                            0, 4)
    lResult = RegSetValueEx(hKeyHandle, "UID", 0&, REG_SZ, _
                            ByVal User, Len(User))
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Open a new key as follows
    lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & _
                                              DataSourceName & "\Engines\Jet", hKeyHandSub)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    lResult = RegSetValueEx(hKeyHandSub, "ImplicitCommitSync", 0&, REG_SZ, _
                            ByVal "", 0)
    lResult = RegSetValueEx(hKeyHandSub, "MaxBufferSize", 0&, REG_DWORD, _
                            2048, 4)
    lResult = RegSetValueEx(hKeyHandSub, "PageTimeout", 0&, REG_DWORD, _
                            5, 4)
    lResult = RegSetValueEx(hKeyHandSub, "Threads", 0&, REG_DWORD, _
                            3, 4)
    lResult = RegSetValueEx(hKeyHandSub, "UserCommitSync", 0&, REG_SZ, _
                            ByVal "Yes", 3)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Close the new Sub key.
    lResult = RegCloseKey(hKeyHandSub)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Close the new DSN key.

    lResult = RegCloseKey(hKeyHandle)

 
    lResult = RegCreateKey(HKEY_CURRENT_USER, _
                           "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
    lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
                            ByVal DriverName, Len(DriverName))
    lResult = RegCloseKey(hKeyHandle)

    INVENTORY_REPORT_CONNECTION = "DSN=INVENTORY;UID=admin;PWD=;DSQ=" & FILNAME
End Sub

Private Sub cmdAddTagNumbers_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_DataEntryTagByPartNo.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdEditPhysCount_Click()
    frmPMIS_Physical_AddPhyCntTicket.Show
End Sub

Private Sub cmdCheckCutOffBalances_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_CutOffCheckPrevBal.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdConsolidatePhysicalCount_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_CreateConsPhyCNT.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdCreateCutOffMasterFile_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_CreateCutOffMaster.Show
    Screen.MousePointer = 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDataEntry_Click()
     
     
     
     
End Sub

Private Sub cmdDisplayTagMasterFilebyPartNumber_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_DataEntryTagByPartNo.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerateLedgerFile_Click()
    Screen.MousePointer = 11
    frmPMISProcess_GenLedgerFile.Show
    Screen.MousePointer = 0
End Sub
 
Private Sub cmdInventoryJustification_Click()
    Screen.MousePointer = 11
    rptInventory.WindowTitle = "Inventory Justification Report"
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "InvJustification.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    NEW_LogAudit "V", "PHYSICAL COUNT", "", "", "", "", "Inventory Justification", ""
    Screen.MousePointer = 0
End Sub
 
     

Private Sub cmdUnaccPartNo_Click()
    Screen.MousePointer = 11
    rptInventory.WindowTitle = "Unaccounted Part Number"
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnAcctPartNo.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    NEW_LogAudit "V", "PHYSICAL COUNT", "", "", "", "", "Unaccounted Part Number", ""
    Screen.MousePointer = 0
End Sub

Private Sub cmdUnacctTagNo_Click()
    Screen.MousePointer = 11
    rptInventory.WindowTitle = "Unaccounted Tag Number"
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "UnacctTagNo.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    NEW_LogAudit "V", "PHYSICAL COUNT", "", "", "", "", "Unaccounted Tag Number", ""
    Screen.MousePointer = 0
End Sub

Private Sub cmdVarianceReport_Click()
    Screen.MousePointer = 11
    rptInventory.WindowTitle = "Variance Report"
    rptInventory.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptInventory.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptInventory, PMIS_REPORT_PATH & "Variance.rpt", "", INVENTORY_REPORT_CONNECTION, 1
    NEW_LogAudit "V", "PHYSICAL COUNT", "", "", "", "", "Variance Report", ""
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    CenterMe frmMain, Me, 1
    TabControl1.SelectedItem = 0
    On Error Resume Next
    Dim MYPATH, PAYLNAME                               As String
    MYPATH = App.Path
    cmdDialogINV.Filter = "Access Files (*.MDB)|*.MDB"
    cmdDialogINV.FilterIndex = 1
    cmdDialogINV.DefaultExt = "MDB"
    cmdDialogINV.DialogTitle = "Open Inventory Database"
    PAYLNAME = cmdDialogINV.FileName
    If MYPATH <> "\" Then
        cmdDialogINV.FileName = MYPATH & "\" & cmdDialogINV.FileName
    End If
    If PAYLNAME = "" Then
        cmdDialogINV.FileName = "*.MDB"
    End If
    cmdDialogINV.Action = 1
    If err = 32755 Then Exit Sub
    FILNAME = cmdDialogINV.FileName
    Dim CS                                             As String
    CS = wizVar.DecryptAccess("50726F@§É¥¥èï_Å∂†d}oNvmbÜÑlëmîßâN∂•èµÆí•®m±Ö¶®ñ±±p")
    On Error GoTo Errorcode
    Set gconINVENTORY = New ADODB.Connection
    gconINVENTORY.ConnectionString = CS & FILNAME
    gconINVENTORY.Open
    If err = 32755 Then Exit Sub
    SETODBC
    Exit Sub

Errorcode:
    ShowADOErrors gconINVENTORY
    On Error Resume Next
    MsgSpeechBox "Warning: Inventory Database is Invalid or Corrupted!" & vbCrLf & _
                 "Inventory Menu will be Unloaded... Contact EDP Immediately."
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    gconINVENTORY.Close
    UnloadForm Me
End Sub

Private Sub mDEPhyTicket_Click()
    frmPMIS_Physical_AddPhyCntTicket.Show
End Sub

Private Sub mDEPhyTicket_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Flat_DE
    'mDEPhyTicket.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub mDETagNumbers_Click()
    Screen.MousePointer = 11
    frmPMIS_Physical_DataEntryTagByPartNo.Show
    Screen.MousePointer = 0
End Sub






