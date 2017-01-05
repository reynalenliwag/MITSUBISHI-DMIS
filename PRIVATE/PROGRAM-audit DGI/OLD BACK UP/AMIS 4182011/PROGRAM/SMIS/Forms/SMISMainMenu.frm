VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMIS Main Menu"
   ClientHeight    =   6255
   ClientLeft      =   210
   ClientTop       =   1890
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Icon            =   "SMISMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   10770
   Begin XtremeSuiteControls.TabControl TabControl3 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   10755
      _Version        =   655364
      _ExtentX        =   18971
      _ExtentY        =   11033
      _StockProps     =   64
      DrawFocusRect   =   0   'False
      AutoResizeClient=   0   'False
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   100
      ItemCount       =   5
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Item(2).Caption =   "Inquiry"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage3"
      Item(3).Caption =   "Reports"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "TabControlPage4"
      Item(3).Control(1)=   "TabControlPage5"
      Item(4).Caption =   "Other Setups"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "TabControlPage6"
      Begin XtremeSuiteControls.TabControlPage TabControlPage6 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   6
         Top             =   585
         Visible         =   0   'False
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
         Begin VB.CommandButton cmdSignatories 
            Height          =   645
            Left            =   270
            MouseIcon       =   "SMISMainMenu.frx":0B9E
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":0CF0
            Style           =   1  'Graphical
            TabIndex        =   203
            Tag             =   "1208"
            ToolTipText     =   "Lead Classifications"
            Top             =   2460
            Width           =   720
         End
         Begin VB.CommandButton Command24 
            Height          =   645
            Left            =   6360
            MouseIcon       =   "SMISMainMenu.frx":155A
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":16AC
            Style           =   1  'Graphical
            TabIndex        =   192
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   720
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_Reminders 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":1F27
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2079
            Style           =   1  'Graphical
            TabIndex        =   178
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   1710
            Width           =   720
         End
         Begin VB.CommandButton Command3 
            Height          =   645
            Left            =   6360
            MouseIcon       =   "SMISMainMenu.frx":28F4
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2A46
            Style           =   1  'Graphical
            TabIndex        =   119
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   1500
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdCompanyProfile 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":32C1
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3413
            Style           =   1  'Graphical
            TabIndex        =   74
            Tag             =   "1163"
            ToolTipText     =   "Company Profile"
            Top             =   210
            Width           =   720
         End
         Begin VB.CommandButton cmdPasswordMaintain 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":3CEB
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3E3D
            Style           =   1  'Graphical
            TabIndex        =   73
            Tag             =   "1164"
            ToolTipText     =   "Password Maintenance"
            Top             =   975
            Width           =   720
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Signatories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1065
            TabIndex        =   204
            Top             =   2670
            Width           =   975
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Update Master File"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   7170
            TabIndex        =   193
            Top             =   795
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reminders"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1065
            TabIndex        =   179
            Top             =   1905
            Width           =   930
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Reminders"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   7170
            TabIndex        =   120
            Top             =   1575
            Visible         =   0   'False
            Width           =   2985
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Profile"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1065
            TabIndex        =   76
            Top             =   405
            Width           =   1395
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password Maintenance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1065
            TabIndex        =   75
            Top             =   1185
            Width           =   2010
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage5 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   5
         Top             =   585
         Visible         =   0   'False
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
         Begin XtremeSuiteControls.TabControl TabControl1 
            Height          =   5625
            Left            =   30
            TabIndex        =   49
            Top             =   -60
            Width           =   10605
            _Version        =   655364
            _ExtentX        =   18706
            _ExtentY        =   9922
            _StockProps     =   64
            AutoResizeClient=   0   'False
            Appearance      =   2
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            PaintManager.FixedTabWidth=   100
            PaintManager.MinTabWidth=   100
            ItemCount       =   4
            SelectedItem    =   1
            Item(0).Caption =   "Inventory"
            Item(0).ControlCount=   20
            Item(0).Control(0)=   "Label13"
            Item(0).Control(1)=   "Label14"
            Item(0).Control(2)=   "Label20"
            Item(0).Control(3)=   "Label23"
            Item(0).Control(4)=   "Label17"
            Item(0).Control(5)=   "cmdMonthlyReport"
            Item(0).Control(6)=   "cmdMonthlyInventoryControl"
            Item(0).Control(7)=   "cmdEndingInvenotory"
            Item(0).Control(8)=   "cmdVehicleOnStock"
            Item(0).Control(9)=   "cmdVehicle_InvReport"
            Item(0).Control(10)=   "cmdDeliveryReport"
            Item(0).Control(11)=   "Label82"
            Item(0).Control(12)=   "cmdStockandSalesTracking"
            Item(0).Control(13)=   "Label83"
            Item(0).Control(14)=   "cmdUnitReleasedReport"
            Item(0).Control(15)=   "Label16"
            Item(0).Control(16)=   "cmdListofUnitsReg"
            Item(0).Control(17)=   "Label37"
            Item(0).Control(18)=   "cmdBirYear"
            Item(0).Control(19)=   "Label31"
            Item(1).Caption =   "Sales"
            Item(1).ControlCount=   24
            Item(1).Control(0)=   "Label15"
            Item(1).Control(1)=   "Label24"
            Item(1).Control(2)=   "Label27"
            Item(1).Control(3)=   "Label28"
            Item(1).Control(4)=   "cmdReport_Sales_VehicleSales"
            Item(1).Control(5)=   "cmdMonthlyVehicleGrossProfit"
            Item(1).Control(6)=   "cmdYearlyGrossProfit"
            Item(1).Control(7)=   "cmdDistributionOfSales"
            Item(1).Control(8)=   "cmdPrint"
            Item(1).Control(9)=   "cmd_UnitCommission"
            Item(1).Control(10)=   "Label88"
            Item(1).Control(11)=   "Label89"
            Item(1).Control(12)=   "Command11"
            Item(1).Control(13)=   "Label90"
            Item(1).Control(14)=   "Command18"
            Item(1).Control(15)=   "Label93"
            Item(1).Control(16)=   "Command23"
            Item(1).Control(17)=   "Label94"
            Item(1).Control(18)=   "Command27"
            Item(1).Control(19)=   "Label99"
            Item(1).Control(20)=   "Command28"
            Item(1).Control(21)=   "Label100"
            Item(1).Control(22)=   "Label101"
            Item(1).Control(23)=   "cmdAfterSalesReport"
            Item(2).Caption =   "Monitoring"
            Item(2).ControlCount=   20
            Item(2).Control(0)=   "Command45"
            Item(2).Control(1)=   "Command12"
            Item(2).Control(2)=   "Command13"
            Item(2).Control(3)=   "Label55"
            Item(2).Control(4)=   "Label67"
            Item(2).Control(5)=   "Label68"
            Item(2).Control(6)=   "Command8"
            Item(2).Control(7)=   "Command14"
            Item(2).Control(8)=   "Command7"
            Item(2).Control(9)=   "cmdOther_SAEPerformance_1"
            Item(2).Control(10)=   "Label57"
            Item(2).Control(11)=   "Label69"
            Item(2).Control(12)=   "Label56"
            Item(2).Control(13)=   "Label36"
            Item(2).Control(14)=   "Command10"
            Item(2).Control(15)=   "Label87"
            Item(2).Control(16)=   "Command25"
            Item(2).Control(17)=   "Label95"
            Item(2).Control(18)=   "cmdHitRatio"
            Item(2).Control(19)=   "Label97"
            Item(3).Caption =   "Customer"
            Item(3).ControlCount=   28
            Item(3).Control(0)=   "cmdSalesAppointment"
            Item(3).Control(1)=   "cmdCustomer_Reminder_And_Task_Internal"
            Item(3).Control(2)=   "cmdServiceAppointment"
            Item(3).Control(3)=   "Label12"
            Item(3).Control(4)=   "Label30"
            Item(3).Control(5)=   "Label41"
            Item(3).Control(6)=   "cmdCustomer_Reminder_And_Task"
            Item(3).Control(7)=   "Label29"
            Item(3).Control(8)=   "cmdVehiclesalesCustomer"
            Item(3).Control(9)=   "cmdCustomerInfoReport"
            Item(3).Control(10)=   "Command4"
            Item(3).Control(11)=   "Command6"
            Item(3).Control(12)=   "cmdMarketting_BirthDay"
            Item(3).Control(13)=   "Command5"
            Item(3).Control(14)=   "cmdMarketting_CustomerDirectory"
            Item(3).Control(15)=   "Label32"
            Item(3).Control(16)=   "Label45"
            Item(3).Control(17)=   "Label52"
            Item(3).Control(18)=   "Label54"
            Item(3).Control(19)=   "Label33"
            Item(3).Control(20)=   "Label53"
            Item(3).Control(21)=   "Label34"
            Item(3).Control(22)=   "cmdCustomerwithInsurance"
            Item(3).Control(23)=   "Label35"
            Item(3).Control(24)=   "cmdCustomerLog"
            Item(3).Control(25)=   "Label6"
            Item(3).Control(26)=   "Command15"
            Item(3).Control(27)=   "Label40"
            Begin VB.CommandButton cmdAfterSalesReport 
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":4761
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":48B3
               Style           =   1  'Graphical
               TabIndex        =   211
               Tag             =   "1088"
               ToolTipText     =   "View Sales Appointment"
               Top             =   4860
               Width           =   720
            End
            Begin VB.CommandButton Command28 
               Height          =   645
               Left            =   240
               MouseIcon       =   "SMISMainMenu.frx":4FBC
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":510E
               Style           =   1  'Graphical
               TabIndex        =   209
               ToolTipText     =   "Ranged Customer Directory By Customer Type"
               Top             =   4860
               Width           =   720
            End
            Begin VB.CommandButton Command27 
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":5A08
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":5B5A
               Style           =   1  'Graphical
               TabIndex        =   207
               ToolTipText     =   "Ranged Customer Directory By Customer Type"
               Top             =   4020
               Width           =   720
            End
            Begin VB.CommandButton cmdHitRatio 
               Height          =   645
               Left            =   -64060
               MouseIcon       =   "SMISMainMenu.frx":6454
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":65A6
               Style           =   1  'Graphical
               TabIndex        =   201
               Tag             =   "1152"
               ToolTipText     =   "Lost Sales Monitoring"
               Top             =   4170
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command25 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":6DAC
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":6EFE
               Style           =   1  'Graphical
               TabIndex        =   194
               Tag             =   "1158"
               ToolTipText     =   "Customer With Insurance Policies"
               Top             =   4140
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command23 
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":78DB
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":7A2D
               Style           =   1  'Graphical
               TabIndex        =   188
               ToolTipText     =   "Ranged Customer Directory By Customer Type"
               Top             =   3195
               Width           =   720
            End
            Begin VB.CommandButton Command18 
               Height          =   645
               Left            =   240
               MouseIcon       =   "SMISMainMenu.frx":8327
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":8479
               Style           =   1  'Graphical
               TabIndex        =   186
               Tag             =   "1152"
               ToolTipText     =   "Vehicle Sales Report"
               Top             =   3990
               Width           =   720
            End
            Begin VB.CommandButton Command15 
               Height          =   645
               Left            =   -62500
               MouseIcon       =   "SMISMainMenu.frx":8DEF
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":8F41
               Style           =   1  'Graphical
               TabIndex        =   182
               Tag             =   "1088"
               ToolTipText     =   "Customer Log"
               Top             =   1650
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command11 
               Height          =   645
               Left            =   240
               Picture         =   "SMISMainMenu.frx":94CD
               Style           =   1  'Graphical
               TabIndex        =   175
               Top             =   3165
               Width           =   720
            End
            Begin VB.CommandButton Command10 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":98F1
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":9A43
               Style           =   1  'Graphical
               TabIndex        =   170
               ToolTipText     =   "Test Drive Evaluation Report"
               Top             =   3330
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdBirYear 
               Height          =   645
               Left            =   -64930
               MouseIcon       =   "SMISMainMenu.frx":A37C
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":A4CE
               Style           =   1  'Graphical
               TabIndex        =   168
               Tag             =   "1162"
               ToolTipText     =   "BIR YEAR - Report"
               Top             =   3150
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdOther_SAEPerformance_1 
               Height          =   645
               Left            =   -64060
               MouseIcon       =   "SMISMainMenu.frx":AEF8
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":B04A
               Style           =   1  'Graphical
               TabIndex        =   163
               Tag             =   "1156"
               ToolTipText     =   "Sales Executive Performance Report"
               Top             =   885
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command7 
               Height          =   645
               Left            =   -64060
               MouseIcon       =   "SMISMainMenu.frx":B865
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":B9B7
               Style           =   1  'Graphical
               TabIndex        =   162
               ToolTipText     =   "Sales Executive Performance Report"
               Top             =   1740
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command14 
               Height          =   645
               Left            =   -64060
               MouseIcon       =   "SMISMainMenu.frx":C1D2
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":C324
               Style           =   1  'Graphical
               TabIndex        =   161
               ToolTipText     =   "Sales Team Performance Report"
               Top             =   2625
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command8 
               Height          =   645
               Left            =   -64060
               MouseIcon       =   "SMISMainMenu.frx":CB3F
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":CC91
               Style           =   1  'Graphical
               TabIndex        =   160
               ToolTipText     =   "Test Drive Evaluation Report"
               Top             =   3420
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustomerLog 
               Height          =   645
               Left            =   -62530
               MouseIcon       =   "SMISMainMenu.frx":D5CA
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":D71C
               Style           =   1  'Graphical
               TabIndex        =   158
               Tag             =   "1088"
               ToolTipText     =   "Customer Log"
               Top             =   870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustomerwithInsurance 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":DCA8
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":DDFA
               Style           =   1  'Graphical
               TabIndex        =   156
               Tag             =   "1158"
               ToolTipText     =   "Customer With Insurance Policies"
               Top             =   4590
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdMarketting_CustomerDirectory 
               Height          =   645
               Left            =   -69910
               MouseIcon       =   "SMISMainMenu.frx":E7D7
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":E929
               Style           =   1  'Graphical
               TabIndex        =   148
               Tag             =   "1159"
               ToolTipText     =   "Customer Directory"
               Top             =   3120
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command5 
               Height          =   645
               Left            =   -69880
               MouseIcon       =   "SMISMainMenu.frx":F1EC
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":F33E
               Style           =   1  'Graphical
               TabIndex        =   147
               ToolTipText     =   "Yearly Customer Directory by Customer Type"
               Top             =   2400
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdMarketting_BirthDay 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":FAD2
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":FC24
               Style           =   1  'Graphical
               TabIndex        =   146
               Tag             =   "1160"
               ToolTipText     =   "Birthday Celebrant of the Month"
               Top             =   3870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command6 
               Height          =   645
               Left            =   -69880
               MouseIcon       =   "SMISMainMenu.frx":107DC
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1092E
               Style           =   1  'Graphical
               TabIndex        =   145
               ToolTipText     =   "Ranged Customer Directory By Customer Type"
               Top             =   870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command4 
               Height          =   645
               Left            =   -69880
               MouseIcon       =   "SMISMainMenu.frx":11228
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1137A
               Style           =   1  'Graphical
               TabIndex        =   144
               ToolTipText     =   "Monthly Customer Directory By Customer Type"
               Top             =   1650
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustomerInfoReport 
               Height          =   645
               Left            =   -69910
               MouseIcon       =   "SMISMainMenu.frx":11AE9
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":11C3B
               Style           =   1  'Graphical
               TabIndex        =   143
               Tag             =   "1088"
               ToolTipText     =   "Customer Information Report"
               Top             =   3870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdVehiclesalesCustomer 
               Height          =   645
               Left            =   -69910
               MouseIcon       =   "SMISMainMenu.frx":12505
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":12657
               Style           =   1  'Graphical
               TabIndex        =   142
               Tag             =   "1157"
               ToolTipText     =   "Vehicle Sales Customer Summary"
               Top             =   4590
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustomer_Reminder_And_Task 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":13068
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":131BA
               Style           =   1  'Graphical
               TabIndex        =   140
               Tag             =   "1088"
               ToolTipText     =   "Customer Reminders And Tasks"
               Top             =   2400
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdListofUnitsReg 
               Height          =   645
               Left            =   -64930
               MouseIcon       =   "SMISMainMenu.frx":13873
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":139C5
               Style           =   1  'Graphical
               TabIndex        =   138
               Tag             =   "1155"
               ToolTipText     =   "List of Units Registered"
               Top             =   2370
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdUnitReleasedReport 
               Height          =   645
               Left            =   -64900
               MouseIcon       =   "SMISMainMenu.frx":143EC
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1453E
               Style           =   1  'Graphical
               TabIndex        =   136
               Tag             =   "1149"
               ToolTipText     =   "Units Released Report"
               Top             =   1560
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdPrint 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":14FAF
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":15101
               Style           =   1  'Graphical
               TabIndex        =   135
               ToolTipText     =   "Print Report"
               Top             =   720
               Width           =   720
            End
            Begin VB.CommandButton cmdReport_Sales_VehicleSales 
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":155A0
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":156F2
               Style           =   1  'Graphical
               TabIndex        =   134
               Tag             =   "1152"
               ToolTipText     =   "Vehicle Sales Report"
               Top             =   2370
               Width           =   720
            End
            Begin VB.CommandButton cmd_UnitCommission 
               Height          =   645
               Left            =   6450
               MouseIcon       =   "SMISMainMenu.frx":16068
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":161BA
               Style           =   1  'Graphical
               TabIndex        =   131
               Tag             =   "1155"
               ToolTipText     =   "List of Units Registered"
               Top             =   1545
               Width           =   720
            End
            Begin VB.CommandButton cmdDistributionOfSales 
               Height          =   645
               Left            =   240
               MouseIcon       =   "SMISMainMenu.frx":16BE1
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":16D33
               Style           =   1  'Graphical
               TabIndex        =   130
               Tag             =   "1153"
               ToolTipText     =   "Distribution of Sales as to Mode of Payment"
               Top             =   2355
               Width           =   720
            End
            Begin VB.CommandButton cmdStockandSalesTracking 
               Height          =   645
               Left            =   -64930
               MouseIcon       =   "SMISMainMenu.frx":17572
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":176C4
               Style           =   1  'Graphical
               TabIndex        =   124
               Tag             =   "1145"
               ToolTipText     =   "Ending Inventory Report"
               Top             =   780
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdServiceAppointment 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":1809E
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":181F0
               Style           =   1  'Graphical
               TabIndex        =   109
               Tag             =   "1088"
               ToolTipText     =   "Service Appointment"
               Top             =   1650
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustomer_Reminder_And_Task_Internal 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":188F9
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":18A4B
               Style           =   1  'Graphical
               TabIndex        =   108
               Tag             =   "1088"
               ToolTipText     =   "Internal Reminders"
               Top             =   3120
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdSalesAppointment 
               Height          =   645
               Left            =   -65920
               MouseIcon       =   "SMISMainMenu.frx":19104
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":19256
               Style           =   1  'Graphical
               TabIndex        =   107
               Tag             =   "1088"
               ToolTipText     =   "Sales Appointment"
               Top             =   870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command13 
               Height          =   645
               Left            =   -69760
               MaskColor       =   &H00FFFFFF&
               MouseIcon       =   "SMISMainMenu.frx":1995F
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":19AB1
               Style           =   1  'Graphical
               TabIndex        =   69
               Tag             =   "1152"
               ToolTipText     =   "Sales Trend"
               Top             =   2520
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command12 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":1A01F
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1A171
               Style           =   1  'Graphical
               TabIndex        =   68
               Tag             =   "1152"
               ToolTipText     =   "Progress Monitoring"
               Top             =   1710
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command45 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":1A4E4
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1A636
               Style           =   1  'Graphical
               TabIndex        =   67
               Tag             =   "1152"
               ToolTipText     =   "Lost Sales Monitoring"
               Top             =   870
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdYearlyGrossProfit 
               Height          =   645
               Left            =   240
               MouseIcon       =   "SMISMainMenu.frx":1AE3C
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1AF8E
               Style           =   1  'Graphical
               TabIndex        =   62
               Tag             =   "1151"
               ToolTipText     =   "Yearly Vehicle Gross Profile Report"
               Top             =   1530
               Width           =   720
            End
            Begin VB.CommandButton cmdMonthlyVehicleGrossProfit 
               Height          =   645
               Left            =   240
               MouseIcon       =   "SMISMainMenu.frx":1B9D0
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1BB22
               Style           =   1  'Graphical
               TabIndex        =   61
               Tag             =   "1150"
               ToolTipText     =   "Monthly Vehicle Gross Profile Report"
               Top             =   720
               Width           =   720
            End
            Begin VB.CommandButton cmdDeliveryReport 
               Height          =   645
               Left            =   -69640
               MouseIcon       =   "SMISMainMenu.frx":1C5D0
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1C722
               Style           =   1  'Graphical
               TabIndex        =   59
               Tag             =   "1148"
               ToolTipText     =   "Delivery Unit Report"
               Top             =   3930
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdEndingInvenotory 
               Height          =   645
               Left            =   -69640
               MouseIcon       =   "SMISMainMenu.frx":1D0F0
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1D242
               Style           =   1  'Graphical
               TabIndex        =   56
               Tag             =   "1145"
               ToolTipText     =   "Ending Inventory Report"
               Top             =   2415
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdVehicleOnStock 
               Height          =   645
               Left            =   -64900
               MouseIcon       =   "SMISMainMenu.frx":1DC1C
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1DD6E
               Style           =   1  'Graphical
               TabIndex        =   55
               Tag             =   "1146"
               ToolTipText     =   "Vehicle On Stock"
               Top             =   3900
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdVehicle_InvReport 
               Height          =   645
               Left            =   -69640
               MouseIcon       =   "SMISMainMenu.frx":1E71F
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1E871
               Style           =   1  'Graphical
               TabIndex        =   54
               Tag             =   "1147"
               ToolTipText     =   "Vehicle Inventory Report"
               Top             =   3180
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdMonthlyInventoryControl 
               Height          =   645
               Left            =   -69640
               MouseIcon       =   "SMISMainMenu.frx":1F1FD
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1F34F
               Style           =   1  'Graphical
               TabIndex        =   52
               Tag             =   "1144"
               ToolTipText     =   "Monthly Inventory Control"
               Top             =   1635
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdMonthlyReport 
               Height          =   645
               Left            =   -69640
               MouseIcon       =   "SMISMainMenu.frx":1FD30
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":1FE82
               Style           =   1  'Graphical
               TabIndex        =   50
               Tag             =   "1143"
               ToolTipText     =   "Monthly Purchase Report"
               Top             =   855
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.Label Label101 
               BackStyle       =   0  'Transparent
               Caption         =   "After Sales Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   7260
               TabIndex        =   212
               Top             =   5085
               Width           =   2295
            End
            Begin VB.Label Label100 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Sales Projection"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   210
               Top             =   5055
               Width           =   2070
            End
            Begin VB.Label Label99 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   7290
               TabIndex        =   208
               Top             =   4275
               Width           =   1095
            End
            Begin VB.Label Label97 
               AutoSize        =   -1  'True
               Caption         =   "Hit Ratio"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -63220
               TabIndex        =   202
               Top             =   4380
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Label Label95 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Transaction Status Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68965
               TabIndex        =   195
               Top             =   4305
               Visible         =   0   'False
               Width           =   2235
            End
            Begin VB.Label Label94 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net Sales Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   7290
               TabIndex        =   189
               Top             =   3465
               Width           =   1425
            End
            Begin VB.Label Label93 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Profit Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   187
               Top             =   4230
               Width           =   1635
            End
            Begin VB.Label Label40 
               BackStyle       =   0  'Transparent
               Caption         =   "Log Summary Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   450
               Left            =   -61720
               TabIndex        =   183
               Top             =   1815
               Visible         =   0   'False
               Width           =   2235
            End
            Begin VB.Label Label90 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hyundai Sales Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   176
               Top             =   3330
               Width           =   1815
            End
            Begin VB.Label Label87 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LTO Status Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68965
               TabIndex        =   171
               Top             =   3615
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BIR YEAR - Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -64090
               TabIndex        =   169
               Top             =   3345
               Visible         =   0   'False
               Width           =   1500
            End
            Begin VB.Label Label36 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive Performance Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   600
               Left            =   -63220
               TabIndex        =   167
               Top             =   1005
               Visible         =   0   'False
               Width           =   3885
            End
            Begin VB.Label Label56 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive Performance Sales Productivity Per SC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   600
               Left            =   -63220
               TabIndex        =   166
               Top             =   1890
               Visible         =   0   'False
               Width           =   3555
            End
            Begin VB.Label Label69 
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive Performance Sales Productivity Per Group"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   570
               Left            =   -63220
               TabIndex        =   165
               Top             =   2730
               Visible         =   0   'False
               Width           =   3540
            End
            Begin VB.Label Label57 
               BackStyle       =   0  'Transparent
               Caption         =   "Test Drive Evaluation Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Left            =   -63220
               TabIndex        =   164
               Top             =   3675
               Visible         =   0   'False
               Width           =   4305
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Log"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   -61720
               TabIndex        =   159
               Top             =   1110
               Visible         =   0   'False
               Width           =   1725
            End
            Begin VB.Label Label35 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer With Insurance Policies"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   540
               Left            =   -65095
               TabIndex        =   157
               Top             =   4635
               Visible         =   0   'False
               Width           =   2340
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Directory"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -69115
               TabIndex        =   155
               Top             =   3390
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label53 
               BackStyle       =   0  'Transparent
               Caption         =   "Yearly Customer Directory by Customer Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   -69115
               TabIndex        =   154
               Top             =   2535
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Birthday Celebrant of the Month"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   450
               Left            =   -65095
               TabIndex        =   153
               Top             =   3960
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Label Label54 
               BackStyle       =   0  'Transparent
               Caption         =   "Ranged Customer Directory By Customer Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   465
               Left            =   -69115
               TabIndex        =   152
               Top             =   1005
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label52 
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Customer Directory By Customer Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   450
               Left            =   -69115
               TabIndex        =   151
               Top             =   1770
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Information Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -69115
               TabIndex        =   150
               Top             =   4110
               Visible         =   0   'False
               Width           =   2475
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Sales Customer Directory"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Left            =   -69115
               TabIndex        =   149
               Top             =   4710
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label29 
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Reminders And Tasks"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   480
               Left            =   -65095
               TabIndex        =   141
               Top             =   2475
               Visible         =   0   'False
               Width           =   2355
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "List of Units Registered"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -64090
               TabIndex        =   139
               Top             =   2580
               Visible         =   0   'False
               Width           =   1980
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Units Released Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -64090
               TabIndex        =   137
               Top             =   1710
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.Label Label89 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Sales Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   7290
               TabIndex        =   133
               Top             =   870
               Width           =   1770
            End
            Begin VB.Label Label88 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Commisison"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   7290
               TabIndex        =   132
               Top             =   1770
               Width           =   1440
            End
            Begin VB.Label Label83 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales && Stock Tracking"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -64090
               TabIndex        =   125
               Top             =   990
               Visible         =   0   'False
               Width           =   1995
            End
            Begin VB.Label Label82 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Order Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -64090
               TabIndex        =   123
               Top             =   4050
               Visible         =   0   'False
               Width           =   1980
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Service Appointment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -65095
               TabIndex        =   112
               Top             =   1860
               Visible         =   0   'False
               Width           =   1770
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Internal Reminders"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -65095
               TabIndex        =   111
               Top             =   3330
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Appointment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -65095
               TabIndex        =   110
               Top             =   1080
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Trend"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68965
               TabIndex        =   72
               Top             =   2790
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label67 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Progress Monitoring"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68965
               TabIndex        =   71
               Top             =   1995
               Visible         =   0   'False
               Width           =   1740
            End
            Begin VB.Label Label55 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lost Sales Monitoring"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68965
               TabIndex        =   70
               Top             =   1155
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Distribution of Sales as to Mode of Payment"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   66
               Top             =   2550
               Width           =   3690
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Sales Summary  Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   7290
               TabIndex        =   65
               Top             =   2550
               Width           =   2685
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Yearly Vehicle Gross Profit Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   64
               Top             =   1815
               Width           =   2880
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Vehicle Gross Profit Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1080
               TabIndex        =   63
               Top             =   945
               Width           =   3015
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delivery Unit Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68815
               TabIndex        =   60
               Top             =   4185
               Visible         =   0   'False
               Width           =   1665
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Inventory Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68815
               TabIndex        =   58
               Top             =   3390
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ending Inventory Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68815
               TabIndex        =   57
               Top             =   2625
               Visible         =   0   'False
               Width           =   2010
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Inventory Control"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68815
               TabIndex        =   53
               Top             =   1875
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Purchase Report"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68815
               TabIndex        =   51
               Top             =   1065
               Visible         =   0   'False
               Width           =   2145
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage4 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   4
         Top             =   585
         Visible         =   0   'False
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage3 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   3
         Top             =   585
         Visible         =   0   'False
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
         Begin XtremeSuiteControls.TabControl TabControl2 
            Height          =   5685
            Left            =   -60
            TabIndex        =   48
            Top             =   -60
            Width           =   10755
            _Version        =   655364
            _ExtentX        =   18971
            _ExtentY        =   10028
            _StockProps     =   64
            Appearance      =   2
            Color           =   4
            PaintManager.BoldSelected=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            PaintManager.FixedTabWidth=   100
            PaintManager.MinTabWidth=   100
            ItemCount       =   3
            Item(0).Caption =   "Prospect && Sales Executives"
            Item(0).ControlCount=   8
            Item(0).Control(0)=   "Label50"
            Item(0).Control(1)=   "Label51"
            Item(0).Control(2)=   "Label38"
            Item(0).Control(3)=   "Label10"
            Item(0).Control(4)=   "cmdInq_ProspInq"
            Item(0).Control(5)=   "cmdSaleAPbySae"
            Item(0).Control(6)=   "cmdMonthlyAppointmentCal"
            Item(0).Control(7)=   "cmdSAEPerf"
            Item(1).Caption =   "Vehicles && Inventory"
            Item(1).ControlCount=   16
            Item(1).Control(0)=   "cmdVehicleMaster"
            Item(1).Control(1)=   "Command20"
            Item(1).Control(2)=   "Command21"
            Item(1).Control(3)=   "Command22"
            Item(1).Control(4)=   "Label63"
            Item(1).Control(5)=   "Label42"
            Item(1).Control(6)=   "Label43"
            Item(1).Control(7)=   "Label44"
            Item(1).Control(8)=   "Label47"
            Item(1).Control(9)=   "Label76"
            Item(1).Control(10)=   "Label77"
            Item(1).Control(11)=   "Label78"
            Item(1).Control(12)=   "cmdAllocatedCar"
            Item(1).Control(13)=   "cmdInvoicedCa"
            Item(1).Control(14)=   "cmdReleaedVehiInqui"
            Item(1).Control(15)=   "cmdVechOnStockInq"
            Item(2).Caption =   "Customer"
            Item(2).ControlCount=   12
            Item(2).Control(0)=   "Command16"
            Item(2).Control(1)=   "Label62"
            Item(2).Control(2)=   "Label61"
            Item(2).Control(3)=   "Label58"
            Item(2).Control(4)=   "Label71"
            Item(2).Control(5)=   "Label72"
            Item(2).Control(6)=   "Label74"
            Item(2).Control(7)=   "cmdAction(22)"
            Item(2).Control(8)=   "cmdServiceHistory"
            Item(2).Control(9)=   "cmdCustSalesHist"
            Item(2).Control(10)=   "cmdCUstVisitCall"
            Item(2).Control(11)=   "cmdCustVehInfoInq"
            Begin VB.CommandButton cmdSAEPerf 
               Height          =   645
               Left            =   285
               MouseIcon       =   "SMISMainMenu.frx":20813
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":20965
               Style           =   1  'Graphical
               TabIndex        =   115
               Tag             =   "1108"
               ToolTipText     =   "Sales Executive Performance"
               Top             =   3015
               Width           =   720
            End
            Begin VB.CommandButton cmdMonthlyAppointmentCal 
               Height          =   645
               Left            =   285
               MouseIcon       =   "SMISMainMenu.frx":213DA
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":2152C
               Style           =   1  'Graphical
               TabIndex        =   114
               Tag             =   "1220"
               ToolTipText     =   "Monthly Appointment Calendar"
               Top             =   2220
               Width           =   720
            End
            Begin VB.CommandButton cmdSaleAPbySae 
               Height          =   645
               Left            =   285
               MouseIcon       =   "SMISMainMenu.frx":21EE4
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":22036
               Style           =   1  'Graphical
               TabIndex        =   113
               Tag             =   "1197"
               ToolTipText     =   "Sales Appointment By SAE"
               Top             =   1425
               Width           =   720
            End
            Begin VB.CommandButton cmdCUstVisitCall 
               Height          =   645
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":228E2
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":22A34
               Style           =   1  'Graphical
               TabIndex        =   100
               Tag             =   "1229"
               ToolTipText     =   "Customer Visit/Call"
               Top             =   2985
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAction 
               Height          =   645
               Index           =   22
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":2310C
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":2325E
               Style           =   1  'Graphical
               TabIndex        =   99
               ToolTipText     =   "Customer Transaction History"
               Top             =   2250
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdServiceHistory 
               Height          =   645
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":23928
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":23A7A
               Style           =   1  'Graphical
               TabIndex        =   98
               Tag             =   "1228"
               ToolTipText     =   "Customer Service History"
               Top             =   1500
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustSalesHist 
               Height          =   645
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":24147
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":24299
               Style           =   1  'Graphical
               TabIndex        =   97
               Tag             =   "1227"
               ToolTipText     =   "Customer Sales History"
               Top             =   750
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdCustVehInfoInq 
               Height          =   645
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":24959
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":24AAB
               Style           =   1  'Graphical
               TabIndex        =   96
               Tag             =   "1229"
               ToolTipText     =   "Customer Vehicle Information"
               Top             =   3735
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command16 
               Height          =   645
               Left            =   -69685
               MouseIcon       =   "SMISMainMenu.frx":25175
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":252C7
               Style           =   1  'Graphical
               TabIndex        =   95
               Tag             =   "1102"
               ToolTipText     =   "Customer Reminders/Tasks"
               Top             =   4485
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command22 
               Height          =   645
               Left            =   -64600
               MouseIcon       =   "SMISMainMenu.frx":25980
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":25C8A
               Style           =   1  'Graphical
               TabIndex        =   86
               Tag             =   "1107"
               ToolTipText     =   "OverDue Order"
               Top             =   2130
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command21 
               Height          =   645
               Left            =   -64600
               MouseIcon       =   "SMISMainMenu.frx":263B4
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":266BE
               Style           =   1  'Graphical
               TabIndex        =   85
               Tag             =   "1107"
               ToolTipText     =   "Pending Order"
               Top             =   1410
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton Command20 
               Height          =   645
               Left            =   -64600
               MouseIcon       =   "SMISMainMenu.frx":26DBB
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":270C5
               Style           =   1  'Graphical
               TabIndex        =   84
               Tag             =   "1107"
               ToolTipText     =   "Served Order"
               Top             =   690
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdAllocatedCar 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":277BE
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":27910
               Style           =   1  'Graphical
               TabIndex        =   83
               Tag             =   "1104"
               ToolTipText     =   "Allocated Cars"
               Top             =   690
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdVechOnStockInq 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":28017
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":28321
               Style           =   1  'Graphical
               TabIndex        =   82
               Tag             =   "1107"
               ToolTipText     =   "Vehicles On Stock"
               Top             =   2850
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdReleaedVehiInqui 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":28A3E
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":28D48
               Style           =   1  'Graphical
               TabIndex        =   81
               Tag             =   "1106"
               ToolTipText     =   "Total Vehicle Released"
               Top             =   2130
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdInvoicedCa 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":29400
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":2970A
               Style           =   1  'Graphical
               TabIndex        =   80
               Tag             =   "1105"
               ToolTipText     =   "Invoiced Cars"
               Top             =   1410
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdVehicleMaster 
               Height          =   645
               Left            =   -69760
               MouseIcon       =   "SMISMainMenu.frx":29DDB
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":2A0E5
               Style           =   1  'Graphical
               TabIndex        =   79
               Tag             =   "1107"
               ToolTipText     =   "Vehicle Master Inquiry"
               Top             =   3570
               Visible         =   0   'False
               Width           =   720
            End
            Begin VB.CommandButton cmdInq_ProspInq 
               Height          =   645
               Left            =   285
               MouseIcon       =   "SMISMainMenu.frx":2A80F
               MousePointer    =   99  'Custom
               Picture         =   "SMISMainMenu.frx":2A961
               Style           =   1  'Graphical
               TabIndex        =   77
               Tag             =   "1196"
               ToolTipText     =   "Prospect Inquiry"
               Top             =   645
               Width           =   720
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive Performance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1110
               TabIndex        =   118
               Top             =   3240
               Width           =   2520
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Appointment Calendar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1110
               TabIndex        =   117
               Top             =   2445
               Width           =   2595
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Appointment By SAE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1110
               TabIndex        =   116
               Top             =   1620
               Width           =   2250
            End
            Begin VB.Label Label74 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Transaction History"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   106
               Top             =   2475
               Visible         =   0   'False
               Width           =   2550
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Vehicle Information"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   105
               Top             =   3945
               Visible         =   0   'False
               Width           =   2535
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Reminders/Task"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   104
               Top             =   4665
               Visible         =   0   'False
               Width           =   2280
            End
            Begin VB.Label Label58 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Service History"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   103
               Top             =   1725
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.Label Label61 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Sales History"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   102
               Top             =   960
               Visible         =   0   'False
               Width           =   2010
            End
            Begin VB.Label Label62 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer Visit/Call"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68875
               TabIndex        =   101
               Top             =   3180
               Visible         =   0   'False
               Width           =   1620
            End
            Begin VB.Label Label78 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Overdue Purchase Order"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -63790
               TabIndex        =   94
               Top             =   2280
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.Label Label77 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pending Purchase Order"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -63790
               TabIndex        =   93
               Top             =   1590
               Visible         =   0   'False
               Width           =   2100
            End
            Begin VB.Label Label76 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Served Purchase Order"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -63790
               TabIndex        =   92
               Top             =   870
               Visible         =   0   'False
               Width           =   2010
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Allocated Cars"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68950
               TabIndex        =   91
               Top             =   855
               Visible         =   0   'False
               Width           =   1245
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Invoiced Cars"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68950
               TabIndex        =   90
               Top             =   1575
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicles On Stock"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68950
               TabIndex        =   89
               Top             =   3030
               Visible         =   0   'False
               Width           =   1560
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Vehicle Released"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68950
               TabIndex        =   88
               Top             =   2310
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.Label Label63 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Master Inquiry"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   -68950
               TabIndex        =   87
               Top             =   3765
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.Label Label50 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prospect Inquiry"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   225
               Left            =   1110
               TabIndex        =   78
               Top             =   855
               Width           =   1395
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5640
         Left            =   -69970
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
         Begin VB.CommandButton cmdMake 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2B218
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2B36A
            Style           =   1  'Graphical
            TabIndex        =   205
            ToolTipText     =   "Test Drive Evaluation Report"
            Top             =   2850
            Width           =   720
         End
         Begin VB.CommandButton Command26 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2BCA3
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2BDF5
            Style           =   1  'Graphical
            TabIndex        =   199
            Tag             =   "1233"
            ToolTipText     =   "PDI Inspection List"
            Top             =   4950
            Width           =   720
         End
         Begin VB.CommandButton Command19 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":2C444
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2C596
            Style           =   1  'Graphical
            TabIndex        =   190
            Tag             =   "1234"
            ToolTipText     =   "PDI Set Up"
            Top             =   4920
            Width           =   720
         End
         Begin VB.CommandButton cmdjob 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2CA3A
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2CB8C
            Style           =   1  'Graphical
            TabIndex        =   184
            Tag             =   "1233"
            ToolTipText     =   "PDI Inspection List"
            Top             =   3555
            Width           =   720
         End
         Begin VB.CommandButton cmdModelATC 
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
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2D1DB
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2D32D
            Style           =   1  'Graphical
            TabIndex        =   122
            Tag             =   "1020"
            ToolTipText     =   "View Chart Of Accounts"
            Top             =   4245
            Width           =   720
         End
         Begin VB.CommandButton cmdTab_CustInfo 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":2D961
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2DAB3
            Style           =   1  'Graphical
            TabIndex        =   36
            Tag             =   "1082"
            ToolTipText     =   "Customer Information"
            Top             =   60
            Width           =   720
         End
         Begin VB.CommandButton cmdSAE 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2E14E
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2E2A0
            Style           =   1  'Graphical
            TabIndex        =   35
            Tag             =   "1087"
            ToolTipText     =   "Sales Account Executives"
            Top             =   1470
            Width           =   720
         End
         Begin VB.CommandButton cmdFinCompany 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":2E8F0
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2EA42
            Style           =   1  'Graphical
            TabIndex        =   34
            Tag             =   "1086"
            ToolTipText     =   "Financing Company"
            Top             =   90
            Width           =   720
         End
         Begin VB.CommandButton cmdTab_Model 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":2F1A6
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2F2F8
            Style           =   1  'Graphical
            TabIndex        =   33
            Tag             =   "1136"
            ToolTipText     =   "Vehicle Model"
            Top             =   1440
            Width           =   720
         End
         Begin VB.CommandButton cmdTab_Color 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":2FAF4
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":2FC46
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "1084"
            ToolTipText     =   "Vehicle Color"
            Top             =   750
            Width           =   720
         End
         Begin VB.CommandButton cmdFinDoc 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":30309
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3045B
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "1203"
            ToolTipText     =   "Financial Document"
            Top             =   2130
            Width           =   720
         End
         Begin VB.CommandButton cmdVehClass 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":30B00
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":30C52
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "1206"
            ToolTipText     =   "Vehicle Class"
            Top             =   780
            Width           =   720
         End
         Begin VB.CommandButton cmdTestDriveVeh 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":311BF
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":31311
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "1205"
            ToolTipText     =   "Test Drive Vehicles"
            Top             =   2835
            Width           =   720
         End
         Begin VB.CommandButton cmdLeadClassification 
            Height          =   645
            Left            =   5610
            MouseIcon       =   "SMISMainMenu.frx":31A85
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":31BD7
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "1208"
            ToolTipText     =   "Lead Classifications"
            Top             =   2160
            Width           =   720
         End
         Begin VB.CommandButton cmdPDIInspectionList 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":32441
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":32593
            Style           =   1  'Graphical
            TabIndex        =   27
            Tag             =   "1233"
            ToolTipText     =   "PDI Inspection List"
            Top             =   3525
            Width           =   720
         End
         Begin VB.CommandButton cmdPDISetup 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":32BE2
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":32D34
            Style           =   1  'Graphical
            TabIndex        =   26
            Tag             =   "1234"
            ToolTipText     =   "PDI Set Up"
            Top             =   4215
            Width           =   720
         End
         Begin VB.Label Label98 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Make"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   206
            Top             =   3090
            Width           =   1140
         End
         Begin VB.Label Label96 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Job Request"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   200
            Top             =   5100
            Width           =   1590
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AOR/OMA/DI Setup"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   191
            Top             =   5100
            Width           =   1575
         End
         Begin VB.Label Label92 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Job Master File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   185
            Top             =   3795
            Width           =   1815
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allocation Slip"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   121
            Top             =   4485
            Width           =   1200
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   47
            Top             =   1620
            Width           =   1185
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   46
            Top             =   270
            Width           =   1860
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   45
            Top             =   930
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Account Executives"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   44
            Top             =   1680
            Width           =   2205
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financing Company"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   43
            Top             =   270
            Width           =   1650
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Financial Document"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   42
            Top             =   2340
            Width           =   1665
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Class"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   41
            Top             =   960
            Width           =   1155
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Classifications"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6465
            TabIndex        =   40
            Top             =   2400
            Width           =   1725
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Drive Vehicles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   39
            Top             =   3045
            Width           =   1635
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PDI Inspection List"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   38
            Top             =   3735
            Width           =   1575
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PDI Set Up"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1095
            TabIndex        =   37
            Top             =   4425
            Width           =   885
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5640
         Left            =   30
         TabIndex        =   1
         Top             =   585
         Width           =   10695
         _Version        =   655364
         _ExtentX        =   18865
         _ExtentY        =   9948
         _StockProps     =   0
         Begin VB.CommandButton cmdMM_StockTransfer 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":33378
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":334CA
            Style           =   1  'Graphical
            TabIndex        =   197
            Tag             =   "1240"
            ToolTipText     =   "Stock Transfer"
            Top             =   2784
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_LoanIndiv 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":33B15
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":33C67
            Style           =   1  'Graphical
            TabIndex        =   196
            Tag             =   "1096"
            ToolTipText     =   "Loan Application ( Individual)"
            Top             =   2790
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_ProspectLog 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":342D3
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":34425
            Style           =   1  'Graphical
            TabIndex        =   180
            Tag             =   "1102"
            ToolTipText     =   "Prospect Logs"
            Top             =   741
            Width           =   720
         End
         Begin VB.CommandButton Command2 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":34AA0
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":34BF2
            Style           =   1  'Graphical
            TabIndex        =   173
            Tag             =   "1086"
            ToolTipText     =   "Financing Company"
            Top             =   4146
            Width           =   720
         End
         Begin VB.CommandButton Command17 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":35356
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":354A8
            Style           =   1  'Graphical
            TabIndex        =   172
            Tag             =   "1102"
            ToolTipText     =   "Prospects"
            Top             =   60
            Width           =   735
         End
         Begin VB.CommandButton Command9 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":35B9A
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":35CEC
            Style           =   1  'Graphical
            TabIndex        =   128
            Tag             =   "1086"
            ToolTipText     =   "Financing Company"
            Top             =   4830
            Width           =   720
         End
         Begin VB.CommandButton Command1 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":36450
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":365A2
            Style           =   1  'Graphical
            TabIndex        =   126
            Tag             =   "1139"
            ToolTipText     =   "Vehicle Sales Monitoring"
            Top             =   4830
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_VSA 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":36C1C
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":36D6E
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "1139"
            ToolTipText     =   "Vehicle Sales Monitoring"
            Top             =   60
            Width           =   735
         End
         Begin VB.CommandButton cmdMM_Invoicing 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":373E8
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3753A
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "1140"
            ToolTipText     =   "Vehicle Invoicing"
            Top             =   1425
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_VSO 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":37BC0
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":37D12
            Style           =   1  'Graphical
            TabIndex        =   13
            Tag             =   "1207"
            ToolTipText     =   "Vehicle Sales Order"
            Top             =   2100
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_SalesCalc 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":38399
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":384EB
            Style           =   1  'Graphical
            TabIndex        =   12
            Tag             =   "1102"
            ToolTipText     =   "Sales Calculator"
            Top             =   2103
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_CustomerLog 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":38B8A
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":38CDC
            Style           =   1  'Graphical
            TabIndex        =   11
            Tag             =   "1102"
            ToolTipText     =   "Customer Logs"
            Top             =   1422
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_Quotation 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":39268
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":393BA
            Style           =   1  'Graphical
            TabIndex        =   10
            Tag             =   "1088"
            ToolTipText     =   "Quotation"
            Top             =   4146
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_Receving 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":39B77
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":39CC9
            Style           =   1  'Graphical
            TabIndex        =   9
            Tag             =   "1137"
            ToolTipText     =   "Vehicle Receiving"
            Top             =   750
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_LoanCorp 
            Height          =   645
            Left            =   5580
            MouseIcon       =   "SMISMainMenu.frx":3A2F5
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3A447
            Style           =   1  'Graphical
            TabIndex        =   8
            Tag             =   "1097"
            ToolTipText     =   "Loan Application ( Corporate) "
            Top             =   3465
            Width           =   720
         End
         Begin VB.CommandButton cmdMM_VPO 
            Height          =   645
            Left            =   240
            MouseIcon       =   "SMISMainMenu.frx":3AA61
            MousePointer    =   99  'Custom
            Picture         =   "SMISMainMenu.frx":3ABB3
            Style           =   1  'Graphical
            TabIndex        =   7
            Tag             =   "1088"
            ToolTipText     =   "Vehicle Purchase Order"
            Top             =   3480
            Width           =   720
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Transfer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   198
            Top             =   3000
            Width           =   1275
         End
         Begin VB.Label Label91 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prospect Logs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   181
            Top             =   945
            Width           =   1245
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Prospects (Sales Diary)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   177
            Top             =   270
            Width           =   2010
         End
         Begin VB.Label Label85 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LTO Status  Monitoring"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   174
            Top             =   4365
            Width           =   1935
         End
         Begin VB.Label Label86 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commission"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   129
            Top             =   5040
            Width           =   1065
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Pre Delivery Inspection"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   127
            Top             =   5040
            Width           =   2625
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sales Monitoring"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   25
            Top             =   270
            Width           =   2100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sales Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   24
            Top             =   2310
            Width           =   1695
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Application (Individual)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   23
            Top             =   3015
            Width           =   2370
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Invoicing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   22
            Top             =   1620
            Width           =   1425
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Calculator"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   21
            Top             =   2310
            Width           =   1395
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Logs"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   20
            Top             =   1620
            Width           =   1305
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Sales Quotation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   19
            Top             =   4350
            Width           =   2025
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Receiving"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   18
            Top             =   945
            Width           =   1500
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Application (Corporate) "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   6435
            TabIndex        =   17
            Top             =   3705
            Width           =   2475
         End
         Begin VB.Label Label79 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Purchase Order"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1080
            TabIndex        =   16
            Top             =   3675
            Width           =   2040
         End
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6630
      Top             =   3870
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAction_Click(Index As Integer)
    If Module_Access(LOGID, "CUSTOMER TRANSACTION HISTORY", "INQUIRY") = False Then Exit Sub
    frmCRIS_Inquiry_CustomerTransHistory.Show
    frmCRIS_Inquiry_CustomerTransHistory.ZOrder 0

End Sub

Private Sub cmdAfterSalesReport_Click()
    If Module_Access(LOGID, "AFTER SALES REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_AfterSales.Show
    frmCRIS_Report_AfterSales.ZOrder 0
End Sub

Private Sub cmdHitRatio_Click()
    If Module_Access(LOGID, "HIT RATIO", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_HitRatio.Show    '*******RYAN CULAWAY MAY 8 2:07PM
End Sub

Private Sub cmdjob_Click()
    frmSMIS_FILE_jobMasterFile.Show
End Sub

Private Sub cmdMake_Click()
    If Module_Access(LOGID, "MAKE", "DATA ENTRY") = False Then Exit Sub
    frmCSMS_MAKE.Show
End Sub

'Private Sub cmdLeadSource_Click()
'If Module_Access(LOGID, "LEAD SOURCE", "REPORTS") = False Then Exit Sub ' **************RYAN CULAWAY MAY 08 2008
'frmSMIS_Report_LeadSource.Show
'End Sub

Private Sub cmdVehicleOnStock_Click()
    If Module_Access(LOGID, "PURCHASE ORDER REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VehiclePurchase.Show
End Sub

Private Sub Command15_Click()
    If Module_Access(LOGID, "LOG SUMMARY REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_LogSummary.Show
    frmSMIS_Report_LogSummary.ZOrder 0
End Sub

Private Sub Command18_Click()
    If Module_Access(LOGID, "VEHICLE SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_GrossProfit.Show
    frmSMIS_Report_GrossProfit.ZOrder 0
End Sub

Private Sub Command19_Click()
    If Module_Access(LOGID, "AOR-OMA-DI MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_File_AORRate.Show
End Sub

Private Sub Command23_Click()
    If Module_Access(LOGID, "NET SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_NetSales.Show
End Sub

Private Sub Command25_Click()
    If Module_Access(LOGID, "TRANSACTION STATUS REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_TransactionStatus.Show
End Sub

Private Sub Command26_Click()
    On Error Resume Next
    If Module_Access(LOGID, "SALES JOB REQUEST", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_jobRequest.Show

End Sub

Private Sub Command27_Click()
    If COMPANY_CODE <> "HAS" Then Exit Sub
    If Module_Access(LOGID, "SALES REPORTS", "REPORTS") = False Then Exit Sub
    HASREPORT.Show
End Sub

Private Sub Command28_Click()
    If Module_Access(LOGID, "VEHICLE SALES PROJECTION", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VehicleSalesProjection.Show
    frmSMIS_Report_VehicleSalesProjection.ZOrder 0
End Sub


Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    TabControl3.SelectedItem = 0
    TabControl2.SelectedItem = 0
    TabControl1.SelectedItem = 0
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub cmd_UnitCommission_Click()
    If Module_Access(LOGID, "UNIT COMMISSION", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_UnitCommission.Show
    frmSMIS_Report_UnitCommission.ZOrder 0
End Sub

Private Sub cmdAllocatedCar_Click()
    If Module_Access(LOGID, "ALLOCATED CARS", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry.Show
    frmSMIS_Inquiry.optAllCars.Value = True
    frmSMIS_Inquiry.ZOrder 0
End Sub

Private Sub cmdBirYear_Click()
    If Module_Access(LOGID, "BIR YEAR REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_BIRYearEnd.Show
    frmSMIS_Report_BIRYearEnd.ZOrder 0
End Sub

Private Sub cmdCustomerwithInsurance_Click()
    If Module_Access(LOGID, "CUSTOMERS WITH INSURANCE POLICIES", "REPORTS") = False Then Exit Sub
    CUST_REPT_TYPE = "2"
    frmSMIS_Report_CustSummary.Show
    frmSMIS_Report_CustSummary.ZOrder 0
End Sub

Private Sub cmdCustSalesHist_Click()
    If Module_Access(LOGID, "CUSTOMER SALES HISTORY", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_CustomerSalesHistory.Show
    frmSMIS_Inquiry_CustomerSalesHistory.ZOrder 0
End Sub

Private Sub cmdCustVehInfoInq_Click()
    If Module_Access(LOGID, "CUSTOMER CALL VISIT HISTORY", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_CallVisit_History.Show
End Sub

Private Sub cmdCUstVisitCall_Click()
    If Module_Access(LOGID, "CUSTOMER CALL VISIT HISTORY", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_CallVisit_History.Show
End Sub

Private Sub cmdDeliveryReport_Click()
    If Module_Access(LOGID, "DELIVERY UNITS REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_DelReport.Show
    frmSMIS_Report_DelReport.ZOrder 0
End Sub

Private Sub cmdDistributionOfSales_Click()
    If Module_Access(LOGID, "SALES DISTRIBUTION OF SALES AS TO MODE OF PAYMENT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_DistModePayment.Show
    frmSMIS_Report_DistModePayment.ZOrder 0
End Sub

Private Sub cmdEndingInvenotory_Click()
    If Module_Access(LOGID, "MONTHLY ENDING INVENTORY", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_InvControl2.Show
    frmSMIS_Report_InvControl2.ZOrder 0
End Sub

Private Sub cmdFinCompany_Click()
    If Module_Access(LOGID, "FINANCING COMPANY", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_FinancingCo.Show
    frmSMIS_Files_FinancingCo.ZOrder 0
End Sub

Private Sub cmdFinDoc_Click()
    If Module_Access(LOGID, "FINANCIAL DOCUMENTS", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Document.Show
    frmSMIS_Files_Document.ZOrder 0
End Sub

Private Sub cmdInq_ProspInq_Click()
    If Module_Access(LOGID, "PROSPECT INQUIRY", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_InquiryMain.optAdvSearch(0).Value = True
    Load frmSMIS_Inquiry_InquiryMain
    frmSMIS_Inquiry_InquiryMain.Show
End Sub

Private Sub cmdInvoicedCa_Click()
    If Module_Access(LOGID, "INVOICED CARS", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry.Show
    frmSMIS_Inquiry.optInvCars.Value = True
    frmSMIS_Inquiry.ZOrder 0

End Sub

Private Sub cmdLeadClassification_Click()
    If Module_Access(LOGID, "CLASSIFY LEADS", "DATA ENTRY") = False Then Exit Sub
    frmCRIS_ClassifyLeads.Show
    frmCRIS_ClassifyLeads.ZOrder 0
End Sub

Private Sub cmdListofUnitsReg_Click()
    If Module_Access(LOGID, "LIST OF UNITS REGISTERED", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_ListReg.Show
    frmSMIS_Report_ListReg.ZOrder 0
End Sub

Private Sub cmdMarketting_BirthDay_Click()
    If Module_Access(LOGID, "BIRTHDAY CELEBRANTS", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_CustomerBDays.Show
    frmSMIS_Report_CustomerBDays.ZOrder 0
End Sub

Private Sub cmdMarketting_CustomerDirectory_Click()
    If Module_Access(LOGID, "REPORT CUSTOMERS DIRECTORY", "REPORTS") = False Then Exit Sub
    frmMain.rptMain.Formulas(0) = "CompanyName='" & COMPANY_NAME & "'"
    frmMain.rptMain.Formulas(1) = "CompanyName='" & COMPANY_ADDRESS & "'"
    frmMain.rptMain.ReportFileName = SMIS_REPORT_PATH & "invoices.rpt"
    frmMain.rptMain.Connect = DMIS_REPORT_Connection
    frmMain.rptMain.Action = 1
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
     Call NEW_LogAudit("V", "AFTER SALES", "", "", "", "CUSTOMER INFORMATION REPORT", "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'LogAudit "V", "CUSTOMER INFORMATION REPORT"
End Sub

Private Sub cmdMM_Invoicing_Click()
    If Module_Access(LOGID, "VEHICLES INVOICING", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    frmSMIS_Trans_VehicleInvoice.Show
    If FormExist("frmSMIS_Trans_VehicleInvoice") Then
        frmSMIS_Trans_VehicleInvoice.WindowState = 0
        frmSMIS_Trans_VehicleInvoice.ZOrder 0
    End If

End Sub

Private Sub cmdMM_LoanIndiv_Click()

    If Module_Access(LOGID, "INDIVIDUAL LOAN APPLICATION", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    frmSMIS_Trans_ApplicationIndividual.Show
    If FormExist("frmSMIS_Trans_ApplicationIndividual") Then
        frmSMIS_Trans_ApplicationIndividual.WindowState = 0
        frmSMIS_Trans_ApplicationIndividual.ZOrder 0
    End If
End Sub

Private Sub cmdMM_LoanCorp_Click()
    If Module_Access(LOGID, "CORPORATE LOAN APPLICATION", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    frmSMIS_Trans_ApplicationCorporate.Show
    If FormExist("frmSMIS_Trans_ApplicationCorporate") Then
        frmSMIS_Trans_ApplicationCorporate.WindowState = 0
        frmSMIS_Trans_ApplicationCorporate.ZOrder 0
    End If


End Sub

Private Sub cmdMM_Receving_Click()
    If Module_Access(LOGID, "VEHICLE RECIEVING", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_MRR.Show
    If FormExist("frmSMIS_Trans_MRR") Then
        frmSMIS_Trans_MRR.WindowState = 0
        frmSMIS_Trans_MRR.ZOrder 0
    End If
End Sub

Private Sub cmdMM_StockTransfer_Click()
    If Module_Access(LOGID, "STOCK TRANSFER", "TRANSACTION") = False Then Exit Sub
    frmSMIS_StockTransferOption.Show
'    frmSMIS_Trans_MRR1.Show
'    If FormExist("frmSMIS_Trans_MRR1") Then
'        frmSMIS_Trans_MRR1.WindowState = 0
'        frmSMIS_Trans_MRR1.ZOrder 0
'    End If

End Sub

Private Sub cmdMM_VSO_Click()
    On Error Resume Next
    If Module_Access(LOGID, "SALES ADMIN SALES ORDER", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_SalesOrder.Show
    If FormExist("frmSMIS_Trans_SalesOrder") Then
        frmSMIS_Trans_SalesOrder.WindowState = 0
        frmSMIS_Trans_SalesOrder.ZOrder 0
    End If

End Sub

Private Sub cmdModelATC_Click()
    If Module_Access(LOGID, "ALLOCATION SLIP", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_FILE_ALLOCATIONSLIP.Show

End Sub

Private Sub cmdMonthlyAppointmentCal_Click()
    If Module_Access(LOGID, "SALES APPOINTMENT CALENDAR", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_SalesAppointment.Show

End Sub

Private Sub cmdMonthlyInventoryControl_Click()
    If Module_Access(LOGID, "MONTHLY INVENTORY CONTROL", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_InvControl.Show
End Sub

Private Sub cmdMonthlyReport_Click()
    If Module_Access(LOGID, "MONTHLY PURCHASES REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_Purchases2.Show
    frmSMIS_Report_Purchases2.ZOrder 0
End Sub

Private Sub cmdMonthlyVehicleGrossProfit_Click()
    If Module_Access(LOGID, "MONTHLY VEHICLE GROSS PROFIT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_GenRep.Show
    frmSMIS_Report_GenRep.ZOrder 0
End Sub

Private Sub cmdOther_SAEPerformance_1_Click()
    If Module_Access(LOGID, "SALES EXECUTIVE PERFORMANCE", "REPORTS") = False Then Exit Sub

    frmSMIS_Report_SAEPer.Show
    frmSMIS_Report_SAEPer.ZOrder 0
End Sub

Private Sub cmdPDIInspectionList_Click()
    If Module_Access(LOGID, "PDI CHECKLIST", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_PDICheckList.Show
    frmSMIS_Files_PDICheckList.ZOrder 0

End Sub

Private Sub cmdPDISetup_Click()

    If Module_Access(LOGID, "PDI SETUP", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_PDISetup.Show
    frmSMIS_Files_PDISetup.ZOrder 0
End Sub

Private Sub cmdPrint_Click()
    If Module_Access(LOGID, "SALES REPORTS", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VehicleSales.Show
End Sub

Private Sub cmdReleaedVehiInqui_Click()
    If Module_Access(LOGID, "TOTAL RELEASED VEHICLES", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry.Show
    frmSMIS_Inquiry.optCarRelease.Value = True
    frmSMIS_Inquiry.ZOrder 0
End Sub

Private Sub cmdReport_Sales_VehicleSales_Click()
    If Module_Access(LOGID, "VEHICLE SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VSRep.Show
    frmSMIS_Report_VSRep.ZOrder 0
End Sub

Private Sub cmdSAE_Click()
    If Module_Access(LOGID, "SALES ACCOUNT EXECUTIVE", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_SalesAE.Show
    frmSMIS_Files_SalesAE.ZOrder 0
End Sub

Private Sub cmdSAEPerf_Click()
    If Module_Access(LOGID, "SALES EXECUTIVE PERFORMANCE", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry.Show
    frmSMIS_Inquiry.optSalesPer.Value = True
    frmSMIS_Inquiry.ZOrder 0
End Sub

Private Sub cmdSaleAPbySae_Click()
    If Module_Access(LOGID, "SALES APPOINTMENT BY SAE", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_InquiryMain.optAdvSearch(1).Value = True
    Load frmSMIS_Inquiry_InquiryMain
    frmSMIS_Inquiry_InquiryMain.Show
    frmSMIS_Inquiry_InquiryMain.ZOrder 0

End Sub

Private Sub cmdServiceHistory_Click()
    If Module_Access(LOGID, "CUSTOMER SERVICE HISTORY", "INQUIRY") = False Then Exit Sub
    frmCSMSCustomerHistory.Show
    frmCSMSCustomerHistory.ZOrder 0
End Sub

Private Sub cmdSignatories_Click()
    If Module_Access(LOGID, "SIGNATORIES", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Signatories.Show
End Sub

Private Sub cmdStockandSalesTracking_Click()
    If Module_Access(LOGID, "SALES AND STOCK TRACKING REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_GenDSSR.Show
    frmSMIS_Report_GenDSSR.ZOrder 0

End Sub

Private Sub cmdTab_Color_Click()
    If Module_Access(LOGID, "VEHICLE COLOR", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Color.Show
    frmSMIS_Files_Color.ZOrder 0
End Sub

Private Sub cmdTab_CustInfo_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
    Call frmAllCustomer.AddEditCustomer("")
    frmAllCustomer.Show
    frmAllCustomer.ZOrder 0
End Sub

Private Sub cmdTab_Model_Click()
    If Module_Access(LOGID, "VEHICLE MODEL", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_Model.Show
    frmSMIS_Files_Model.ZOrder 0
End Sub

Private Sub cmdTestDriveVeh_Click()
    If Module_Access(LOGID, "TEST DRIVE VEHICLES", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Trans_TestVehicles.Show
    frmSMIS_Trans_TestVehicles.ZOrder 0
End Sub

Private Sub cmdUnitReleasedReport_Click()
    If Module_Access(LOGID, "UNITS RELEASED REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_Released.Show
    frmSMIS_Report_Released.ZOrder 0
End Sub

Private Sub cmdVechOnStockInq_Click()
    If Module_Access(LOGID, "VEHICLES ON STOCK", "INQUIRY") = False Then Exit Sub

    frmSMIS_Inquiry.Show
    frmSMIS_Inquiry.optVehStock.Value = True
    frmSMIS_Inquiry.ZOrder 0
End Sub

Private Sub cmdVehClass_Click()
    If Module_Access(LOGID, "VEHICLE CLASS", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Files_VehiclesClass.Show
    frmSMIS_Files_VehiclesClass.ZOrder 0
End Sub

Private Sub cmdVehicle_InvReport_Click()
    If Module_Access(LOGID, "VEHICLE INVENTORY REPORT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VehiclesInventory.Show
    frmSMIS_Report_VehiclesInventory.ZOrder 0
End Sub

Private Sub cmdVehicleMaster_Click()
    If Module_Access(LOGID, "VEHICLE MASTER INQUIRY", "INQUIRY") = False Then Exit Sub
    frmSMIS_Inquiry_VehicleMaster.Show
    frmSMIS_Inquiry_VehicleMaster.ZOrder 0
End Sub

Private Sub cmdMM_VSA_Click()
    If Module_Access(LOGID, "VEHICLES SALES MONITORING", "SYSTEM") = False Then Exit Sub
    MainForm.Show
    If FormExist("MainForm") Then
        MainForm.WindowState = 0
        MainForm.ZOrder 0
    End If
End Sub

Private Sub cmdMM_ProspectLog_Click()
    If Module_Access(LOGID, "PROSPECT LOG", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Log_Menu.picLogCustomer.Visible = False
    frmSMIS_Log_Menu.picLogProspect.Visible = True
    frmSMIS_Log_Menu.Show
End Sub

Private Sub cmdMM_Reminders_Click()
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub cmdVehiclesalesCustomer_Click()
    If Module_Access(LOGID, "VEHICLE SALES CUSTOMERS SUMMARY", "REPORTS") = False Then Exit Sub
    CUST_REPT_TYPE = "1"
    frmSMIS_Report_CustSummary.Show
    frmSMIS_Report_CustSummary.ZOrder 0
End Sub

Private Sub cmdYearlyGrossProfit_Click()
    If Module_Access(LOGID, "YEARLY VEHICLE GROSS PROFIT", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_RepYearly.Show
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "PDI CHECKLIST", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    frmSMIS_Trans_VehiclesCheckList.Show

    If FormExist("frmSMIS_Trans_VehiclesCheckList") Then
        frmSMIS_Trans_VehiclesCheckList.WindowState = 0
        frmSMIS_Trans_VehiclesCheckList.ZOrder 0
    End If

    Err.Clear
End Sub

Private Sub Command10_Click()
    If Module_Access(LOGID, "LTO STATUS", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_LTOStatus.Show
End Sub

Private Sub Command11_Click()
    If Module_Access(LOGID, "HYUNDAI REPORTS", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_VehicleSalesHyundai.Show
    frmSMIS_Report_VehicleSalesHyundai.ZOrder 0
End Sub

Private Sub Command12_Click()
    If Module_Access(LOGID, "PROGRESS MONITORING", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_SalesLead1.Show
    frmSMIS_Report_SalesLead1.ZOrder 0
End Sub

Private Sub Command13_Click()
    If Module_Access(LOGID, "SALES TREND", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_SalesLead2.Show
    frmSMIS_Report_SalesLead2.ZOrder 0
End Sub

Private Sub Command14_Click()
    If Module_Access(LOGID, "SALES EXECUTIVE TEAM PERFORMANCE", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_SAEPer.ShowSAETeamVsPROSPECT
    frmSMIS_Report_SAEPer.Show
End Sub

Private Sub Command16_Click()
    frmCRIS_Inquiry_TaskList.ShowTaskType ("C")
    frmCRIS_Inquiry_TaskList.Show
End Sub

Private Sub cmdMM_SalesCalc_Click()
    If Module_Access(LOGID, "SALES CALCULATOR", "SYSTEM") = False Then Exit Sub
    frmSMIS_Mis_AOR.ShowonlyComputation
    frmSMIS_Mis_AOR.Show
End Sub

Private Sub Command17_Click()
    If LOGSAE = "" Then
        frmSMIS_MISLogInSE.Show
    Else
        MainSAE.Show
    End If
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Module_Access(LOGID, "LTO STATUS", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_LTOStatus.Show
    If FormExist("frmSMIS_LTOStatus") Then
        frmSMIS_LTOStatus.WindowState = 0
        frmSMIS_LTOStatus.ZOrder 0
    End If
End Sub

Private Sub Command20_Click()
    If Module_Access(LOGID, "SERVED PO", "INQUIRY") = False Then Exit Sub
    If FormExist("frmSMIS_Inquiry_OverDuePending") Then
        Unload frmSMIS_Inquiry_OverDuePending
    End If
    frmSMIS_Inquiry_OverDuePending.ShowServerdOrders
    frmSMIS_Inquiry_OverDuePending.Show
End Sub

Private Sub Command21_Click()
    If Module_Access(LOGID, "PENDING PO", "INQUIRY") = False Then Exit Sub
    If FormExist("frmSMIS_Inquiry_OverDuePending") Then
        Unload frmSMIS_Inquiry_OverDuePending
    End If
    frmSMIS_Inquiry_OverDuePending.ShowPendingOrders
    frmSMIS_Inquiry_OverDuePending.Show
End Sub

Private Sub Command22_Click()
    If Module_Access(LOGID, "OVERDUE PO", "INQUIRY") = False Then Exit Sub
    If FormExist("frmSMIS_Inquiry_OverDuePending") Then
        Unload frmSMIS_Inquiry_OverDuePending
    End If
    frmSMIS_Inquiry_OverDuePending.ShowOverDueOrders
    frmSMIS_Inquiry_OverDuePending.Show
End Sub

Private Sub cmdMM_VPO_Click()
    If Module_Access(LOGID, "PURCHASE ORDER", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_Ordering.Show
    If FormExist("frmSMIS_Trans_Ordering") Then
        frmSMIS_Trans_Ordering.WindowState = 0
        frmSMIS_Trans_Ordering.ZOrder 0
    End If
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "CUSTOMER REMINDER", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Log_CustomerReminder.Show
    frmSMIS_Log_CustomerReminder.ZOrder 0
End Sub

Private Sub Command4_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchMonth
    frmSMIS_Report_AfterSales.Show
End Sub

Private Sub Command45_Click()
    If Module_Access(LOGID, "LOST SALES MONITORING", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_SalesLead.Show
    frmSMIS_Report_SalesLead.ZOrder 0
End Sub

Private Sub Command5_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchYear
    frmSMIS_Report_AfterSales.Show
End Sub

Private Sub Command6_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchRange
    frmSMIS_Report_AfterSales.Show
    frmSMIS_Report_AfterSales.ZOrder 0
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "SALES EXECUTIVE PERFORMANCE", "REPORTS") = False Then Exit Sub

    frmSMIS_Report_SAEPer.ShowSAEVsPROSPECT
    frmSMIS_Report_SAEPer.Show
    frmSMIS_Report_SAEPer.ZOrder 0
End Sub

Private Sub Command8_Click()
    If Module_Access(LOGID, "TEST DRIVE EVALUATION", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_TestDriveEvaluation.Show
    frmSMIS_Report_TestDriveEvaluation.ZOrder 0
End Sub

Private Sub cmdMM_CustomerLog_Click()
    If Module_Access(LOGID, "CUSTOMER LOG", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Log_Menu.picLogCustomer.Visible = True
    frmSMIS_Log_Menu.picLogProspect.Visible = False
    frmSMIS_Log_Menu.Show

End Sub

Private Sub Command9_Click()
    On Error Resume Next
    If Module_Access(LOGID, "COMMISSION", "TRANSACTION") = False Then Exit Sub
    frmSMIS_Trans_Commission.Show
    If FormExist("frmSMIS_Trans_Commission") Then
        frmSMIS_Trans_Commission.WindowState = 0
        frmSMIS_Trans_Commission.ZOrder 0
    End If




End Sub

Private Sub cmdCompanyProfile_Click()
    If Module_Access(LOGID, "MAINTAIN COMPANY PROFILE", "SYSTEM") = False Then Exit Sub
    frmSMIS_Files_Profile.Show
    frmSMIS_Files_Profile.ZOrder 0
End Sub

Private Sub cmdCustomer_Reminder_And_Task_Click()
    If Module_Access(LOGID, "CUSTOMER REMINDERS AND TASKS", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_CustomerRemindersAndTask.Show
    frmCRIS_Report_CustomerRemindersAndTask.ZOrder 0
End Sub

Private Sub cmdCustomer_Reminder_And_Task_Internal_Click()
    If Module_Access(LOGID, "INTERNAL REMINDERS", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_InternalReminder.Show
    frmCRIS_Report_InternalReminder.ZOrder 0
End Sub

Private Sub cmdCustomerInfoReport_Click()
    If Module_Access(LOGID, "CUSTOMER INFORMATION", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_CustomerInformationReport.Show
    frmCRIS_Report_CustomerInformationReport.ZOrder 0
End Sub

Private Sub cmdCustomerLog_Click()
    If Module_Access(LOGID, "CUSTOMER LOG REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_Log.Show
End Sub

Private Sub cmdPasswordMaintain_Click()
    frmAccMaintenance.Show
    frmAccMaintenance.ZOrder 0
End Sub

Private Sub cmdSalesAppointment_Click()
    If Module_Access(LOGID, "SALES APPOINTMENT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_SalesAppointment.Show
End Sub

Private Sub cmdServiceAppointment_Click()
    If Module_Access(LOGID, "SERVICE APPOINTMENT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_ServiceAppointment.Show
    frmCRIS_Report_ServiceAppointment.ZOrder 0
End Sub

Private Sub cmdMM_Quotation_Click()
    If Module_Access(LOGID, "QUOTATION", "TRANSACTION") = False Then Exit Sub
    On Error Resume Next
    frmSMIS_Trans_Quotation.Show
    If FormExist("frmSMIS_Trans_Quotation") Then
        frmSMIS_Trans_Quotation.WindowState = 0
        frmSMIS_Trans_Quotation.ZOrder 0
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Me.Hide
        Cancel = 1
    End If
End Sub



Private Sub Label84_Click()
    Dim RS
    Dim gconDMIS
    Set gconDMIS = CreateObject("ADODB.Connection")
    gconDMIS.Open ("DSN=DMIS")
    Set RS = CreateObject("ADODB.Recordset")


    If MsgBox("Are You Sure You Want To Update Master File ", vbInformation + vbYesNo) = vbYes Then


        Set RS = gconDMIS.Execute("SELECT * FROM SMIS_SALESORDER WHERE STATUS='P'")
        While Not RS.EOF

            gconDMIS.Execute ("update smis_MRRINV_TABLE SET ISTATUS='R', RELEASED=1 , DateReleased='" & RS!DateReleased & "',InvoicedDate='" & RS!InvoicedDate & "' ,vi_no='" & RS!VI_NO & "' WHERE IGNKEY='" & RS!IGNKEY_NO & "'")




            RS.MoveNext
        Wend
    End If
    MsgBox "SIR JUNE DONE NA EH"

End Sub

Private Sub TabControl1_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
    If Item.Index = 1 Then
        If COMPANY_CODE = "HAS" Then
            Command27.Visible = True
            Label99.Visible = True
        Else
            Command27.Visible = False
            Label99.Visible = False
        End If
    End If
End Sub
