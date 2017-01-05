VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CRIS Main Menu"
   ClientHeight    =   5955
   ClientLeft      =   990
   ClientTop       =   1065
   ClientWidth     =   10170
   ForeColor       =   &H8000000F&
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   10170
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      _Version        =   655364
      _ExtentX        =   17965
      _ExtentY        =   10557
      _StockProps     =   64
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   4
      Item(0).Caption =   "Main Modules"
      Item(0).ImageIndex=   920
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tbPageMainModules"
      Item(1).Caption =   "History"
      Item(1).ImageIndex=   921
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tbPageFileMaintenance"
      Item(2).Caption =   "Customer Logs"
      Item(2).ImageIndex=   922
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabControlPage1"
      Item(3).Caption =   "Reports"
      Item(3).Tooltip =   "Shortcut for Reports of Customer Relation Management Information System"
      Item(3).ImageIndex=   923
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   5355
         Left            =   -69970
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   10125
         _Version        =   655364
         _ExtentX        =   17859
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton cmdAction 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":6852
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":69A4
            Style           =   1  'Graphical
            TabIndex        =   32
            Tag             =   "1196 "
            ToolTipText     =   "View Customer Log Inquiry"
            Top             =   3390
            Width           =   720
         End
         Begin VB.CommandButton Command4 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":6FC3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":7115
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "1088"
            ToolTipText     =   "View Log Call"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton Command5 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":7637
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":7789
            Style           =   1  'Graphical
            TabIndex        =   30
            Tag             =   "1088"
            ToolTipText     =   "View Log Visit"
            Top             =   1125
            Width           =   720
         End
         Begin VB.CommandButton Command8 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":7D80
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":7ED2
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "1088"
            ToolTipText     =   "View Log Email"
            Top             =   2625
            Width           =   720
         End
         Begin VB.CommandButton Command7 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":856A
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":86BC
            Style           =   1  'Graphical
            TabIndex        =   28
            Tag             =   "1088"
            ToolTipText     =   "View Log Letter"
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Log Inquiry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1425
            TabIndex        =   37
            Top             =   3480
            Width           =   2040
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log Call"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1425
            TabIndex        =   36
            Top             =   2070
            Width           =   795
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log Visit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1425
            TabIndex        =   35
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log Email"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1425
            TabIndex        =   34
            Top             =   2820
            Width           =   960
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Log Letter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1425
            TabIndex        =   33
            Top             =   555
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageFileMaintenance 
         Height          =   5355
         Left            =   -69970
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   10125
         _Version        =   655364
         _ExtentX        =   17859
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton Command6 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":8CA4
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":8DF6
            Style           =   1  'Graphical
            TabIndex        =   38
            Tag             =   "1102"
            ToolTipText     =   "View Customer Reminders/Tasks"
            Top             =   3465
            Width           =   720
         End
         Begin VB.CommandButton Command1 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":94AF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":9601
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Customer Vehicle Information Details "
            Top             =   2700
            Width           =   720
         End
         Begin VB.CommandButton cmdDeductions 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":9CCB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":9E1D
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "View Customer Sales History"
            Top             =   405
            Width           =   720
         End
         Begin VB.CommandButton cmdCommission 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":A4DD
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":A62F
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "View Customer Service History"
            Top             =   1170
            Width           =   720
         End
         Begin VB.CommandButton cmdAdjustments 
            Height          =   645
            Left            =   570
            MouseIcon       =   "MainMenu.frx":ACFC
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":AE4E
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "View Customer Transaction History"
            Top             =   4245
            Width           =   720
         End
         Begin VB.CommandButton cmdGenPay 
            Height          =   645
            Left            =   600
            MouseIcon       =   "MainMenu.frx":B518
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":B66A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Customer Call/Visit History"
            Top             =   1935
            Width           =   720
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Reminders/ Tasks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1410
            TabIndex        =   39
            Top             =   3600
            Width           =   5070
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Vehicle Information Details "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1410
            TabIndex        =   22
            Top             =   2895
            Width           =   3600
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Call/Visit History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1410
            TabIndex        =   20
            Top             =   2085
            Width           =   2505
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Transaction History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1380
            TabIndex        =   19
            Top             =   4395
            Width           =   2775
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Service History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1410
            TabIndex        =   18
            Top             =   1335
            Width           =   2385
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Sales History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   1410
            TabIndex        =   17
            Top             =   585
            Width           =   2190
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageMainModules 
         Height          =   5355
         Left            =   30
         TabIndex        =   1
         Top             =   600
         Width           =   10125
         _Version        =   655364
         _ExtentX        =   17859
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.Timer Timer1 
            Left            =   5370
            Top             =   4740
         End
         Begin VB.CommandButton Command12 
            Height          =   645
            Left            =   5715
            MouseIcon       =   "MainMenu.frx":BD42
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "MainMenu.frx":BE94
            Style           =   1  'Graphical
            TabIndex        =   54
            Tag             =   "1102"
            ToolTipText     =   "View Sales Calculator"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdPDIInspectionList 
            Height          =   645
            Left            =   270
            MouseIcon       =   "MainMenu.frx":C313
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":C465
            Style           =   1  'Graphical
            TabIndex        =   52
            Tag             =   "1233"
            ToolTipText     =   "PDI Inspection List"
            Top             =   3450
            Width           =   720
         End
         Begin VB.CommandButton Command3 
            Height          =   645
            Left            =   5715
            MouseIcon       =   "MainMenu.frx":CAB4
            MousePointer    =   99  'Custom
            OLEDropMode     =   1  'Manual
            Picture         =   "MainMenu.frx":CC06
            Style           =   1  'Graphical
            TabIndex        =   25
            Tag             =   "1102"
            ToolTipText     =   "View Sales Calculator"
            Top             =   2685
            Width           =   720
         End
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   5715
            MouseIcon       =   "MainMenu.frx":D2A5
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":D3F7
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "View Customers"
            Top             =   1125
            Width           =   720
         End
         Begin VB.CommandButton Command10 
            Height          =   645
            Left            =   5715
            MouseIcon       =   "MainMenu.frx":DA92
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":DBE4
            Style           =   1  'Graphical
            TabIndex        =   11
            Tag             =   "1102"
            ToolTipText     =   "View Reminders/Tasks"
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton cmdConfiential201 
            Height          =   645
            Left            =   270
            MouseIcon       =   "MainMenu.frx":E2AC
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":E3FE
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "View Service Appointments"
            Top             =   1890
            Width           =   720
         End
         Begin VB.CommandButton cmdContract 
            Height          =   645
            Left            =   285
            MouseIcon       =   "MainMenu.frx":EB1D
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":EC6F
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Vehicles Catalogue/Brochure  "
            Top             =   360
            Width           =   720
         End
         Begin VB.CommandButton cmdAllowance 
            Height          =   645
            Left            =   285
            MouseIcon       =   "MainMenu.frx":F323
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":F475
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Vehicles Inventory Information"
            Top             =   1155
            Width           =   720
         End
         Begin VB.CommandButton cmdEmployeeAttendance 
            Height          =   645
            Left            =   285
            MouseIcon       =   "MainMenu.frx":FA94
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":FBE6
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "View Sales Appointments"
            Top             =   2685
            Width           =   720
         End
         Begin VB.Label LABDSA 
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5700
            TabIndex        =   56
            Top             =   4140
            Width           =   4125
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Possible Duplicate Customer"
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
            Height          =   270
            Left            =   6525
            TabIndex        =   55
            Top             =   2070
            Width           =   3075
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Vehicle Information"
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
            Height          =   270
            Index           =   2
            Left            =   1140
            TabIndex        =   53
            Top             =   3690
            Width           =   3135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Calculator"
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
            Height          =   270
            Left            =   6510
            TabIndex        =   26
            Top             =   2872
            Width           =   1725
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customers"
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
            Height          =   270
            Left            =   6525
            TabIndex        =   24
            Top             =   1290
            Width           =   1155
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Reminders/Tasks"
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
            Left            =   6525
            TabIndex        =   12
            Top             =   510
            Width           =   2700
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicles Catalogue/ Brochure  "
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
            Height          =   270
            Left            =   1110
            TabIndex        =   6
            Top             =   555
            Width           =   3285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicles Inventory Information"
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
            Height          =   270
            Left            =   1110
            TabIndex        =   5
            Top             =   1350
            Width           =   3225
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Service Appointments"
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
            Height          =   270
            Index           =   0
            Left            =   1110
            TabIndex        =   4
            Top             =   2100
            Width           =   2340
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Appointments"
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
            Height          =   270
            Left            =   1110
            TabIndex        =   3
            Top             =   2872
            Width           =   2100
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   5355
         Left            =   -69970
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   10125
         _Version        =   655364
         _ExtentX        =   17859
         _ExtentY        =   9446
         _StockProps     =   0
         Begin VB.CommandButton Command16 
            Height          =   645
            Left            =   420
            MouseIcon       =   "MainMenu.frx":102EF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":10441
            Style           =   1  'Graphical
            TabIndex        =   66
            Tag             =   "1088"
            ToolTipText     =   "Customer Log"
            Top             =   945
            Width           =   720
         End
         Begin VB.CommandButton Command15 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":109CD
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":10B1F
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Monthly Customer Directory By Customer Type"
            Top             =   3138
            Width           =   720
         End
         Begin VB.CommandButton Command14 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":1128E
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":113E0
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Ranged Customer Directory By Customer Type"
            Top             =   2406
            Width           =   720
         End
         Begin VB.CommandButton cmdMarketting_BirthDay 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":11CDA
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":11E2C
            Style           =   1  'Graphical
            TabIndex        =   59
            Tag             =   "1160"
            ToolTipText     =   "Birthday Celebrant of the Month"
            Top             =   4605
            Width           =   720
         End
         Begin VB.CommandButton Command13 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":129E4
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":12B36
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Yearly Customer Directory by Customer Type"
            Top             =   3870
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerwithInsurance 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":132CA
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":1341C
            Style           =   1  'Graphical
            TabIndex        =   57
            Tag             =   "1158"
            ToolTipText     =   "Customer With Insurance Policies"
            Top             =   4605
            Width           =   720
         End
         Begin VB.CommandButton Command11 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":13DF9
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":13F4B
            Style           =   1  'Graphical
            TabIndex        =   50
            Tag             =   "1088"
            ToolTipText     =   "View Sales Appointment"
            Top             =   1674
            Width           =   720
         End
         Begin VB.CommandButton Command9 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":14654
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":147A6
            Style           =   1  'Graphical
            TabIndex        =   48
            Tag             =   "1088"
            ToolTipText     =   "View Sales Appointment"
            Top             =   942
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerLog 
            Height          =   645
            Left            =   435
            MouseIcon       =   "MainMenu.frx":14EAF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":15001
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "1088"
            ToolTipText     =   "View Customer Log"
            Top             =   210
            Width           =   720
         End
         Begin VB.CommandButton cmdSalesAppointment 
            Height          =   645
            Left            =   405
            MouseIcon       =   "MainMenu.frx":1558D
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":156DF
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "1088"
            ToolTipText     =   "View Sales Appointment"
            Top             =   1680
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomer_Reminder_And_Task 
            Height          =   645
            Left            =   405
            MouseIcon       =   "MainMenu.frx":15DE8
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":15F3A
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "1088"
            ToolTipText     =   "View Customer Reminders And Tasks"
            Top             =   3135
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomer_Reminder_And_Task_Internal 
            Height          =   645
            Left            =   405
            MouseIcon       =   "MainMenu.frx":165F3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":16745
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "1088"
            ToolTipText     =   "View Internal Reminders"
            Top             =   3870
            Width           =   720
         End
         Begin VB.CommandButton cmdServiceAppointment 
            Height          =   645
            Left            =   405
            MouseIcon       =   "MainMenu.frx":16DFE
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":16F50
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "1088"
            ToolTipText     =   "View Service Appointment"
            Top             =   2400
            Width           =   720
         End
         Begin VB.CommandButton cmdCustomerInfoReport 
            Height          =   645
            Left            =   4500
            MouseIcon       =   "MainMenu.frx":17659
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":177AB
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "1088"
            ToolTipText     =   "View Customer Information Report"
            Top             =   210
            Width           =   720
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
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   1245
            TabIndex        =   73
            Top             =   1155
            Width           =   2085
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   72
            Top             =   4815
            Width           =   2670
         End
         Begin VB.Label Label24 
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   71
            Top             =   2640
            Width           =   1770
         End
         Begin VB.Label Label22 
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   70
            Top             =   4140
            Width           =   1620
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   69
            Top             =   3390
            Width           =   2760
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   68
            Top             =   1890
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1245
            TabIndex        =   67
            Top             =   405
            Width           =   1200
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
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   5280
            TabIndex        =   65
            Top             =   3374
            Width           =   4635
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
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   5280
            TabIndex        =   64
            Top             =   2656
            Width           =   4590
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
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   5295
            TabIndex        =   63
            Top             =   4092
            Width           =   4500
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   5355
            TabIndex        =   62
            Top             =   4815
            Width           =   2895
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Service Customer"
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
            Height          =   240
            Left            =   5325
            TabIndex        =   51
            Top             =   1908
            Width           =   2490
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Monthly Sales Customer"
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
            Height          =   240
            Left            =   5310
            TabIndex        =   49
            Top             =   1160
            Width           =   2295
         End
         Begin VB.Label Label25 
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
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   5325
            TabIndex        =   47
            Top             =   412
            Width           =   2745
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
Dim WithEvents FormSearch               As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Dim LOGACTION                           As String
'Upating Code       : AXP-0713200715:16

Private Sub cmdAction_Click()
    If Module_Access(LOGID, "CUSTOMERS LOG INQUIRY", "INQUIRY") = False Then Exit Sub
    On Error GoTo Errorcode:
    Call FormSearch.SearchForProspects(vbNullString)
    LOGACTION = "PROS:LOGINQ"
    FormSearch.Show 1
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdAdjustments_Click()
    If Module_Access(LOGID, "CUSTOMER TRANSACTION HISTORY", "INQUIRY") = False Then Exit Sub
    frmCRIS_Inquiry_CustomerTransHistory.Show
End Sub

Private Sub cmdAllowance_Click()
    If Module_Access(LOGID, "VEHICLE INVENTORY INFORMATION", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11

    frmSMIS_Inquiry_VehicleMaster.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdCommission_Click()
    If Module_Access(LOGID, "CUSTOMER SERVICE HISTORY", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmCSMSCustomerHistory.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdConfiential201_Click()
    If Module_Access(LOGID, "APPOINTMENT", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    '    frmCRIS_Inquiry_ServiceAppointment.Show
    frmCSMSAppointment.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdContract_Click()
    If Module_Access(LOGID, "VEHICLE CATALOGUE", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmCRIS_InquiryVehicleCatalogue.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdCustomer_Reminder_And_Task_Click()
    If Module_Access(LOGID, "CUSTOMER REMINDERS/TASK", "REPORTS") = False Then Exit Sub

    frmCRIS_Report_CustomerRemindersAndTask.Show
End Sub

Private Sub cmdCustomer_Reminder_And_Task_Internal_Click()
    If Module_Access(LOGID, "INTERNAL REMINDERS/TASKS", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_InternalReminder.Show
End Sub

Private Sub cmdCustomerInfoReport_Click()
    If Module_Access(LOGID, "CUSTOMER INFORMATION REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_CustomerInformationReport.Show
End Sub

Private Sub cmdCustomerLog_Click()
    If Module_Access(LOGID, "CUSTOMER LOG REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_Log.Show
End Sub

Private Sub cmdCustomerwithInsurance_Click()
    If Module_Access(LOGID, "CUSTOMERS WITH INSURANCE POLICIES", "REPORTS") = False Then Exit Sub
    CUST_REPT_TYPE = "2"
    frmSMIS_Report_CustSummary.Show
    frmSMIS_Report_CustSummary.ZOrder 0
End Sub

Private Sub cmdDeductions_Click()
    If Module_Access(LOGID, "CUSTOMER SALES HISTORY", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11
    
    If FormExist("frmSMIS_Inquiry_CustomerSalesHistory") Then
        Unload frmSMIS_Inquiry_CustomerSalesHistory
    End If
    frmSMIS_Inquiry_CustomerSalesHistory.Show
    frmSMIS_Inquiry_CustomerSalesHistory.MyTab.SelectedItem = 0
    Screen.MousePointer = 0
End Sub

Private Sub cmdEmployeeAttendance_Click()
    If Module_Access(LOGID, "SALES APPOINTMENT", "INQUIRY") = False Then Exit Sub
    frmCRIS_Inquiry_SalesAppointment.Show
End Sub

Private Sub cmdGenPay_Click()
    If Module_Access(LOGID, "CUSTOMER CALL/VISIT HISTORY", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmSMIS_Inquiry_CallVisit_History.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdMarketting_BirthDay_Click()
    If Module_Access(LOGID, "BIRTHDAY CELEBRANTS", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_CustomerBDays.Show
    frmSMIS_Report_CustomerBDays.ZOrder 0
End Sub

Private Sub cmdPDIInspectionList_Click()
    If Module_Access(LOGID, "CUSTOMER VEHICLE INFORMATION", "DATA ENTRY") = False Then Exit Sub
    frmCSMSSearchCustomerVehicle.Show
End Sub

Private Sub cmdSalesAppointment_Click()
    If Module_Access(LOGID, "SALES APPOINTMENT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_SalesAppointment.Show
End Sub

Private Sub cmdServiceAppointment_Click()
    If Module_Access(LOGID, "SERVICE APPOINTMENT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_ServiceAppointment.Show

End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "CUSTOMER VEHICLE INFORMATION DETAIL", "INQUIRY") = False Then Exit Sub
    Screen.MousePointer = 11

    If FormExist("frmSMIS_Inquiry_CustomerSalesHistory") Then
        Unload frmSMIS_Inquiry_CustomerSalesHistory
    End If
    frmSMIS_Inquiry_CustomerSalesHistory.Show
    frmSMIS_Inquiry_CustomerSalesHistory.MyTab.SelectedItem = 1
    Screen.MousePointer = 0
End Sub

Private Sub Command10_Click()
    If Module_Access(LOGID, "REMINDER/TASK", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmSMIS_Log_CustomerReminder.AddReminder ("C")
    frmSMIS_Log_CustomerReminder.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command11_Click()
    If Module_Access(LOGID, "MONTHLY SERVICE CUSTOMER", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_AfterSales.ServiceReport
    frmCRIS_Report_AfterSales.Show
End Sub

Private Sub Command12_Click()
    If Module_Access(LOGID, "DUPLICATE CUSTOMER", "SYSTEM") = False Then Exit Sub
    frmCRIS_Inquiry_PossibleDuplicates.Show
End Sub

Private Sub Command13_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchYear
    frmSMIS_Report_AfterSales.Show
End Sub

Private Sub Command14_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchRange
    frmSMIS_Report_AfterSales.Show
    frmSMIS_Report_AfterSales.ZOrder 0
End Sub

Private Sub Command15_Click()
    If Module_Access(LOGID, "AFTER SALES", "REPORTS") = False Then Exit Sub
    frmSMIS_Report_AfterSales.SearchMonth
    frmSMIS_Report_AfterSales.Show
End Sub

Private Sub Command16_Click()
    If Module_Access(LOGID, "CUSTOMER LOG REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_Log.Show
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "CUSTOMER", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmAllCustomer.Show
    frmAllCustomer.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "SALES CALCULATOR", "SYSTEM") = False Then Exit Sub
    Screen.MousePointer = 11
    frmSMIS_Mis_AOR.ShowonlyComputation
    frmSMIS_Mis_AOR.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command4_Click()
    If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:CALL"
    FormSearch.Show 1
End Sub

Private Sub Command5_Click()
    If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:VISIT"
    FormSearch.Show 1
End Sub

Private Sub Command6_Click()
    If Module_Access(LOGID, "CUSTOMER REMINDERS/TASK", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmCRIS_Inquiry_TaskList.ShowTaskType ("C")
    frmCRIS_Inquiry_TaskList.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:LETTER"
    FormSearch.Show 1
End Sub

Private Sub Command8_Click()
    If Module_Access(LOGID, "CUSTOMERS LOGS", "DATA ENTRY") = False Then Exit Sub
    Call FormSearch.SearchForCustomers
    LOGACTION = "CUS:EMAIL"
    FormSearch.Show 1
End Sub

Private Sub Command9_Click()
    If Module_Access(LOGID, "MONTHLY SALES REPORT", "REPORTS") = False Then Exit Sub
    frmCRIS_Report_AfterSales.SalesReport
    frmCRIS_Report_AfterSales.Show

End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    TabControl1.Icons = frmMain.CommandBars1.Icons
    TabControl1.SelectedItem = 0
    CenterMe frmMain, Me, 1
    Set FormSearch = New frmSMIS_Mis_SearchMaster
    Screen.MousePointer = 0
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    
    Select Case LOGACTION
        Case "CUS:LETTER"
            Call frmCRIS_Log_Letter.AddLetter(0, oCusRs!CUSCDE)
            frmCRIS_Log_Letter.Show
        Case "CUS:VISIT"
            Call frmCRIS_Log_Visits.AddVisit(0, oCusRs!CUSCDE)
            frmCRIS_Log_Visits.Show
        Case "CUS:CALL"
            Call frmCRIS_Log_Call.AddCall(0, oCusRs!CUSCDE)
            frmCRIS_Log_Call.Show
        Case "CUS:EMAIL"
            Call frmCRIS_Log_Email.AddEmail(0, oCusRs!CUSCDE)
            frmCRIS_Log_Email.Show
        Case "CUS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWCUSTOMERLOG(oCusRs!CUSCDE, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
        Case "PROS:LOGINQ"
            Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(oCusRs!ProspectID, oCusRs!AcctName)
            frmSMIS_Inquiry_ViewLog.Show
            
    End Select
    
End Sub

Private Sub Label10_Click()
    OpenCDRoom
End Sub

Private Sub Timer1_Timer()
    If LABDSA.Caption <> "" Then
        If LABDSA.Visible = True Then
            LABDSA.Visible = False
        Else
            LABDSA.Visible = True
        End If
    End If
End Sub

