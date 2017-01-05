VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HRMS Main Menu"
   ClientHeight    =   6630
   ClientLeft      =   990
   ClientTop       =   1170
   ClientWidth     =   10350
   ForeColor       =   &H8000000F&
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   10350
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _Version        =   655364
      _ExtentX        =   18230
      _ExtentY        =   11721
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
      SelectedItem    =   2
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "tbPageMainModules"
      Item(1).Caption =   "Tables && File Maintenance"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tbPageTable"
      Item(2).Caption =   "Reports"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tbPageFileMaintenance"
      Item(3).Caption =   "Other Setups"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "tbPageReport"
      Item(3).Control(1)=   "tbPageOtherSetup"
      Begin XtremeSuiteControls.TabControlPage tbPageOtherSetup 
         Height          =   6015
         Left            =   -69970
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   10275
         _Version        =   655364
         _ExtentX        =   18124
         _ExtentY        =   10610
         _StockProps     =   0
         Begin VB.CommandButton cmdFileCompanyProfile 
            Height          =   675
            Left            =   300
            MouseIcon       =   "MainMenu.frx":6852
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":69A4
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "View Company Profile"
            Top             =   210
            Width           =   735
         End
         Begin VB.CommandButton Command29 
            Height          =   645
            Left            =   300
            MouseIcon       =   "MainMenu.frx":6E93
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":6FE5
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Password Maintenance "
            Top             =   1710
            Width           =   720
         End
         Begin VB.CommandButton Command27 
            Height          =   645
            Left            =   300
            MouseIcon       =   "MainMenu.frx":8967
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":8AB9
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Signatories and Headers"
            Top             =   990
            Width           =   720
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Profile/Pay Period Set Up (F12)"
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
            Height          =   285
            Left            =   1140
            TabIndex        =   35
            Top             =   420
            Width           =   4665
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password Maintenance "
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
            Height          =   285
            Left            =   1170
            TabIndex        =   24
            Top             =   1905
            Width           =   2730
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Signatories and Headers"
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
            Height          =   285
            Left            =   1170
            TabIndex        =   17
            Top             =   1185
            Width           =   2820
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageReport 
         Height          =   6015
         Left            =   -69970
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   10275
         _Version        =   655364
         _ExtentX        =   18124
         _ExtentY        =   10610
         _StockProps     =   0
      End
      Begin XtremeSuiteControls.TabControlPage tbPageFileMaintenance 
         Height          =   6015
         Left            =   30
         TabIndex        =   3
         Top             =   600
         Width           =   10275
         _Version        =   655364
         _ExtentX        =   18124
         _ExtentY        =   10610
         _StockProps     =   0
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   6015
            Left            =   7080
            ScaleHeight     =   5985
            ScaleWidth      =   3165
            TabIndex        =   95
            Top             =   0
            Width           =   3195
            Begin VB.CommandButton Command15 
               Caption         =   "Create Diskette Layout"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1245
               Left            =   570
               Picture         =   "MainMenu.frx":8EFB
               Style           =   1  'Graphical
               TabIndex        =   96
               ToolTipText     =   "Create Diskette Layout"
               Top             =   1380
               Width           =   1815
            End
            Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
               Height          =   435
               Left            =   0
               TabIndex        =   104
               Top             =   0
               Width           =   3195
               _Version        =   655364
               _ExtentX        =   5636
               _ExtentY        =   767
               _StockProps     =   14
               Caption         =   "Generate Diskette"
               ForeColor       =   8388608
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               GradientColorLight=   16777215
               GradientColorDark=   16744576
               ForeColor       =   8388608
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Diskette Layout Generator"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   855
               Left            =   240
               TabIndex        =   103
               Top             =   480
               Width           =   2505
            End
            Begin VB.Label Label56 
               BackStyle       =   0  'Transparent
               Caption         =   "For:"
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
               Height          =   345
               Left            =   90
               TabIndex        =   102
               Top             =   2700
               Width           =   645
            End
            Begin VB.Label Label57 
               BackStyle       =   0  'Transparent
               Caption         =   "*  SSS"
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
               Height          =   285
               Left            =   630
               TabIndex        =   101
               Top             =   3030
               Width           =   1545
            End
            Begin VB.Label Label58 
               BackStyle       =   0  'Transparent
               Caption         =   "*  SSS Loan"
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
               Height          =   285
               Left            =   630
               TabIndex        =   100
               Top             =   3330
               Width           =   1545
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "*  PhilHealth"
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
               Height          =   285
               Left            =   630
               TabIndex        =   99
               Top             =   3660
               Width           =   1785
            End
            Begin VB.Label Label60 
               BackStyle       =   0  'Transparent
               Caption         =   "*  Pag-ibig"
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
               Height          =   285
               Left            =   630
               TabIndex        =   98
               Top             =   3990
               Width           =   1545
            End
            Begin VB.Label Label61 
               BackStyle       =   0  'Transparent
               Caption         =   "*  Pag-ibig Loan"
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
               Height          =   285
               Left            =   630
               TabIndex        =   97
               Top             =   4320
               Width           =   2265
            End
         End
         Begin XtremeSuiteControls.TabControl TabControl2 
            Height          =   6945
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   7095
            _Version        =   655364
            _ExtentX        =   12515
            _ExtentY        =   12250
            _StockProps     =   64
            Appearance      =   1
            Color           =   16
            PaintManager.BoldSelected=   -1  'True
            PaintManager.DisableLunaColors=   0   'False
            PaintManager.HotTracking=   -1  'True
            PaintManager.ShowIcons=   -1  'True
            PaintManager.LargeIcons=   -1  'True
            ItemCount       =   5
            SelectedItem    =   2
            Item(0).Caption =   "Every Payroll"
            Item(0).ControlCount=   1
            Item(0).Control(0)=   "tbPageEveryPayroll"
            Item(1).Caption =   "Monthly Reports"
            Item(1).ControlCount=   1
            Item(1).Control(0)=   "tbPageMonthlyReports"
            Item(2).Caption =   "Quarterly Reports"
            Item(2).ControlCount=   1
            Item(2).Control(0)=   "tbPageQuaterlyReports"
            Item(3).Caption =   "Other Reports"
            Item(3).ControlCount=   1
            Item(3).Control(0)=   "tbPageYearlyReports"
            Item(4).Caption =   "Schedules"
            Item(4).ControlCount=   1
            Item(4).Control(0)=   "TabControlPage1"
            Begin XtremeSuiteControls.TabControlPage TabControlPage1 
               Height          =   6315
               Left            =   -69970
               TabIndex        =   146
               Top             =   600
               Visible         =   0   'False
               Width           =   7035
               _Version        =   655364
               _ExtentX        =   12409
               _ExtentY        =   11139
               _StockProps     =   0
               Begin VB.CommandButton cmdReport_Sched_TAXDUE 
                  Height          =   645
                  Left            =   3780
                  MouseIcon       =   "MainMenu.frx":933D
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":948F
                  Style           =   1  'Graphical
                  TabIndex        =   166
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   3390
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_13MONTH 
                  Height          =   645
                  Left            =   3780
                  MouseIcon       =   "MainMenu.frx":98F1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":9A43
                  Style           =   1  'Graphical
                  TabIndex        =   165
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   2625
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_COMMISSIONTAX 
                  Height          =   645
                  Left            =   3780
                  MouseIcon       =   "MainMenu.frx":9EA5
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":9FF7
                  Style           =   1  'Graphical
                  TabIndex        =   164
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   1875
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_PAYROLL 
                  Height          =   645
                  Left            =   3780
                  MouseIcon       =   "MainMenu.frx":A459
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":A5AB
                  Style           =   1  'Graphical
                  TabIndex        =   163
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   360
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_COMMISSION 
                  Height          =   645
                  Left            =   3780
                  MouseIcon       =   "MainMenu.frx":AA0D
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":AB5F
                  Style           =   1  'Graphical
                  TabIndex        =   162
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   1110
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_OVERTIMEPAY 
                  Height          =   645
                  Left            =   315
                  MouseIcon       =   "MainMenu.frx":AFC1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":B113
                  Style           =   1  'Graphical
                  TabIndex        =   161
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   3390
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_TAXWHELD 
                  Height          =   645
                  Left            =   315
                  MouseIcon       =   "MainMenu.frx":B575
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":B6C7
                  Style           =   1  'Graphical
                  TabIndex        =   160
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   2625
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_PAGIBIG 
                  Height          =   645
                  Left            =   315
                  MouseIcon       =   "MainMenu.frx":BB29
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":BC7B
                  Style           =   1  'Graphical
                  TabIndex        =   159
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   1875
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_PHIC 
                  Height          =   645
                  Left            =   315
                  MouseIcon       =   "MainMenu.frx":C0DD
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":C22F
                  Style           =   1  'Graphical
                  TabIndex        =   158
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   1110
                  Width           =   720
               End
               Begin VB.CommandButton cmdReport_Sched_SSS 
                  Height          =   645
                  Left            =   315
                  MouseIcon       =   "MainMenu.frx":C691
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":C7E3
                  Style           =   1  'Graphical
                  TabIndex        =   157
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   360
                  Width           =   720
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Tax Due/Refund"
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
                  Height          =   540
                  Index           =   9
                  Left            =   4590
                  TabIndex        =   156
                  Top             =   3442
                  Width           =   1560
               End
               Begin VB.Label labSched 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "13th Month Pay Schedule"
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
                  Index           =   8
                  Left            =   4620
                  TabIndex        =   155
                  Top             =   2827
                  Width           =   2415
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Commission Tax"
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
                  Height          =   540
                  Index           =   7
                  Left            =   4590
                  TabIndex        =   154
                  Top             =   1927
                  Width           =   2385
               End
               Begin VB.Label labSched 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Commission"
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
                  Index           =   6
                  Left            =   4590
                  TabIndex        =   153
                  Top             =   1312
                  Width           =   2355
               End
               Begin VB.Label labSched 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Payroll"
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
                  Index           =   5
                  Left            =   4590
                  TabIndex        =   152
                  Top             =   562
                  Width           =   1875
               End
               Begin VB.Label labSched 
                  Caption         =   "Schedule of SSS Premium Contribution"
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
                  Height          =   510
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   151
                  Top             =   420
                  Width           =   2565
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Philhealth Premium Contribution"
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
                  Height          =   510
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   150
                  Top             =   1170
                  Width           =   2565
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Pag-Ibig Premium Contribution"
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
                  Height          =   510
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   149
                  Top             =   1935
                  Width           =   2565
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Tax Withheld"
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
                  Height          =   300
                  Index           =   3
                  Left            =   1080
                  TabIndex        =   148
                  Top             =   2790
                  Width           =   2565
               End
               Begin VB.Label labSched 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Schedule of Overtime Pay"
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
                  Height          =   330
                  Index           =   4
                  Left            =   1110
                  TabIndex        =   147
                  Top             =   3540
                  Width           =   2565
               End
            End
            Begin XtremeSuiteControls.TabControlPage tbPageEveryPayroll 
               Height          =   6315
               Left            =   -69970
               TabIndex        =   53
               Top             =   600
               Visible         =   0   'False
               Width           =   7035
               _Version        =   655364
               _ExtentX        =   12409
               _ExtentY        =   11139
               _StockProps     =   0
               Begin VB.CommandButton Command36 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":CC45
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":CD97
                  Style           =   1  'Graphical
                  TabIndex        =   60
                  ToolTipText     =   "Print Commission Breakdown"
                  Top             =   3780
                  Width           =   720
               End
               Begin VB.CommandButton Command20 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":D1F9
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":D34B
                  Style           =   1  'Graphical
                  TabIndex        =   59
                  ToolTipText     =   "Print Adjustment Breakdown"
                  Top             =   3060
                  Width           =   720
               End
               Begin VB.CommandButton cmdOTBreakDown 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":D7AD
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":D8FF
                  Style           =   1  'Graphical
                  TabIndex        =   58
                  ToolTipText     =   "Print Overtime Breakdown "
                  Top             =   2340
                  Width           =   720
               End
               Begin VB.CommandButton Command18 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":DD61
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":DEB3
                  Style           =   1  'Graphical
                  TabIndex        =   57
                  ToolTipText     =   "Print Deductions Breakdown"
                  Top             =   1605
                  Width           =   720
               End
               Begin VB.CommandButton Command17 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":E315
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":E467
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  ToolTipText     =   "Print Payroll PaySlips"
                  Top             =   885
                  Width           =   720
               End
               Begin VB.CommandButton cmdEV_ATMAdvice 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":E8C9
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":EA1B
                  Style           =   1  'Graphical
                  TabIndex        =   55
                  ToolTipText     =   "Print Payroll ATM Advice"
                  Top             =   165
                  Width           =   720
               End
               Begin VB.CommandButton cmdEV_DTRSummary 
                  Height          =   645
                  Left            =   510
                  MouseIcon       =   "MainMenu.frx":EE7D
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":EFCF
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  ToolTipText     =   "Print Commission Breakdown"
                  Top             =   4530
                  Width           =   720
               End
               Begin VB.Label Label82 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Commission Breakdown"
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   67
                  Top             =   3960
                  Width           =   3675
               End
               Begin VB.Label Label66 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Adjustment Breakdown"
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   66
                  Top             =   3255
                  Width           =   3675
               End
               Begin VB.Label Label65 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Overtime Breakdown "
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   65
                  Top             =   2505
                  Width           =   3225
               End
               Begin VB.Label Label64 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Deductions Breakdown"
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   64
                  Top             =   1800
                  Width           =   3855
               End
               Begin VB.Label Label63 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Payroll PaySlips"
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   63
                  Top             =   1080
                  Width           =   3225
               End
               Begin VB.Label Label62 
                  BackColor       =   &H80000009&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Payroll ATM Advice"
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
                  Height          =   375
                  Left            =   1350
                  TabIndex        =   62
                  Top             =   375
                  Width           =   3225
               End
               Begin VB.Label Label14 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Print Daily Time Record"
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
                  Height          =   375
                  Left            =   1335
                  TabIndex        =   61
                  Top             =   4710
                  Width           =   3675
               End
            End
            Begin XtremeSuiteControls.TabControlPage tbPageYearlyReports 
               Height          =   6315
               Left            =   -69970
               TabIndex        =   68
               Top             =   600
               Visible         =   0   'False
               Width           =   7035
               _Version        =   655364
               _ExtentX        =   12409
               _ExtentY        =   11139
               _StockProps     =   0
               Begin Crystal.CrystalReport rptReports 
                  Left            =   5940
                  Top             =   2790
                  _ExtentX        =   741
                  _ExtentY        =   741
                  _Version        =   348160
                  PrintFileLinesPerPage=   60
               End
               Begin VB.CommandButton cmd_ReportBlankEmployye 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":F431
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":F583
                  Style           =   1  'Graphical
                  TabIndex        =   173
                  ToolTipText     =   "Other Reports"
                  Top             =   60
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_YearToDate 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":F9E5
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":FB37
                  Style           =   1  'Graphical
                  TabIndex        =   171
                  ToolTipText     =   "Other Reports"
                  Top             =   4560
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_ResingedEmployee 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":FF99
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":100EB
                  Style           =   1  'Graphical
                  TabIndex        =   169
                  ToolTipText     =   "Other Reports"
                  Top             =   3918
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_OtherEmployeeListing 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":1054D
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":1069F
                  Style           =   1  'Graphical
                  TabIndex        =   167
                  ToolTipText     =   "Other Reports"
                  Top             =   3275
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_EmployeeList 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":10B01
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":10C53
                  Style           =   1  'Graphical
                  TabIndex        =   72
                  ToolTipText     =   "Other Reports"
                  Top             =   2632
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_AlphalistwioutPrevEmployer 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":110B5
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":11207
                  Style           =   1  'Graphical
                  TabIndex        =   71
                  ToolTipText     =   "Year to Date Reports"
                  Top             =   1989
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_ALphalistWPrevEmployer 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":11669
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":117BB
                  Style           =   1  'Graphical
                  TabIndex        =   70
                  ToolTipText     =   "Form 2316 && Alphalist"
                  Top             =   1346
                  Width           =   720
               End
               Begin VB.CommandButton cmdReports_AlphaListTerminated 
                  Height          =   645
                  Left            =   600
                  MouseIcon       =   "MainMenu.frx":11C1D
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":11D6F
                  Style           =   1  'Graphical
                  TabIndex        =   69
                  ToolTipText     =   "13th Month Pay and Others"
                  Top             =   703
                  Width           =   720
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Blank Employee Information  Sheet "
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   174
                  Top             =   180
                  Width           =   4065
               End
               Begin VB.Label Label37 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Year-To-Date Details"
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   172
                  Top             =   4770
                  Width           =   2325
               End
               Begin VB.Label Label36 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   " Resigned Employee Listing"
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   170
                  Top             =   4110
                  Width           =   3195
               End
               Begin VB.Label Label35 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Other Employee Listing Wage"
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   168
                  Top             =   3450
                  Width           =   3375
               End
               Begin VB.Label Label81 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Employee Listing "
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   76
                  Top             =   2805
                  Width           =   2040
               End
               Begin VB.Label Label79 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alpha List W/Out Previous Employee"
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   75
                  Top             =   2145
                  Width           =   4170
               End
               Begin VB.Label Label78 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alpha List With Previous Employer"
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   74
                  Top             =   1485
                  Width           =   3945
               End
               Begin VB.Label Label77 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Alpha List Terminated "
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
                  Height          =   285
                  Left            =   1470
                  TabIndex        =   73
                  Top             =   840
                  Width           =   2790
               End
            End
            Begin XtremeSuiteControls.TabControlPage tbPageQuaterlyReports 
               Height          =   6315
               Left            =   30
               TabIndex        =   77
               Top             =   600
               Width           =   7035
               _Version        =   655364
               _ExtentX        =   12409
               _ExtentY        =   11139
               _StockProps     =   0
               Begin VB.CommandButton Command37 
                  Height          =   645
                  Left            =   540
                  MouseIcon       =   "MainMenu.frx":121D1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":12323
                  Style           =   1  'Graphical
                  TabIndex        =   80
                  ToolTipText     =   "Quarterly Loans"
                  Top             =   2100
                  Width           =   720
               End
               Begin VB.CommandButton Command30 
                  Height          =   645
                  Left            =   540
                  MouseIcon       =   "MainMenu.frx":140AD
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":141FF
                  Style           =   1  'Graphical
                  TabIndex        =   79
                  ToolTipText     =   "R3 - SSS Contribution Collection List"
                  Top             =   1290
                  Width           =   720
               End
               Begin VB.CommandButton Command26 
                  Height          =   645
                  Left            =   540
                  MouseIcon       =   "MainMenu.frx":15D31
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":15E83
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  ToolTipText     =   "RF1 - Employer Quarterly Remittance"
                  Top             =   510
                  Width           =   720
               End
               Begin VB.Label Label83 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Quarterly Loans"
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
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   83
                  Top             =   2310
                  Width           =   1830
               End
               Begin VB.Label Label76 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "R3 - SSS Contribution Collection List"
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
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   82
                  Top             =   1500
                  Width           =   4200
               End
               Begin VB.Label Label72 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "RF1 - Employer Quarterly Remittance"
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
                  Height          =   285
                  Left            =   1410
                  TabIndex        =   81
                  Top             =   720
                  Width           =   4215
               End
            End
            Begin XtremeSuiteControls.TabControlPage tbPageMonthlyReports 
               Height          =   6315
               Left            =   -69970
               TabIndex        =   84
               Top             =   600
               Visible         =   0   'False
               Width           =   7035
               _Version        =   655364
               _ExtentX        =   12409
               _ExtentY        =   11139
               _StockProps     =   0
               Begin VB.CommandButton Command24 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":17161
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":172B3
                  Style           =   1  'Graphical
                  TabIndex        =   89
                  ToolTipText     =   "PAG-IBIG Contributions"
                  Top             =   1911
                  Width           =   720
               End
               Begin VB.CommandButton Command23 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":1BCD1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":1BE23
                  Style           =   1  'Graphical
                  TabIndex        =   88
                  ToolTipText     =   "SSS Contributions"
                  Top             =   1173
                  Width           =   720
               End
               Begin VB.CommandButton Command22 
                  Height          =   645
                  Left            =   540
                  MouseIcon       =   "MainMenu.frx":1D955
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":1DAA7
                  Style           =   1  'Graphical
                  TabIndex        =   87
                  ToolTipText     =   "TAX Withheld"
                  Top             =   2649
                  Width           =   720
               End
               Begin VB.CommandButton Command21 
                  Height          =   645
                  Left            =   525
                  MouseIcon       =   "MainMenu.frx":1FFB1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":20103
                  Style           =   1  'Graphical
                  TabIndex        =   86
                  ToolTipText     =   "PHILHEALTH Contributions"
                  Top             =   435
                  Width           =   720
               End
               Begin VB.CommandButton Command25 
                  Height          =   645
                  Left            =   540
                  MouseIcon       =   "MainMenu.frx":213E1
                  MousePointer    =   99  'Custom
                  Picture         =   "MainMenu.frx":21533
                  Style           =   1  'Graphical
                  TabIndex        =   85
                  ToolTipText     =   "LOANS Remitted"
                  Top             =   3390
                  Width           =   720
               End
               Begin VB.Label Label71 
                  BackStyle       =   0  'Transparent
                  Caption         =   "LOANS Remitted"
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
                  Height          =   375
                  Left            =   1410
                  TabIndex        =   94
                  Top             =   3585
                  Width           =   3225
               End
               Begin VB.Label Label70 
                  BackStyle       =   0  'Transparent
                  Caption         =   "TAX Withheld"
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
                  Height          =   375
                  Left            =   1410
                  TabIndex        =   93
                  Top             =   2865
                  Width           =   3225
               End
               Begin VB.Label Label69 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PAG-IBIG Contributions"
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
                  Height          =   375
                  Left            =   1410
                  TabIndex        =   92
                  Top             =   2100
                  Width           =   3225
               End
               Begin VB.Label Label68 
                  BackStyle       =   0  'Transparent
                  Caption         =   "SSS Contributions"
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
                  Height          =   375
                  Left            =   1410
                  TabIndex        =   91
                  Top             =   1365
                  Width           =   3225
               End
               Begin VB.Label Label67 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PHILHEALTH Contributions"
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
                  Height          =   375
                  Left            =   1410
                  TabIndex        =   90
                  Top             =   630
                  Width           =   3225
               End
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageTable 
         Height          =   6015
         Left            =   -69970
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   10275
         _Version        =   655364
         _ExtentX        =   18124
         _ExtentY        =   10610
         _StockProps     =   0
         Begin VB.CommandButton Command7 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":232BD
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2340F
            Style           =   1  'Graphical
            TabIndex        =   144
            ToolTipText     =   "View Beginning Balances"
            Top             =   5175
            Width           =   615
         End
         Begin VB.CommandButton cmdFileDepartmentalCodes 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":23719
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2386B
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "View Department Codes"
            Top             =   120
            Width           =   615
         End
         Begin VB.CommandButton cmdFileOverTimeCodes 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":23B75
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":23CC7
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "View Overtime Codes"
            Top             =   3765
            Width           =   615
         End
         Begin VB.CommandButton cmdFileDeductionCode 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":23FD1
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":24123
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "View Deduction Codes"
            Top             =   2325
            Width           =   615
         End
         Begin VB.CommandButton cmdFileSalaryGradeCode 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":2442D
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2457F
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "View Salary Grade Codes"
            Top             =   1620
            Width           =   615
         End
         Begin VB.CommandButton cmdFileLoanCode 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":24889
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":249DB
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "View Loan Codes"
            Top             =   3030
            Width           =   615
         End
         Begin VB.CommandButton cmdFileAdjustmentCode 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":26765
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":268B7
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "View Adjustment Code"
            Top             =   900
            Width           =   615
         End
         Begin VB.CommandButton Command2 
            Height          =   645
            Left            =   5130
            MouseIcon       =   "MainMenu.frx":26BC1
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":26D13
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "View Overtime Codes"
            Top             =   4470
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Height          =   645
            Left            =   7980
            MouseIcon       =   "MainMenu.frx":273BA
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2750C
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "View Overtime Module"
            Top             =   4425
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdFileWorkingDaysSetup 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":27A10
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":27B62
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "View Working Days Setup"
            Top             =   3030
            Width           =   615
         End
         Begin VB.CommandButton cmdShiftModule 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":27E6C
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":27FBE
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "View Overtime Module"
            Top             =   5175
            Width           =   615
         End
         Begin VB.CommandButton Command28 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":284C2
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":28614
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Schedule of Deductions"
            Top             =   4470
            Width           =   615
         End
         Begin VB.CommandButton cmdTableOverTime 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":28A56
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":28BA8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Overtime Setup"
            Top             =   3765
            Width           =   615
         End
         Begin VB.CommandButton cmdTablePagIbig 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":29022
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":29174
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "PAGIBIG Table"
            Top             =   1620
            Width           =   615
         End
         Begin VB.CommandButton cmdTablePhilHealth 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":2DB92
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2DCE4
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "PHILHEALTH Table"
            Top             =   885
            Width           =   615
         End
         Begin VB.CommandButton cmdTableWithHolding 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":2EFC2
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2F114
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "WITHHOLDING TAX Table"
            Top             =   2325
            Width           =   615
         End
         Begin VB.CommandButton cmdTableSSS 
            Height          =   645
            Left            =   390
            MouseIcon       =   "MainMenu.frx":3161E
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":31770
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "SSS Table"
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Holiday Setup"
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
            Height          =   285
            Left            =   5970
            TabIndex        =   145
            Top             =   5355
            Width           =   1605
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Department Codes"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   51
            Top             =   300
            Width           =   2145
         End
         Begin VB.Label Label84 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Codes"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   50
            Top             =   3945
            Width           =   1830
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deduction Codes"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   49
            Top             =   2505
            Width           =   1995
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Codes"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   48
            Top             =   1800
            Width           =   1530
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Codes"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   47
            Top             =   3210
            Width           =   1395
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment Code"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   46
            Top             =   1065
            Width           =   1980
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Shift Code"
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
            Height          =   285
            Left            =   5940
            TabIndex        =   45
            Top             =   4650
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Codes"
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
            Height          =   285
            Left            =   8790
            TabIndex        =   44
            Top             =   4605
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll Setup"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   33
            Top             =   3210
            Width           =   1530
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Shift Set Up"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   32
            Top             =   5355
            Width           =   1965
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Deductions"
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
            Height          =   285
            Index           =   0
            Left            =   1110
            TabIndex        =   31
            Top             =   4650
            Width           =   2745
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAGIBIG Table"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   15
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PHILHEALTH Table"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   14
            Top             =   1065
            Width           =   2235
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WITHHOLDING TAX Table"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   13
            Top             =   2505
            Width           =   2955
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SSS Table"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   12
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Setup"
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
            Height          =   285
            Left            =   1110
            TabIndex        =   11
            Top             =   3945
            Width           =   1755
         End
      End
      Begin XtremeSuiteControls.TabControlPage tbPageMainModules 
         Height          =   6015
         Left            =   -69970
         TabIndex        =   1
         Top             =   600
         Visible         =   0   'False
         Width           =   10275
         _Version        =   655364
         _ExtentX        =   18124
         _ExtentY        =   10610
         _StockProps     =   0
         Begin VB.CommandButton cmdPrintPay 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":332A2
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":333F4
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Print Payroll Sheet"
            Top             =   5340
            Width           =   615
         End
         Begin VB.CommandButton cmdGenPay 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":33856
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":339A8
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Generate Payroll"
            Top             =   4695
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":33DEA
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":33F3C
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Generate Payroll"
            Top             =   4050
            Width           =   615
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   5715
            Left            =   4980
            TabIndex        =   26
            Top             =   90
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   10081
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            WordWrap        =   0   'False
            ShowFocusRect   =   0   'False
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Maintenance"
            TabPicture(0)   =   "MainMenu.frx":34584
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label16"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label18"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label13"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label9"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label15"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label6"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label20"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label11"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "Label44"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "Label46"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "cmdRequestOTLEAVE"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "cmdApprovalOTLEAVE"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "cmdCommission"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "cmdOvertime"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "cmdAdjustments"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "cmdAdvance"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).Control(16)=   "Command3"
            Tab(0).Control(16).Enabled=   0   'False
            Tab(0).Control(17)=   "cmdDeductions"
            Tab(0).Control(17).Enabled=   0   'False
            Tab(0).Control(18)=   "Command8"
            Tab(0).Control(18).Enabled=   0   'False
            Tab(0).Control(19)=   "Command11"
            Tab(0).Control(19).Enabled=   0   'False
            Tab(0).ControlCount=   20
            TabCaption(1)   =   "Others"
            TabPicture(1)   =   "MainMenu.frx":345A0
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Command9"
            Tab(1).Control(1)=   "cmdEmployeeAttendance"
            Tab(1).Control(2)=   "cmdLedger"
            Tab(1).Control(3)=   "Command5"
            Tab(1).Control(4)=   "Command4"
            Tab(1).Control(5)=   "cmdFilePrevEmployer"
            Tab(1).Control(6)=   "cmdFileBegBal"
            Tab(1).Control(7)=   "Label45"
            Tab(1).Control(8)=   "Label23"
            Tab(1).Control(9)=   "Label17"
            Tab(1).Control(10)=   "Label24"
            Tab(1).Control(11)=   "Label22"
            Tab(1).Control(12)=   "Label28"
            Tab(1).Control(13)=   "Label29"
            Tab(1).ControlCount=   14
            Begin VB.CommandButton Command11 
               Height          =   585
               Left            =   2490
               MouseIcon       =   "MainMenu.frx":345BC
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3470E
               Style           =   1  'Graphical
               TabIndex        =   179
               ToolTipText     =   "View Deductions Module"
               Top             =   1710
               Width           =   615
            End
            Begin VB.CommandButton Command9 
               Height          =   585
               Left            =   -74640
               MouseIcon       =   "MainMenu.frx":34ABB
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":34C0D
               Style           =   1  'Graphical
               TabIndex        =   177
               ToolTipText     =   "Print Payroll Sheet"
               Top             =   4455
               Width           =   615
            End
            Begin VB.CommandButton Command8 
               Height          =   585
               Left            =   2490
               MouseIcon       =   "MainMenu.frx":352E6
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":35438
               Style           =   1  'Graphical
               TabIndex        =   175
               ToolTipText     =   "View Deductions Module"
               Top             =   2940
               Width           =   615
            End
            Begin VB.CommandButton cmdDeductions 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":357E5
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":35937
               Style           =   1  'Graphical
               TabIndex        =   130
               ToolTipText     =   "View Deductions Module"
               Top             =   2955
               Width           =   615
            End
            Begin VB.CommandButton Command3 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":35CE4
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":35E36
               Style           =   1  'Graphical
               TabIndex        =   129
               ToolTipText     =   "View Adjustment Code"
               Top             =   3570
               Width           =   615
            End
            Begin VB.CommandButton cmdAdvance 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":363D0
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":36522
               Style           =   1  'Graphical
               TabIndex        =   128
               ToolTipText     =   "View Beginning Balances"
               Top             =   4200
               Width           =   615
            End
            Begin VB.CommandButton cmdAdjustments 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":3682C
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3697E
               Style           =   1  'Graphical
               TabIndex        =   127
               ToolTipText     =   "View Adjustments Module"
               Top             =   4830
               Width           =   615
            End
            Begin VB.CommandButton cmdOvertime 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":38300
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":38452
               Style           =   1  'Graphical
               TabIndex        =   122
               ToolTipText     =   "View Overtime Module"
               Top             =   1710
               Width           =   615
            End
            Begin VB.CommandButton cmdCommission 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":38956
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":38AA8
               Style           =   1  'Graphical
               TabIndex        =   121
               ToolTipText     =   "View Commission Module"
               Top             =   2325
               Width           =   615
            End
            Begin VB.CommandButton cmdApprovalOTLEAVE 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":38F4C
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3909E
               Style           =   1  'Graphical
               TabIndex        =   120
               ToolTipText     =   "View Overtime Module"
               Top             =   1080
               Width           =   615
            End
            Begin VB.CommandButton cmdRequestOTLEAVE 
               Height          =   585
               Left            =   330
               MouseIcon       =   "MainMenu.frx":395A2
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":396F4
               Style           =   1  'Graphical
               TabIndex        =   119
               ToolTipText     =   "View Overtime Module"
               Top             =   450
               Width           =   615
            End
            Begin VB.CommandButton cmdEmployeeAttendance 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":39F70
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3A0C2
               Style           =   1  'Graphical
               TabIndex        =   112
               ToolTipText     =   "Employee Time Cards"
               Top             =   450
               Width           =   615
            End
            Begin VB.CommandButton cmdLedger 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":3C434
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3C586
               Style           =   1  'Graphical
               TabIndex        =   111
               ToolTipText     =   "Employee Ledger"
               Top             =   1104
               Width           =   615
            End
            Begin VB.CommandButton Command5 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":3C9FD
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3CB4F
               Style           =   1  'Graphical
               TabIndex        =   110
               ToolTipText     =   "Employee Ledger"
               Top             =   1758
               Width           =   615
            End
            Begin VB.CommandButton Command4 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":3D1F1
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3D343
               Style           =   1  'Graphical
               TabIndex        =   109
               ToolTipText     =   "Print Payroll Sheet"
               Top             =   3780
               Width           =   615
            End
            Begin VB.CommandButton cmdFilePrevEmployer 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":3DA1C
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3DB6E
               Style           =   1  'Graphical
               TabIndex        =   108
               ToolTipText     =   "View Previous Employer"
               Top             =   3090
               Width           =   615
            End
            Begin VB.CommandButton cmdFileBegBal 
               Height          =   585
               Left            =   -74670
               MouseIcon       =   "MainMenu.frx":3DE78
               MousePointer    =   99  'Custom
               Picture         =   "MainMenu.frx":3DFCA
               Style           =   1  'Graphical
               TabIndex        =   107
               ToolTipText     =   "View Beginning Balances"
               Top             =   2430
               Width           =   615
            End
            Begin VB.Label Label46 
               BackStyle       =   0  'Transparent
               Caption         =   "Leave Modules"
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
               Height          =   570
               Left            =   3210
               TabIndex        =   180
               Top             =   1717
               Width           =   1290
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ATM Entry"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   178
               Top             =   3870
               Width           =   1185
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Other Deductions"
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
               Height          =   570
               Left            =   3210
               TabIndex        =   176
               Top             =   2940
               Width           =   1320
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Attendance Deductions"
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
               Height          =   570
               Left            =   1050
               TabIndex        =   134
               Top             =   2910
               Width           =   1395
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Employee Loans "
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
               Height          =   285
               Left            =   1050
               TabIndex        =   133
               Top             =   3735
               Width           =   1980
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Salary Advance Module"
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
               Height          =   285
               Left            =   1050
               TabIndex        =   132
               Top             =   4350
               Width           =   2685
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Adjustments Module"
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
               Height          =   285
               Left            =   1050
               TabIndex        =   131
               Top             =   4980
               Width           =   2340
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Overtime Module"
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
               Height          =   570
               Left            =   1050
               TabIndex        =   126
               Top             =   1717
               Width           =   1290
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Commission Module"
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
               Height          =   285
               Left            =   1050
               TabIndex        =   125
               Top             =   2490
               Width           =   2340
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Approval For Leave/OT"
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
               Height          =   285
               Left            =   1050
               TabIndex        =   124
               Top             =   1245
               Width           =   2625
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Request For Leave/OT"
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
               Height          =   285
               Left            =   1050
               TabIndex        =   123
               Top             =   630
               Width           =   2550
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Employee Ledger"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   118
               Top             =   1206
               Width           =   2010
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Employee Time Cards"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   117
               Top             =   540
               Width           =   2505
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Employee Attendance"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   116
               Top             =   1872
               Width           =   2505
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ATM Summary"
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
               Height          =   285
               Left            =   -73890
               TabIndex        =   115
               Top             =   4605
               Width           =   1665
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Beginning Balances"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   114
               Top             =   2580
               Width           =   2310
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Previous Employer"
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
               Height          =   285
               Left            =   -73860
               TabIndex        =   113
               Top             =   3240
               Width           =   2175
            End
         End
         Begin VB.CommandButton cmdAPP 
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
            Left            =   270
            MouseIcon       =   "MainMenu.frx":3E2D4
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3E426
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Applicant Information System"
            Top             =   3330
            Width           =   615
         End
         Begin VB.CommandButton Command10 
            Height          =   615
            Left            =   270
            MouseIcon       =   "MainMenu.frx":3EA48
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3EB9A
            Style           =   1  'Graphical
            TabIndex        =   22
            Tag             =   "1102"
            ToolTipText     =   "Reminders"
            Top             =   2655
            Width           =   615
         End
         Begin VB.CommandButton cmdEMPINFO_RegProb 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":3F7DC
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3F92E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Regular/Probationary 201 File"
            Top             =   90
            Width           =   615
         End
         Begin VB.CommandButton cmdEMPINFO_Confiential201 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":41CA0
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":41DF2
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Confidential 201 File"
            Top             =   2010
            Width           =   615
         End
         Begin VB.CommandButton cmdEMPINFO_Contract 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":43774
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":438C6
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Contractual 201 File"
            Top             =   735
            Width           =   615
         End
         Begin VB.CommandButton cmdEMPINFO_Allowance 
            Height          =   585
            Left            =   270
            MouseIcon       =   "MainMenu.frx":45C38
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":45D8A
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Allowance Base 201 File"
            Top             =   1365
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confidential 201 File"
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
            Height          =   285
            Left            =   990
            TabIndex        =   143
            Top             =   2166
            Width           =   2310
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allowance-Based 201 File"
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
            Height          =   285
            Left            =   990
            TabIndex        =   142
            Top             =   1514
            Width           =   2925
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contractual 201 File"
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
            Height          =   285
            Left            =   990
            TabIndex        =   141
            Top             =   862
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Regular/Probationary 201 File"
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
            Height          =   285
            Left            =   990
            TabIndex        =   140
            Top             =   210
            Width           =   3375
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Reminders"
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
            Height          =   375
            Left            =   990
            TabIndex        =   139
            Top             =   2818
            Width           =   2490
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Applicant Information"
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
            Height          =   285
            Left            =   990
            TabIndex        =   138
            Top             =   3560
            Width           =   2445
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Update Attendance"
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
            Height          =   285
            Left            =   990
            TabIndex        =   137
            Top             =   4212
            Width           =   2190
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Generate Payroll"
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
            Height          =   285
            Left            =   990
            TabIndex        =   136
            Top             =   4864
            Width           =   1890
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Payroll Sheet"
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
            Height          =   285
            Left            =   990
            TabIndex        =   135
            Top             =   5520
            Width           =   2115
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

Function CheckIfTheresEmployee() As Boolean
    Dim rsTmp                           As New ADODB.Recordset
    Set rsTmp = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        CheckIfTheresEmployee = True
    Else
        CheckIfTheresEmployee = False
    End If
    Set rsTmp = Nothing
End Function

Private Sub cmdAdvance_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ADVANCE", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    frmHRMS_Advance.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdEMPINFO_Allowance_Click()
    If Module_Access(LOGID, "ALLOWANCE BASE INFO", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    EMP_TYPE = "ALLOWANCE BASE"
    HEADOREMP = "EMP_A"
    Unload frmHRMSEmpInfo
    frmHRMSEmpInfo.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdEMPINFO_Confiential201_Click()
    If Module_Access(LOGID, "MANAGERS INFO", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    HEADOREMP = "HEAD"
    Unload frmHRMSEmpInfo
    frmHRMSEmpInfo.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdEMPINFO_Contract_Click()
    If Module_Access(LOGID, "CONTRACTUAL INFO", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    EMP_TYPE = "CONTRACTUAL"
    HEADOREMP = "EMP_A"
    Unload frmHRMSEmpInfo
    frmHRMSEmpInfo.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdEMPINFO_RegProb_Click()
    If Module_Access(LOGID, "EMPLOYEE INFO", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    HEADOREMP = "EMP_A"
    Unload frmHRMSEmpInfo
    frmHRMSEmpInfo.Show
    FormExistsShow frmHRMSEmpInfo
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdjustments_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ADJUSTMENTS", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    frmHRMSAdjustment.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdAPP_Click()
    If Module_Access(LOGID, "APPLICANT INFO", "DATA ENTRY") = False Then Exit Sub
    frmAISMAIN2.Show
End Sub

Private Sub cmdCommission_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN COMMISSION", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    frmHRMSCommission.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdDeductions_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN DEDUCTIONS", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    Unload frmHRMSDeductions
    EMP_TYPE = "EMPLOYEE"
    DEDUCTION_OPTION = "ATTENDANCE DEDUCTION"
    frmHRMSDeductions.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdEmployeeAttendance_Click()
    If Module_Access(LOGID, "EMPLOYEE TIME CARD", "SYSTEM") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSEditCards.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdFileAdjustmentCode_Click()
    If Module_Access(LOGID, "FILES ADJUSTMENTS", "DATA ENTRY") = False Then Exit Sub
    frmHRMSCodes_Adjustment.Show
End Sub

Private Sub cmdFileBegBal_Click()
    If Module_Access(LOGID, "EMPLOYEE BEGINNING BALANCE", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    frmHRMSBeginingBalance.Show
End Sub

Private Sub cmdFileCompanyProfile_Click()
    If Module_Access(LOGID, "HRMS PROFILE", "SYSTEM") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSProfile.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdFileDeductionCode_Click()
    If Module_Access(LOGID, "FILES DEDUCTION CODES", "DATA ENTRY") = False Then Exit Sub
    frmHRMSDeductionCodeMaterFile.Show
End Sub

Private Sub cmdFileDepartmentalCodes_Click()
    If Module_Access(LOGID, "FILES DEPARTMENT", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSDepartment.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdFileLoanCode_Click()
    If Module_Access(LOGID, "FILES LOAN CODES", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_LoanCodes.Show
End Sub

Private Sub cmdFileOverTimeCodes_Click()
    If Module_Access(LOGID, "FILES OVERTIME CODES", "DATA ENTRY") = False Then Exit Sub
    frmHRMSOTCodes.Show
End Sub

Private Sub cmdFilePrevEmployer_Click()
    If Module_Access(LOGID, "EMPLOYEE PREVIOUS EMPLOYER", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    frmHRMSPrevEmp.Show
End Sub

Private Sub cmdFileSalaryGradeCode_Click()
    If Module_Access(LOGID, "FILES SALARY GRADE CODES", "DATA ENTRY") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSSalaryGrade.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdFileWorkingDaysSetup_Click()
    If Module_Access(LOGID, "WORKING DAY SETUP", "SYSTEM") = False Then Exit Sub
    frmHRMSPayrollSetup.Show
End Sub

Private Sub cmdGenPay_Click()
    If Module_Access(LOGID, "PROCESS GENERATE PAYROLL", "PROCESSING") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    frmHRMSGenerate.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdLedger_Click()
    If Module_Access(LOGID, "EMPLOYEE LEDGER", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    HEADOREMP = "EMP_A"
    frmHRMSLedger.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdOTBreakDown_Click()
    If Module_Access(LOGID, "REPORT PRINT OVERTIME BREAKDOWN", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "Print OverTime BreakDown"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub cmdOvertime_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN OVERTIME", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    frmHRMSOvertime.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdPrintPay_Click()
    If Module_Access(LOGID, "REPORT PRINT PAYROLL SHEET", "REPORTS") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    frmHRMSPrintPayroll.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdRequestOTLEAVE_Click()
    If Module_Access(LOGID, "REQUEST FOR LEAVE/OT", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    On Error Resume Next
    frmHRMS_RequestForLeave.Show
End Sub

Private Sub cmdTableOverTime_Click()
    If Module_Access(LOGID, "OVER TIME SETUP", "SYSTEM") = False Then Exit Sub
    frmSETUP_Overtime.Show
End Sub

Private Sub cmdTablePagIbig_Click()
    If Module_Access(LOGID, "TABLE PAGIBIG", "DATA ENTRY") = False Then Exit Sub
    frmHRMSTables_PAGIBIG.Show
End Sub

Private Sub cmdTablePhilHealth_Click()
    If Module_Access(LOGID, "TABLE PHILHEALTH", "DATA ENTRY") = False Then Exit Sub
    frmHRMSTables_PHIC.Show
End Sub

Private Sub cmdTableSSS_Click()
    If Module_Access(LOGID, "TABLE SSS", "DATA ENTRY") = False Then Exit Sub
    frmHRMSTables_SSS.Show
End Sub

Private Sub cmdTableWithHolding_Click()
    If Module_Access(LOGID, "TABLE BIR", "DATA ENTRY") = False Then Exit Sub
    frmHRMSTables_Tax.Show
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "FILES LEAVE CODES", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_Leave.Show
End Sub

Private Sub Command10_Click()
    If Module_Access(LOGID, "REMINDERS", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub cmdReport_Sched_PAGIBIG_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF PAGIBIG PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDPAGIBIG"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_TAXWHELD_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF TAX WITHHELD", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDTAX"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_OVERTIMEPAY_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF OVERTIME PAY", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDOVERTIME"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_COMMISSION_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF COMMISSION", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDCOMMISSION"
    frmHRMSYearly.Show
End Sub

Private Sub Command11_Click()
    frmHRMS_Leave.Show
End Sub

Private Sub Command15_Click()
    If Module_Access(LOGID, "CREATE DISKETTE LAYOUT", "PROCESSING") = False Then Exit Sub
    frmHRMSDISKGenerator.Show 1
End Sub

Private Sub cmdEV_ATMAdvice_Click()
    SavePicture cmdEV_ATMAdvice.Picture, "C:/text.bmp"
    If Module_Access(LOGID, "REPORT PRINT ATM ADVICE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPrintATM.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdReport_Sched_PAYROLL_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF PAYROLL", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDPAYROLL"
    frmHRMSYearly.Show
End Sub

Private Sub Command17_Click()
    If Module_Access(LOGID, "REPORT PRINT PAYROLL SHEET", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPrintPayroll.Show
    frmHRMSPrintPayroll.chkPaySlip.Value = 1
    Screen.MousePointer = 0
End Sub

Private Sub Command18_Click()
    If Module_Access(LOGID, "REPORT PRINT DEDUCTION BREAKDOWN", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "Print Deduction BreakDown"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub cmdReport_Sched_COMMISSIONTAX_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF COMMISSION TAX", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDCOMMISSIONTAX"
    frmHRMSYearly.Show
End Sub

Private Sub Command2_Click()
    If Module_Access(LOGID, "FILES TIME SHIFT CODES", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_TimeShift.Show
End Sub

Private Sub Command20_Click()
    If Module_Access(LOGID, "REPORT PRINT AJUSTMENT BREAKDOWN", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "Print Adjustment BreakDown"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub Command21_Click()
    If Module_Access(LOGID, "REPORT PHILHEALTH MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPHMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command22_Click()
    If Module_Access(LOGID, "REPORT WITHHOLDING TAX MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSTAXMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command23_Click()
    If Module_Access(LOGID, "REPORT SSS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSSSSMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command24_Click()
    If Module_Access(LOGID, "REPORT PAG-IBIG MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPagibigMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command25_Click()
    If Module_Access(LOGID, "REPORT LOANS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSLoansMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command26_Click()
frmHRMS_GovermentForms.Show

    If Module_Access(LOGID, "REPORT RF1-EMPLOYER QUARTERLY REMITTANCE", "REPORTS") = False Then Exit Sub
    
End Sub

Private Sub Command27_Click()
    If Module_Access(LOGID, "OTHER SETUP SIGNATORIES AND HEADERS", "SYSTEM") = False Then Exit Sub
End Sub

Private Sub Command28_Click()
    If Module_Access(LOGID, "OTHER SETUP SCHEDULE OF DEDUCTIONS", "SYSTEM") = False Then Exit Sub
    frmSETUP_Deduction.Show
End Sub

Private Sub Command29_Click()
    If Module_Access(LOGID, "PASSWORD MAINTENANCE", "SYSTEM") = False Then Exit Sub
    frmAccMaintenance.Show
End Sub

Private Sub cmdShiftModule_Click()
    If Module_Access(LOGID, "SHIFT MAINTENANCE", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_Shift_Management.Show
End Sub

Private Sub cmdEV_DTRSummary_Click()
    If Module_Access(LOGID, "DTR SUMMARY REPORTS", "REPORTS") = False Then Exit Sub
    frmHRMSDTRSummary.Show
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN LOANS", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    frmHRMSLoans.Show
    FormExistsShow frmHRMSLoans
End Sub

Private Sub Command30_Click()
    If Module_Access(LOGID, "REPORT R3-SSS CONTRIBUTION COLLECTION LIST", "REPORTS") = False Then Exit Sub
End Sub

Private Sub cmdReports_AlphaListTerminated_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST TERMINATED", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "ALTERMINATED"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReports_ALphalistWPrevEmployer_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST WITH PREVIOUS EMP", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "ALWITHEMP"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReports_AlphalistwioutPrevEmployer_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST W/OUT PREVIOUS EMPLOYEE", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "ALWITHNOEMP"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReports_OtherEmployeeListing_Click()
    If Module_Access(LOGID, "REPORT OTHER EMPLOYEE LISTING WAGE", "REPORTS") = False Then Exit Sub
    rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptReports, HRMS_REPORT_PATH & "empAge.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub cmdReports_EmployeeList_Click()
    If Module_Access(LOGID, "REPORT EMPLOYEE LISTING", "REPORTS") = False Then Exit Sub
    rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptReports, HRMS_REPORT_PATH & "emplist.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub Command36_Click()
    If Module_Access(LOGID, "REPORT PRINT COMMISSION BREAKDOWN", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "Print Commission BreakDown"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub Command37_Click()
    If Module_Access(LOGID, "REPORT QUARTERLY LOANS", "REPORTS") = False Then Exit Sub
    frmHRMS_PrintQuarterlyLoan.Show
End Sub

Private Sub cmdApprovalOTLEAVE_Click()
    If Module_Access(LOGID, "LEAVE/OT APPROVAL", "DATA ENTRY") = False Then Exit Sub
    On Error Resume Next
    frmHRMS_ApprovalForLeave.Show
End Sub

Private Sub cmdReport_Sched_13MONTH_Click()
    If Module_Access(LOGID, "REPORT 13TH MONTH PAY SCHEDULE", "REPORTS") = False Then Exit Sub
    frmHRMSPrint13thMonth.Show
End Sub

Private Sub cmdReport_Sched_TAXDUE_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF TAXDUE/REFUND", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDTAXDUEREFUND"
    frmHRMSYearly.Show
End Sub

Private Sub Command4_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ATM ENTRY", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    frmHRMSATM.Show
End Sub

Private Sub cmdReports_ResingedEmployee_Click()
    If Module_Access(LOGID, "REPORT RESIGNESS LISTING", "REPORTS") = False Then Exit Sub
    rptReports.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptReports.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptReports.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptReports, HRMS_REPORT_PATH & "resignees.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub cmdReports_YearToDate_Click()
    If Module_Access(LOGID, "REPORT YEAR-TO-DATE DETAILS", "REPORTS") = False Then Exit Sub
    frmHRMSPrintYTDProcessing.Show
End Sub

Private Sub cmd_ReportBlankEmployye_Click()
    If Module_Access(LOGID, "REPORT PRINT BLANK EMP. INFO SHEET", "REPORTS") = False Then Exit Sub
    PrintSQLReport rptReports, HRMS_REPORT_PATH & "blankempinfo.rpt", "", DMIS_REPORT_Connection, 1
End Sub

Private Sub Command5_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ATTENDANCE", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    frmHRMSDailyMonitoring.Show
End Sub

Private Sub Command6_Click()
    If Module_Access(LOGID, "PROCESS UPDATE ATTENDANCE", "PROCESSING") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    On Error Resume Next
    frmHRMSUpDateAttendance.Show
End Sub

Private Sub Command7_Click()
    If Module_Access(LOGID, "HOLIDAY SETUP", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_HolidaySetup.Show
End Sub

Private Sub cmdReport_Sched_SSS_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF SSS PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDSSS"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_PHIC_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF PHILHEALTH PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub
    FORMYEARLYREQUEST = "SCHEDPHIC"
    frmHRMSYearly.Show
End Sub

Private Sub Command8_Click()
    'If Module_Access(LOGID, "EMPLOYEE MAINTAIN OTHER DEDUCTIONS", "DATA ENTRY") = False Then Exit Sub
    If CheckIfTheresEmployee = False Then
        MessagePop InfoVoid, "No Record", "There's no Employee Record on the Database"
        Exit Sub
    End If
    Screen.MousePointer = 11
    EMP_TYPE = "EMPLOYEE"
    DEDUCTION_OPTION = "OTHER DEDUCTIONS"
    frmHRMSDeductions.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command9_Click()
    'frmHRMS_ATM_Summary.Show
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SSTab1.Tab = 0
    TabControl1.SelectedItem = 0
    GetThePayrollCode
End Sub

Sub FormExistsShow(frmx As Form)
    On Error GoTo Errorcode
    Dim m_Exists                        As Boolean
    Dim FRM                             As Form
    For Each FRM In Forms
        If (UCase(FRM.Name) = UCase(frmx.Name)) Then
            m_Exists = True
        End If
    Next
    Set FRM = Nothing
    If m_Exists = True Then
        frmx.WindowState = 0
        frmx.ZOrder 0
    End If
    Exit Sub
Errorcode:
    Err.Clear
End Sub

