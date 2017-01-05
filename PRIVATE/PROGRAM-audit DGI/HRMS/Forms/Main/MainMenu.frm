VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HRMS Main Menu"
   ClientHeight    =   6690
   ClientLeft      =   990
   ClientTop       =   1170
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   10260
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6675
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10245
      _Version        =   655364
      _ExtentX        =   18071
      _ExtentY        =   11774
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      PaintManager.MinTabWidth=   100
      ItemCount       =   4
      Item(0).Caption =   "Main Modules"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "cmdPrintPay"
      Item(0).Control(1)=   "cmdGenPay"
      Item(0).Control(2)=   "Command6"
      Item(0).Control(3)=   "cmdAPP"
      Item(0).Control(4)=   "Command10"
      Item(0).Control(5)=   "cmdEMPINFO_RegProb"
      Item(0).Control(6)=   "cmdEMPINFO_Confiential201"
      Item(0).Control(7)=   "cmdEMPINFO_Contract"
      Item(0).Control(8)=   "cmdEMPINFO_Allowance"
      Item(0).Control(9)=   "Label1(2)"
      Item(0).Control(10)=   "Label1(3)"
      Item(0).Control(11)=   "Label1(1)"
      Item(0).Control(12)=   "Label1(0)"
      Item(0).Control(13)=   "Label1(4)"
      Item(0).Control(14)=   "Label1(5)"
      Item(0).Control(15)=   "Label1(6)"
      Item(0).Control(16)=   "Label1(7)"
      Item(0).Control(17)=   "Label1(8)"
      Item(0).Control(18)=   "TabControl3"
      Item(1).Caption =   "Tables && File Maintenance"
      Item(1).ControlCount=   38
      Item(1).Control(0)=   "Command19"
      Item(1).Control(1)=   "Command7"
      Item(1).Control(2)=   "cmdFileDepartmentalCodes"
      Item(1).Control(3)=   "cmdFileOverTimeCodes"
      Item(1).Control(4)=   "cmdFileDeductionCode"
      Item(1).Control(5)=   "cmdFileSalaryGradeCode"
      Item(1).Control(6)=   "cmdFileLoanCode"
      Item(1).Control(7)=   "cmdFileAdjustmentCode"
      Item(1).Control(8)=   "Command2"
      Item(1).Control(9)=   "Command1"
      Item(1).Control(10)=   "cmdFileWorkingDaysSetup"
      Item(1).Control(11)=   "cmdShiftModule"
      Item(1).Control(12)=   "Command28"
      Item(1).Control(13)=   "cmdTableOverTime"
      Item(1).Control(14)=   "cmdTablePagIbig"
      Item(1).Control(15)=   "cmdTablePhilHealth"
      Item(1).Control(16)=   "cmdTableWithHolding"
      Item(1).Control(17)=   "cmdTableSSS"
      Item(1).Control(18)=   "Label1(39)"
      Item(1).Control(19)=   "Label1(38)"
      Item(1).Control(20)=   "Label1(32)"
      Item(1).Control(21)=   "Label1(36)"
      Item(1).Control(22)=   "Label1(34)"
      Item(1).Control(23)=   "Label41"
      Item(1).Control(24)=   "Label1(35)"
      Item(1).Control(25)=   "Label1(33)"
      Item(1).Control(26)=   "Label1(37)"
      Item(1).Control(27)=   "Label1(31)"
      Item(1).Control(28)=   "Label1(28)"
      Item(1).Control(29)=   "Label1(30)"
      Item(1).Control(30)=   "Label1(29)"
      Item(1).Control(31)=   "Label1(26)"
      Item(1).Control(32)=   "Label1(25)"
      Item(1).Control(33)=   "Label1(27)"
      Item(1).Control(34)=   "Label1(24)"
      Item(1).Control(35)=   "Label47"
      Item(1).Control(36)=   "Command42"
      Item(1).Control(37)=   "Label4"
      Item(2).Caption =   "Reports"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "Picture1"
      Item(2).Control(1)=   "TabControl2"
      Item(3).Caption =   "Other Setups"
      Item(3).ControlCount=   8
      Item(3).Control(0)=   "Command38"
      Item(3).Control(1)=   "cmdFileCompanyProfile"
      Item(3).Control(2)=   "Command29"
      Item(3).Control(3)=   "Command27"
      Item(3).Control(4)=   "Label19(1)"
      Item(3).Control(5)=   "Label43"
      Item(3).Control(6)=   "Label75"
      Item(3).Control(7)=   "Label73"
      Begin VB.CommandButton Command42 
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
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":6852
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":69A4
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Employee Payroll Set Up"
         Top             =   5790
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6045
         Left            =   -63040
         ScaleHeight     =   6015
         ScaleWidth      =   3195
         TabIndex        =   94
         Top             =   600
         Visible         =   0   'False
         Width           =   3225
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
            Picture         =   "MainMenu.frx":6FC6
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Create Diskette Layout"
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label Label61 
            BackStyle       =   0  'Transparent
            Caption         =   "*  Pag-ibig Loan"
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
            Left            =   630
            TabIndex        =   103
            Top             =   4320
            Width           =   2265
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "*  Pag-ibig"
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
            Left            =   630
            TabIndex        =   102
            Top             =   3990
            Width           =   1545
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "*  PhilHealth"
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
            Left            =   630
            TabIndex        =   101
            Top             =   3660
            Width           =   1785
         End
         Begin VB.Label Label58 
            BackStyle       =   0  'Transparent
            Caption         =   "*  SSS Loan"
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
            Left            =   630
            TabIndex        =   100
            Top             =   3330
            Width           =   1545
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "*  SSS"
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
            Left            =   630
            TabIndex        =   99
            Top             =   3030
            Width           =   1545
         End
         Begin VB.Label Label56 
            BackStyle       =   0  'Transparent
            Caption         =   "For:"
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
            Height          =   345
            Left            =   90
            TabIndex        =   98
            Top             =   2700
            Width           =   645
         End
         Begin VB.Label Label54 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Diskette Layout Generator"
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   97
            Top             =   480
            Width           =   2505
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   435
            Left            =   0
            TabIndex        =   96
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
            SubItemCaption  =   -1  'True
            ForeColor       =   8388608
         End
      End
      Begin VB.CommandButton cmdTableSSS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":7408
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":755A
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "SSS Table"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdTableWithHolding 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":908C
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":91DE
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "WITHHOLDING TAX Table"
         Top             =   2925
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdTablePhilHealth 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":B6E8
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":B83A
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "PHILHEALTH Table"
         Top             =   1485
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdTablePagIbig 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":CB18
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":CC6A
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "PAGIBIG Table"
         Top             =   2220
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdTableOverTime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -62680
         MouseIcon       =   "MainMenu.frx":11688
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":117DA
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Overtime Setup"
         Top             =   2250
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command28 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":11C54
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":11DA6
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Schedule of Deductions"
         Top             =   4350
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdShiftModule 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":121E8
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":1233A
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "View Time Shift Set Up"
         Top             =   5055
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileWorkingDaysSetup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":1283E
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":12990
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "View Payroll Setup"
         Top             =   3630
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69670
         MouseIcon       =   "MainMenu.frx":12C9A
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":12DEC
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "View Leave Code"
         Top             =   5790
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -62680
         MouseIcon       =   "MainMenu.frx":132F0
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":13442
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "View Time Shift Code"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileAdjustmentCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":13AE9
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":13C3B
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "View Adjustment Code"
         Top             =   1500
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileLoanCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":13F45
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":14097
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "View Loan Codes"
         Top             =   2910
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileSalaryGradeCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -62680
         MouseIcon       =   "MainMenu.frx":15E21
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":15F73
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "View Salary Grade Codes"
         Top             =   1500
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileDeductionCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":1627D
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":163CF
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "View Deduction Codes"
         Top             =   2205
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileOverTimeCodes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":166D9
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":1682B
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "View Overtime Codes"
         Top             =   3645
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdFileDepartmentalCodes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":16B35
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":16C87
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "View Department Codes"
         Top             =   720
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":16F91
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":170E3
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "View Holiday Setup"
         Top             =   4365
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -65830
         MouseIcon       =   "MainMenu.frx":173ED
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":1753F
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Leave Maintenance"
         Top             =   5070
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdEMPINFO_Allowance 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":179B9
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":17B0B
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Allowance Base 201 File"
         Top             =   1995
         Width           =   615
      End
      Begin VB.CommandButton cmdEMPINFO_Contract 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":19E7D
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":19FCF
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Contractual 201 File"
         Top             =   1365
         Width           =   615
      End
      Begin VB.CommandButton cmdEMPINFO_Confiential201 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":1C341
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":1C493
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Confidential 201 File"
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton cmdEMPINFO_RegProb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":1DE15
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":1DF67
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Regular/Probationary 201 File"
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   300
         MouseIcon       =   "MainMenu.frx":202D9
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":2042B
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "1102"
         ToolTipText     =   "Reminders"
         Top             =   3285
         Width           =   615
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
         Left            =   300
         MouseIcon       =   "MainMenu.frx":2106D
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":211BF
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Applicant Information System"
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":217E1
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":21933
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Update Attendance"
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdGenPay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":21F7B
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":220CD
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Generate Payroll"
         Top             =   5325
         Width           =   615
      End
      Begin VB.CommandButton cmdPrintPay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   300
         MouseIcon       =   "MainMenu.frx":2250F
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":22661
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Print Payroll Sheet"
         Top             =   5970
         Width           =   615
      End
      Begin VB.CommandButton Command27 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69700
         MouseIcon       =   "MainMenu.frx":22AC3
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":22C15
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Signatories and Headers"
         Top             =   1530
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69700
         MouseIcon       =   "MainMenu.frx":23057
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":231A9
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Password Maintenance "
         Top             =   2250
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton cmdFileCompanyProfile 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -69700
         MouseIcon       =   "MainMenu.frx":24B2B
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":24C7D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "View Company Profile"
         Top             =   750
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command38 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   -69700
         MouseIcon       =   "MainMenu.frx":2516C
         MousePointer    =   99  'Custom
         Picture         =   "MainMenu.frx":252BE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Generate Payroll"
         Top             =   2970
         Visible         =   0   'False
         Width           =   720
      End
      Begin XtremeSuiteControls.TabControl TabControl3 
         Height          =   6075
         Left            =   4680
         TabIndex        =   27
         Top             =   570
         Width           =   5535
         _Version        =   655364
         _ExtentX        =   9763
         _ExtentY        =   10716
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   100
         ItemCount       =   2
         Item(0).Caption =   "Maintenance"
         Item(0).ControlCount=   16
         Item(0).Control(0)=   "Command8"
         Item(0).Control(1)=   "cmdDeductions"
         Item(0).Control(2)=   "Command3"
         Item(0).Control(3)=   "cmdAdvance"
         Item(0).Control(4)=   "cmdAdjustments"
         Item(0).Control(5)=   "cmdOvertime"
         Item(0).Control(6)=   "cmdApprovalOTLEAVE"
         Item(0).Control(7)=   "cmdRequestOTLEAVE"
         Item(0).Control(8)=   "Label1(13)"
         Item(0).Control(9)=   "Label1(12)"
         Item(0).Control(10)=   "Label1(14)"
         Item(0).Control(11)=   "Label1(15)"
         Item(0).Control(12)=   "Label1(16)"
         Item(0).Control(13)=   "Label1(11)"
         Item(0).Control(14)=   "Label1(10)"
         Item(0).Control(15)=   "Label1(9)"
         Item(1).Caption =   "Others"
         Item(1).ControlCount=   16
         Item(1).Control(0)=   "Command9"
         Item(1).Control(1)=   "cmdEmployeeAttendance"
         Item(1).Control(2)=   "cmdLedger"
         Item(1).Control(3)=   "Command5"
         Item(1).Control(4)=   "Command4"
         Item(1).Control(5)=   "cmdFilePrevEmployer"
         Item(1).Control(6)=   "cmdFileBegBal"
         Item(1).Control(7)=   "Label1(21)"
         Item(1).Control(8)=   "Label1(18)"
         Item(1).Control(9)=   "Label1(17)"
         Item(1).Control(10)=   "Label1(19)"
         Item(1).Control(11)=   "Label1(22)"
         Item(1).Control(12)=   "Label1(20)"
         Item(1).Control(13)=   "Label1(23)"
         Item(1).Control(14)=   "Command11"
         Item(1).Control(15)=   "Label1(40)"
         Begin VB.CommandButton Command11 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":26340
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":26492
            Style           =   1  'Graphical
            TabIndex        =   191
            ToolTipText     =   "View Leave Convention"
            Top             =   5340
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":27514
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":27666
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "View Other Deduction Module"
            Top             =   2685
            Width           =   615
         End
         Begin VB.CommandButton cmdDeductions 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":27A13
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":27B65
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "View Deductions Module"
            Top             =   2040
            Width           =   615
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":27F12
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":28064
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "View Employee Loan"
            Top             =   3300
            Width           =   615
         End
         Begin VB.CommandButton cmdAdvance 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":285FE
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":28750
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "View Salary Advance Module"
            Top             =   3930
            Width           =   615
         End
         Begin VB.CommandButton cmdAdjustments 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":28A5A
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":28BAC
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "View Adjustments Module"
            Top             =   4560
            Width           =   615
         End
         Begin VB.CommandButton cmdOvertime 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":2A52E
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2A680
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "View Overtime Module"
            Top             =   1425
            Width           =   615
         End
         Begin VB.CommandButton cmdApprovalOTLEAVE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   5940
            MouseIcon       =   "MainMenu.frx":2AB84
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2ACD6
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "View Overtime Module"
            Top             =   5280
            Width           =   615
         End
         Begin VB.CommandButton cmdRequestOTLEAVE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   180
            MouseIcon       =   "MainMenu.frx":2B1DA
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2B32C
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Request & Approval For Leave/OT"
            Top             =   780
            Width           =   615
         End
         Begin VB.CommandButton Command9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2BBA8
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2BCFA
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "View ATM Summary"
            Top             =   4020
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdEmployeeAttendance 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2C3D3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2C525
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Employee Time Cards"
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdLedger 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2E897
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2E9E9
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Employee Ledger"
            Top             =   1380
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2EE60
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2EFB2
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Employee Attendance"
            Top             =   2040
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2F654
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2F7A6
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "View Atm Enrty"
            Top             =   3360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdFilePrevEmployer 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":2FE7F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":2FFD1
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "View Previous Employer"
            Top             =   4680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdFileBegBal 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   -69820
            MouseIcon       =   "MainMenu.frx":302DB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3042D
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "View Beginning Balances"
            Top             =   2700
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Convertion"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   40
            Left            =   -69100
            TabIndex        =   192
            Top             =   5520
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Deductions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   13
            Left            =   900
            TabIndex        =   57
            Top             =   2820
            Width           =   2910
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Attendance Deductions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   12
            Left            =   900
            TabIndex        =   56
            Top             =   2190
            Width           =   2955
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Loans "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   14
            Left            =   900
            TabIndex        =   55
            Top             =   3465
            Width           =   1635
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Advance Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   15
            Left            =   900
            TabIndex        =   54
            Top             =   4080
            Width           =   2265
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustments Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   16
            Left            =   900
            TabIndex        =   53
            Top             =   4710
            Width           =   1920
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Overtime Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   570
            Index           =   11
            Left            =   900
            TabIndex        =   52
            Top             =   1590
            Width           =   2160
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Approval For Leave/OT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   10
            Left            =   6660
            TabIndex        =   51
            Top             =   5445
            Width           =   2205
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Request && Approval For Leave/OT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   9
            Left            =   900
            TabIndex        =   50
            Top             =   960
            Width           =   3225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ATM Entry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   21
            Left            =   -69100
            TabIndex        =   49
            Top             =   3540
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Ledger"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   18
            Left            =   -69100
            TabIndex        =   48
            Top             =   1500
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Time Cards"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   17
            Left            =   -69100
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Attendance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   19
            Left            =   -69100
            TabIndex        =   46
            Top             =   2175
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ATM Summary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   22
            Left            =   -69100
            TabIndex        =   45
            Top             =   4185
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Beginning Balances"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   20
            Left            =   -69100
            TabIndex        =   44
            Top             =   2880
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Previous Employer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   23
            Left            =   -69100
            TabIndex        =   43
            Top             =   4860
            Visible         =   0   'False
            Width           =   1785
         End
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   6075
         Left            =   -69970
         TabIndex        =   104
         Top             =   600
         Visible         =   0   'False
         Width           =   6915
         _Version        =   655364
         _ExtentX        =   12197
         _ExtentY        =   10716
         _StockProps     =   64
         Appearance      =   2
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         PaintManager.LargeIcons=   -1  'True
         PaintManager.FixedTabWidth=   100
         ItemCount       =   5
         Item(0).Caption =   "Every Payroll"
         Item(0).ControlCount=   26
         Item(0).Control(0)=   "Command39"
         Item(0).Control(1)=   "Command35"
         Item(0).Control(2)=   "Command12"
         Item(0).Control(3)=   "Command36"
         Item(0).Control(4)=   "Command20"
         Item(0).Control(5)=   "cmdOTBreakDown"
         Item(0).Control(6)=   "Command18"
         Item(0).Control(7)=   "Command17"
         Item(0).Control(8)=   "cmdEV_ATMAdvice"
         Item(0).Control(9)=   "cmdEV_DTRSummary"
         Item(0).Control(10)=   "Label82(1)"
         Item(0).Control(11)=   "Label53(0)"
         Item(0).Control(12)=   "Label48"
         Item(0).Control(13)=   "Label82(0)"
         Item(0).Control(14)=   "Label66"
         Item(0).Control(15)=   "Label65"
         Item(0).Control(16)=   "Label64"
         Item(0).Control(17)=   "Label63"
         Item(0).Control(18)=   "Label62"
         Item(0).Control(19)=   "Label14"
         Item(0).Control(20)=   "Command40"
         Item(0).Control(21)=   "Label2"
         Item(0).Control(22)=   "Command41"
         Item(0).Control(23)=   "Label3"
         Item(0).Control(24)=   "Command43"
         Item(0).Control(25)=   "Label5"
         Item(1).Caption =   "Monthly Reports"
         Item(1).ControlCount=   10
         Item(1).Control(0)=   "Command24"
         Item(1).Control(1)=   "Command23"
         Item(1).Control(2)=   "Command22"
         Item(1).Control(3)=   "Command21"
         Item(1).Control(4)=   "Command25"
         Item(1).Control(5)=   "Label71"
         Item(1).Control(6)=   "Label70"
         Item(1).Control(7)=   "Label69"
         Item(1).Control(8)=   "Label68"
         Item(1).Control(9)=   "Label67"
         Item(2).Caption =   "Quarterly Reports"
         Item(2).ControlCount=   6
         Item(2).Control(0)=   "Command37"
         Item(2).Control(1)=   "Command30"
         Item(2).Control(2)=   "Command26"
         Item(2).Control(3)=   "Label83"
         Item(2).Control(4)=   "Label76"
         Item(2).Control(5)=   "Label72"
         Item(3).Caption =   "Other Reports"
         Item(3).ControlCount=   18
         Item(3).Control(0)=   "Command34"
         Item(3).Control(1)=   "Command13"
         Item(3).Control(2)=   "Command16"
         Item(3).Control(3)=   "Command14"
         Item(3).Control(4)=   "cmd_ReportBlankEmployye"
         Item(3).Control(5)=   "cmdReports_ResingedEmployee"
         Item(3).Control(6)=   "cmdReports_OtherEmployeeListing"
         Item(3).Control(7)=   "cmdReports_EmployeeList"
         Item(3).Control(8)=   "Label52"
         Item(3).Control(9)=   "labSched(10)"
         Item(3).Control(10)=   "Label50"
         Item(3).Control(11)=   "Label49"
         Item(3).Control(12)=   "Label38"
         Item(3).Control(13)=   "Label36"
         Item(3).Control(14)=   "Label35"
         Item(3).Control(15)=   "Label81"
         Item(3).Control(16)=   "Command44"
         Item(3).Control(17)=   "Label6"
         Item(4).Caption =   "Schedules"
         Item(4).ControlCount=   34
         Item(4).Control(0)=   "Command33"
         Item(4).Control(1)=   "Command32"
         Item(4).Control(2)=   "cmdReports_AlphaListTerminated"
         Item(4).Control(3)=   "cmdReports_ALphalistWPrevEmployer"
         Item(4).Control(4)=   "cmdReports_AlphalistwioutPrevEmployer"
         Item(4).Control(5)=   "cmdReports_YearToDate"
         Item(4).Control(6)=   "Command31"
         Item(4).Control(7)=   "cmdReport_Sched_TAXDUE"
         Item(4).Control(8)=   "cmdReport_Sched_13MONTH"
         Item(4).Control(9)=   "cmdReport_Sched_COMMISSIONTAX"
         Item(4).Control(10)=   "cmdReport_Sched_PAYROLL"
         Item(4).Control(11)=   "cmdReport_Sched_COMMISSION"
         Item(4).Control(12)=   "cmdReport_Sched_OVERTIMEPAY"
         Item(4).Control(13)=   "cmdReport_Sched_TAXWHELD"
         Item(4).Control(14)=   "cmdReport_Sched_PAGIBIG"
         Item(4).Control(15)=   "cmdReport_Sched_PHIC"
         Item(4).Control(16)=   "cmdReport_Sched_SSS"
         Item(4).Control(17)=   "labSched(13)"
         Item(4).Control(18)=   "labSched(12)"
         Item(4).Control(19)=   "Label77"
         Item(4).Control(20)=   "Label78"
         Item(4).Control(21)=   "Label79"
         Item(4).Control(22)=   "Label37"
         Item(4).Control(23)=   "labSched(11)"
         Item(4).Control(24)=   "labSched(9)"
         Item(4).Control(25)=   "labSched(8)"
         Item(4).Control(26)=   "labSched(7)"
         Item(4).Control(27)=   "labSched(6)"
         Item(4).Control(28)=   "labSched(5)"
         Item(4).Control(29)=   "labSched(0)"
         Item(4).Control(30)=   "labSched(1)"
         Item(4).Control(31)=   "labSched(2)"
         Item(4).Control(32)=   "labSched(3)"
         Item(4).Control(33)=   "labSched(4)"
         Begin VB.CommandButton Command44 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":30737
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":30889
            Style           =   1  'Graphical
            TabIndex        =   201
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   5160
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command43 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3630
            MouseIcon       =   "MainMenu.frx":30CEB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":30E3D
            Style           =   1  'Graphical
            TabIndex        =   199
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   1800
            Width           =   570
         End
         Begin VB.CommandButton Command41 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3630
            MouseIcon       =   "MainMenu.frx":3129F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":313F1
            Style           =   1  'Graphical
            TabIndex        =   195
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   1290
            Width           =   570
         End
         Begin VB.CommandButton Command40 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3630
            MouseIcon       =   "MainMenu.frx":31853
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":319A5
            Style           =   1  'Graphical
            TabIndex        =   193
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   750
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_SSS 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":31E07
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":31F59
            Style           =   1  'Graphical
            TabIndex        =   173
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   765
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_PHIC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":323BB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3250D
            Style           =   1  'Graphical
            TabIndex        =   172
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   1420
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_PAGIBIG 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3296F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":32AC1
            Style           =   1  'Graphical
            TabIndex        =   171
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   2075
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_TAXWHELD 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":32F23
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":33075
            Style           =   1  'Graphical
            TabIndex        =   170
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   2730
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_OVERTIMEPAY 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":334D7
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":33629
            Style           =   1  'Graphical
            TabIndex        =   169
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   3385
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_COMMISSION 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -67480
            MouseIcon       =   "MainMenu.frx":33A8B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":33BDD
            Style           =   1  'Graphical
            TabIndex        =   168
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_PAYROLL 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3403F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":34191
            Style           =   1  'Graphical
            TabIndex        =   167
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   4040
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_COMMISSIONTAX 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -63010
            MouseIcon       =   "MainMenu.frx":345F3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":34745
            Style           =   1  'Graphical
            TabIndex        =   166
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   4140
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_13MONTH 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":34BA7
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":34CF9
            Style           =   1  'Graphical
            TabIndex        =   165
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   765
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReport_Sched_TAXDUE 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3515B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":352AD
            Style           =   1  'Graphical
            TabIndex        =   164
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   4695
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command31 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":3570F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":35861
            Style           =   1  'Graphical
            TabIndex        =   163
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   3930
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReports_YearToDate 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":35CC3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":35E15
            Style           =   1  'Graphical
            TabIndex        =   162
            ToolTipText     =   "Other Reports"
            Top             =   3309
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.CommandButton cmdReports_AlphalistwioutPrevEmployer 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":36277
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":363C9
            Style           =   1  'Graphical
            TabIndex        =   161
            ToolTipText     =   "Year to Date Reports"
            Top             =   2688
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.CommandButton cmdReports_ALphalistWPrevEmployer 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":3682B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3697D
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "Form 2316 && Alphalist"
            Top             =   2067
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.CommandButton cmdReports_AlphaListTerminated 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":36DDF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":36F31
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "13th Month Pay and Others"
            Top             =   1446
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.CommandButton Command32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -66220
            MouseIcon       =   "MainMenu.frx":37393
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":374E5
            Style           =   1  'Graphical
            TabIndex        =   158
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   4611
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command33 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":37947
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":37A99
            Style           =   1  'Graphical
            TabIndex        =   157
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   5355
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReports_EmployeeList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":37EFB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3804D
            Style           =   1  'Graphical
            TabIndex        =   148
            ToolTipText     =   "Other Reports"
            Top             =   1320
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReports_OtherEmployeeListing 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":384AF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":38601
            Style           =   1  'Graphical
            TabIndex        =   147
            ToolTipText     =   "Other Reports"
            Top             =   1875
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdReports_ResingedEmployee 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":38A63
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":38BB5
            Style           =   1  'Graphical
            TabIndex        =   146
            ToolTipText     =   "Other Reports"
            Top             =   2415
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmd_ReportBlankEmployye 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":39017
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":39169
            Style           =   1  'Graphical
            TabIndex        =   145
            ToolTipText     =   "Other Reports"
            Top             =   780
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":395CB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3971D
            Style           =   1  'Graphical
            TabIndex        =   144
            ToolTipText     =   "Other Reports"
            Top             =   2955
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":39B7F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":39CD1
            Style           =   1  'Graphical
            TabIndex        =   143
            ToolTipText     =   "Other Reports"
            Top             =   3495
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3A133
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3A285
            Style           =   1  'Graphical
            TabIndex        =   142
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   4050
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command34 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3A6E7
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3A839
            Style           =   1  'Graphical
            TabIndex        =   141
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   4590
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton Command26 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3AC9B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3ADED
            Style           =   1  'Graphical
            TabIndex        =   137
            ToolTipText     =   "RF1 - Employer Quarterly Remittance"
            Top             =   780
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command30 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3C0CB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3C21D
            Style           =   1  'Graphical
            TabIndex        =   136
            ToolTipText     =   "R3 - SSS Contribution Collection List"
            Top             =   1560
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command37 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3DD4F
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3DEA1
            Style           =   1  'Graphical
            TabIndex        =   135
            ToolTipText     =   "Quarterly Loans"
            Top             =   2370
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command25 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":3FC2B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":3FD7D
            Style           =   1  'Graphical
            TabIndex        =   129
            ToolTipText     =   "LOANS Remitted"
            Top             =   3765
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command21 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":41B07
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":41C59
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "PHILHEALTH Contributions"
            Top             =   810
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command22 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":42F37
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":43089
            Style           =   1  'Graphical
            TabIndex        =   127
            ToolTipText     =   "TAX Withheld"
            Top             =   3030
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command23 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":45593
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":456E5
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "SSS Contributions"
            Top             =   1545
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton Command24 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   -69685
            MouseIcon       =   "MainMenu.frx":47217
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":47369
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "PAG-IBIG Contributions"
            Top             =   2280
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton cmdEV_DTRSummary 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   300
            MouseIcon       =   "MainMenu.frx":4BD87
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4BED9
            Style           =   1  'Graphical
            TabIndex        =   114
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   3395
            Width           =   570
         End
         Begin VB.CommandButton cmdEV_ATMAdvice 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   315
            MouseIcon       =   "MainMenu.frx":4C33B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4C48D
            Style           =   1  'Graphical
            TabIndex        =   113
            ToolTipText     =   "Print Payroll ATM Advice"
            Top             =   720
            Width           =   570
         End
         Begin VB.CommandButton Command17 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   315
            MouseIcon       =   "MainMenu.frx":4C8EF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4CA41
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Print Payroll PaySlips"
            Top             =   1255
            Width           =   570
         End
         Begin VB.CommandButton Command18 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   315
            MouseIcon       =   "MainMenu.frx":4CEA3
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4CFF5
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Print Deductions Breakdown"
            Top             =   1790
            Width           =   570
         End
         Begin VB.CommandButton cmdOTBreakDown 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   315
            MouseIcon       =   "MainMenu.frx":4D457
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4D5A9
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Print Overtime Breakdown "
            Top             =   2325
            Width           =   570
         End
         Begin VB.CommandButton Command20 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   315
            MouseIcon       =   "MainMenu.frx":4DA0B
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4DB5D
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   "Print Adjustment Breakdown"
            Top             =   2860
            Width           =   570
         End
         Begin VB.CommandButton Command36 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   285
            MouseIcon       =   "MainMenu.frx":4DFBF
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4E111
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   5535
            Width           =   570
         End
         Begin VB.CommandButton Command12 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   300
            MouseIcon       =   "MainMenu.frx":4E573
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4E6C5
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   3930
            Width           =   570
         End
         Begin VB.CommandButton Command35 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   300
            MouseIcon       =   "MainMenu.frx":4EB27
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4EC79
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   4465
            Width           =   570
         End
         Begin VB.CommandButton Command39 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   300
            MouseIcon       =   "MainMenu.frx":4F0DB
            MousePointer    =   99  'Custom
            Picture         =   "MainMenu.frx":4F22D
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Print Commission Breakdown"
            Top             =   5000
            Width           =   570
         End
         Begin Crystal.CrystalReport rptReports 
            Left            =   630
            Top             =   6240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label Label6 
            Caption         =   "Print Payroll Summary"
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
            Height          =   255
            Left            =   -69040
            TabIndex        =   202
            Top             =   5280
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Print ATM Summary"
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
            Height          =   255
            Left            =   4320
            TabIndex        =   200
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Print Late/Absent Record"
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
            Left            =   4320
            TabIndex        =   196
            Top             =   1380
            Width           =   2130
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Print Monthly Time Record"
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
            Left            =   4320
            TabIndex        =   194
            Top             =   840
            Width           =   2235
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Overtime Pay"
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
            Height          =   210
            Index           =   4
            Left            =   -69010
            TabIndex        =   190
            Top             =   3555
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Tax Withheld"
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
            Height          =   180
            Index           =   3
            Left            =   -69010
            TabIndex        =   189
            Top             =   2895
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Pag-Ibig Premium Contribution"
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
            Index           =   2
            Left            =   -69010
            TabIndex        =   188
            Top             =   2100
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Philhealth Premium Contribution"
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
            Index           =   1
            Left            =   -69010
            TabIndex        =   187
            Top             =   1455
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label labSched 
            Caption         =   "Schedule of SSS Premium Contribution"
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
            Height          =   510
            Index           =   0
            Left            =   -69010
            TabIndex        =   186
            Top             =   795
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Label labSched 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Payroll"
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
            Index           =   5
            Left            =   -69010
            TabIndex        =   185
            Top             =   4200
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label labSched 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Commission"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   6
            Left            =   -66760
            TabIndex        =   184
            Top             =   6360
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Commission Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Index           =   7
            Left            =   -62020
            TabIndex        =   183
            Top             =   4245
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.Label labSched 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "13th Month Pay Schedule"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Index           =   8
            Left            =   -65530
            TabIndex        =   182
            Top             =   885
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Tax Due/Refund"
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
            Height          =   420
            Index           =   9
            Left            =   -69010
            TabIndex        =   181
            Top             =   4710
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha List 2008"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   11
            Left            =   -65530
            TabIndex        =   180
            Top             =   4080
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Year-To-Date Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   -65530
            TabIndex        =   179
            Top             =   3450
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label79 
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha List W/Out Previous Employer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   450
            Left            =   -65530
            TabIndex        =   178
            Top             =   2715
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha List With Previous Employer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   510
            Left            =   -65530
            TabIndex        =   177
            Top             =   2115
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alpha List Terminated "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   -65530
            TabIndex        =   176
            Top             =   1530
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "BIR 2316"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Index           =   12
            Left            =   -65530
            TabIndex        =   175
            Top             =   4755
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label labSched 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Schedule of Deduction"
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
            Index           =   13
            Left            =   -69010
            TabIndex        =   174
            Top             =   5490
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label81 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Listing "
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
            Left            =   -69010
            TabIndex        =   156
            Top             =   1455
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Employee Listing Wage"
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
            Left            =   -69010
            TabIndex        =   155
            Top             =   2010
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Resigned Employee Listing"
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
            Left            =   -69010
            TabIndex        =   154
            Top             =   2580
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blank Employee Information  Sheet "
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
            Left            =   -69010
            TabIndex        =   153
            Top             =   900
            Visible         =   0   'False
            Width           =   3000
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Loan Reports"
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
            Left            =   -69010
            TabIndex        =   152
            Top             =   3090
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Daily Log Reports"
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
            Left            =   -69010
            TabIndex        =   151
            Top             =   3600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label labSched 
            BackStyle       =   0  'Transparent
            Caption         =   "Leave Summary Report"
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
            Height          =   330
            Index           =   10
            Left            =   -69010
            TabIndex        =   150
            Top             =   4170
            Visible         =   0   'False
            Width           =   3045
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Yearly Individual Payroll Summary "
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
            Left            =   -69010
            TabIndex        =   149
            Top             =   4680
            Visible         =   0   'False
            Width           =   2910
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RF1 - Employer Quarterly Remittance"
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
            Left            =   -68740
            TabIndex        =   140
            Top             =   990
            Visible         =   0   'False
            Width           =   3120
         End
         Begin VB.Label Label76 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "R3 - SSS Contribution Collection List"
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
            Left            =   -68740
            TabIndex        =   139
            Top             =   1770
            Visible         =   0   'False
            Width           =   3060
         End
         Begin VB.Label Label83 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quarterly Loans"
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
            Left            =   -68740
            TabIndex        =   138
            Top             =   2580
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label67 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PHILHEALTH Contributions"
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
            Left            =   -68740
            TabIndex        =   134
            Top             =   1005
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.Label Label68 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SSS Contributions"
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
            Left            =   -68740
            TabIndex        =   133
            Top             =   1740
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.Label Label69 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PAG-IBIG Contributions"
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
            Left            =   -68740
            TabIndex        =   132
            Top             =   2475
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TAX Withheld"
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
            Left            =   -68740
            TabIndex        =   131
            Top             =   3240
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOANS Remitted"
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
            Left            =   -68740
            TabIndex        =   130
            Top             =   3960
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Daily Time Record"
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
            Left            =   990
            TabIndex        =   124
            Top             =   3525
            Width           =   1980
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Print Payroll ATM Advice"
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
            Left            =   990
            TabIndex        =   123
            Top             =   840
            Width           =   2070
         End
         Begin VB.Label Label63 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Payroll PaySlips"
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
            Left            =   990
            TabIndex        =   122
            Top             =   1395
            Width           =   1815
         End
         Begin VB.Label Label64 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Deductions Breakdown"
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
            Left            =   990
            TabIndex        =   121
            Top             =   1935
            Width           =   2430
         End
         Begin VB.Label Label65 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Overtime Breakdown "
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
            Left            =   990
            TabIndex        =   120
            Top             =   2430
            Width           =   2295
         End
         Begin VB.Label Label66 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Adjustment Breakdown"
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
            Left            =   990
            TabIndex        =   119
            Top             =   3000
            Width           =   2445
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Commission Breakdown"
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
            Index           =   0
            Left            =   990
            TabIndex        =   118
            Top             =   5655
            Width           =   2535
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Allowance Computation"
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
            Left            =   990
            TabIndex        =   117
            Top             =   4035
            Width           =   2010
         End
         Begin VB.Label Label53 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Define 201 Report"
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
            Index           =   0
            Left            =   990
            TabIndex        =   116
            Top             =   4605
            Width           =   1965
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Print Loans Breakdown"
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
            Index           =   1
            Left            =   990
            TabIndex        =   115
            Top             =   5115
            Width           =   1995
         End
      End
      Begin VB.Label Label4 
         Caption         =   "Employee Payroll Set Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   -65110
         TabIndex        =   198
         Top             =   5940
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -61930
         TabIndex        =   93
         Top             =   2415
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   24
         Left            =   -68950
         TabIndex        =   92
         Top             =   900
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WITHHOLDING TAX Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   27
         Left            =   -68950
         TabIndex        =   91
         Top             =   3105
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PHILHEALTH Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   25
         Left            =   -68950
         TabIndex        =   90
         Top             =   1665
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAGIBIG Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   26
         Left            =   -68950
         TabIndex        =   89
         Top             =   2400
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Schedule of Deductions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   29
         Left            =   -68950
         TabIndex        =   88
         Top             =   4530
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Shift Set Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   30
         Left            =   -68950
         TabIndex        =   87
         Top             =   5235
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   28
         Left            =   -68950
         TabIndex        =   86
         Top             =   3810
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   31
         Left            =   -68920
         TabIndex        =   85
         Top             =   5940
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Shift Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   37
         Left            =   -61930
         TabIndex        =   84
         Top             =   900
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   33
         Left            =   -65080
         TabIndex        =   83
         Top             =   1665
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   35
         Left            =   -65080
         TabIndex        =   82
         Top             =   3090
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -61930
         TabIndex        =   81
         Top             =   1680
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   34
         Left            =   -65080
         TabIndex        =   80
         Top             =   2385
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   36
         Left            =   -65080
         TabIndex        =   79
         Top             =   3825
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department Codes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   32
         Left            =   -65080
         TabIndex        =   78
         Top             =   900
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Holiday Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   38
         Left            =   -65080
         TabIndex        =   77
         Top             =   4545
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Leave Maintenance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   39
         Left            =   -65080
         TabIndex        =   76
         Top             =   5220
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print Payroll Sheet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   8
         Left            =   1020
         TabIndex        =   26
         Top             =   6150
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate Payroll"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   7
         Left            =   1020
         TabIndex        =   25
         Top             =   5490
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update Attendance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   1020
         TabIndex        =   24
         Top             =   4845
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Applicant Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   1020
         TabIndex        =   23
         Top             =   4185
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   1020
         TabIndex        =   22
         Top             =   3450
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Regular/Probationary 201 File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   1020
         TabIndex        =   21
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contractual 201 File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   1020
         TabIndex        =   20
         Top             =   1485
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Allowance-Based 201 File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   1020
         TabIndex        =   19
         Top             =   2145
         Width           =   2445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confidential 201 File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   1020
         TabIndex        =   18
         Top             =   2790
         Width           =   1950
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Signatories and Headers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -68860
         TabIndex        =   8
         Top             =   1725
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Maintenance "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -68860
         TabIndex        =   7
         Top             =   2445
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Profile/Pay Period Set Up (F12)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   -68860
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   3900
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generate 13th Month Pay"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   -68860
         TabIndex        =   5
         Top             =   3135
         Visible         =   0   'False
         Width           =   2400
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
    Dim RSTMP                                                         As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfTheresEmployee = True
    Else
        CheckIfTheresEmployee = False
    End If
    Set RSTMP = Nothing
End Function

Sub FormExistsShow(FRMx As Form)
    On Error GoTo Errorcode
    Dim m_Exists                                                      As Boolean
    Dim FRM                                                           As Form
    For Each FRM In Forms
        If (UCase(FRM.NAME) = UCase(FRMx.NAME)) Then
            m_Exists = True
        End If
    Next
    Set FRM = Nothing
    If m_Exists = True Then
        FRMx.WindowState = 0
        FRMx.ZOrder 0
    End If
    Exit Sub
Errorcode:
    Err.Clear
End Sub

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
    'If Module_Access(LOGID, "EMPLOYEE TIME CARD", "DATA ENTRY") = False Then Exit Sub
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
    'frmHRMS_Ledger2.Show
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
    'frmSETUP_Overtime.Show
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
    frmHRMSFiles_leaveMaster.Show
    'frmHRMS_Leave_Codes.Show
End Sub

Private Sub Command10_Click()
    If Module_Access(LOGID, "REMINDERS", "DATA ENTRY") = False Then Exit Sub
    frmSMIS_Log_Reminder.Show
End Sub

Private Sub cmdReport_Sched_PAGIBIG_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF PAGIBIG PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDPAGIBIG"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_TAXWHELD_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF TAX WITHHELD", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDTAX"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_OVERTIMEPAY_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF OVERTIME PAY", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDOVERTIME"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_COMMISSION_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF COMMISSION", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDCOMMISSION"
    frmHRMSYearly.Show
End Sub

Private Sub Command11_Click()
    If Module_Access(LOGID, "LEAVE CONVERSION", "PROCESSING") = False Then Exit Sub
    'frmHRMS_Trans_LeaveConvertion.Show
    'frmHRMS_Leave.Show
End Sub

Private Sub Command12_Click()
    If Module_Access(LOGID, "REPORT ALLOWANCE COMPUTATION", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "Allowance Computation Report"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub Command13_Click()
    If Module_Access(LOGID, "LEAVE SUMMARY REPORT", "REPORTS") = False Then Exit Sub
    'FORMYEARLYREQUEST = "LEAVESUMMARY"
    'frmHRMSYearly.Show
     Screen.MousePointer = 11

    Dim RSTMP                                           As New ADODB.Recordset
    Dim XXX                                             As String
    Dim xlApp                                           As Excel.Application
    Dim xlbook                                          As Excel.Workbook
    Dim xlsheet                                         As Excel.Worksheet
    Dim cmd                                             As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.NamedParameters = True
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SP_LEAVE_SUMMARY"
    cmd.ActiveConnection = gconDMIS
    Set RSTMP = cmd.Execute
    
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        If Len(Dir(HRMS_REPORT_PATH & "Leave.xlt")) = 0 Then
            MessagePop InfoStop, "Error", "Leave_summary.xlt cannot be found in server Report Path." & vbCrLf & "Please contact I.T Department", vbInformation
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Set xlApp = New Excel.Application
        Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "Leave.xlt")
        Set xlsheet = xlbook.Worksheets(1)
        
        xlsheet.Range("A7").CopyFromRecordset RSTMP
        xlApp.Visible = True
        If Not xlbook Is Nothing Then
            Set xlbook = Nothing
            Set xlApp = Nothing
        End If
        Set xlApp = Nothing
    Else
        Call ShowNoRecord
    End If
    Set RSTMP = Nothing
    Screen.MousePointer = 0

        
End Sub

Private Sub Command14_Click()
    If Module_Access(LOGID, "LOAN REPORT", "REPORTS") = False Then Exit Sub
    frmHRMSLoansBalances.Show
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
    FormYearlyRequest = "SCHEDPAYROLL"
    frmHRMSYearly.Show
End Sub

Private Sub Command16_Click()
    If Module_Access(LOGID, "REPORT DAILY LOG REPORT", "REPORTS") = False Then Exit Sub
    frmHRMS_PrintDailyLog.Show
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
    FormYearlyRequest = "SCHEDCOMMISSIONTAX"
    frmHRMSYearly.Show
End Sub

Private Sub Command19_Click()
    If Module_Access(LOGID, "FILES LEAVE MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    frmHRMS_Leave_Maintenance.Show
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
    frmHRMSPHMonthly.Caption = "PHILHEALTH MONTHLY REMITTANCE"
    frmHRMSPHMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command22_Click()
    If Module_Access(LOGID, "REPORT WITHHOLDING TAX MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPHMonthly.Caption = "REPORT WITHHOLDING TAX MONTHLY REMITTANCE"
    frmHRMSPHMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command23_Click()
    If Module_Access(LOGID, "REPORT SSS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPHMonthly.Caption = "REPORT SSS MONTHLY REMITTANCE"
    frmHRMSPHMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command24_Click()
    If Module_Access(LOGID, "REPORT PAG-IBIG MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPHMonthly.Caption = "REPORT PAG-IBIG MONTHLY REMITTANCE"
    frmHRMSPHMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command25_Click()
    If Module_Access(LOGID, "REPORT LOANS MONTHLY REMITTANCE", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSLoansMonthly.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command26_Click()
    If Module_Access(LOGID, "REPORT RF1-EMPLOYER QUARTERLY REMITTANCE", "REPORTS") = False Then Exit Sub
    frmHRMS_GovermentForms.Show
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
    frmHRMS_Reports_R3.Show
    
End Sub

Private Sub cmdReports_AlphaListTerminated_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST TERMINATED", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "ALTERMINATED"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReports_ALphalistWPrevEmployer_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST WITH PREVIOUS EMP", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "ALWITHEMP"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReports_AlphalistwioutPrevEmployer_Click()
    If Module_Access(LOGID, "REPORT ALPHALIST W/OUT PREVIOUS EMPLOYEE", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "ALWITHNOEMP"
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

Private Sub Command31_Click()
    If Module_Access(LOGID, "ALPHALIST REPORT", "REPORTS") = False Then Exit Sub
    frmHRMS_AlphaList2008.Show
End Sub

Private Sub Command32_Click()
    If Module_Access(LOGID, "REPORT 2316", "REPORTS") = False Then Exit Sub
    frmHRMS_2316.Show
End Sub

Private Sub Command33_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF DEDUCTION", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDDEDUCTION"
    frmHRMSYearly.Show
End Sub

Private Sub Command34_Click()
    If Module_Access(LOGID, "REPORTS YEARLY INDIVIDUAL", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "YEARLYSCHEDPAYROLL"
    frmHRMSYearly.Show
End Sub

Private Sub Command35_Click()
    If Module_Access(LOGID, "REPORT USER DEFINE REPORTS", "REPORTS") = False Then Exit Sub
    frmHRMS_Reports_201Reports.Show
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
    FormYearlyRequest = "SCHEDTAXDUEREFUND"
    frmHRMSYearly.Show
End Sub

Private Sub Command38_Click()
    If Module_Access(LOGID, "PROCESS 13TH MONTH PAY", "PROCESSING") = False Then Exit Sub
    frmHRMS_Process13th.Show
End Sub

Private Sub Command39_Click()
    If Module_Access(LOGID, "PRINT LOAN DEDUCTION BREAKDOWN", "REPORTS") = False Then Exit Sub
    frmHRMS_Reports_LaonBreakdown.Show
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

 

Private Sub Command40_Click()
    Screen.MousePointer = 11
    'frmHRMS_Dtrweekly.Show
    frmHRMS_Reports_MonthlyTimeRecord.Show
    Screen.MousePointer = 0
End Sub

Private Sub Command41_Click()
Screen.MousePointer = 11
frmHRMS_Deduction_Report.Show
frmHRMS_Deduction_Report.chklate.Value = 1
Screen.MousePointer = 0
End Sub

Private Sub Command42_Click()
FrmHRMS_Employee_Setup.Show
End Sub

Private Sub Command43_Click()
If Module_Access(LOGID, "ATM SUMMARY LIST", "REPORTS") = False Then Exit Sub
    frmHRMSPRINT_BreakDown.Caption = "ATM SUMMARY LIST"
    frmHRMSPRINT_BreakDown.Show
End Sub

Private Sub Command44_Click()
    If Module_Access(LOGID, "REPORT PRINT PAYROLL SUMMARY", "REPORTS") = False Then Exit Sub
    Screen.MousePointer = 11
    frmHRMSPrintPayroll.Show
    frmHRMSPrintPayroll.chkInclude.Value = 1
    frmHRMSPrintPayroll.chkPaySlip.Visible = False
    Screen.MousePointer = 0
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
    FormYearlyRequest = "SCHEDSSS"
    frmHRMSYearly.Show
End Sub

Private Sub cmdReport_Sched_PHIC_Click()
    If Module_Access(LOGID, "REPORT SCHEDULE OF PHILHEALTH PREMIUM CONTRIBUTION", "REPORTS") = False Then Exit Sub
    FormYearlyRequest = "SCHEDPHIC"
    frmHRMSYearly.Show
End Sub

Private Sub Command8_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN OTHER DEDUCTIONS", "DATA ENTRY") = False Then Exit Sub
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
    If Module_Access(LOGID, "ATM SUMMARY", "SYSTEM") = False Then Exit Sub
    frmHRMS_ATM_Summary.Show
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    TabControl3.SelectedItem = 0
    'SSTab1.Tab = 0
    TabControl1.SelectedItem = 0
    Call GetThePayrollCode
End Sub
'
'Private Sub Label1_Click(Index As Integer)
'    If Index = 16 Then frmTaxComputer.Show
'End Sub
Private Sub TabControl1_SelectedChanged(ByVal ITEM As XtremeSuiteControls.ITabControlItem)

End Sub
