VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F985F9B0-A252-46B5-A444-E023A386B6FE}#1.0#0"; "wizBox.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMSEmpInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Information"
   ClientHeight    =   8700
   ClientLeft      =   345
   ClientTop       =   915
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00D8E9EC&
   Icon            =   "EmpInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8700
   ScaleWidth      =   11700
   Begin VB.PictureBox picOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4725
      Left            =   9150
      ScaleHeight     =   4695
      ScaleWidth      =   2445
      TabIndex        =   138
      Top             =   2580
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FCE2CF&
         Caption         =   "EMPLOYEE LEVEL"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   195
         Top             =   4260
         Width           =   2295
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FCE2CF&
         Caption         =   "RESIGNATION"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   3930
         Width           =   2295
      End
      Begin VB.CommandButton optViewOtherInfo 
         BackColor       =   &H00FCE2CF&
         Caption         =   "OTHER INFO"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton cmdLedger 
         BackColor       =   &H00FCE2CF&
         Caption         =   "LEDGER"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   144
         ToolTipText     =   "View Employee's Ledger"
         Top             =   3270
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeductions 
         BackColor       =   &H00FCE2CF&
         Caption         =   "DEDUCTIONS"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "View Employee's Deductions"
         Top             =   2940
         Width           =   2295
      End
      Begin VB.CommandButton cmdSalaryAdvance 
         BackColor       =   &H00FCE2CF&
         Caption         =   "COMMISSION"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "View Employee's Salary Advance"
         Top             =   2610
         Width           =   2295
      End
      Begin VB.CommandButton cmdOvertime 
         BackColor       =   &H00FCE2CF&
         Caption         =   "OVERTIME"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "View Employee's Overtime"
         Top             =   2280
         Width           =   2295
      End
      Begin VB.CommandButton cmdAdjustment 
         BackColor       =   &H00FCE2CF&
         Caption         =   "ADJUSTMENTS"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "View Adjustment"
         Top             =   1950
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FCE2CF&
         Caption         =   "SET SALARY"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   1620
         Width           =   2295
      End
      Begin VB.CommandButton Command5 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   147
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdAttendance 
         BackColor       =   &H00FCE2CF&
         Caption         =   "ATTENDANCE"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "View Employee's Attendance"
         Top             =   1290
         Width           =   2295
      End
      Begin VB.CommandButton cmdSAType 
         BackColor       =   &H00FCE2CF&
         Caption         =   "SET SA POSITION TYPE"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "View Employee's Attendance"
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdTechPost 
         BackColor       =   &H00FCE2CF&
         Caption         =   "SET TECH POSITION TYPE"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "View Employee's Attendance"
         Top             =   630
         Width           =   2295
      End
      Begin VB.CommandButton cmdEnrollFingerPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENROLL FINGERPRINT"
         Height          =   345
         Left            =   90
         MaskColor       =   &H00C00000&
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "View Employee's Attendance"
         Top             =   300
         Width           =   2295
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Width           =   2985
         _Version        =   655364
         _ExtentX        =   5265
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Options"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox PIC_EMPLEVEL 
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
      Height          =   2865
      Left            =   5430
      ScaleHeight     =   2835
      ScaleWidth      =   2775
      TabIndex        =   187
      Top             =   2760
      Visible         =   0   'False
      Width           =   2805
      Begin VB.OptionButton OPT_CONFI 
         Caption         =   "Confidential"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   1350
         Width           =   2565
      End
      Begin VB.OptionButton OPT_REG 
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   1020
         Width           =   2565
      End
      Begin VB.OptionButton OPT_CONTRACT 
         Caption         =   "Contractual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   192
         Top             =   690
         Width           =   2565
      End
      Begin VB.OptionButton OPT_ALLOWANCE 
         Caption         =   "Allowance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   360
         Width           =   2565
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Cancel"
         Height          =   765
         Left            =   1980
         MouseIcon       =   "EmpInfo.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Cancel Entry"
         Top             =   1920
         Width           =   705
      End
      Begin VB.CommandButton Command13 
         Caption         =   "&Save"
         Height          =   765
         Left            =   1290
         MouseIcon       =   "EmpInfo.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   189
         ToolTipText     =   "Save Entry"
         Top             =   1920
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   2
         Left            =   0
         TabIndex        =   190
         Top             =   0
         Width           =   6615
         _Version        =   655364
         _ExtentX        =   11668
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "EMPLOYEE LEVEL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   8421504
      End
   End
   Begin VB.PictureBox picSalary 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   2655
      Left            =   4110
      ScaleHeight     =   2625
      ScaleWidth      =   3465
      TabIndex        =   130
      Top             =   3023
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ComboBox cboSALCODE 
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
         Left            =   180
         TabIndex        =   172
         Text            =   "cboSALCODE"
         Top             =   2010
         Width           =   1605
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   675
         Left            =   2610
         MouseIcon       =   "EmpInfo.frx":0C3C
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Cancel Entry"
         Top             =   1830
         Width           =   705
      End
      Begin VB.TextBox txtALLOWANCE 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   180
         TabIndex        =   132
         Top             =   1350
         Width           =   3105
      End
      Begin VB.TextBox txtBASICSALARY 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   180
         TabIndex        =   131
         Top             =   645
         Width           =   3105
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Save"
         Height          =   675
         Left            =   1920
         MouseIcon       =   "EmpInfo.frx":10CC
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":121E
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Save Entry"
         Top             =   1830
         Width           =   705
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Grade Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   173
         Top             =   1740
         Width           =   1650
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Width           =   3495
         _Version        =   655364
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "SET SALARY"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Allowance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   134
         Top             =   1095
         Width           =   885
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   133
         Top             =   390
         Width           =   1110
      End
   End
   Begin VB.PictureBox picTECH 
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
      Height          =   2535
      Left            =   3480
      ScaleHeight     =   2505
      ScaleWidth      =   5205
      TabIndex        =   174
      Top             =   2790
      Visible         =   0   'False
      Width           =   5235
      Begin VB.CheckBox chk1 
         Caption         =   "GJ Technician Master"
         Height          =   255
         Left            =   150
         TabIndex        =   185
         Top             =   420
         Width           =   1935
      End
      Begin VB.CheckBox chk2 
         Caption         =   "GJ Technician Expert"
         Height          =   255
         Left            =   150
         TabIndex        =   184
         Top             =   690
         Width           =   1845
      End
      Begin VB.CheckBox chk3 
         Caption         =   "GJ Technician Certified"
         Height          =   255
         Left            =   150
         TabIndex        =   183
         Top             =   990
         Width           =   2055
      End
      Begin VB.CheckBox chk5 
         Caption         =   "GJ Technician New"
         Height          =   255
         Left            =   2310
         TabIndex        =   182
         Top             =   420
         Width           =   1815
      End
      Begin VB.CheckBox chk6 
         Caption         =   "BP In-House Technician Paint"
         Height          =   255
         Left            =   2310
         TabIndex        =   181
         Top             =   690
         Width           =   2865
      End
      Begin VB.CheckBox chk7 
         Caption         =   "BP In-HouseTechnician Tinsmist"
         Height          =   255
         Left            =   2310
         TabIndex        =   180
         Top             =   990
         Width           =   2955
      End
      Begin VB.CheckBox chk8 
         Caption         =   "Contractor Technician"
         Height          =   255
         Left            =   2310
         TabIndex        =   179
         Top             =   1260
         Width           =   2235
      End
      Begin VB.CheckBox chk4 
         Caption         =   "Quick Service Technician"
         Height          =   255
         Left            =   150
         TabIndex        =   178
         Top             =   1260
         Width           =   2955
      End
      Begin VB.CommandButton Command11 
         Caption         =   "&Cancel"
         Height          =   765
         Left            =   4410
         MouseIcon       =   "EmpInfo.frx":156E
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":16C0
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Cancel Entry"
         Top             =   1650
         Width           =   705
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Save"
         Height          =   765
         Left            =   3720
         MouseIcon       =   "EmpInfo.frx":19FE
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":1B50
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Save Entry"
         Top             =   1650
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   177
         Top             =   0
         Width           =   6615
         _Version        =   655364
         _ExtentX        =   11668
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "SET TECHNICIAN TYPE"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8421504
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picTemplate2 
      Height          =   1125
      Left            =   3600
      ScaleHeight     =   1065
      ScaleWidth      =   1125
      TabIndex        =   168
      Top             =   7140
      Width           =   1185
      Begin VB.Label labFP2 
         Alignment       =   2  'Center
         Caption         =   "Register you Finger Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   120
         TabIndex        =   170
         Top             =   150
         Width           =   975
      End
      Begin VB.Image imgTemplate2 
         Height          =   1065
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox picTemplate1 
      Height          =   1125
      Left            =   2340
      ScaleHeight     =   1065
      ScaleWidth      =   1125
      TabIndex        =   167
      Top             =   7140
      Width           =   1185
      Begin VB.Label labFP1 
         Alignment       =   2  'Center
         Caption         =   "Register you Finger Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   885
         Left            =   90
         TabIndex        =   169
         Top             =   150
         Width           =   975
      End
      Begin VB.Image imgTemplate1 
         Height          =   1065
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      ScaleHeight     =   855
      ScaleWidth      =   7830
      TabIndex        =   110
      Top             =   7350
      Width           =   7830
      Begin VB.CommandButton Command4 
         Caption         =   "&Options"
         Height          =   795
         Left            =   6690
         MouseIcon       =   "EmpInfo.frx":1EA0
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":1FF2
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   6000
         MouseIcon       =   "EmpInfo.frx":2074
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":21C6
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   5310
         MouseIcon       =   "EmpInfo.frx":252C
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":267E
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   4620
         MouseIcon       =   "EmpInfo.frx":29E4
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":2B36
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   3930
         MouseIcon       =   "EmpInfo.frx":2E61
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":2FB3
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3240
         MouseIcon       =   "EmpInfo.frx":330F
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":3461
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   2550
         MouseIcon       =   "EmpInfo.frx":3774
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":38C6
         Style           =   1  'Graphical
         TabIndex        =   113
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   1860
         MouseIcon       =   "EmpInfo.frx":3BC0
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":3D12
         Style           =   1  'Graphical
         TabIndex        =   112
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   1170
         MouseIcon       =   "EmpInfo.frx":406A
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":41BC
         Style           =   1  'Graphical
         TabIndex        =   111
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   345
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11610
      TabIndex        =   125
      Top             =   8340
      Width           =   11640
      Begin VB.Label lblBasicSalary 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BASIC SALARY NOT SET"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   30
         TabIndex        =   150
         Top             =   30
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.Label lblAllowance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ALLOWANCE NOT SET"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2550
         TabIndex        =   149
         Top             =   30
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label labSHIFTSCHED 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   9420
         TabIndex        =   128
         Top             =   30
         Width           =   2160
      End
      Begin VB.Label labSHIFTCODE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SHIFT1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   8400
         TabIndex        =   127
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label Label42 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current Employee Shift Schedule :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   4920
         TabIndex        =   126
         Top             =   30
         Width           =   3465
      End
   End
   Begin VB.OptionButton optInActiveEmployee 
      Caption         =   "In-Active Employees"
      Height          =   315
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "View In-Active Employees"
      Top             =   7980
      Width           =   2145
   End
   Begin VB.OptionButton optActiveEmployee 
      Caption         =   "Active Employees"
      Height          =   285
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "View Active Employees"
      Top             =   7710
      Value           =   -1  'True
      Width           =   2145
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   2835
      Left            =   2400
      ScaleHeight     =   2835
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.ComboBox cboBlood 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   3780
         TabIndex        =   24
         Text            =   "cboBlood"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.OptionButton optInActive 
         Caption         =   "In - Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   1455
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1170
         TabIndex        =   22
         Text            =   "cboStatus"
         Top             =   1620
         Width           =   1185
      End
      Begin VB.ComboBox cboSex 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2880
         TabIndex        =   18
         Text            =   "cboSex"
         Top             =   1200
         Width           =   915
      End
      Begin VB.TextBox txtSSSNo 
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
         Left            =   3780
         TabIndex        =   33
         Top             =   2085
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1170
         TabIndex        =   12
         Top             =   810
         Width           =   4905
      End
      Begin VB.TextBox txtMiddleName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7410
         TabIndex        =   9
         Top             =   420
         Width           =   1725
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4110
         TabIndex        =   7
         Top             =   420
         Width           =   1935
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1170
         TabIndex        =   6
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtTelephone 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7410
         TabIndex        =   13
         Top             =   810
         Width           =   1725
      End
      Begin MSMask.MaskEdBox txtCitizen 
         Height          =   345
         Left            =   1170
         TabIndex        =   37
         Top             =   2475
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBirthPlace 
         Height          =   345
         Left            =   5220
         TabIndex        =   19
         Top             =   1200
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHeight 
         Height          =   315
         Left            =   6120
         TabIndex        =   26
         Top             =   1650
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtWeight 
         Height          =   315
         Left            =   8010
         TabIndex        =   29
         Top             =   1650
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBirthDate 
         Height          =   345
         Left            =   1170
         TabIndex        =   16
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtReligion 
         Height          =   345
         Left            =   1170
         TabIndex        =   32
         Top             =   2085
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPHNo 
         Height          =   345
         Left            =   7290
         TabIndex        =   36
         Top             =   2085
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTINNo 
         Height          =   345
         Left            =   3780
         TabIndex        =   40
         Top             =   2475
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPagIbigNo 
         Height          =   345
         Left            =   7290
         TabIndex        =   42
         Top             =   2460
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optActive 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtEmpNo 
         Height          =   375
         Left            =   1170
         TabIndex        =   4
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "lbs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   8820
         TabIndex        =   30
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6810
         TabIndex        =   27
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Blood Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2520
         TabIndex        =   23
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   9570
         X2              =   0
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Pag-Ibig No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   41
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TIN No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3090
         TabIndex        =   38
         Top             =   2520
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PHIC No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5880
         TabIndex        =   35
         Top             =   2130
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SSS No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2970
         TabIndex        =   34
         Top             =   2130
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Religion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   31
         Top             =   2130
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         TabIndex        =   11
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6210
         TabIndex        =   10
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3090
         TabIndex        =   8
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         TabIndex        =   5
         Top             =   450
         Width           =   945
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Citizenship"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         TabIndex        =   39
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2460
         TabIndex        =   17
         Top             =   1260
         Width           =   435
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6210
         TabIndex        =   14
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Place"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3930
         TabIndex        =   20
         Top             =   1215
         Width           =   1275
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   60
         TabIndex        =   15
         Top             =   1230
         Width           =   870
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   25
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7290
         TabIndex        =   28
         Top             =   1680
         Width           =   705
      End
   End
   Begin MSComctlLib.ImageList imlEmpInfo 
      Left            =   165
      Top             =   7710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog comDialogEmpInfo 
      Left            =   1575
      Top             =   7845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport rptEmpInfo 
      Left            =   1365
      Top             =   7845
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   7605
      Left            =   0
      ScaleHeight     =   7605
      ScaleWidth      =   2355
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   0
      Width           =   2355
      Begin wizBox.Box Box4 
         Height          =   1785
         Left            =   210
         Top             =   90
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   3149
      End
      Begin VB.CommandButton cmdDeletePicture 
         Caption         =   "Delete Picture"
         Height          =   435
         Left            =   1140
         TabIndex        =   45
         ToolTipText     =   "Delete Picture of Employee"
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton cmdInsertPic 
         Caption         =   "Insert Picture"
         Height          =   435
         Left            =   210
         TabIndex        =   44
         ToolTipText     =   "Insert Picture of Employee"
         Top             =   1920
         Width           =   945
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   30
         MaxLength       =   35
         TabIndex        =   46
         Top             =   2370
         Width           =   2235
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   4815
         Left            =   0
         TabIndex        =   47
         Top             =   2760
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   8493
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
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "EmpInfo.frx":451B
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.OptionButton optViewBalances 
         Caption         =   "View Employee Balances"
         Height          =   345
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "View Employee Balances"
         Top             =   3300
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Image imgDispPic 
         Height          =   1635
         Left            =   270
         Stretch         =   -1  'True
         Top             =   150
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   60
      Picture         =   "EmpInfo.frx":467D
      ScaleHeight     =   5745
      ScaleWidth      =   2205
      TabIndex        =   49
      Top             =   90
      Width           =   2265
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   1575
      Left            =   2400
      ScaleHeight     =   1575
      ScaleWidth      =   9195
      TabIndex        =   84
      Top             =   5475
      Width           =   9195
      Begin MSMask.MaskEdBox txtSpouse 
         Height          =   315
         Left            =   960
         TabIndex        =   87
         Top             =   90
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSOccupation 
         Height          =   315
         Left            =   7110
         TabIndex        =   90
         Top             =   90
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFather 
         Height          =   315
         Left            =   960
         TabIndex        =   92
         Top             =   450
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMother 
         Height          =   315
         Left            =   960
         TabIndex        =   98
         Top             =   810
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPerson 
         Height          =   345
         Left            =   1500
         TabIndex        =   103
         Top             =   1200
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFOccupation 
         Height          =   315
         Left            =   7110
         TabIndex        =   96
         Top             =   450
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMOccupation 
         Height          =   315
         Left            =   7110
         TabIndex        =   102
         Top             =   810
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSpouseAge 
         Height          =   315
         Left            =   5160
         TabIndex        =   88
         Top             =   90
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFatherAge 
         Height          =   315
         Left            =   5160
         TabIndex        =   93
         Top             =   450
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMotherAge 
         Height          =   315
         Left            =   5160
         TabIndex        =   99
         Top             =   810
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRelation 
         Height          =   345
         Left            =   5160
         TabIndex        =   106
         Top             =   1170
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRelTelNo 
         Height          =   315
         Left            =   7980
         TabIndex        =   108
         Top             =   1200
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   9525
         X2              =   -45
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7080
         TabIndex        =   107
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   105
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   100
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   94
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4680
         TabIndex        =   85
         Top             =   90
         Width           =   555
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5820
         TabIndex        =   101
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5820
         TabIndex        =   95
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Person to Notify"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   104
         Top             =   1230
         Width           =   1725
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mother"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   97
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Father"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   91
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5820
         TabIndex        =   89
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Spouse"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   86
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10170
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   119
      Top             =   7350
      Width           =   1440
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   720
         MouseIcon       =   "EmpInfo.frx":183DA
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":1852C
         Style           =   1  'Graphical
         TabIndex        =   120
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "EmpInfo.frx":1886A
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":189BC
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   2640
      Left            =   2400
      ScaleHeight     =   2640
      ScaleWidth      =   9195
      TabIndex        =   53
      Top             =   2820
      Width           =   9195
      Begin VB.TextBox txtResigned 
         Height          =   315
         Left            =   7320
         TabIndex        =   163
         Top             =   900
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   5520
         Top             =   240
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "View"
         Height          =   285
         Left            =   8460
         TabIndex        =   129
         Top             =   1410
         Width           =   705
      End
      Begin VB.PictureBox picPosition 
         Height          =   345
         Left            =   1140
         ScaleHeight     =   285
         ScaleWidth      =   2805
         TabIndex        =   62
         Top             =   900
         Width           =   2865
         Begin VB.CommandButton cmdPosi 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2460
            TabIndex        =   63
            Top             =   0
            Width           =   345
         End
         Begin MSForms.CheckBox chkPos 
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   67
            Top             =   840
            Width           =   2295
            BackColor       =   16777215
            ForeColor       =   0
            DisplayStyle    =   4
            Size            =   "4048;556"
            Value           =   "0"
            Caption         =   "Account Executive"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPos 
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   66
            Top             =   570
            Width           =   2295
            BackColor       =   16777215
            ForeColor       =   0
            DisplayStyle    =   4
            Size            =   "4048;556"
            Value           =   "0"
            Caption         =   "Service Advisor"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPos 
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   65
            Top             =   270
            Width           =   2295
            BackColor       =   16777215
            ForeColor       =   0
            DisplayStyle    =   4
            Size            =   "4048;556"
            Value           =   "0"
            Caption         =   "Parts Salesman"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.CheckBox chkPos 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   64
            Top             =   0
            Width           =   2565
            BackColor       =   16777215
            ForeColor       =   0
            DisplayStyle    =   4
            Size            =   "4524;556"
            Value           =   "0"
            Caption         =   "Technician"
            FontName        =   "Arial"
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtCompanyName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   80
         Text            =   "Please Enter Previous Company"
         Top             =   2280
         Width           =   8985
      End
      Begin VB.ComboBox cboPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1140
         TabIndex        =   59
         Text            =   "cboPosition"
         Top             =   502
         Width           =   3585
      End
      Begin VB.ComboBox cboPayroll 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4980
         TabIndex        =   71
         Text            =   "cboPayroll"
         Top             =   1440
         Width           =   1065
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Subject to Cola"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3390
         TabIndex        =   83
         Top             =   3000
         Width           =   1965
      End
      Begin VB.ComboBox cboExStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4980
         TabIndex        =   74
         Text            =   "cboExStatus"
         Top             =   1800
         Width           =   1065
      End
      Begin VB.CheckBox chkWithPrevious 
         Caption         =   "w/ Previous Employer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   81
         Top             =   1890
         Width           =   2475
      End
      Begin VB.ComboBox cboSalaryGradeLevel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4050
         TabIndex        =   72
         Text            =   "cboSalaryGradeLevel"
         Top             =   900
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.ComboBox cboDeptName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1140
         TabIndex        =   54
         Text            =   "cboDeptName"
         Top             =   112
         Width           =   3585
      End
      Begin VB.ComboBox cboEmpStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7320
         TabIndex        =   78
         Text            =   "cboEmpStatus"
         Top             =   1800
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtDateHired 
         Height          =   345
         Left            =   7320
         TabIndex        =   61
         Top             =   510
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtAcctNo 
         Height          =   345
         Left            =   7320
         TabIndex        =   56
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCOLA_RATE 
         Height          =   315
         Left            =   5340
         TabIndex        =   82
         Top             =   2940
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboPayrollGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "EmpInfo.frx":18D0C
         Left            =   7320
         List            =   "EmpInfo.frx":18D0E
         TabIndex        =   77
         Text            =   "cboPayrollGroup"
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "****"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   315
         Left            =   120
         TabIndex        =   151
         Top             =   1410
         Width           =   3375
      End
      Begin VB.Label LABSLEVEL 
         Alignment       =   2  'Center
         Caption         =   "xxxxx"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   4020
         TabIndex        =   124
         Top             =   930
         Width           =   1875
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         X1              =   9570
         X2              =   0
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   6060
         TabIndex        =   76
         Top             =   1440
         Width           =   1230
      End
      Begin VB.Label lab4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sys Function"
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
         Height          =   435
         Left            =   30
         TabIndex        =   69
         Top             =   840
         Width           =   1035
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderStyle     =   6  'Inside Solid
         X1              =   9570
         X2              =   0
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Deduction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   3690
         TabIndex        =   70
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Exemption Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2820
         TabIndex        =   75
         Top             =   1845
         Width           =   2085
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Acct. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6240
         TabIndex        =   57
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label Label39 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   55
         Top             =   150
         Width           =   1845
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6090
         TabIndex        =   79
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   58
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Code"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   73
         Top             =   2280
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Hired"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5760
         TabIndex        =   60
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Resigned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5850
         TabIndex        =   68
         Top             =   930
         Width           =   1395
      End
   End
   Begin VB.PictureBox picRegi 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   5100
      ScaleHeight     =   4545
      ScaleWidth      =   4035
      TabIndex        =   152
      Top             =   1320
      Visible         =   0   'False
      Width           =   4065
      Begin MSComCtl2.DTPicker txtReg_Resinged 
         Height          =   375
         Left            =   180
         TabIndex        =   158
         Top             =   1350
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   55115777
         CurrentDate     =   39582
      End
      Begin VB.TextBox txtReg_Notes 
         Height          =   1485
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   157
         Top             =   2100
         Width           =   3705
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   3270
         MouseIcon       =   "EmpInfo.frx":18D10
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":18E62
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Cancel"
         Top             =   3660
         Width           =   705
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Save"
         Height          =   795
         Left            =   2580
         MouseIcon       =   "EmpInfo.frx":191A0
         MousePointer    =   99  'Custom
         Picture         =   "EmpInfo.frx":192F2
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Save this Record"
         Top             =   3660
         Width           =   705
      End
      Begin VB.CommandButton Command6 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3630
         TabIndex        =   153
         Top             =   30
         Width           =   315
      End
      Begin MSComCtl2.DTPicker txtReg_Filed 
         Height          =   375
         Left            =   180
         TabIndex        =   159
         Top             =   660
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   55115777
         CurrentDate     =   39582
      End
      Begin VB.Label Label48 
         Caption         =   "Date Resinged"
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
         Left            =   180
         TabIndex        =   162
         Top             =   1110
         Width           =   3705
      End
      Begin VB.Label Label47 
         Caption         =   "Date Filed"
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
         Left            =   180
         TabIndex        =   161
         Top             =   420
         Width           =   3705
      End
      Begin VB.Label Label46 
         Caption         =   "Reson for Resigning "
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
         Left            =   180
         TabIndex        =   160
         Top             =   1830
         Width           =   3705
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   345
         Left            =   0
         TabIndex        =   154
         Top             =   0
         Width           =   4095
         _Version        =   655364
         _ExtentX        =   7223
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Resignation Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label labPicfilename 
      BackColor       =   &H8000000D&
      Caption         =   "Label39"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   10320
      TabIndex        =   109
      Top             =   5730
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label labIDprev 
      Caption         =   "IDprev"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3300
      TabIndex        =   50
      Top             =   330
      Width           =   615
   End
   Begin VB.Label labEmpNo 
      BackColor       =   &H8000000D&
      Caption         =   "Label42"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   10740
      TabIndex        =   51
      Top             =   480
      Width           =   465
   End
   Begin VB.Label LabID 
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   52
      Top             =   690
      Width           =   255
   End
End
Attribute VB_Name = "frmHRMSEmpInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                           As ADODB.Recordset
Dim rsSSS                                               As ADODB.Recordset
Dim rsSalaryLoan                                        As ADODB.Recordset
Dim rsCalamityLoan                                      As ADODB.Recordset
Dim rsPH                                                As ADODB.Recordset
Dim rsTIN                                               As ADODB.Recordset
Dim rsPagIbig                                           As ADODB.Recordset
Dim rsDepartment                                        As ADODB.Recordset
Attribute rsDepartment.VB_VarUserMemId = 1073938438
Dim rsSalaryGrade                                       As ADODB.Recordset
Dim ADDOREDIT                                           As String
Attribute ADDOREDIT.VB_VarUserMemId = 1073938441
Dim DEPCODE                                             As String
Dim SalCode                                             As String
Dim PICTfilname                                         As String
Attribute PICTfilname.VB_VarUserMemId = 1073938444
Dim EMPLIVIL                                            As String
Dim LocalAcess                                          As String
Dim vIS_TECHNICIAN                                      As Boolean
Dim vIS_SA                                              As Boolean
Dim vIS_SAE                                             As Boolean
Dim vIS_PARTSSA                                         As Boolean
Dim WithEvents FRM                                      As frmHRMSOvertime
Attribute FRM.VB_VarHelpID = -1

Function GetSLevel(BASICSAL)
    Dim rsSalLevel                                     As ADODB.Recordset
    Set rsSalLevel = gconDMIS.Execute("SELECT SDESCRIPTION FROM HRMS_SALARYLEVEL WHERE " & NumericVal(BASICSAL) & " BETWEEN SAL_FROM AND SAL_TO")
    If Not rsSalLevel.EOF Or Not rsSalLevel.BOF Then
        GetSLevel = "**" & Null2String(rsSalLevel!SDESCRIPTION) & "**"
    End If
End Function

Function GetPlevel(DateHired)
    Dim COUNTER                                        As Integer
    If DateHired <> "" Then
        COUNTER = DateDiff("m", DateHired, Now)
        If COUNTER > 6 Then
            GetPlevel = "REGULAR"
        Else
            GetPlevel = "PROBATIONARY"
        End If
    End If
End Function

Function SetCboDepCode(CCC As String)
    If CCC <> "" Then
        Set rsDepartment = New ADODB.Recordset
        rsDepartment.Open "SELECT * FROM HRMS_DEPARTMENT WHERE DEPTNAME = '" & CCC & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsDepartment.EOF And Not rsDepartment.BOF Then
            SetCboDepCode = Null2String(rsDepartment!DEPTCODE)
        End If
    End If
End Function

Function SetCboDepName(XXX As String)
    If XXX <> "" Then
        Set rsDepartment = New ADODB.Recordset
        rsDepartment.Open "SELECT * FROM HRMS_DEPARTMENT WHERE DEPTCODE = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsDepartment.EOF And Not rsDepartment.BOF Then
            SetCboDepName = Null2String(rsDepartment!DEPTNAME)
        End If
    End If
End Function

Function SetCboSalaryCode(CCC As String)
    If CCC <> "" Then
        Set rsSalaryGrade = New ADODB.Recordset
        rsSalaryGrade.Open "SELECT * FROM HRMS_SALARYGRADE WHERE [LEVEL] = '" & CCC & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
            SetCboSalaryCode = Null2String(rsSalaryGrade!CODE)
        End If
    End If
End Function

Function SetCboSalaryLevel(XXX As String)
    If XXX <> "" Then
        Set rsSalaryGrade = New ADODB.Recordset
        rsSalaryGrade.Open "SELECT * FROM HRMS_SALARYGRADE WHERE CODE = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
            SetCboSalaryLevel = Null2String(rsSalaryGrade![LEVEL])
        End If
    End If
End Function

Function SetShiftSched(XXX As String) As String
    Dim rsTIME_SHIFT_CODE                              As ADODB.Recordset
    Set rsTIME_SHIFT_CODE = New ADODB.Recordset
    Set rsTIME_SHIFT_CODE = gconDMIS.Execute("SELECT * FROM HRMS_TIME_SHIFT_CODE WHERE SHIFTCODE ='" & XXX & "'")
    If Not rsTIME_SHIFT_CODE.EOF And Not rsTIME_SHIFT_CODE.BOF Then
        SetShiftSched = Format(Null2String(rsTIME_SHIFT_CODE!FROM1), "HH:MM AM/PM") & " - " & Format(Null2String(rsTIME_SHIFT_CODE!TO1), "HH:MM AM/PM")
    End If
    Set rsTIME_SHIFT_CODE = Nothing
End Function

Sub DMISFunctionOrigTop()
    chkPos(0).Top = 0
    chkPos(1).Top = 270
    chkPos(2).Top = 570
    chkPos(3).Top = 840
End Sub

Sub EnableAddorEdit(XXX As Boolean)
    Picture3.Enabled = XXX
    Picture4.Enabled = XXX
    Picture5.Enabled = XXX
    cmdInsertPic.Enabled = XXX
    cmdDeletePicture.Enabled = XXX
End Sub

Sub FillCBO()
    cboBlood.Clear
    cboBlood.AddItem "O"
    cboBlood.AddItem "A"
    cboBlood.AddItem "B"
    cboBlood.AddItem "AB"
    cboBlood.ListIndex = 0
    cboPayroll.Clear
    cboPayroll.AddItem "Daily Base"
    cboPayroll.AddItem "Weekly Base"
    cboPayroll.AddItem "Semi-Monthly Base"
    cboPayroll.AddItem "Monthly Base"
    cboPayroll.ListIndex = 0
    cboEmpStatus.Clear
    cboEmpStatus.AddItem "Monthly"
    cboEmpStatus.AddItem "Daily"
    cboSex.Clear
    cboSex.AddItem "Male"
    cboSex.AddItem "Female"
    cboStatus.Clear
    cboStatus.AddItem "Single"
    cboStatus.AddItem "Married"
    cboStatus.AddItem "Separated"
    cboExStatus.Clear
    cboExStatus.AddItem "Z"
    cboExStatus.AddItem "S"
    cboExStatus.AddItem "HF"
    cboExStatus.AddItem "ME"
    cboExStatus.AddItem "HF1"
    cboExStatus.AddItem "HF2"
    cboExStatus.AddItem "HF3"
    cboExStatus.AddItem "HF4"
    cboExStatus.AddItem "ME1"
    cboExStatus.AddItem "ME2"
    cboExStatus.AddItem "ME3"
    cboExStatus.AddItem "ME4"
    Combo_Loadval cboPosition, gconDMIS.Execute("SELECT DISTINCT POSITION FROM HRMS_EMPINFO where isnull(POSITION,'')<>''")
    
    FillCboDepName
    Combo_Loadval cboSALCODE, gconDMIS.Execute("SELECT CODE FROM HRMS_SALARYGRADE ORDER BY CODE ASC")
    'FillCboSalaryGradeLevel
End Sub

Sub FillCboDepName()
    Set rsDepartment = New ADODB.Recordset
    rsDepartment.Open "SELECT * FROM HRMS_DEPARTMENT ORDER BY DEPTNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        rsDepartment.MoveFirst
        cboDeptName.Clear
        Do While Not rsDepartment.EOF
            cboDeptName.AddItem Null2String(rsDepartment!DEPTNAME)
            rsDepartment.MoveNext
        Loop
    End If
End Sub

Sub FillCboSalaryGradeLevel()
    Set rsSalaryGrade = New ADODB.Recordset
    rsSalaryGrade.Open "SELECT * FROM HRMS_SALARYGRADE ORDER BY [LEVEL] ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        rsSalaryGrade.MoveFirst
        cboSalaryGradeLevel.Clear
        Do While Not rsSalaryGrade.EOF
            cboSalaryGradeLevel.AddItem Null2String(rsSalaryGrade![LEVEL])
            rsSalaryGrade.MoveNext
        Loop
    End If
End Sub

Sub FillGrid()
    On Error Resume Next
    Dim rsEMPINFO2                                     As ADODB.Recordset
    lsAdjustment.Enabled = False
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    If optActiveEmployee.Value = True Then
        Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND ACTIVEINACTIVE ='A' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Else
        Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND ACTIVEINACTIVE ='I' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    End If
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
    lsAdjustment.Enabled = True
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    On Error GoTo Errorcode:
    Dim rsEMPINFO2                                     As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    If optActiveEmployee.Value = True Then
        Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND ACTIVEINACTIVE ='A' AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Else
        Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND ACTIVEINACTIVE='I' AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    End If
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub DisableForm(COND As Boolean)
    Picture5.Enabled = COND
    Picture6.Enabled = COND
    Picture4.Enabled = COND
    Picture3.Enabled = COND
    Picture2.Enabled = COND
    Picture1.Enabled = COND
    optActiveEmployee.Enabled = COND
    optInActiveEmployee.Enabled = COND
End Sub

Sub FillPayrollGroup()
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEDUCTION_SET FROM HRMS_SETUPDEDUCTION ORDER BY DEDUCTION_SET")
    cboPayrollGroup.Clear
    cboPayrollGroup.AddItem ""
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboPayrollGroup.AddItem NumericVal(RSTMP!Deduction_set)
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub InitMemvars()
    txtReg_Notes = ""
    txtReg_Resinged.Value = ""
    txtReg_Filed = ""
    labPicfilename.Caption = ""
    PICTfilname = ""
    LoadPic imgDispPic, ""
    'imgDispPic.Picture = LoadPicture(PICTfilname)
    Dim rsEmpInfoDup                                   As ADODB.Recordset
    Set rsEmpInfoDup = New ADODB.Recordset
    rsEmpInfoDup.Open "SELECt EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' ORDER BY EMPNO DESC", gconDMIS
    If Not rsEmpInfoDup.EOF And Not rsEmpInfoDup.BOF Then
        rsEmpInfoDup.MoveFirst
        txtEmpNo.Text = Format(NumericVal(N2Str2Zero(rsEmpInfoDup!EMPNO)) + 1, "0000")
    Else
        txtEmpNo.Text = "0001"
    End If
    If COMPANY_CODE = "HARI" Then
        lab4 = "ADMS Function"
    Else
        lab4 = "DMIS Function"
    End If
    txtAcctNo.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAddress.Text = ""
    cboPayrollGroup.ListIndex = -1
    txtTelephone.Text = ""
    txtBirthDate.Text = ""
    txtBirthPlace.Text = ""
    txtHeight.Text = ""
    txtWeight.Text = ""
    txtReligion.Text = ""
    txtCitizen.Text = ""
    txtSSSNo.Text = ""
    txtTINNo.Text = ""
    txtPHNo.Text = ""
    txtPagIbigNo.Text = ""
    txtCompanyName.Text = ""
    cboPosition.Text = ""
    txtDateHired.Text = ""
    txtResigned.Text = ""
    txtSpouse.Text = ""
    txtSpouseAge.Text = ""
    txtSOccupation.Text = ""
    txtFather.Text = ""
    txtFatherAge.Text = ""
    txtFOccupation.Text = ""
    txtMother.Text = ""
    txtMotherAge.Text = ""
    txtMOccupation.Text = ""
    txtPerson.Text = ""
    txtRelation.Text = ""
    txtRelTelNo.Text = ""
    chkWithPrevious.Value = 0
    optActive.Value = True
    txtALLOWANCE = "0.00"
    txtBASICSALARY = "0.00"
    FillCBO
    DMISFunctionOrigTop
End Sub

Sub rsrefresh()
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            If optActiveEmployee.Value = True Then
                Set rsEmpInfo = New ADODB.Recordset
                rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'M' AND ACTIVEINACTIVE = 'A' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
                EMPLIVIL = "M"
            Else
                Set rsEmpInfo = New ADODB.Recordset
                rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'M' AND ACTIVEINACTIVE = 'I' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
                EMPLIVIL = "M"
            End If
        End If
        If HEADOREMP = "EMP_A" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND ACTIVEINACTIVE= 'A' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "E"
        End If
        If HEADOREMP = "EMP_U" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'E' AND ACTIVEINACTIVE= 'I' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "E"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then
        If HEADOREMP = "EMP_A" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'C' AND ACTIVEINACTIVE= 'A' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "C"
        End If
        If HEADOREMP = "EMP_U" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'C' AND ACTIVEINACTIVE= 'I' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "C"
        End If
    End If
    If EMP_TYPE = "ALLOWANCE BASE" Then
        If HEADOREMP = "EMP_A" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'A' AND ACTIVEINACTIVE= 'A' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "A"
        End If
        If HEADOREMP = "EMP_U" Then
            Set rsEmpInfo = New ADODB.Recordset
            rsEmpInfo.Open "SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = 'A' AND ACTIVEINACTIVE= 'I' ORDER BY LASTNAME, FIRSTNAME, MIDDLENAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
            EMPLIVIL = "A"
        End If
    End If
End Sub

Sub SetDMISFunc()
    If vIS_TECHNICIAN = True Then
        chkPos(0).Visible = True
        chkPos(0).Value = 1
    Else
        chkPos(0).Visible = False
        chkPos(0).Value = 0
    End If
    If vIS_PARTSSA = True Then
        chkPos(1).Visible = True
        chkPos(1).Top = chkPos(0).Top: chkPos(2).ZOrder 0
        chkPos(1).Value = 1
    Else
        chkPos(1).Visible = False
        chkPos(1).Value = 0
    End If
    If vIS_SA = True Then
        chkPos(2).Visible = True
        chkPos(2).Top = chkPos(0).Top
        chkPos(2).ZOrder 0
        chkPos(2).Value = 1
    Else
        chkPos(2).Visible = False
        chkPos(2).Value = 0
    End If
    If vIS_SAE = True Then
        chkPos(3).Visible = True
        chkPos(3).Top = chkPos(0).Top
        chkPos(2).ZOrder 0
        chkPos(3).Value = 1
    Else
        chkPos(3).Visible = False
        chkPos(3).Value = 0
    End If
End Sub

Sub StoreMemVars()
    picPosition.HEIGHT = 345
    imgDispPic.Picture = Nothing
    If Not (rsEmpInfo.EOF And rsEmpInfo.BOF) Then
        Call RefreshFingerPrint
        EnableAddorEdit False
        cboSALCODE.Text = Null2String(rsEmpInfo!SalaryCode)
        cboBlood.Text = Trim(Null2String(rsEmpInfo!BloodType))
        cboPayroll.Text = Null2String(rsEmpInfo!payrolltype)
        LabID.Caption = rsEmpInfo!ID
        labEmpNo.Caption = Null2String(rsEmpInfo!EMPNO)
        txtEmpNo.Text = Null2String(rsEmpInfo!EMPNO)
        txtAcctNo.Text = Null2String(rsEmpInfo!ACCOUNTNO)
        If IsNull(rsEmpInfo!DEPTCODE) = True Then
            FillCboDepName
        Else
            cboDeptName.Text = SetCboDepName(rsEmpInfo!DEPTCODE)
        End If
        If IsNull(rsEmpInfo!SalaryCode) = True Then
            FillCboSalaryGradeLevel
        Else
            cboSalaryGradeLevel.Text = SetCboSalaryLevel(rsEmpInfo!SalaryCode)
        End If
        txtLastName.Text = Null2String(rsEmpInfo!lastname)
        txtFirstName.Text = Null2String(rsEmpInfo!FIRSTNAME)
        txtMiddleName.Text = Null2String(rsEmpInfo!MIDDLENAME)
        txtAddress.Text = Null2String(rsEmpInfo!ADDRESS)
        txtTelephone.Text = Null2String(rsEmpInfo!TELEPHONE)
        txtBirthDate.Text = Null2String(rsEmpInfo!BIRTHDATE)
        If Null2String(rsEmpInfo!SEX) = "M" Then
            cboSex.Text = "Male"
        Else
            cboSex.Text = "Female"
        End If
        cboStatus.Text = Null2String(rsEmpInfo!STATUS)
        txtBirthPlace.Text = Null2String(rsEmpInfo!BIRTHPLACE)
        txtHeight.Text = Null2String(rsEmpInfo!HEIGHT)
        txtWeight.Text = Null2String(rsEmpInfo!WEIGHT)
        txtReligion.Text = Null2String(rsEmpInfo!RELIGION)
        txtCitizen.Text = Null2String(rsEmpInfo!CITIZEN)
        txtSSSNo.Text = Null2String(rsEmpInfo!SSSNO)
        txtTINNo.Text = Null2String(rsEmpInfo!tinno)
        txtPHNo.Text = Null2String(rsEmpInfo!PHNO)
        txtPagIbigNo.Text = Null2String(rsEmpInfo!pagibigno)
        If Null2String(rsEmpInfo!withprevious) = "Y" Then
            chkWithPrevious.Value = 1
        Else
            chkWithPrevious.Value = 0
        End If
        If chkWithPrevious.Value = 1 Then
            txtCompanyName.Text = Null2String(rsEmpInfo!PreviousCompany)
        Else
            txtCompanyName.Enabled = False
        End If
        cboExStatus.Text = Null2String(rsEmpInfo!EXSTATUS)
        cboPosition.Text = Null2String(rsEmpInfo!Position)
        If Null2String(rsEmpInfo!EMPSTATUS) = "M" Then
            cboEmpStatus.Text = "Monthly"
        Else
            cboEmpStatus.Text = "Daily"
        End If
        txtDateHired.Text = Null2String(rsEmpInfo!DateHired)
        txtResigned.Text = Null2String(rsEmpInfo!RESIGNED)
        txtSpouse.Text = Null2String(rsEmpInfo!SPOUSE)
        txtSpouseAge.Text = Null2String(rsEmpInfo!SPOUSEAGE)
        txtSOccupation.Text = Null2String(rsEmpInfo!SOCCUPATION)
        txtFather.Text = Null2String(rsEmpInfo!FATHER)
        txtFatherAge.Text = Null2String(rsEmpInfo!FATHERAGE)
        txtFOccupation.Text = Null2String(rsEmpInfo!FOCCUPATION)
        txtMother.Text = Null2String(rsEmpInfo!MOTHER)
        txtBASICSALARY = FormatNumber(NumericVal(rsEmpInfo!BASICSALARY))
        txtMotherAge.Text = Null2String(rsEmpInfo!MOTHERAGE)
        txtMOccupation.Text = Null2String(rsEmpInfo!MOCCUPATION)
        txtPerson.Text = Null2String(rsEmpInfo!person)
        txtRelation.Text = Null2String(rsEmpInfo!Relation)
        txtRelTelNo.Text = Null2String(rsEmpInfo!reltelno)
        labPicfilename.Caption = Null2String(rsEmpInfo!PICFILNAME)
        If Null2String(rsEmpInfo!PayrollGroup) <> "" Then
            cboPayrollGroup = Null2String(rsEmpInfo!PayrollGroup)
        Else
            cboPayrollGroup = ""
        End If
        txtALLOWANCE = FormatNumber(NumericVal(rsEmpInfo!ALLOWANCE))
        labSHIFTCODE.Caption = Null2String(rsEmpInfo!Shift)
        labSHIFTSCHED.Caption = SetShiftSched(Null2String(rsEmpInfo!Shift))
        If Null2Bool(rsEmpInfo!IS_TECHNICIAN) = False And Null2Bool(rsEmpInfo!IS_SAE) = False And Null2Bool(rsEmpInfo!IS_PARTS_SALESMAN) = False And Null2Bool(rsEmpInfo!IS_SERVICE_ADVISER) = False Then
            chkPos(0).Visible = False
            chkPos(1).Visible = False
            chkPos(2).Visible = False
            chkPos(3).Visible = False
        End If
        vIS_TECHNICIAN = Null2Bool(rsEmpInfo!IS_TECHNICIAN)
        vIS_PARTSSA = Null2Bool(rsEmpInfo!IS_PARTS_SALESMAN)
        vIS_SA = Null2Bool(rsEmpInfo!IS_SERVICE_ADVISER)
        vIS_SAE = Null2Bool(rsEmpInfo!IS_SAE)
                
        If Null2Bool(rsEmpInfo!IS_SERVICE_ADVISER) = True Then
            cmdSAType.Enabled = True
        Else
            cmdSAType.Enabled = False
        End If
        If Null2Bool(rsEmpInfo!IS_TECHNICIAN) = True Then
            cmdTechPost.Enabled = True
        Else
            cmdTechPost.Enabled = False
        End If
        
        Call SetDMISFunc
        If Null2Bool(rsEmpInfo!SUBJECT_TO_COLA) = True Then
            Check1.Value = 1
            txtCOLA_RATE.Text = N2Str2Zero(rsEmpInfo!COLA_RATE)
        Else
            Check1.Value = 0
            txtCOLA_RATE.Text = 0
        End If
        If Null2String(rsEmpInfo!ACTIVEINACTIVE) = "A" Then
            optActive.Value = True
        Else
            optInActive.Value = True
        End If
        If Null2String(rsEmpInfo!PICFILNAME) <> "" Then
            On Error Resume Next
            If Len(Dir(HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME))) <= 0 Then
                Exit Sub
            End If
            LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(rsEmpInfo!PICFILNAME)
        Else
            LoadPic imgDispPic, ""
        End If
        LABSLEVEL.Caption = GetSLevel(NumericVal(rsEmpInfo!BASICSALARY))
        Label45.Caption = GetPlevel(rsEmpInfo!DateHired)
        PICTfilname = ""
        If NumericVal(rsEmpInfo!BASICSALARY) = 0 Then
            lblBasicSalary.Visible = True
        End If
        If Not NumericVal(rsEmpInfo!BASICSALARY) = 0 Then
            lblBasicSalary.Visible = False
        End If
        If NumericVal(rsEmpInfo!ALLOWANCE) = 0 Then
            lblAllowance.Visible = True
        End If
        If Not NumericVal(rsEmpInfo!ALLOWANCE) = 0 Then
            lblAllowance.Visible = False
        End If
        If Null2String(rsEmpInfo!EMPSTATUS) = "D" Then
            cmdAttendance.Enabled = True
        End If
        If Not Null2String(rsEmpInfo!EMPSTATUS) = "D" Then
            cmdAttendance.Enabled = False
        End If
        If IsDate(rsEmpInfo!RESIGNED) = True Then
            txtReg_Filed = Null2String(rsEmpInfo!RESIGNED_FILED)
            txtReg_Resinged = Null2String(rsEmpInfo!RESIGNED)
            txtReg_Notes = Null2String(rsEmpInfo!RESIGNED_NOTES)
        End If
        
        'UPDATE BY   : MJP012609 0536PM
        'DESCRIPTION : TO DISPLAY THE TECHNICIAN TYPE
            Call DisplayPosition(Null2String(rsEmpInfo!EMPNO))
        'UPDATE BY   : MJP012609 0536PM
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then
            cmdAdd.Value = True
         
        Else
            Unload Me
        End If
    End If
End Sub

Sub RefreshFingerPrint()
    Dim rsEmpInfo                                      As ADODB.Recordset
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("Select FPIMAGE1,FPIMAGE2,EMPNO,LASTNAME,FIRSTNAME,MIDDLENAME from HRMS_EmpInfo WHERE EMPNO = '" & EMPINFOEMPNO.Caption & "'")
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Dim st                                         As New ADODB.Stream
        Dim strTemp                                    As String
        st.Type = adTypeBinary
        If IsNull(rsEmpInfo!FPIMAGE1) = False Then
            st.Open: labFP1.Visible = False
            st.Write rsEmpInfo.FIELDS("FPIMAGE1").Value
            st.SaveToFile Environ("TEMP") & "\FPIMAGE1", adSaveCreateOverWrite
            imgTemplate1.Picture = LoadPicture(Environ("TEMP") & "\FPIMAGE1")
            Kill (Environ("TEMP") & "\FPIMAGE1")
            st.Close
        Else
            LoadPic imgTemplate1, "": labFP1.Visible = True
        End If
        If IsNull(rsEmpInfo!FPIMAGE2) = False Then
            st.Open: labFP2.Visible = False
            st.Write rsEmpInfo.FIELDS("FPIMAGE2").Value
            st.SaveToFile Environ("TEMP") & "\FPIMAGE2", adSaveCreateOverWrite
            imgTemplate2.Picture = LoadPicture(Environ("TEMP") & "\FPIMAGE2")
            Kill (Environ("TEMP") & "\FPIMAGE2")
            st.Close
        Else
            LoadPic imgTemplate2, "": labFP2.Visible = True
        End If
        Set st = Nothing
    End If
    Set rsEmpInfo = Nothing
End Sub

Private Sub cboDeptName_GotFocus()
    If cboDeptName.Text = "" Then
        FillCboDepName
    End If
End Sub

Private Sub cboPosition_LostFocus()
    If cboPosition.Text <> "" Then
        cboPosition.Text = Cap1st(cboPosition.Text)
    End If
End Sub

Private Sub cboSalaryGradeLevel_GotFocus()
    If cboSalaryGradeLevel.Text = "" Then
        FillCboSalaryGradeLevel
    End If
End Sub

Private Sub cboStatus_Change()
    If cboStatus.Text = "Single" Then
        cboExStatus.Text = "S"
    End If
    If cboStatus.Text = "Married" Then
        cboExStatus.Text = "ME"
    End If
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtCOLA_RATE.Enabled = True
    Else
        txtCOLA_RATE.Enabled = False
    End If
End Sub

Private Sub chkWithPrevious_Click()
    If chkWithPrevious.Value = 1 Then
        txtCompanyName.Enabled = True
        If ADDOREDIT = "ADD" Then
            txtCompanyName.Text = "Please Enter Previous Company"
        Else
            txtCompanyName.Text = Null2String(rsEmpInfo!PreviousCompany)
        End If
    Else
        txtCompanyName.Enabled = False
        txtCompanyName.Text = ""
    End If
End Sub

Private Sub cmdAdd_Click()
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "EMP_A" Then
            If Function_Access(LOGID, "Acess_Add", "EMPLOYEE INFO") = False Then Exit Sub
        Else
            If Function_Access(LOGID, "Acess_Add", "MANAGERS INFO") = False Then Exit Sub
        End If
    ElseIf EMP_TYPE = "CONTRACTUAL" Then
        If Function_Access(LOGID, "Acess_Add", "CONTRACTUAL INFO") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Add", "ALLOWANCE BASE INFO") = False Then Exit Sub
    End If

    ADDOREDIT = "ADD"
    EnableAddorEdit True
    Picture1.Visible = False
    Picture2.Visible = True
    lsAdjustment.Enabled = False
    txtSearch.Enabled = False
    InitMemvars
    optActive.Value = True
    Dim rsADDEmpNo                                     As ADODB.Recordset
    Set rsADDEmpNo = New ADODB.Recordset
    Set rsADDEmpNo = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' ORDER BY EMPNO DESC")
    If Not rsADDEmpNo.EOF And Not rsADDEmpNo.BOF Then
        txtEmpNo.Text = Format(NumericVal(rsADDEmpNo!EMPNO) + 1, "000")
    Else
        txtEmpNo.Text = "001"
    End If
End Sub

Private Sub cmdAdjustment_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ADJUSTMENTS", "Data Entry") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSAdjustment
    frmHRMSAdjustment.cmdFind.Enabled = False
    frmHRMSAdjustment.cmdPrevious.Enabled = False
    frmHRMSAdjustment.cmdNext.Enabled = False
    frmHRMSAdjustment.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdAttendance_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN ATTENDANCE", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSDailyMonitoring
    frmHRMSDailyMonitoring.cmdFind.Enabled = False
    frmHRMSDailyMonitoring.cmdPrevious.Enabled = False
    frmHRMSDailyMonitoring.cmdNext.Enabled = False
    frmHRMSDailyMonitoring.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    txtSearch.Enabled = True
    txtSearch.BackColor = vbWhite
    EnableAddorEdit False
    Picture1.Visible = True
    Picture2.Visible = False
    lsAdjustment.Enabled = True
    txtSearch.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCredits_Click()
    frmHRMS_Leave.Show
End Sub

Private Sub cmdDeductions_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN DEDUCTIONS", "Data Entry") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSDeductions
    frmHRMSDeductions.cmdFind.Enabled = False
    frmHRMSDeductions.cmdPrevious.Enabled = False
    frmHRMSDeductions.cmdNext.Enabled = False
    frmHRMSDeductions.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdDelete_Click()
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "EMP_A" Then
            If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE INFO") = False Then Exit Sub
        Else
            If Function_Access(LOGID, "Acess_Delete", "MANAGERS INFO") = False Then Exit Sub
        End If
    ElseIf EMP_TYPE = "CONTRACTUAL" Then
        If Function_Access(LOGID, "Acess_Delete", "CONTRACTUAL INFO") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Delete", "ALLOWANCE BASE INFO") = False Then Exit Sub
    End If
    On Error GoTo Errorcode
    If Not rsEmpInfo.BOF Or Not rsEmpInfo.EOF Then
        If ShowConfirmDelete = True Then

            SQL_STATEMENT = "DELETE FROM HRMS_EMPINFO WHERE ID = " & LabID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "EMPLOYEE INFO", SQL_STATEMENT, LabID, Null2String(EMP_TYPE), txtEmpNo, "", ""
            ShowDeletedMsg
            rsrefresh
            StoreMemVars
            FillSearchGrid ""
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    Exit Sub
Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdDeletePicture_Click()
    If MsgBox("Delete this picture? Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
        DeletePic imgDispPic
        labPicfilename.Caption = ""
    End If
End Sub

Private Sub cmdEdit_Click()
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "EMP_A" Or HEADOREMP = "EMP_U" Then
            If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE INFO") = False Then Exit Sub
        Else
            If Function_Access(LOGID, "Acess_Edit", "MANAGERS INFO") = False Then Exit Sub
        End If
    ElseIf EMP_TYPE = "CONTRACTUAL" Then
        If Function_Access(LOGID, "Acess_Edit", "CONTRACTUAL INFO") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_Edit", "ALLOWANCE BASE INFO") = False Then Exit Sub
    End If
    txtSearch.Enabled = False
    txtSearch.BackColor = &HC0C0C0
    ADDOREDIT = "EDIT"
    EnableAddorEdit True
    Picture1.Visible = False
    Picture2.Visible = True
    lsAdjustment.Enabled = False
End Sub

Private Sub cmdEnrollFingerPrint_Click()
    If Module_Access(LOGID, "EMPLOYEE ENROLL FINGERPRINT", "Data Entry") = False Then Exit Sub
    On Error GoTo Errorcode
    Unload frmHRMSEmpEnrollBio
    frmHRMSEmpEnrollBio.Show 1
    Call RefreshFingerPrint
    Exit Sub
    
Errorcode:
    MessagePop InfoStop, "Error", "Unknow Error" & " Error Code: " & Err.NUMBER & " " & Err.Description
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdInsertPic_Click()
    labPicfilename.Caption = ShowInsertpic(comDialogEmpInfo, imgDispPic)
End Sub

Private Sub cmdLedger_Click()
    If Module_Access(LOGID, "EMPLOYEE LEDGER", "Data Entry") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSLedger
    frmHRMSLedger.cmdFind.Enabled = False
    frmHRMSLedger.cmdPrevious.Enabled = False
    frmHRMSLedger.cmdNext.Enabled = False
    frmHRMSLedger.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdOvertime_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN OVERTIME", "DATA ENTRY") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSOvertime
    frmHRMSOvertime.cmdFind.Enabled = False
    frmHRMSOvertime.cmdPrevious.Enabled = False
    frmHRMSOvertime.cmdNext.Enabled = False
    frmHRMSOvertime.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdPosi_Click()
    DMISFunctionOrigTop
    If picPosition.HEIGHT = 345 Then
        DMISFunctionOrigTop
        picPosition.HEIGHT = 1215
        chkPos(0).Visible = True
        chkPos(1).Visible = True
        chkPos(2).Visible = True
        chkPos(3).Visible = True
    Else
        picPosition.HEIGHT = 345
        SetDMISFunc
        chkPos(0).Visible = True
        chkPos(1).Visible = True
        chkPos(2).Visible = True
        chkPos(3).Visible = True
    End If
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "EMP_A" Then
            If Function_Access(LOGID, "Acess_Print", "EMPLOYEE INFO") = False Then Exit Sub
        Else
            If Function_Access(LOGID, "Acess_Print", "MANAGERS INFO") = False Then Exit Sub
        End If
    ElseIf EMP_TYPE = "CONTRACTUAL" Then
        If Function_Access(LOGID, "Acess_Print", "CONTRACTUAL INFO") = False Then Exit Sub
    Else
        If Function_Access(LOGID, "Acess_print", "ALLOWANCE BASE INFO") = False Then Exit Sub
    End If

    Screen.MousePointer = 11
    rptEmpInfo.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    rptEmpInfo.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    rptEmpInfo.Formulas(2) = "COMPANY_TIN = '" & COMPANY_TIN & "'"
    PrintSQLReport rptEmpInfo, HRMS_REPORT_PATH & "empinfo.rpt", "{empinfo.empno} = " & N2Str2Null(Null2String(rsEmpInfo!EMPNO)), DMIS_REPORT_Connection, 1
    LogAudit "V", "EMPLOYEE INFO", txtEmpNo.Text
    Screen.MousePointer = 0
End Sub

Private Sub cmdSalaryAdvance_Click()
    If Module_Access(LOGID, "EMPLOYEE MAINTAIN COMMISSION", "Data Entry") = False Then Exit Sub
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Unload frmHRMSCommission
    frmHRMSCommission.cmdFind.Enabled = False
    frmHRMSCommission.cmdPrevious.Enabled = False
    frmHRMSCommission.cmdNext.Enabled = False
    frmHRMSCommission.Show
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    'On Error GoTo Errorcode
    Dim NewName, WITHPREV                              As String
    If chkWithPrevious.Value = 1 Then
        WITHPREV = "'Y'"
    Else
        WITHPREV = "'N'"
    End If
    If txtLastName.Text = "" Then
        ShowIsRequiredMsg "Last Name"
        txtLastName.SetFocus
        Exit Sub
    End If

    If txtFirstName.Text = "" Then
        ShowIsRequiredMsg "First Name"
        txtFirstName.SetFocus
        Exit Sub
    End If



    If txtMiddleName.Text = "" Then
        ShowIsRequiredMsg "Middle Name"
        txtMiddleName.SetFocus
        Exit Sub
    End If

      
    If txtDateHired.Text = "" Then
        ShowIsRequiredMsg "Date Hired"
        txtDateHired.SetFocus
        Exit Sub
    End If



    If txtEmpNo.Text = "" Then
        ShowIsRequiredMsg "Employee Number"
        txtEmpNo.SetFocus
        Exit Sub
    End If

    Dim rsEmpInfoDup                                   As New ADODB.Recordset
    If ADDOREDIT = "ADD" Then
        rsEmpInfoDup.Open "SELECT id FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND EMPNO = '" & txtEmpNo.Text & "' ORDER BY ID ASC", gconDMIS
        If Not rsEmpInfoDup.EOF And Not rsEmpInfoDup.BOF Then
            ShowAlreadyExistMsg "Employee Number"
            Exit Sub
        End If

        Set rsEmpInfoDup = New ADODB.Recordset
        rsEmpInfoDup.Open "SELECT id FROM HRMS_EMPINFO ORDER BY ID ASC", gconDMIS
        If Not rsEmpInfoDup.EOF And Not rsEmpInfoDup.BOF Then
            rsEmpInfoDup.MoveLast
            LabID.Caption = NumericVal(rsEmpInfoDup!ID) + 1
        End If
    Else
        rsEmpInfoDup.Open "SELECT id FROM HRMS_EMPINFO WHERE EMPLEVEL = '" & EMPLIVIL & "' AND EMPNO = '" & txtEmpNo.Text & "' ORDER BY ID ASC", gconDMIS
        If Not (rsEmpInfoDup.BOF And rsEmpInfoDup.EOF) Then
            If LabID.Caption <> rsEmpInfoDup!ID Then
                ShowAlreadyExistMsg "Employee Number"
                Exit Sub
            End If
        End If
    End If
    
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CODE FROM CMIS_SBOOK WHERE BOOK = 'I' AND CODE = " & N2Str2Null(txtEmpNo) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        MsgBox "Duplicate Employee no is not Allowed. Employee no already Exist in the Cashier Employee Record", vbInformation, "Info"
        Exit Sub
    End If
    
    If cboSex.Text = "" Then
        ShowIsRequiredMsg "Sex"
        cboSex.SetFocus
        Exit Sub
    End If
    If cboStatus.Text = "" Then
        ShowIsRequiredMsg "Civil Status"
        cboStatus.SetFocus
        Exit Sub
    End If
    If cboExStatus.Text = "" Then
        ShowIsRequiredMsg "Exemption Status"
        cboExStatus.SetFocus
        Exit Sub
    End If
    If cboEmpStatus.Text = "" Then
        ShowIsRequiredMsg "Employment Status"
        cboEmpStatus.SetFocus
        Exit Sub
    End If
    If (cboDeptName.Text) = "" Then
        ShowIsRequiredMsg "Department Name"
        cboDeptName.SetFocus
        Exit Sub
    End If
    
    If UCase(txtSSSNo.Text) = "N/A" Then
        MsgBoxXP "Invalid SSS Number!", "Error", XP_OKOnly, msg_Information
        Exit Sub
    End If
    Dim ACTIVESTATUS                                   As String
    Dim SUBJECT_TO_COLA                                As Integer
    Dim VCOLA_RATE                                     As Double

    If Check1.Value = 1 Then
        SUBJECT_TO_COLA = 1
        VCOLA_RATE = N2Str2Zero(txtCOLA_RATE.Text)
    Else
        SUBJECT_TO_COLA = 0
        VCOLA_RATE = 0
    End If
    If optActive.Value = True Then
        ACTIVESTATUS = "A"
    Else
        ACTIVESTATUS = "I"
    End If

    Dim vtxtEmpNo                                      As String
    Dim VtxtAcctNo                                     As String
    Dim vtxtLastName                                   As String
    Dim vtxtFirstName                                  As String
    Dim vtxtMiddleName                                 As String
    Dim vtxtAddress                                    As String
    Dim vtxtTelephone                                  As String
    Dim vtxtBirthDate                                  As String
    Dim vtxtBirthPlace                                 As String
    Dim vtxtHeight                                     As String
    Dim vtxtWeight                                     As String
    Dim vtxtReligion                                   As String
    Dim vtxtCitizen                                    As String
    Dim vtxtSSSNo                                      As String
    Dim vtxtTINNo                                      As String
    Dim vtxtPHNo                                       As String
    Dim vtxtPagIbigNo                                  As String
    Dim vtxtPreviousCompany                            As String
    Dim vtxtPosition                                   As String
    Dim vtxtDateHired                                  As String
    Dim vtxtResigned                                   As String
    Dim vtxtSpouse                                     As String
    Dim vtxtSpouseAge                                  As String
    Dim vtxtSOccupation                                As String
    Dim vtxtFather                                     As String
    Dim vtxtFatherAge                                  As String
    Dim vtxtFOccupation                                As String
    Dim vtxtMother                                     As String
    Dim vtxtMotherAge                                  As String
    Dim vtxtMOccupation                                As String
    Dim vtxtPerson                                     As String
    Dim vtxtRelation                                   As String
    Dim vtxtRelTelNo                                   As String
    Dim vcbosex                                        As String
    Dim vcboSTATUS                                     As String
    Dim vcboexStatus                                   As String
    Dim vcboempStatus                                  As String
    Dim vtxtPayrollGroup                               As String
    Dim vEmpPicture                                    As Object
    Dim IS_SAE                                         As Integer
    Dim IS_SER_ADV                                     As Integer
    Dim IS_PAR_SAL                                     As Integer
    Dim IS_TECHA                                       As Integer
    Dim vtxtMonthlyAllowance                           As Double
    Dim vTXTBASICSALARY                                As Double

    'vtxtEmpNo = N2Str2Null(Format(txtEmpNo.Text, "000"))
    vtxtEmpNo = N2Str2Null(txtEmpNo.Text)
    DEPCODE = N2Str2Null(SetCboDepCode(cboDeptName.Text))
    SalCode = N2Str2Null(SetCboSalaryCode(cboSalaryGradeLevel.Text))
    VtxtAcctNo = N2Str2Null(txtAcctNo.Text)
    vtxtLastName = N2Str2Null(Cap1st(txtLastName.Text))
    vtxtFirstName = N2Str2Null(Cap1st(txtFirstName.Text))
    vtxtMiddleName = N2Str2Null(Cap1st(txtMiddleName.Text))
    vtxtAddress = N2Str2Null(txtAddress.Text)
    vtxtTelephone = N2Str2Null(txtTelephone.Text)
    vtxtBirthDate = N2Date2Null(txtBirthDate.Text)
    vtxtBirthPlace = N2Str2Null(txtBirthPlace.Text)
    vtxtHeight = N2Str2Null(txtHeight.Text)
    vtxtWeight = N2Str2Null(txtWeight.Text)
    vtxtReligion = N2Str2Null(txtReligion.Text)
    vtxtCitizen = N2Str2Null(txtCitizen.Text)
    vtxtSSSNo = N2Str2Null(txtSSSNo.Text)
    vtxtTINNo = N2Str2Null(txtTINNo.Text)
    vtxtPHNo = N2Str2Null(txtPHNo.Text)
    vtxtPagIbigNo = N2Str2Null(txtPagIbigNo.Text)
    vtxtPreviousCompany = N2Str2Null(txtCompanyName.Text)
    vtxtPosition = N2Str2Null(cboPosition.Text)
    vtxtDateHired = N2Date2Null(txtDateHired.Text)
    vtxtResigned = N2Date2Null(txtResigned.Text)
    vtxtSpouse = N2Str2Null(Cap1st(txtSpouse.Text))
    vtxtSpouseAge = N2Str2Null(txtSpouseAge.Text)
    vtxtSOccupation = N2Str2Null(Cap1st(txtSOccupation.Text))
    vtxtFather = N2Str2Null(Cap1st(txtFather.Text))
    vtxtFatherAge = N2Str2Null(txtFatherAge.Text)
    vtxtFOccupation = N2Str2Null(Cap1st(txtFOccupation.Text))
    vtxtMother = N2Str2Null(Cap1st(txtMother.Text))
    vtxtMotherAge = N2Str2Null(txtMotherAge.Text)
    vtxtMOccupation = N2Str2Null(Cap1st(txtMOccupation.Text))
    vtxtPerson = N2Str2Null(Cap1st(txtPerson.Text))
    vtxtRelation = N2Str2Null(txtRelation.Text)
    vtxtRelTelNo = N2Str2Null(txtRelTelNo.Text)
    vTXTBASICSALARY = NumericVal(txtBASICSALARY)
    vtxtPayrollGroup = N2Str2Null(cboPayrollGroup)
    vtxtMonthlyAllowance = NumericVal(txtALLOWANCE)

    If chkPos(0).Value = True Then
        IS_TECHA = 1
    End If
    If chkPos(1).Value = True Then
        IS_PAR_SAL = 1
    End If
    If chkPos(2).Value = True Then
        IS_SER_ADV = 1
    End If
    If chkPos(3).Value = True Then
        IS_SAE = 1
    End If
    If cboSex.Text = "Male" Then
        vcbosex = "'M'"
    Else
        vcbosex = "'F'"
    End If

    vcboSTATUS = N2Str2Null(cboStatus.Text)
    vcboexStatus = N2Str2Null(cboExStatus.Text)
    If cboEmpStatus.Text = "Monthly" Then
        vcboempStatus = "'M'"
    Else
        vcboempStatus = "'D'"
    End If
    If labPicfilename.Caption <> "" Then
        On Error Resume Next
        NewName = getpicfilename(labPicfilename.Caption)
        SavePic imgDispPic, HRMS_PICTURES_PATH & NewName
        PICTfilname = N2Str2Null(NewName)
    Else
        PICTfilname = "NULL"
    End If
    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO HRMS_EMPINFO" & _
            " (BASICSALARY, ALLOWANCE, PayrollGroup, empno,deptcode,accountno,lastname,firstname,middlename,address,telephone,birthdate,sex,BloodType,PayrollType,status,birthplace" & _
            ",height,weight,religion,citizen,sssno,tinno,phno,pagibigno,exstatus,[position]" & _
            ",empstatus,salarycode,datehired, resigned,spouse,spouseage,soccupation" & _
            ",father,fatherage,foccupation,mother,motherage,moccupation,person,relation,reltelno,picfilname,EMPLEVEL,previouscompany,withprevious,activeinactive,SUBJECT_TO_COLA,COLA_RATE" & _
            ",IS_TECHNICIAN,IS_PARTS_SALESMAN,IS_SAE,IS_SERVICE_ADVISER)" & _
            " values (" & vTXTBASICSALARY & "," & vtxtMonthlyAllowance & "," & vtxtPayrollGroup & "," & vtxtEmpNo & ", " & DEPCODE & ", " & VtxtAcctNo & ", " & vtxtLastName & ", " & vtxtFirstName & ", " & vtxtMiddleName & ", " & vtxtAddress & "," & _
            " " & vtxtTelephone & ", " & vtxtBirthDate & ", " & vcbosex & ",'" & cboBlood & "','" & cboPayroll & "'," & vcboSTATUS & ", " & vtxtBirthPlace & ", " & vtxtHeight & ", " & vtxtWeight & ", " & vtxtReligion & _
            ", " & vtxtCitizen & ", " & vtxtSSSNo & ", " & vtxtTINNo & ", " & vtxtPHNo & ", " & vtxtPagIbigNo & ", " & vcboexStatus & ", " & vtxtPosition & _
            ", " & vcboempStatus & ", " & SalCode & ", " & vtxtDateHired & ", " & vtxtResigned & ", " & vtxtSpouse & ", " & vtxtSpouseAge & _
            ", " & vtxtSOccupation & ", " & vtxtFather & ", " & vtxtFatherAge & ", " & vtxtFOccupation & ", " & vtxtMother & ", " & vtxtMotherAge & ", " & vtxtMOccupation & ", " & vtxtPerson & _
            ", " & vtxtRelation & ", " & vtxtRelTelNo & ", " & PICTfilname & ", '" & EMPLIVIL & "', " & vtxtPreviousCompany & ", " & WITHPREV & ", '" & ACTIVESTATUS & "'," & SUBJECT_TO_COLA & "," & VCOLA_RATE & _
            "," & IS_TECHA & "," & IS_PAR_SAL & "," & IS_SAE & "," & IS_SER_ADV & ")"

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "EMPLOYEE INFO", SQL_STATEMENT, LabID, Null2String(EMP_TYPE), Null2String(vtxtEmpNo), "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET" & _
            " EMPNO = " & vtxtEmpNo & ", DEPTCODE = " & DEPCODE & ", ACCOUNTNO = " & VtxtAcctNo & "," & _
            " LASTNAME = " & vtxtLastName & ", FIRSTNAME = " & vtxtFirstName & ", MIDDLENAME = " & vtxtMiddleName & "," & _
            " ADDRESS = " & vtxtAddress & ", TELEPHONE = " & vtxtTelephone & ", BIRTHDATE = " & vtxtBirthDate & "," & _
            " SEX = " & vcbosex & ",BLOODTYPE = '" & cboBlood & "',PAYROLLTYPE = '" & cboPayroll.Text & "',STATUS = " & vcboSTATUS & ", BIRTHPLACE = " & vtxtBirthPlace & "," & _
            " HEIGHT = " & vtxtHeight & ", WEIGHT = " & vtxtWeight & ", RELIGION = " & vtxtReligion & "," & _
            " CITIZEN = " & vtxtCitizen & ", SSSNO = " & vtxtSSSNo & ", TINNO = " & vtxtTINNo & "," & _
            " PHNO = " & vtxtPHNo & ", PAGIBIGNO = " & vtxtPagIbigNo & "," & _
            " EXSTATUS = " & vcboexStatus & ", [POSITION] = " & vtxtPosition & ", EMPSTATUS = " & vcboempStatus & "," & _
            " SALARYCODE = " & SalCode & "," & _
            " BASICSALARY = " & vTXTBASICSALARY & "," & _
            " ALLOWANCE = " & vtxtMonthlyAllowance & "," & _
            " PAYROLLGROUP = " & vtxtPayrollGroup & "," & _
            " DATEHIRED = " & vtxtDateHired & ", RESIGNED = " & vtxtResigned & ", SPOUSE = " & vtxtSpouse & "," & _
            " SPOUSEAGE = " & vtxtSpouseAge & ", SOCCUPATION = " & vtxtSOccupation & ", FATHER = " & vtxtFather & "," & _
            " FATHERAGE = " & vtxtFatherAge & ", FOCCUPATION = " & vtxtFOccupation & ", MOTHER = " & vtxtMother & "," & _
            " MOTHERAGE = " & vtxtMotherAge & ", MOCCUPATION = " & vtxtMOccupation & ", PERSON = " & vtxtPerson & "," & _
            " RELATION = " & vtxtRelation & ", RELTELNO = " & vtxtRelTelNo & ", PICFILNAME = " & PICTfilname & "," & _
            " EMPLEVEL = '" & EMPLIVIL & "', PREVIOUSCOMPANY = " & vtxtPreviousCompany & ", WITHPREVIOUS = " & WITHPREV & ", ACTIVEINACTIVE = '" & ACTIVESTATUS & "', SUBJECT_TO_COLA = " & SUBJECT_TO_COLA & ", COLA_RATE = " & VCOLA_RATE & _
            ", IS_TECHNICIAN = " & IS_TECHA & ",IS_SAE = " & IS_SAE & ",IS_PARTS_SALESMAN = " & IS_PAR_SAL & ",IS_SERVICE_ADVISER = " & IS_SER_ADV & _
            " WHERE ID = " & LabID.Caption
        gconDMIS.Execute SQL_STATEMENT
        
        NEW_LogAudit "A", "EMPLOYEE INFO", SQL_STATEMENT, LabID, Null2String(EMP_TYPE), Null2String(vtxtEmpNo), "", ""
        
        'UPDATE BY   : MJP07252009 0128PM
            Call UpdateRelatedTabletoHR(EMPLIVIL, txtEmpNo, Null2String(rsEmpInfo!EMPNO))
        'UPDATE BY   : MJP07252009 0128PM
        
        SQL_STATEMENT = ""
        Call ShowSuccessFullyUpdated
    End If
    
' JBF: 08/31/2010 - to refresh data after save

'    If optActive.Value = False Then
'        If EMPLIVIL = "M" Then
'            HEADOREMP = "HEAD"
'        ElseIf EMPLIVIL = "E" Then
'            HEADOREMP = "EMP_U"
'        End If
'        optInActiveEmployee.Value = True
'    Else
'        If EMPLIVIL = "M" Then
'            HEADOREMP = "HEAD"
'        ElseIf EMPLIVIL = "E" Then
'            HEADOREMP = "EMP_A"
'        End If
'        optActiveEmployee.Value = True
'    End If
'
'    rsrefresh
'    On Error Resume Next
'    rsEmpInfo.Find "id = " & LabID.Caption
'    cmdCancel.Value = True
'
'    txtSearch.Enabled = True
'    txtSearch.BackColor = vbWhite
'    FillSearchGrid txtSearch.Text
'
'    Exit Sub
' ********************************************************************

 'JBF: 08/31/2010 - to refresh data after save
    rsrefresh
    FillGrid
    EnableAddorEdit False
    EMPINFOSHOW = True
    cmdCancel.Value = True
    txtSearch.Enabled = True
    txtSearch.BackColor = vbWhite
    Exit Sub

 ' ******************************************

Errorcode:
    ShowVBError
    Exit Sub
End Sub

'UPDATE BY   : MJP07252009 0128PM
'DESCRIPTION : TO UPDATE ALL TABLE RELATED TO HRMS_EMPINFO
Sub UpdateRelatedTabletoHR(XLEVEL As String, XEMPNO As String, xOLDEMPNO As String)
    Dim xxxEmpno                                            As String
    Dim xxxEMPNO_IN_PETTY                                   As String
    
    xxxEmpno = N2Str2Null(XLEVEL & XEMPNO)
    xxxEMPNO_IN_PETTY = N2Str2Null(XLEVEL & xOLDEMPNO)
    gconDMIS.Execute ("UPDATE CMIS_PETTY                 SET EMPLOYEE = " & xxxEmpno & " WHERE EMPLOYEE = " & N2Str2Null(xxxEMPNO_IN_PETTY) & "")
    gconDMIS.Execute ("UPDATE CMIS_LTOPondo              SET EMPLOYEE = " & xxxEmpno & " WHERE EMPLOYEE = " & N2Str2Null(xxxEMPNO_IN_PETTY) & "")
    
    gconDMIS.Execute ("UPDATE HRMS_Adjustment            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Advance               SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Atm                   SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_AtmDET                SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Attend                SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_BegBalance            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Commission            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_DailyMonitoring       SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_DedDetails            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Deductions            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Education             SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_ExamsPassed           SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Leave                 SET EMPLNO = " & N2Str2Null(XEMPNO) & " WHERE EMPLNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_LeaveDet              SET EMPLNO = " & N2Str2Null(XEMPNO) & " WHERE EMPLNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_LoanMas               SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_LoanMasDet            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Memorandum            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Overtime              SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PagIbig               SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PagIbigDet            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PastEmployment        SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Payroll               SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Payroll_Det           SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PerformanceEvaluation SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PersonalAction        SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PhilHealth            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PhilHealthDet         SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_PrevEmp               SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_RequestLeave_OT       SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_SalaryAdvance         SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_SSS                   SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_SSSDet                SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_TAX_REMITANCE         SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Tin                   SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_TinDet                SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_Training              SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_TrainingPlan          SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
    gconDMIS.Execute ("UPDATE HRMS_YTDDetails            SET EMPNO = " & N2Str2Null(XEMPNO) & " WHERE EMPNO = " & N2Str2Null(xOLDEMPNO) & "")
End Sub
'UPDATE BY   : MJP07252009 0128PM

Private Sub cmdTechPost_Click()
    If Module_Access(LOGID, "EMPLOYEE SET TECHNICIAN TYPE", "SYSTEM") = False Then Exit Sub
    picTECH.ZOrder 0
    picTECH.Visible = True
    
    DisableForm False
    picOption.Visible = False
    picRegi.Visible = False
End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "EMPLOYEE SALARY-ALLOWANCE", "SYSTEM") = False Then Exit Sub
    picSalary.ZOrder 0
    picSalary.Visible = True
    DisableForm False
    picOption.Visible = False
    picRegi.Visible = False
End Sub

Private Sub cmdView_Click()
    If Not cboPayrollGroup = "" Then
        frmSETUP_Deduction.Show
        frmSETUP_Deduction.StoreMemVars cboPayrollGroup
    End If
End Sub

Private Sub Command10_Click()
    PIC_EMPLEVEL.Visible = False
    Call DisableForm(True)
     
End Sub

Private Sub Command11_Click()
    picTECH.Visible = False
    Call DisableForm(True)
    Call StoreMemVars
End Sub

Private Sub Command12_Click()
    Dim xCount                                          As Integer
    Dim TECH_POSITION                                   As String
    xCount = 0

    If chk1.Value = 1 Then xCount = xCount + 1
    If chk2.Value = 1 Then xCount = xCount + 1
    If chk3.Value = 1 Then xCount = xCount + 1
    If chk4.Value = 1 Then xCount = xCount + 1
    If chk5.Value = 1 Then xCount = xCount + 1
    If chk6.Value = 1 Then xCount = xCount + 1
    If chk7.Value = 1 Then xCount = xCount + 1
    If chk8.Value = 1 Then xCount = xCount + 1

    If xCount > 1 Then
        MsgBox "You are only allowed to select one Position.", vbInformation, "INFORMATION"
        Exit Sub
    End If
    
    TECH_POSITION = N2Str2Null(EX_POSITION)
    
    gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET CSMS_POSITION = " & TECH_POSITION & _
        " WHERE ID = " & LabID & "")
    
    Call ShowSuccessFullyUpdated
    Call DisableForm(True)
    Call Command11_Click
End Sub

Function EX_POSITION() As String
    Dim vPOSITION                                      As String
    Dim xPOSITION                                      As String
    vPOSITION = "00000000"

    If chk1.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk1.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk2.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk2.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk3.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk3.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk4.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk4.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk5.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk5.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk6.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk6.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk7.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk7.Value = 0 Then xPOSITION = xPOSITION & "0"
    If chk8.Value = 1 Then xPOSITION = xPOSITION & "1"
    If chk8.Value = 0 Then xPOSITION = xPOSITION & "0"

    EX_POSITION = xPOSITION & vPOSITION
End Function

Private Sub Command13_Click()
    If MsgBox("Are you Sure?", vbInformation + vbYesNo) = vbNo Then Exit Sub

    If OPT_REG.Value = True Then
        gconDMIS.Execute ("update hrms_empinfo set EMPLEVEL='E' WHERE ID=" & LabID)
    ElseIf OPT_CONTRACT.Value = True Then
        gconDMIS.Execute ("update hrms_empinfo set EMPLEVEL='C' WHERE ID=" & LabID)
    ElseIf OPT_CONFI.Value = True Then
        gconDMIS.Execute ("update hrms_empinfo set EMPLEVEL='M' WHERE ID=" & LabID)
    ElseIf OPT_ALLOWANCE.Value = True Then
        gconDMIS.Execute ("update hrms_empinfo set EMPLEVEL='A' WHERE ID=" & LabID)
    End If
    rsrefresh
    StoreMemVars
    Command10.Value = True
    txtsearch_Change
End Sub

Private Sub Command14_Click()
    If Null2String(rsEmpInfo!EMPLEVEL) = "E" Then
        OPT_REG.Value = True
    ElseIf Null2String(rsEmpInfo!EMPLEVEL) = "C" Then
        OPT_CONTRACT.Value = True
    ElseIf Null2String(rsEmpInfo!EMPLEVEL) = "M" Then
        OPT_CONFI.Value = True
    ElseIf Null2String(rsEmpInfo!EMPLEVEL) = "A" Then
        OPT_ALLOWANCE.Value = True
    End If
    
    picOption.Visible = False
    PIC_EMPLEVEL.Visible = True
    PIC_EMPLEVEL.ZOrder 0
    
End Sub

Private Sub Command2_Click()
    picSalary.Visible = False
    DisableForm True
    StoreMemVars
End Sub

Private Sub Command3_Click()
    Dim vtxtMonthlyAllowance                           As Double
    Dim vTXTBASICSALARY                                As Double
    picSalary.Visible = False
    vTXTBASICSALARY = NumericVal(txtBASICSALARY)
    vtxtMonthlyAllowance = NumericVal(txtALLOWANCE)
    SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET BASICSALARY = " & vTXTBASICSALARY & _
                    ", ALLOWANCE = " & vtxtMonthlyAllowance & _
                    ", SALARYCODE = " & N2Str2Null(cboSALCODE) & _
                    " WHERE ID = " & LabID.Caption & ""
    gconDMIS.Execute SQL_STATEMENT
    
    NEW_LogAudit "E", "EMPLOYEE INFO", SQL_STATEMENT, LabID, "SL", Null2String(txtEmpNo), "", ""
    SQL_STATEMENT = ""
    
    Call rsrefresh
    rsEmpInfo.Find "ID = " & LabID & ""
    Call StoreMemVars
    Call DisableForm(True)
    Call Command2_Click
End Sub

Sub DisplayPosition(VEMPNO As String)
   Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT CSMS_POSITION FROM HRMS_EMPINFO WHERE EMPNO = '" & VEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!CSMS_POSITION) = "" Then
            chk1.Value = 0
            chk2.Value = 0
            chk3.Value = 0
            chk4.Value = 0
            chk5.Value = 0
            chk6.Value = 0
            chk7.Value = 0
            chk8.Value = 0
        Else
            If Len(Null2String(RSTMP!CSMS_POSITION)) = 16 Then
                chk1.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 1, 1)
                chk2.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 2, 1)
                chk3.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 3, 1)
                chk4.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 4, 1)
                chk5.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 5, 1)
                chk6.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 6, 1)
                chk7.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 7, 1)
                chk8.Value = Mid(Null2String(RSTMP!CSMS_POSITION), 8, 1)
            Else
                chk1.Value = 0: chk2.Value = 0
                chk3.Value = 0: chk4.Value = 0
                chk5.Value = 0: chk6.Value = 0
                chk7.Value = 0: chk8.Value = 0
            End If
        End If
    End If
    Set RSTMP = Nothing
End Sub

Private Sub Command4_Click()
    picOption.Visible = True
    picOption.ZOrder 0
    DisableForm False
End Sub

Private Sub Command5_Click()
    picOption.Visible = False
    DisableForm True
End Sub

Private Sub Command6_Click()
    picRegi.Visible = False
    picOption.Visible = True
End Sub

Private Sub Command7_Click()
    picRegi.Visible = False
    picOption.Visible = True
End Sub

Private Sub Command8_Click()
    If IsNull(txtReg_Resinged) = True Then
        MessagePop InfoVoid, "Invalid Filing  Date", "Please Check Proper Filing Date"
        Exit Sub
    End If
    If IsNull(txtReg_Resinged) = True Then
        MessagePop InfoVoid, "Invalid Resignation  Date", "Please Check Proper Resignation Date"
        Exit Sub
    End If
    If DateDiff("d", txtDateHired, txtReg_Resinged) < 0 Then
        MessagePop InfoVoid, "Invalid Resigned Date", "Date Hired Is Less than Date Resigned"
        Exit Sub
    End If
    If DateDiff("d", txtReg_Resinged.Value, txtReg_Filed.Value) < 0 Then
        If MsgBox("Filed Date is Less than Resigned Date." & vbCrLf & "Are you Sure?", vbInformation + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("Are You Sure You Want to File Resignation Details", vbInformation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ACTIVEINACTIVE='I' , RESIGNED_FILED=" & N2Str2Null(txtReg_Filed.Value) & ",RESIGNED=" & N2Str2Null(txtReg_Resinged.Value) & ", RESIGNED_NOTES=" & N2Str2Null(txtReg_Notes) & " WHERE ID=" & LabID)
    rsrefresh
    rsEmpInfo.Find ("ID=" & LabID)
    StoreMemVars
    picRegi.Visible = False
    picOption.Visible = False
    MessagePop RecSaveOk, "Employee Information Updated", "Employee Resignation Details Sucessfully Updated"
End Sub

Private Sub Command9_Click()
    If IsDate(txtResigned) = False Then
        txtReg_Notes = ""
        txtReg_Resinged.Value = ""
        txtReg_Filed = ""
    End If
    picOption.Visible = False
    picRegi.Visible = True
    picRegi.ZOrder 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (EMPLOYEE INFO)"
            Call frmALL_AuditInquiry.DisplayHistory(LabID, "EMPLOYEE INFO")
            'End If

        Case vbKeyF1: cmdPrevious_Click
        Case vbKeyF2: cmdNext_Click
        Case vbKeyF3: cmdFind_Click
            'Case vbKeyF4: cmdAdd_Click' USED IN DROP DOWN OF COMBO
        Case vbKeyF5: cmdEdit_Click
        Case vbKeyF6: cmdDelete_Click
        Case vbKeyF7: cmdPrint_Click
        Case vbKeyF8: cmdExit_Click
        Case vbKeyF9: cmdSave_Click
        Case vbKeyF10: cmdCancel_Click
        Case vbKeyEscape:
            If picSalary.Visible = True Then
                picSalary.Visible = False
                DisableForm True
            ElseIf picRegi.Visible = True Then
                picRegi.Visible = False
                picOption.Visible = True
                DisableForm False
            ElseIf picOption.Visible = True Then
                picOption.Visible = False
                DisableForm True
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'DrawXPCtl Me
    Set EMPINFOEMPNO = frmHRMSEmpInfo.labEmpNo
    rsrefresh
    InitMemvars
    FillPayrollGroup
    FillGrid
    EnableAddorEdit False
    EMPINFOSHOW = True
    Screen.MousePointer = 0
    StoreMemVars
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EMPINFOSHOW = False
    UnloadForm Me
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    'orig
    'On Error Resume Next
    'rsEmpInfo.Bookmark = rsFIND(rsEmpInfo.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    'Call StoreMemVars
    
    rsEmpInfo.Requery
    rsEmpInfo.Find ("EMPNO=" & ITEM.ListSubItems(1).Text)
    StoreMemVars
    
    'ki mark working
    'Call DisplayInformation(ITEM.ListSubItems(1).Text)
End Sub

Sub DisplayInformation(XEMPNO As String)
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("select * from hrms_Empinfo where empno = '" & XEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
         Call RefreshFingerPrint
        EnableAddorEdit False
        cboSALCODE.Text = Null2String(RSTMP!SalaryCode)
        cboBlood.Text = Trim(Null2String(RSTMP!BloodType))
        cboPayroll.Text = Null2String(RSTMP!payrolltype)
        LabID.Caption = RSTMP!ID
        labEmpNo.Caption = Null2String(RSTMP!EMPNO)
        txtEmpNo.Text = Null2String(RSTMP!EMPNO)
        txtAcctNo.Text = Null2String(RSTMP!ACCOUNTNO)
        If IsNull(RSTMP!DEPTCODE) = True Then
            FillCboDepName
        Else
            cboDeptName.Text = SetCboDepName(RSTMP!DEPTCODE)
        End If
        If IsNull(RSTMP!SalaryCode) = True Then
            FillCboSalaryGradeLevel
        Else
            cboSalaryGradeLevel.Text = SetCboSalaryLevel(RSTMP!SalaryCode)
        End If
        txtLastName.Text = Null2String(RSTMP!lastname)
        txtFirstName.Text = Null2String(RSTMP!FIRSTNAME)
        txtMiddleName.Text = Null2String(RSTMP!MIDDLENAME)
        txtAddress.Text = Null2String(RSTMP!ADDRESS)
        txtTelephone.Text = Null2String(RSTMP!TELEPHONE)
        txtBirthDate.Text = Null2String(RSTMP!BIRTHDATE)
        If Null2String(RSTMP!SEX) = "M" Then
            cboSex.Text = "Male"
        Else
            cboSex.Text = "Female"
        End If
        cboStatus.Text = Null2String(RSTMP!STATUS)
        txtBirthPlace.Text = Null2String(RSTMP!BIRTHPLACE)
        txtHeight.Text = Null2String(RSTMP!HEIGHT)
        txtWeight.Text = Null2String(RSTMP!WEIGHT)
        txtReligion.Text = Null2String(RSTMP!RELIGION)
        txtCitizen.Text = Null2String(RSTMP!CITIZEN)
        txtSSSNo.Text = Null2String(RSTMP!SSSNO)
        txtTINNo.Text = Null2String(RSTMP!tinno)
        txtPHNo.Text = Null2String(RSTMP!PHNO)
        txtPagIbigNo.Text = Null2String(RSTMP!pagibigno)
        If Null2String(RSTMP!withprevious) = "Y" Then
            chkWithPrevious.Value = 1
        Else
            chkWithPrevious.Value = 0
        End If
        If chkWithPrevious.Value = 1 Then
            txtCompanyName.Text = Null2String(RSTMP!PreviousCompany)
        Else
            txtCompanyName.Enabled = False
        End If
        cboExStatus.Text = Null2String(RSTMP!EXSTATUS)
        cboPosition.Text = Null2String(RSTMP!Position)
        If Null2String(RSTMP!EMPSTATUS) = "M" Then
            cboEmpStatus.Text = "Monthly"
        Else
            cboEmpStatus.Text = "Daily"
        End If
        txtDateHired.Text = Null2String(RSTMP!DateHired)
        txtResigned.Text = Null2String(RSTMP!RESIGNED)
        txtSpouse.Text = Null2String(RSTMP!SPOUSE)
        txtSpouseAge.Text = Null2String(RSTMP!SPOUSEAGE)
        txtSOccupation.Text = Null2String(RSTMP!SOCCUPATION)
        txtFather.Text = Null2String(RSTMP!FATHER)
        txtFatherAge.Text = Null2String(RSTMP!FATHERAGE)
        txtFOccupation.Text = Null2String(RSTMP!FOCCUPATION)
        txtMother.Text = Null2String(RSTMP!MOTHER)
        txtBASICSALARY = FormatNumber(NumericVal(RSTMP!BASICSALARY))
        txtMotherAge.Text = Null2String(RSTMP!MOTHERAGE)
        txtMOccupation.Text = Null2String(RSTMP!MOCCUPATION)
        txtPerson.Text = Null2String(RSTMP!person)
        txtRelation.Text = Null2String(RSTMP!Relation)
        txtRelTelNo.Text = Null2String(RSTMP!reltelno)
        labPicfilename.Caption = Null2String(RSTMP!PICFILNAME)
        If Null2String(RSTMP!PayrollGroup) <> "" Then
            cboPayrollGroup = Null2String(RSTMP!PayrollGroup)
        Else
            cboPayrollGroup = ""
        End If
        txtALLOWANCE = FormatNumber(NumericVal(RSTMP!ALLOWANCE))
        labSHIFTCODE.Caption = Null2String(RSTMP!Shift)
        labSHIFTSCHED.Caption = SetShiftSched(Null2String(RSTMP!Shift))
        If Null2Bool(RSTMP!IS_TECHNICIAN) = False And Null2Bool(RSTMP!IS_SAE) = False And Null2Bool(RSTMP!IS_PARTS_SALESMAN) = False And Null2Bool(RSTMP!IS_SERVICE_ADVISER) = False Then
            chkPos(0).Visible = False
            chkPos(1).Visible = False
            chkPos(2).Visible = False
            chkPos(3).Visible = False
        End If
        vIS_TECHNICIAN = Null2Bool(RSTMP!IS_TECHNICIAN)
        vIS_PARTSSA = Null2Bool(RSTMP!IS_PARTS_SALESMAN)
        vIS_SA = Null2Bool(RSTMP!IS_SERVICE_ADVISER)
        vIS_SAE = Null2Bool(RSTMP!IS_SAE)
                
        If Null2Bool(RSTMP!IS_SERVICE_ADVISER) = True Then
            cmdSAType.Enabled = True
        Else
            cmdSAType.Enabled = False
        End If
        If Null2Bool(RSTMP!IS_TECHNICIAN) = True Then
            cmdTechPost.Enabled = True
        Else
            cmdTechPost.Enabled = False
        End If
        
        Call SetDMISFunc
        If Null2Bool(RSTMP!SUBJECT_TO_COLA) = True Then
            Check1.Value = 1
            txtCOLA_RATE.Text = N2Str2Zero(RSTMP!COLA_RATE)
        Else
            Check1.Value = 0
            txtCOLA_RATE.Text = 0
        End If
        If Null2String(RSTMP!ACTIVEINACTIVE) = "A" Then
            optActive.Value = True
        Else
            optInActive.Value = True
        End If
        If Null2String(RSTMP!PICFILNAME) <> "" Then
            On Error Resume Next
            If Len(Dir(HRMS_PICTURES_PATH & Null2String(RSTMP!PICFILNAME))) <= 0 Then
                Exit Sub
            End If
            LoadPic imgDispPic, HRMS_PICTURES_PATH & Null2String(RSTMP!PICFILNAME)
        Else
            LoadPic imgDispPic, ""
        End If
        LABSLEVEL.Caption = GetSLevel(NumericVal(RSTMP!BASICSALARY))
        Label45.Caption = GetPlevel(RSTMP!DateHired)
        PICTfilname = ""
        If NumericVal(RSTMP!BASICSALARY) = 0 Then
            lblBasicSalary.Visible = True
        End If
        If Not NumericVal(RSTMP!BASICSALARY) = 0 Then
            lblBasicSalary.Visible = False
        End If
        If NumericVal(RSTMP!ALLOWANCE) = 0 Then
            lblAllowance.Visible = True
        End If
        If Not NumericVal(RSTMP!ALLOWANCE) = 0 Then
            lblAllowance.Visible = False
        End If
        If Null2String(RSTMP!EMPSTATUS) = "D" Then
            cmdAttendance.Enabled = True
        End If
        If Not Null2String(RSTMP!EMPSTATUS) = "D" Then
            cmdAttendance.Enabled = False
        End If
        If IsDate(RSTMP!RESIGNED) = True Then
            txtReg_Filed = Null2String(RSTMP!RESIGNED_FILED)
            txtReg_Resinged = Null2String(RSTMP!RESIGNED)
            txtReg_Notes = Null2String(RSTMP!RESIGNED_NOTES)
        End If
        
        'UPDATE BY   : MJP012609 0536PM
        'DESCRIPTION : TO DISPLAY THE TECHNICIAN TYPE
            Call DisplayPosition(Null2String(RSTMP!EMPNO))
        'UPDATE BY   : MJP012609 0536PM
    End If
    Set RSTMP = Nothing
End Sub

Private Sub optActiveEmployee_Click()
    If HEADOREMP = "HEAD" Then
        HEADOREMP = "HEAD"
    Else
        If HEADOREMP = "EMP_A" Then
            HEADOREMP = "EMP_U"
        Else
            HEADOREMP = "EMP_A"
        End If
    End If
    rsrefresh
    StoreMemVars
    FillGrid
End Sub

Private Sub optInActiveEmployee_Click()
    If HEADOREMP = "HEAD" Then
        HEADOREMP = "HEAD"
    Else
        If HEADOREMP = "EMP_U" Then
            HEADOREMP = "EMP_A"
        Else
            HEADOREMP = "EMP_U"
        End If
    End If
    rsrefresh
    StoreMemVars
    FillGrid
End Sub

Private Sub optViewOtherInfo_Click()
    EMPLOYEE_NO = N2Str2Null(txtEmpNo.Text)
    frmOTHERInfoMain.Show vbModal
End Sub

Private Sub Timer1_Timer()
    If lblBasicSalary.ForeColor = vbBlack Then
        lblBasicSalary.ForeColor = vbRed
        lblAllowance.ForeColor = vbRed
    Else
        lblBasicSalary.ForeColor = vbBlack
        lblAllowance.ForeColor = vbBlack
    End If
End Sub

Private Sub txtAddress_LostFocus()
    If txtAddress.Text <> "" Then
        txtAddress.Text = Cap1st(txtAddress.Text)
    End If
End Sub

Private Sub txtALLOWANCE_GotFocus()
    If txtALLOWANCE = "0.00" Then
        txtALLOWANCE = ""
    End If
End Sub

Private Sub txtALLOWANCE_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtALLOWANCE_LostFocus()
    If IsNumeric(txtALLOWANCE) = False Then
        txtALLOWANCE = "0.00"
    Else
        txtALLOWANCE = FormatNumber(NumericVal(txtALLOWANCE))
    End If
End Sub

Private Sub txtBASICSALARY_GotFocus()
    If txtBASICSALARY = "0.00" Then
        txtBASICSALARY = ""
        LABSLEVEL = ""
    End If
End Sub

Private Sub txtBASICSALARY_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtBASICSALARY_LostFocus()
    If IsNumeric(txtBASICSALARY) = False Then
        txtBASICSALARY = "0.00"
        LABSLEVEL = ""
    Else
        txtBASICSALARY = FormatNumber(NumericVal(txtBASICSALARY))
        LABSLEVEL = GetSLevel(NumericVal(txtBASICSALARY))
    End If
End Sub

Private Sub txtBirthPlace_LostFocus()
    If txtBirthPlace.Text <> "" Then
        txtBirthPlace.Text = Cap1st(txtBirthPlace.Text)
    End If
End Sub

Private Sub txtCitizen_LostFocus()
    If txtCitizen.Text <> "" Then
        txtCitizen.Text = Cap1st(txtCitizen.Text)
    End If
End Sub

Private Sub txtCompanyName_GotFocus()
    If txtCompanyName.Text = "Please Enter Previous Company" Then
        txtCompanyName.Text = ""
    End If
End Sub

Private Sub txtFather_LostFocus()
    If txtFather.Text <> "" Then
        txtFather.Text = Cap1st(txtFather.Text)
    End If
End Sub

Private Sub txtFirstName_LostFocus()
    If txtFirstName.Text <> "" Then
        txtFirstName.Text = Cap1st(txtFirstName.Text)
    End If
End Sub

Private Sub txtFOccupation_LostFocus()
    If txtFOccupation.Text <> "" Then
        txtFOccupation.Text = Cap1st(txtFOccupation.Text)
    End If
End Sub

Private Sub txtLastName_LostFocus()
    If txtLastName.Text <> "" Then
        txtLastName.Text = Cap1st(txtLastName.Text)
    End If
End Sub

Private Sub txtMiddleName_LostFocus()
    If txtMiddleName.Text <> "" Then
        txtMiddleName.Text = Cap1st(txtMiddleName.Text)
    End If
End Sub

Private Sub txtMOccupation_LostFocus()
    If txtMOccupation.Text <> "" Then
        txtMOccupation.Text = Cap1st(txtMOccupation.Text)
    End If
End Sub

Private Sub txtMother_LostFocus()
    If txtMother.Text <> "" Then
        txtMother.Text = Cap1st(txtMother.Text)
    End If
End Sub

Private Sub txtPerson_LostFocus()
    If txtPerson.Text <> "" Then
        txtPerson.Text = Cap1st(txtPerson.Text)
    End If
End Sub

Private Sub txtRelation_LostFocus()
    If txtRelation.Text <> "" Then
        txtRelation.Text = Cap1st(txtRelation.Text)
    End If
End Sub

Private Sub txtReligion_LostFocus()
    If txtReligion.Text <> "" Then
        txtReligion.Text = Cap1st(txtReligion.Text)
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSOccupation_LostFocus()
    If txtSOccupation.Text <> "" Then
        txtSOccupation.Text = Cap1st(txtSOccupation.Text)
    End If
End Sub

Private Sub txtSpouse_LostFocus()
    If txtSpouse.Text <> "" Then
        txtSpouse.Text = Cap1st(txtSpouse.Text)
    End If
End Sub

