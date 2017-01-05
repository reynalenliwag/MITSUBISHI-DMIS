VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmHRMS_Process13th 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process 13th Month Pay"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   14370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Process13th.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   14370
   Begin FlexCell.Grid Grid1 
      Height          =   6615
      Left            =   0
      TabIndex        =   48
      Top             =   480
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11668
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   12632256
      Rows            =   30
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   6090
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picadd 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   10740
      ScaleHeight     =   795
      ScaleWidth      =   3615
      TabIndex        =   12
      Top             =   7140
      Width           =   3615
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   2520
         Top             =   0
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   2880
         MouseIcon       =   "frmHRMS_Process13th.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   2160
         MouseIcon       =   "frmHRMS_Process13th.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1440
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmHRMS_Process13th.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post "
         Height          =   795
         Left            =   720
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmHRMS_Process13th.frx":1E89
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":1FDB
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Genrate"
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmHRMS_Process13th.frx":2300
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":2452
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.ComboBox cboYear1 
      Height          =   330
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   60
      Width           =   1755
   End
   Begin VB.PictureBox picGEN 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   4343
      ScaleHeight     =   3525
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   2198
      Visible         =   0   'False
      Width           =   5685
      Begin VB.ComboBox cboMONTH 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   510
         Width           =   4095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   4860
         MouseIcon       =   "frmHRMS_Process13th.frx":2765
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":28B7
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   1140
         Width           =   705
      End
      Begin MSComctlLib.ProgressBar prbD 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   2070
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.ComboBox cboyear 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   1935
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Ok"
         Height          =   795
         Left            =   4140
         MouseIcon       =   "frmHRMS_Process13th.frx":2BF5
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":2D47
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Process Generation of Payroll"
         Top             =   1140
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar prbN 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   2700
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TO MONTH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   27
         Top             =   600
         Width           =   1380
      End
      Begin VB.Label lblNAME 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE NAME: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   150
         TabIndex        =   8
         Top             =   3060
         Width           =   1455
      End
      Begin VB.Label lblDEPT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPARTMENT  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YEAR"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Index           =   0
         Left            =   750
         TabIndex        =   4
         Top             =   1260
         Width           =   690
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Index           =   0
         Left            =   -30
         TabIndex        =   2
         Top             =   0
         Width           =   5715
         _Version        =   655364
         _ExtentX        =   10081
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   " Generate 13th Month Pay"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   5078
      ScaleHeight     =   2175
      ScaleWidth      =   4185
      TabIndex        =   18
      Top             =   2873
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox cboYEAR2 
         Height          =   330
         Left            =   2520
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   630
         Width           =   1575
      End
      Begin VB.ComboBox cboMONTH1 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   630
         Width           =   2355
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1395
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Payslip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   1050
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   3360
         MouseIcon       =   "frmHRMS_Process13th.frx":30B5
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":3207
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Cancel"
         Top             =   1290
         Width           =   705
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         Height          =   795
         Left            =   2640
         MouseIcon       =   "frmHRMS_Process13th.frx":3545
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":3697
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print this Record"
         Top             =   1290
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Confidential"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   21
         Top             =   1590
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Regular"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   20
         Top             =   1860
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   1
         Left            =   2580
         TabIndex        =   29
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   28
         Top             =   390
         Width           =   525
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   5025
         _Version        =   655364
         _ExtentX        =   8864
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   " Print 13th Month Pay"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
   Begin VB.PictureBox picDeduction 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4088
      ScaleHeight     =   2505
      ScaleWidth      =   6165
      TabIndex        =   34
      Top             =   2708
      Visible         =   0   'False
      Width           =   6195
      Begin VB.ComboBox cboDed 
         Height          =   330
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1080
         Width           =   2985
      End
      Begin VB.TextBox txtEmpno 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   360
         Width           =   1905
      End
      Begin VB.TextBox txtDetAmount 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1500
         TabIndex        =   41
         Top             =   1440
         Width           =   1905
      End
      Begin VB.TextBox txtEmployeeName 
         Height          =   315
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   4545
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   795
         Left            =   5340
         MouseIcon       =   "frmHRMS_Process13th.frx":39FD
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":3B4F
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cancel"
         Top             =   1620
         Width           =   705
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   795
         Left            =   4650
         MouseIcon       =   "frmHRMS_Process13th.frx":3E8D
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Process13th.frx":3FDF
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancel"
         Top             =   1620
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deduction For"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   5
         Left            =   285
         TabIndex        =   46
         Top             =   1170
         Width           =   1140
      End
      Begin VB.Label lblOldAmount 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   3930
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   42
         Top             =   450
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   3
         Left            =   105
         TabIndex        =   39
         Top             =   825
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   2
         Left            =   765
         TabIndex        =   38
         Top             =   1530
         Width           =   660
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   6315
         _Version        =   655364
         _ExtentX        =   11139
         _ExtentY        =   529
         _StockProps     =   14
         Caption         =   " Input Deductions"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   795
      Index           =   1
      Left            =   30
      TabIndex        =   44
      Top             =   7140
      Width           =   14355
      _Version        =   655364
      _ExtentX        =   25321
      _ExtentY        =   1402
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label labStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10980
      TabIndex        =   33
      Top             =   60
      Width           =   3330
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F77C48&
      BackStyle       =   1  'Opaque
      Height          =   945
      Left            =   300
      Shape           =   4  'Rounded Rectangle
      Top             =   1260
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View Existing Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   150
      Width           =   1740
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   465
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   14355
      _Version        =   655364
      _ExtentX        =   25321
      _ExtentY        =   820
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuPSlip 
         Caption         =   "Print Payslip"
      End
      Begin VB.Menu mnuPSum 
         Caption         =   "Print Summary"
      End
   End
End
Attribute VB_Name = "frmHRMS_Process13th"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS13TH As ADODB.Recordset
Dim xlApp                                           As Excel.Application
Dim xlbook                                          As Excel.Workbook
Dim xlsheet                                         As Excel.Worksheet
Dim xSEC_FROM_CUTDATE                               As Integer
Dim xSEC_TO_CUTDATE                                 As Integer

Private Sub cboYear1_Change()
    cboYear1.Enabled = False
    Call CheckIf13thMonthExist
    cboYear1.Enabled = True
End Sub

Private Sub cboYear1_Click()
    cboYear1.Enabled = False
    Call CheckIf13thMonthExist
    cboYear1.Enabled = True
End Sub

Sub CheckIf13thMonthExist()
    Dim RSTMP                                       As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT EMPNO, PAYROLLSTATUS FROM HRMS_PAYROLL WHERE PAY_MONTH = 13 AND PAY_YEAR = " & cboYear1 & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!payrollstatus) = "P" Then
            labStatus.Caption = "** POSTED **"
            cmdPost.Enabled = False
            cmdUnPost.Enabled = True
        Else
            labStatus.Caption = ""
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
        End If
        Screen.MousePointer = 11
        Call FillGrid2
        Screen.MousePointer = 0
    Else
        Call ShowNoRecord
        labStatus.Caption = ""
        Grid1.Rows = 1
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdAdd_Click()
    Grid1.Rows = 1
    cboYear1.Enabled = False
    picadd.Enabled = False
    Grid1.Enabled = False
    
    cboMONTH.Text = MonthName(MONTH(Date))
    prbD.Value = 0
    prbN.Value = 0
    lblDEPT.Caption = "":    lblNAME.Caption = "":    labStatus.Caption = ""
    picGEN.Visible = True
    picGEN.ZOrder 0
End Sub

Private Sub cmdCancel_Click()
    picGEN.Visible = False
    picGEN.ZOrder 1
    picadd.Enabled = True
    Grid1.Enabled = True
    cboYear1.Enabled = True
    
    cboYear1.Text = YEAR(Date)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    If MsgBox("Process 13th Month Pay for this Year", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    Call GetSecondCutOffDate
    Call FillGrid
    Screen.MousePointer = 0
End Sub

Private Sub cmdPost_Click()
    If Grid1.Rows = 1 Then Exit Sub
    
    If MsgBox("Post this 13th Month Pay", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("UPDATE HRMS_PAYROLL SET " & _
        " PAYROLLSTATUS = 'P' " & _
        " WHERE PAY_MONTH = 13 " & _
        " AND PAY_YEAR = " & cboYear1 & "")
    labStatus.Caption = "** POSTED **"
    cmdUnPost.Enabled = True
    cmdPost.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    If Grid1.Rows = 1 Then
        ShowNoRecord
        cboYear1.SetFocus
        Exit Sub
    End If
    cboMONTH1.Text = MonthName(MONTH(Date))
    cboYear1.Enabled = False
    picadd.Enabled = False
    Grid1.Enabled = False
    
    picPrint.Visible = True
    picPrint.ZOrder 0
End Sub

Private Sub cmdUnPost_Click()
    If Grid1.Rows = 1 Then Exit Sub
    
    If MsgBox("UnPost this 13th Month Pay", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("UPDATE HRMS_PAYROLL SET " & _
        " PAYROLLSTATUS = NULL " & _
        " WHERE PAY_MONTH = 13 " & _
        " AND PAY_YEAR = " & cboYear1 & "")
    labStatus.Caption = ""
    cmdPost.Enabled = True
    cmdUnPost.Enabled = False
End Sub

Private Sub Command1_Click()
    If Grid1.Rows = 1 Then Exit Sub
    
    Screen.MousePointer = 11
    If Option1.Value = True Then
        If Check2.Value = 1 Then
            rpt1.WindowTitle = "13th Month Pay"
            rpt1.ReportTitle = "13th Month Pay"
            rpt1.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rpt1.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rpt1, HRMS_REPORT_PATH & "13thMonthPay_REG.rpt", "MONTH({HRMS_EMPINFO.DATEHIRED}) <= " & What_month(cboMONTH1) & " AND YEAR({HRMS_EMPINFO.DATEHIRED}) <= " & cboYEAR2 & " AND {HRMS_payroll.PAY_MONTH} = " & 13 & " AND {HRMS_payroll.PAY_YEAR} = " & cboYear1 & "", DMIS_REPORT_Connection, 1
            rpt1.Reset
        End If
        
        If Check1.Value = 1 Then
            rpt1.WindowTitle = "13th Month Pay"
            rpt1.ReportTitle = "13th Month Pay"
            rpt1.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
            rpt1.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rpt1, HRMS_REPORT_PATH & "13thMonthPay.rpt", "MONTH({HRMS_EMPINFO.DATEHIRED}) <= " & What_month(cboMONTH1) & " AND YEAR({HRMS_EMPINFO.DATEHIRED}) <= " & cboYEAR2 & " AND {HRMS_payroll.PAY_MONTH} = " & 13 & " AND {HRMS_payroll.PAY_YEAR} = " & cboYear1 & "", DMIS_REPORT_Connection, 1
            rpt1.Reset
        End If
    Else
        Call PRINTEXCEL
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
    picPrint.Visible = False
    picPrint.ZOrder 1
    
    picadd.Enabled = True
    Grid1.Enabled = True
    cboYear1.Enabled = True
End Sub

Private Sub Command3_Click()
    cboYear1.Enabled = True
    Grid1.Enabled = True
    picadd.Enabled = True
    
    picDeduction.Visible = False
    picDeduction.ZOrder 1
End Sub

Private Sub Command4_Click()
    If MsgBox("Save this 13th Month Pay deduction", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub

    gconDMIS.Execute ("UPDATE HRMS_PAYROLL SET " & _
        " DEDUCTABLE13 = " & NumericVal(txtDetAmount) & _
        ", NET_AMT13 = ROUND(NET13,2) - " & NumericVal(txtDetAmount) & _
        ", DEDUCTION_CODE = " & N2Str2Null(GetDeductionCode(cboDed)) & _
        " WHERE EMPNO = " & N2Str2Null(txtEmpno) & _
        " AND PAY_YEAR = " & cboYear1 & _
        " AND PAY_MONTH = 13 ")
    
    Grid1.Cell(Grid1.ActiveCell.Row, 17).Text = Format(NumericVal(txtDetAmount), "#,###,##0.00")
    Grid1.Cell(Grid1.ActiveCell.Row, 18).Text = Format((NumericVal(Grid1.Cell(Grid1.ActiveCell.Row, 18).Text) + NumericVal(lblOldAmount)) - NumericVal(txtDetAmount), "#,###,##0.00")
    Call Command3_Click
    Call ShowSuccessFullyUpdated
End Sub

Function GetDeductionCode(XXX As String) As String
    Dim RSTMP                                   As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CODE FROM HRMS_DEDUCTIONCODE WHERE DESCRIPTION = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetDeductionCode = Null2String(RSTMP.FIELDS(0))
    End If
    Set RSTMP = Nothing
End Function

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
            
    Call fillcbomonth(cboMONTH1)
    Call fillcbomonth(cboMONTH)
    'Call FillcboYear(cboyear)
    'Call FillcboYear(cboYear1)
    'Call FillcboYear(cboYEAR2)
    Call FillCboDeduction
    
    
    Call fillcombo_up(cboyear)
    Call fillcombo_up(cboYear1)
    Call fillcombo_up(cboYEAR2)
    
End Sub

Sub FillCboDeduction()
    Dim RSTMP                                       As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DESCRIPTION FROM HRMS_DEDUCTIONCODE WHERE " & _
        " ISNULL(CODE,'') NOT IN ('LT','WD') ORDER BY DESCRIPTION ASC")
    cboDed.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        cboDed.AddItem ""
        Do While Not RSTMP.EOF
            cboDed.AddItem Null2String(RSTMP.FIELDS(0))
            RSTMP.MoveNext
        Loop
        cboDed.ListIndex = 0
    End If
    Set RSTMP = Nothing
End Sub

Sub FillGrid()
    Dim RSTMP                                       As New ADODB.Recordset
    Dim RSDEPT                                      As New ADODB.Recordset
    Dim RSEMPL                                      As New ADODB.Recordset
    Dim CNT                                         As Integer
    Dim RG                                          As FlexCell.Range
    Dim GRANDTOTAL                                  As Double
    Dim XVAT                                        As Double
    Dim X13THPAY                                    As Double
    Dim XOVERWRITE                                  As String
    Dim xJAN As Double: Dim XFEB As Double: Dim XMAR As Double: Dim XAPR As Double
    Dim XMAY As Double: Dim XJUN As Double: Dim XJUL As Double: Dim XAUG As Double
    Dim XSEP As Double: Dim XOCT As Double: Dim XNOV As Double: Dim XDEC As Double
    Dim LEAVE_DAY_CNT                               As Integer
    Dim XBASIC_T                                    As Double
    Dim XBASIC_NT                                   As Double
    Dim XALLOWANCE                                  As Double
    Dim xLASTDATE                                   As String
    
    XOVERWRITE = "NO"
    Set RSTMP = gconDMIS.Execute("SELECT PAYROLLSTATUS FROM HRMS_PAYROLL WHERE PAY_MONTH = 13 AND PAY_YEAR = " & cboyear & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!payrollstatus) = "P" Then
            MsgBox "13th Month pay already generated and posted", vbInformation, "HRMS"
            Exit Sub
        Else
            If MsgBox("13th Month pay already generated, do you want to Overwrite", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
            gconDMIS.Execute ("UPDATE HRMS_PAYROLL SET ALLOWANCE13 = 0, TAX13 = 0, NETPAY = 0, BASIC_T = 0, BASIC_NT = 0 WHERE PAY_MONTH = 13 AND PAY_YEAR = " & cboyear & "")
            XOVERWRITE = "YES"
        End If
    Else
    
    End If
    Set RSTMP = Nothing
    
    xLASTDATE = lastDay(What_month(cboMONTH) & "/1/" & cboyear)
    Set RSDEPT = gconDMIS.Execute("SELECT DISTINCT DEPTCODE FROM HRMS_EMPINFO ORDER BY DEPTCODE")
    
    Call InitGrid
    Grid1.Rows = 1
    CNT = 0
    Dim ITEM_CNT As Integer
    prbD.Value = 0
                        
    If Not (RSDEPT.BOF And RSDEPT.EOF) Then
        prbD.Max = RSDEPT.RecordCount
        Do While Not RSDEPT.EOF
            CNT = CNT + 1
            Grid1.AddItem CNT & ". " & FindDepartmentName(Null2String(RSDEPT!DEPTCODE))
            ITEM_CNT = ITEM_CNT + 1
            Set RG = Grid1.Range(ITEM_CNT, 1, ITEM_CNT, 19)
            DoEvents
            prbD.Value = prbD.Value + 1
            lblDEPT.Caption = "DEPARTMENT : " & FindDepartmentName(Null2String(RSDEPT!DEPTCODE))
            DoEvents
            
            Dim X As Integer
            For X = 0 To 19
                Grid1.Column(X).Locked = False
            Next
            RG.Merge
            RG.BackColor = &HFCE2CF
            RG.BackColor = &HFCE2CF
            RG.Borders(cellEdgeTop) = cellThin
            RG.Borders(cellEdgeBottom) = cellThin
            RG.Locked = True
            
            For X = 0 To 19
                Grid1.Column(X).Locked = True
            Next
            
'
'            Set RSEMPL = gconDMIS.Execute("SELECT ALLOWANCE, BASICSALARY, DATEHIRED, EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE DEPTCODE = '" & Null2String(RSDEPT!DEPTCODE) & _
'                "' AND RESIGNED IS NULL " & _
'                " AND YEAR(DATEHIRED) <= " & cboYear & _
'                " AND MONTH(DATEHIRED) <= " & What_month(cboMOnth) & _
'                " ORDER BY LASTNAME")
            Set RSEMPL = gconDMIS.Execute("SELECT ALLOWANCE, BASICSALARY, DATEHIRED, EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE DEPTCODE = '" & Null2String(RSDEPT!DEPTCODE) & _
                "' AND RESIGNED IS NULL " & _
                " AND DATEHIRED <= " & N2Str2Null(xLASTDATE) & _
                " ORDER BY LASTNAME")
            If Not (RSEMPL.BOF And RSEMPL.EOF) Then
                prbN.Max = RSEMPL.RecordCount
                prbN.Value = 0
                Do While Not RSEMPL.EOF
                    DoEvents
                    prbN.Value = prbN.Value + 1
                    
                    
                   ' If RSEMPL!FULLNAME = "Edonio, Armina" Then Stop
                    
                    
                    
                    lblNAME.Caption = "EMPLOYEE NAME : " & Null2String(RSEMPL!FULLNAME)
                    xJAN = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 1, NumericVal(RSEMPL!BASICSALARY))
                    XFEB = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 2, NumericVal(RSEMPL!BASICSALARY))
                    XMAR = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 3, NumericVal(RSEMPL!BASICSALARY))
                    XAPR = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 4, NumericVal(RSEMPL!BASICSALARY))
                    XMAY = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 5, NumericVal(RSEMPL!BASICSALARY))
                    XJUN = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 6, NumericVal(RSEMPL!BASICSALARY))
                    XJUL = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 7, NumericVal(RSEMPL!BASICSALARY))
                    XAUG = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 8, NumericVal(RSEMPL!BASICSALARY))
                    XSEP = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 9, NumericVal(RSEMPL!BASICSALARY))
                    XOCT = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 10, NumericVal(RSEMPL!BASICSALARY))
                    XNOV = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 11, NumericVal(RSEMPL!BASICSALARY))
                    XDEC = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 12, NumericVal(RSEMPL!BASICSALARY))
                    
                    GRANDTOTAL = ((NumericVal(RSEMPL!BASICSALARY)) / 12) * (GetGrandTotal(Null2String(RSEMPL!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSEMPL!EMPNO), cboyear))
                    XALLOWANCE = ((NumericVal(RSEMPL!ALLOWANCE)) / 12) * (GetGrandTotal(Null2String(RSEMPL!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSEMPL!EMPNO), cboyear))
                    If GRANDTOTAL > 30000 Then
                        XBASIC_T = Format(NumericVal(RSEMPL!BASICSALARY) - 30000, "#,###,##0.00")
                        XBASIC_NT = Format(30000, "#,###,##0.00")
                        XVAT = Format(GetFromTaxTable(GetEmployeeField(Null2String(RSEMPL!EMPNO), "EXSTATUS"), XBASIC_T, "MONTHLY BASE"), "#,###,##0.00")
                    Else
                        XBASIC_T = Format(0, "#,###,##0.00")
                        XBASIC_NT = GRANDTOTAL
                        XVAT = Format(0, "#,###,##0.00")
                    End If
                    X13THPAY = Format((XBASIC_NT + XBASIC_T + XALLOWANCE) - XVAT, "#,###,##0.00")
                    
                    Grid1.AddItem "     " & Null2String(RSEMPL!FULLNAME) & Chr(9) & _
                        Format(xJAN, "#,###,##0.00") & Chr(9) & _
                        Format(XFEB, "#,###,##0.00") & Chr(9) & _
                        Format(XMAR, "#,###,##0.00") & Chr(9) & _
                        Format(XAPR, "#,###,##0.00") & Chr(9) & _
                        Format(XMAY, "#,###,##0.00") & Chr(9) & _
                        Format(XJUN, "#,###,##0.00") & Chr(9) & _
                        Format(XJUL, "#,###,##0.00") & Chr(9) & _
                        Format(XAUG, "#,###,##0.00") & Chr(9) & _
                        Format(XSEP, "#,###,##0.00") & Chr(9) & _
                        Format(XOCT, "#,###,##0.00") & Chr(9) & _
                        Format(XNOV, "#,###,##0.00") & Chr(9) & _
                        Format(XDEC, "#,###,##0.00") & Chr(9) & _
                        Format(GRANDTOTAL, "#,###,##0.00") & Chr(9) & _
                        Format(XVAT, "#,###,##0.00") & Chr(9) & _
                        Format(X13THPAY, "#,###,##0.00") & Chr(9) & _
                        Format(0, "#,###,##0.00") & Chr(9) & _
                        Format(X13THPAY, "#,###,##0.00") & Chr(9) & _
                        Null2String(RSEMPL!EMPNO), False
                    DoEvents
                    
                    If XOVERWRITE = "NO" Then
                        gconDMIS.Execute ("INSERT INTO HRMS_PAYROLL (EMPNO, NETPAY, PAY_MONTH, PAY_YEAR, TAX13, NET13, BASIC_T, BASIC_NT, ALLOWANCE13, DEDUCTABLE13, NET_AMT13) " & _
                            " VALUES(" & N2Str2Null(RSEMPL!EMPNO) & _
                            ", " & X13THPAY & _
                            ", " & 13 & _
                            ", " & cboyear & _
                            ", " & XVAT & _
                            ", " & X13THPAY & _
                            ", " & XBASIC_T & _
                            ", " & XBASIC_NT & _
                            ", " & XALLOWANCE & _
                            ", 0 " & _
                            ", " & X13THPAY & ")")
                    Else
                        gconDMIS.Execute ("UPDATE HRMS_PAYROLL SET NETPAY = " & X13THPAY & _
                            ", TAX13 = " & XVAT & _
                            ", NET13 = " & X13THPAY & _
                            ", BASIC_T = " & XBASIC_T & _
                            ", BASIC_NT = " & XBASIC_NT & _
                            ", ALLOWANCE13 = " & XALLOWANCE & _
                            ", DEDUCTABLE13 = 0 " & _
                            ", NET_AMT13 = " & X13THPAY & _
                            " WHERE PAY_MONTH = " & 13 & _
                            " AND PAY_YEAR = " & cboyear & _
                            " AND EMPNO = " & N2Str2Null(RSEMPL!EMPNO) & "")
                    End If
                    
                    ITEM_CNT = ITEM_CNT + 1
                    RSEMPL.MoveNext
                Loop
            End If
            Set RSEMPL = Nothing
            RSDEPT.MoveNext
        Loop
    End If
    Set RSDEPT = Nothing
    Grid1.Refresh
    
    Grid1.Enabled = True
    cboYear1.Enabled = True
    picGEN.Visible = False
    picGEN.ZOrder 1
    picadd.Enabled = True
End Sub

Function GetGrandTotal(xDATEHIRED As String) As Double
    Dim X                                               As Integer
    Dim MONTH_CNT                                       As Double
    Dim XTOTAL                                          As Double
    MONTH_CNT = 0
    If (xDATEHIRED) = "" Then
        GetGrandTotal = 0
        Exit Function
    End If
    If YEAR(Date) = YEAR(xDATEHIRED) Then
        For X = MONTH(xDATEHIRED) To 12
            If X = MONTH(xDATEHIRED) Then
                If Day(xDATEHIRED) < 16 Then
                    MONTH_CNT = MONTH_CNT + 1
                Else
                    MONTH_CNT = MONTH_CNT + 0.5
                End If
            Else
                MONTH_CNT = MONTH_CNT + 1
            End If
        Next
    Else
        MONTH_CNT = 12
    End If
    GetGrandTotal = MONTH_CNT
End Function

Function CheckIfTheresMaternityLeave(XEMPNO As String, XYEAR As Integer) As Double
    Dim RSTMP                                       As New ADODB.Recordset
    Dim LEAVE_CNT                                   As Double
    Set RSTMP = gconDMIS.Execute("SELECT * FROM HRMS_REQUESTLEAVE_OT WHERE " & _
        " EMPNO = " & N2Str2Null(XEMPNO) & _
        " AND PAY_YEAR = " & XYEAR & _
        " AND REQCODE = 'ML' " & _
        " AND STATUS = 'A'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            If Null2String(RSTMP!ML_TYPE) = "Cesarean" Then
                LEAVE_CNT = LEAVE_CNT + 2.6
            Else
                LEAVE_CNT = LEAVE_CNT + 2
            End If
            RSTMP.MoveNext
        Loop
    End If
    If LEAVE_CNT = 0 Then
        CheckIfTheresMaternityLeave = 0
    Else
        CheckIfTheresMaternityLeave = LEAVE_CNT
    End If
    Set RSTMP = Nothing
End Function

Function GetEmployeePerDayRate(XEMPNO As String) As Double
    
End Function

Sub GetSecondCutOffDate()
    Dim RSTMP                                   As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT FROMDATE2, TODATE2 FROM HRMS_PAYROLLSETUP")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        xSEC_FROM_CUTDATE = RSTMP.FIELDS(0)
        xSEC_TO_CUTDATE = RSTMP.FIELDS(1)
    End If
    Set RSTMP = Nothing
End Sub

Function GetEmployeeBasic(xDATEHIRED As String, XMONTH As Integer, XSALARY As Double) As Double
    If Not xDATEHIRED = "" Then
        If YEAR(Now) = YEAR(xDATEHIRED) Then
            'If XMONTH >= MONTH(xDATEHIRED) Then
            '    GetEmployeeBasic = XSALARY
            'Else
            '    GetEmployeeBasic = 0
            'End If
            
            If XMONTH > MONTH(xDATEHIRED) Then
                GetEmployeeBasic = XSALARY
            ElseIf XMONTH = MONTH(xDATEHIRED) Then
                If (Day(xDATEHIRED) < 6) Or (Day(xDATEHIRED) > 20) Then
                    GetEmployeeBasic = XSALARY
                Else
                    GetEmployeeBasic = XSALARY / 2
                End If
            Else
                GetEmployeeBasic = 0
            End If
        Else
            GetEmployeeBasic = XSALARY
        End If
    Else
        GetEmployeeBasic = XSALARY
    End If
End Function

Function GetEmployeeField(XEMPNO As String, xFIELD As String) As String
    Dim RSTMP                                           As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT " & xFIELD & " AS XFIELD FROM HRMS_EMPINFO " & _
        " WHERE EMPNO = '" & XEMPNO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetEmployeeField = Null2String(RSTMP!xFIELD)
    End If
    Set RSTMP = Nothing
End Function

Sub FillGrid2()
    Dim RSDEPT                                          As New ADODB.Recordset
    Dim RSEMPL                                          As New ADODB.Recordset
    Dim CNT                                             As Integer
    Dim RG                                              As FlexCell.Range
    Dim GRANDTOTAL                                      As Double
    Dim XVAT                                            As Double
    Dim X13THPAY                                        As Double
    Dim xJAN As Double: Dim XFEB As Double: Dim XMAR As Double: Dim XAPR As Double
    Dim XMAY As Double: Dim XJUN As Double: Dim XJUL As Double: Dim XAUG As Double
    Dim XSEP As Double: Dim XOCT As Double: Dim XNOV As Double: Dim XDEC As Double
    Dim XBASIC_NT                                       As Double
    Dim XBASIC_T                                        As Double
    Dim RS13TH                                          As New ADODB.Recordset
    Dim XALLOWANCE                                      As Double
    Set RSDEPT = gconDMIS.Execute("SELECT DISTINCT DEPTCODE FROM HRMS_EMPINFO ORDER BY DEPTCODE")
    Call InitGrid
    Grid1.Rows = 1
    
    CNT = 0
    Dim ITEM_CNT As Integer
    Dim X As Integer
    For X = 0 To 19
        Grid1.Column(X).Locked = False
    Next
    If Not (RSDEPT.BOF And RSDEPT.EOF) Then
        Do While Not RSDEPT.EOF
            CNT = CNT + 1
            Grid1.AddItem CNT & ". " & FindDepartmentName(Null2String(RSDEPT!DEPTCODE))
            ITEM_CNT = ITEM_CNT + 1
            Set RG = Grid1.Range(ITEM_CNT, 1, ITEM_CNT, 18)

            RG.Merge
            RG.BackColor = &HFCE2CF
            RG.BackColor = &HFCE2CF
            RG.Borders(cellEdgeTop) = cellThin
            RG.Borders(cellEdgeBottom) = cellThin
            
            Set RSEMPL = gconDMIS.Execute("SELECT ALLOWANCE, BASICSALARY, DATEHIRED, EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME FROM HRMS_EMPINFO WHERE " & _
                " DEPTCODE = '" & Null2String(RSDEPT!DEPTCODE) & _
                "' AND RESIGNED IS NULL ORDER BY LASTNAME")
            If Not (RSEMPL.BOF And RSEMPL.EOF) Then
                Do While Not RSEMPL.EOF
                    DoEvents
                    Set RS13TH = gconDMIS.Execute("SELECT DEDUCTABLE13, NET13 FROM HRMS_PAYROLL WHERE PAY_MONTH = 13 " & _
                        " AND PAY_YEAR = " & cboYear1 & _
                        " AND NET13 <> 0 AND EMPNO = " & N2Str2Null(RSEMPL!EMPNO) & "")
                    If Not (RS13TH.BOF And RS13TH.EOF) Then
                        ITEM_CNT = ITEM_CNT + 1
                        xJAN = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 1, NumericVal(RSEMPL!BASICSALARY))
                        XFEB = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 2, NumericVal(RSEMPL!BASICSALARY))
                        XMAR = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 3, NumericVal(RSEMPL!BASICSALARY))
                        XAPR = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 4, NumericVal(RSEMPL!BASICSALARY))
                        XMAY = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 5, NumericVal(RSEMPL!BASICSALARY))
                        XJUN = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 6, NumericVal(RSEMPL!BASICSALARY))
                        XJUL = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 7, NumericVal(RSEMPL!BASICSALARY))
                        XAUG = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 8, NumericVal(RSEMPL!BASICSALARY))
                        XSEP = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 9, NumericVal(RSEMPL!BASICSALARY))
                        XOCT = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 10, NumericVal(RSEMPL!BASICSALARY))
                        XNOV = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 11, NumericVal(RSEMPL!BASICSALARY))
                        XDEC = GetEmployeeBasic(Null2String(RSEMPL!DateHired), 12, NumericVal(RSEMPL!BASICSALARY))
                        
                        GRANDTOTAL = (NumericVal(RSEMPL!BASICSALARY) / 12) * (GetGrandTotal(Null2String(RSEMPL!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSEMPL!EMPNO), cboYear1))
                        XALLOWANCE = ((NumericVal(RSEMPL!ALLOWANCE)) / 12) * (GetGrandTotal(Null2String(RSEMPL!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSEMPL!EMPNO), cboYear1))
                        If GRANDTOTAL > 30000 Then
                            XBASIC_T = NumericVal(RSEMPL!BASICSALARY) - 30000
                            XBASIC_NT = 30000
                            XVAT = GetFromTaxTable(GetEmployeeField(Null2String(RSEMPL!EMPNO), "EXSTATUS"), XBASIC_T, "MONTHLY BASE")
                        Else
                            XVAT = Format(0, "#,###,##0.00")
                            XBASIC_T = Format(0, "#,###,##0.00")
                            XBASIC_NT = GRANDTOTAL
                        End If
                        X13THPAY = Format((XBASIC_NT + XBASIC_T + XALLOWANCE) - XVAT, "#,###,##0.00")
                        
                        Grid1.AddItem "     " & Null2String(RSEMPL!FULLNAME) & Chr(9) & _
                            Format(xJAN, "#,###,##0.00") & Chr(9) & _
                            Format(XFEB, "#,###,##0.00") & Chr(9) & _
                            Format(XMAR, "#,###,##0.00") & Chr(9) & _
                            Format(XAPR, "#,###,##0.00") & Chr(9) & _
                            Format(XMAY, "#,###,##0.00") & Chr(9) & _
                            Format(XJUN, "#,###,##0.00") & Chr(9) & _
                            Format(XJUL, "#,###,##0.00") & Chr(9) & _
                            Format(XAUG, "#,###,##0.00") & Chr(9) & _
                            Format(XSEP, "#,###,##0.00") & Chr(9) & _
                            Format(XOCT, "#,###,##0.00") & Chr(9) & _
                            Format(XNOV, "#,###,##0.00") & Chr(9) & _
                            Format(XDEC, "#,###,##0.00") & Chr(9) & _
                            Format(GRANDTOTAL, "#,###,##0.00") & Chr(9) & _
                            Format(XVAT, "#,###,##0.00") & Chr(9) & _
                            Format(X13THPAY, "#,###,##0.00") & Chr(9) & _
                            Format(NumericVal(RS13TH!DEDUCTABLE13), "#,###,##0.00") & Chr(9) & _
                            Format(X13THPAY - NumericVal(RS13TH!DEDUCTABLE13), "#,###,##0.00") & Chr(9) & _
                            RSEMPL!EMPNO, False
                    End If
                    Set RS13TH = Nothing
                    
                    RSEMPL.MoveNext
                Loop
            End If
            Set RSEMPL = Nothing
            RSDEPT.MoveNext
        Loop
    End If
    Grid1.AutoRedraw = True
    Set RSDEPT = Nothing
    For X = 0 To 19
        Grid1.Column(X).Locked = True
    Next
    Grid1.Refresh
    
    Grid1.Enabled = True
    picGEN.Visible = False
    picGEN.ZOrder 1
    picadd.Enabled = True

    Exit Sub
    
exit_me:
    
End Sub

Function Get13thMonthPay(XEMPNO As String) As Double
    Dim Amount13thPay As Double
    Dim X As Integer
    
    For X = 1 To 12
        Amount13thPay = Amount13thPay + GetBasicSalary(XEMPNO, X, cboyear)
    Next
    Get13thMonthPay = Amount13thPay
End Function

Function GetBasicSalary(XEMPNO As String, XMONTH As Integer, XYEAR As Integer) As Double
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT SUM(RATE) AS PAY_MONTH, SUM(UNDERTIME) AS LATE_MONTH, SUM(ABSENT) AS ABSENT_MONTH FROM HRMS_PAYROLL WHERE EMPNO = '" & XEMPNO & _
        "' AND PAY_MONTH = " & XMONTH & _
        " AND PAY_YEAR = " & XYEAR & _
        " AND PAYROLLSTATUS = 'P'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetBasicSalary = NumericVal(RSTMP!PAY_MONTH)  '- (NumericVal(RSTMP!LATE_MONTH) + NumericVal(RSTMP!ABSENT_MONTH))
    End If
    Set RSTMP = Nothing
End Function

Function FindDepartmentName(XXX As String) As String
    Dim RSTMP As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DEPTNAME FROM HRMS_DEPARTMENT WHERE DEPTCODE = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindDepartmentName = Null2String(RSTMP!DEPTNAME)
    End If
    Set RSTMP = Nothing
End Function

Sub InitGrid()
    With Grid1
        .Rows = 1
        .Cols = 20
        
        .DefaultRowHeight = 18
        .RowHeight(0) = 20
        .Cell(0, 0).Text = "L/N"
        .Column(0).Width = 20
        .Column(0).Locked = True

        .Cell(0, 1).Text = "DEPARTMENT"
        .Column(1).Width = 217
        .Column(1).Locked = True

        .Cell(0, 2).Text = "JAN"
        .Column(2).Width = 70
        .Column(2).Alignment = cellRightGeneral
        .Column(2).Locked = True

        .Cell(0, 3).Text = "FEB"
        .Column(3).Width = 70
        .Column(3).Alignment = cellRightGeneral
        .Column(3).Locked = True

        .Cell(0, 4).Text = "MAR"
        .Column(4).Width = 70
        .Column(4).Alignment = cellRightGeneral
        .Column(4).Locked = True

        .Cell(0, 5).Text = "APR"
        .Column(5).Width = 70
        .Column(5).Alignment = cellRightGeneral
        .Column(5).Locked = True

        .Cell(0, 6).Text = "MAY"
        .Column(6).Width = 70
        .Column(6).Alignment = cellRightGeneral
        .Column(6).Locked = True

        .Cell(0, 7).Text = "JUN"
        .Column(7).Width = 70
        .Column(7).Alignment = cellRightGeneral
        .Column(7).Locked = True

        .Cell(0, 8).Text = "JUL"
        .Column(8).Width = 70
        .Column(8).Alignment = cellRightGeneral
        .Column(8).Locked = True

        .Cell(0, 9).Text = "AUG"
        .Column(9).Width = 70
        .Column(9).Alignment = cellRightGeneral
        .Column(9).Locked = True
        
        .Cell(0, 10).Text = "SEP"
        .Column(10).Width = 70
        .Column(10).Alignment = cellRightGeneral
        .Column(10).Locked = True
        
        .Cell(0, 11).Text = "OCT"
        .Column(11).Width = 70
        .Column(11).Alignment = cellRightGeneral
        .Column(11).Locked = True
        
        .Cell(0, 12).Text = "NOV"
        .Column(12).Width = 70
        .Column(12).Alignment = cellRightGeneral
        .Column(12).Locked = True
        
        .Cell(0, 13).Text = "DEC"
        .Column(13).Width = 70
        .Column(13).Alignment = cellRightGeneral
        .Column(13).Locked = True
        
        .Cell(0, 14).Text = "AMOUNT"
        .Column(14).Width = 70
        .Column(14).Alignment = cellRightGeneral
        .Column(14).Locked = True
        
        .Cell(0, 15).Text = "W/TAX"
        .Column(15).Width = 80
        .Column(15).Alignment = cellRightGeneral
        .Column(15).Locked = True
                        
        .Cell(0, 16).Text = "13TH MONTH PAY"
        .Column(16).Width = 100
        .Column(16).Alignment = cellRightGeneral
        .Column(16).Locked = True
        
        .Cell(0, 17).Text = "DEDUCTABLE"
        .Column(17).Width = 100
        .Column(17).Alignment = cellRightGeneral
        .Column(17).Locked = True
        
        .Cell(0, 18).Text = "NET AMOUNT"
        .Column(18).Width = 100
        .Column(18).Alignment = cellRightGeneral
        .Column(18).Locked = True
        
        .Cell(0, 19).Text = "EMPNO"
        .Column(19).Width = 0
        .Column(19).Alignment = cellCenterCenter
        .Column(19).Locked = True
        
        .Range(0, 0, 0, 10).WrapText = True
   End With
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Grid1_DblClick()
    If labStatus.Caption = "** POSTED **" Then
        MessagePop InfoFriend, "Action Denied", "13th Month pay is already posted"
        Exit Sub
    End If
    
    txtEmployeeName.Text = LTrim(RTrim(Grid1.Cell(Grid1.ActiveCell.Row, 1).Text))
    txtEmpno.Text = LTrim(RTrim(Grid1.Cell(Grid1.ActiveCell.Row, 19).Text))
    
    Dim RSTMP                               As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = " & N2Str2Null(txtEmpno) & "")
    If (RSTMP.BOF And RSTMP.EOF) Then
        Exit Sub
    End If
    Set RSTMP = Nothing
    
    picDeduction.Visible = True
    picDeduction.ZOrder 0
    Set RSTMP = gconDMIS.Execute("SELECT DEDUCTABLE13, ISNULL(DEDUCTION_CODE,'') FROM HRMS_PAYROLL WHERE " & _
        " PAY_MONTH = 13 " & _
        " AND PAY_YEAR = " & cboYear1 & _
        " AND EMPNO = " & N2Str2Null(txtEmpno) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblOldAmount.Caption = NumericVal(RSTMP!DEDUCTABLE13)
        txtDetAmount.Text = NumericVal(RSTMP!DEDUCTABLE13)
        If GetDeductionDesc(Null2String(RSTMP.FIELDS(1))) <> "" Then
            cboDed.Text = GetDeductionDesc(Null2String(RSTMP.FIELDS(1)))
        Else
            cboDed.ListIndex = 0
        End If
    Else
        lblOldAmount.Caption = "0.00"
        txtDetAmount.Text = "0.00"
        cboDed.ListIndex = 0
    End If
    Set RSTMP = Nothing
    txtDetAmount.SetFocus
    
    cboYear1.Enabled = False
    picadd.Enabled = False
    Grid1.Enabled = False
End Sub

Private Sub txtDetAmount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Function GetDeductionDesc(XXX As String) As String
    Dim RSTMP                                       As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT DESCRIPTION FROM HRMS_DEDUCTIONCODE WHERE " & _
        " CODE = " & N2Str2Null(XXX) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetDeductionDesc = Null2String(RSTMP.FIELDS(0))
    End If
    Set RSTMP = Nothing
End Function

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Private Sub Timer1_Timer()
    If labStatus.ForeColor = vbRed Then
        labStatus.ForeColor = vbBlack
    Else
        labStatus.ForeColor = vbRed
    End If
End Sub

Function GetFromTaxTable(TAXCODE As String, EMPSAL_GROSS As Variant, EMP_PAYROLL_TYPE As String)
    GetFromTaxTable = 0
    Dim RSTAX                                                         As New ADODB.Recordset
    Dim COLNO                                                         As Integer
    Dim RESULT_TAX                                                    As Double
    Dim rsTemporaryRecordSet As New ADODB.Recordset
    Set rsTemporaryRecordSet = gconDMIS.Execute("SELECT * FROM HRMS_TAXTABLEDETAILS WHERE TAXBASIS = '" & EMP_PAYROLL_TYPE & "' AND TAXCODE = '" & TAXCODE & "'")
    If Not (rsTemporaryRecordSet.BOF And rsTemporaryRecordSet.EOF) Then
        If EMPSAL_GROSS >= 1 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col2 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col1
            COLNO = 1
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col2 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col3 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col2
            COLNO = 2
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col3 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col4 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col3
            COLNO = 3
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col4 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col5 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col4
            COLNO = 4
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col5 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col6 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col5
            COLNO = 5
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col6 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col7 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col6
            COLNO = 6
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col7 And EMPSAL_GROSS <= (rsTemporaryRecordSet!Col8 - 1) Then
            RESULT_TAX = rsTemporaryRecordSet!Col7
            COLNO = 7
        End If
        If EMPSAL_GROSS >= rsTemporaryRecordSet!Col8 Then
            RESULT_TAX = rsTemporaryRecordSet!Col8
            COLNO = 8
        End If
        Set RSTAX = gconDMIS.Execute("SELECT * FROM HRMS_TAXTABLE WHERE TAXBASIS = '" & EMP_PAYROLL_TYPE & "'")
        If Not (RSTAX.BOF And RSTAX.EOF) Then
            If COLNO = 1 Then
                GetFromTaxTable = 1
            End If
            If COLNO = 2 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per2)
            End If
            If COLNO = 3 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per3) + RSTAX!EXp3
            End If
            If COLNO = 4 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per4) + RSTAX!EXp4
            End If
            If COLNO = 5 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per5) + RSTAX!EXp5
            End If
            If COLNO = 6 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per6) + RSTAX!EXp6
            End If
            If COLNO = 7 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per7) + RSTAX!EXp7
            End If
            If COLNO = 8 Then
                GetFromTaxTable = ((EMPSAL_GROSS - RESULT_TAX) * RSTAX!Per8) + RSTAX!EXp8
            End If
        End If
    End If
    Set rsTemporaryRecordSet = Nothing
End Function

Sub PRINTEXCEL()
    Dim RSTMP                                       As New ADODB.Recordset
    Dim RSDEPT                                      As New ADODB.Recordset
    Dim XLOC                                        As Integer
    Dim XTAX                                        As Double
    Dim xcnt                                        As Integer
    Dim XTOTAL                                      As Double
    Dim XCOND                                       As String
    Dim XBASIC_T                                    As Double
    Dim XBASIC_NT                                   As Double
    Dim XALLOWANCE                                  As Double
    Dim XDEPT_CNT                                   As Integer
    Dim X1                                          As Double
    Dim X2                                          As Double
    Dim X3                                          As Double
    Dim X4                                          As Double
    Dim X5                                          As Double
    Dim X6                                          As Double
    
    
    If Check2.Value = 1 Then
        If Check1.Value = 1 Then
            XCOND = ""
        Else
            XCOND = " AND EMPLEVEL = 'E' "
        End If
    End If
    
    If Check1.Value = 1 Then
        If Check2.Value = 1 Then
            XCOND = ""
        Else
            XCOND = " AND EMPLEVEL = 'M' "
        End If
    End If
    
    
    Set xlApp = CreateObject("Excel.Application")
    Set xlbook = xlApp.Workbooks.Open(HRMS_REPORT_PATH & "13thMonthSummary.xlt")
    Set xlsheet = xlbook.Worksheets(1)
    xlsheet.Cells(1, "A") = COMPANY_NAME
    xlsheet.Cells(2, "A") = COMPANY_ADDRESS
    xlsheet.Cells(3, "A") = "13th Month Pay for the year " & cboYear1
    XLOC = 5
    Set RSDEPT = gconDMIS.Execute("SELECT DISTINCT DEPTCODE FROM HRMS_EMPINFO ORDER BY DEPTCODE")
    If Not (RSDEPT.BOF And RSDEPT.EOF) Then
        Do While Not RSDEPT.EOF
            xlsheet.Cells(XLOC, "A") = FindDepartmentName(Null2String(RSDEPT!DEPTCODE))
            xlsheet.Range("A" & XLOC & ":" & "I" & XLOC).BorderAround ColorIndex:=1, WEIGHT:=xlThin, Color:=&H800080
            xlsheet.Range("A" & XLOC & ":" & "I" & XLOC).Interior.Color = &HFADDB6
            
            XDEPT_CNT = 0
            X1 = 0: X2 = 0: X3 = 0: X4 = 0: X5 = 0: X6 = 0
            XLOC = XLOC + 1
            Set RSTMP = gconDMIS.Execute("SELECT ALLOWANCE, DATEHIRED, EMPNO, LASTNAME + ', ' + FIRSTNAME AS FULLNAME, DATEHIRED, BASICSALARY, ALLOWANCE, PAYROLLTYPE, EXSTATUS FROM HRMS_EMPINFO WHERE DEPTCODE = '" & Null2String(RSDEPT!DEPTCODE) & "' " & XCOND & " AND RESIGNED IS NULL ORDER BY LASTNAME")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                Do While Not RSTMP.EOF
                    xlsheet.Cells(XLOC, "B") = Null2String(RSTMP!FULLNAME)
                    xlsheet.Cells(XLOC, "C") = Null2String(RSTMP!DateHired)
                    XTOTAL = (NumericVal(RSTMP!BASICSALARY) / 12) * (GetGrandTotal(Null2String(RSTMP!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSTMP!EMPNO), cboYear1))
                    XALLOWANCE = ((NumericVal(RSTMP!ALLOWANCE)) / 12) * (GetGrandTotal(Null2String(RSTMP!DateHired)) - CheckIfTheresMaternityLeave(Null2String(RSTMP!EMPNO), cboYear1))
                    If XTOTAL > 30000 Then
                        XBASIC_T = Format(NumericVal(RSTMP!BASICSALARY) - 30000, "#,###,##0.00")
                        XBASIC_NT = Format(30000, "#,###,##0.00")
                        XTAX = Format(GetFromTaxTable(GetEmployeeField(Null2String(RSTMP!EMPNO), "EXSTATUS"), XBASIC_T, "MONTHLY BASE"), "#,###,##0.00")
                    Else
                        XTAX = Format(0, "#,###,##0.00")
                        XBASIC_T = Format(0, "#,###,##0.00")
                        XBASIC_NT = NumericVal(XTOTAL)
                    End If
'                    xlsheet.Cells(XLOC, "D") = Format(NumericVal(XTOTAL), "#,###,##0.00")
'                    xlsheet.Cells(XLOC, "E") = Format(NumericVal(XALLOWANCE), "#,###,##0.00")
'                    xlsheet.Cells(XLOC, "F") = Format(XTAX, "#,###,##0.00")
'                    xlsheet.Cells(XLOC, "G") = Format((XBASIC_NT + XBASIC_T + NumericVal(XALLOWANCE)) - XTAX, "#,###,##0.00")
'                    xlsheet.Cells(XLOC, "H") = ReturnAmountFrom13monthPay("DEDUCTABLE13", RSTMP!EMPNO, cboYEAR2)
'                    xlsheet.Cells(XLOC, "I") = ReturnAmountFrom13monthPay("NET_AMT13", RSTMP!EMPNO, cboYEAR2)
                    
                    xlsheet.Cells(XLOC, "D") = ReturnAmountFrom13monthPay("BASIC_NT", RSTMP!EMPNO, cboYEAR2)
                    xlsheet.Cells(XLOC, "E") = ReturnAmountFrom13monthPay("BASIC_T", RSTMP!EMPNO, cboYEAR2)
                    xlsheet.Cells(XLOC, "F") = ReturnAmountFrom13monthPay("ALLOWANCE13", RSTMP!EMPNO, cboYEAR2)
                    xlsheet.Cells(XLOC, "G") = ReturnAmountFrom13monthPay("TAX13", RSTMP!EMPNO, cboYEAR2)
                    xlsheet.Cells(XLOC, "H") = ReturnAmountFrom13monthPay("DEDUCTABLE13", RSTMP!EMPNO, cboYEAR2)
                    xlsheet.Cells(XLOC, "I") = ReturnAmountFrom13monthPay("NET_AMT13", RSTMP!EMPNO, cboYEAR2)
                    
                    X1 = X1 + ReturnAmountFrom13monthPay("BASIC_NT", RSTMP!EMPNO, cboYEAR2)
                    X2 = X2 + ReturnAmountFrom13monthPay("BASIC_T", RSTMP!EMPNO, cboYEAR2)
                    X3 = X3 + ReturnAmountFrom13monthPay("ALLOWANCE13", RSTMP!EMPNO, cboYEAR2)
                    X4 = X4 + ReturnAmountFrom13monthPay("TAX13", RSTMP!EMPNO, cboYEAR2)
                    X5 = X5 + ReturnAmountFrom13monthPay("DEDUCTABLE13", RSTMP!EMPNO, cboYEAR2)
                    X6 = X6 + ReturnAmountFrom13monthPay("NET_AMT13", RSTMP!EMPNO, cboYEAR2)
                    XDEPT_CNT = XDEPT_CNT + 1
                    XLOC = XLOC + 1
                    RSTMP.MoveNext
                Loop
            End If
            Set RSTMP = Nothing
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "D") = Format(X1, "#,###,##0.00")
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "E") = Format(X2, "#,###,##0.00")
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "F") = Format(X3, "#,###,##0.00")
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "G") = Format(X4, "#,###,##0.00")
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "H") = Format(X5, "#,###,##0.00")
            xlsheet.Cells(XLOC - (XDEPT_CNT + 1), "I") = Format(X6, "#,###,##0.00")
            RSDEPT.MoveNext
        Loop
    End If
    Set RSDEPT = Nothing
    
    xlApp.Windows.ITEM(1).Caption = "13th Month Pay"
    xlApp.Visible = True
    Set xlApp = Nothing
    Set xlsheet = Nothing
    Set xlbook = Nothing
End Sub

Function ReturnAmountFrom13monthPay(xTABLE As String, XEMPNO As String, XYEAR As Integer) As Double
    Dim RSTMP                                       As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT " & xTABLE & " FROM HRMS_PAYROLL WHERE " & _
        " EMPNO = " & N2Str2Null(XEMPNO) & _
        " AND PAY_MONTH = 13 " & _
        " AND PAY_YEAR = " & XYEAR & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ReturnAmountFrom13monthPay = (NumericVal(RSTMP.FIELDS(0)))
    End If
    Set RSTMP = Nothing
End Function

Function GetSalaryFromPayroll(XEMPNO As String, SDATEHIRED As String, xBASIC As Double) As Double
    Dim RSTMP As New ADODB.Recordset
    Dim XTOTAL As Double
    Dim xmon As Integer
    Dim X As Integer
    
    If Not SDATEHIRED = "" Then
        If YEAR(SDATEHIRED) = YEAR(Now) Then
            xmon = MONTH(SDATEHIRED)
        Else
            xmon = 1
        End If
    Else
        xmon = 1
    End If
    
    For X = xmon To 12
        XTOTAL = XTOTAL + xBASIC
    Next
    GetSalaryFromPayroll = XTOTAL / 12
    Set RSTMP = Nothing
End Function
