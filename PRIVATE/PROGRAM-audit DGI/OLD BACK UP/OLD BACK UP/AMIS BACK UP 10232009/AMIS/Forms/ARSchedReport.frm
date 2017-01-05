VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmAMISARSchedReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERATE AR REPORT"
   ClientHeight    =   2490
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "ARSchedReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   4170
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   5310
      TabIndex        =   51
      Top             =   2040
      Width           =   585
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   1425
      Left            =   4320
      TabIndex        =   49
      Top             =   3540
      Width           =   2805
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   1275
      Left            =   6180
      TabIndex        =   48
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtdescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8460
      TabIndex        =   18
      Top             =   450
      Width           =   4275
   End
   Begin VB.ComboBox cboacctcode 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8460
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   450
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Refresh Accounts Receivables"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   8460
      TabIndex        =   7
      Top             =   450
      Width           =   4965
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   4035
         MouseIcon       =   "ARSchedReport.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Process Import of Accounts Payable"
         Top             =   1020
         Width           =   705
      End
      Begin wizProgBar.Prg progCPB 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   630
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   556
         Picture         =   "ARSchedReport.frx":122F
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ARSchedReport.frx":124B
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpTranDate 
         Height          =   405
         Left            =   930
         TabIndex        =   10
         Top             =   1020
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   38216
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "AS OF : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   150
         TabIndex        =   12
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label labCPB 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   150
         TabIndex        =   11
         Top             =   330
         Width           =   5835
      End
   End
   Begin VB.OptionButton optForthePeriod 
      Caption         =   "Accounts Receivable for the Period"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   8670
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.OptionButton optAsOf 
      Caption         =   "Accounts Receivable as Of"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00701E2A&
      Height          =   255
      Left            =   8460
      TabIndex        =   0
      Top             =   450
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3585
   End
   Begin Crystal.CrystalReport rptAMISDueReport 
      Left            =   90
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Accounts Receivable Aging Report"
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
   Begin VB.Frame picPeriod 
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
      Left            =   8460
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpFrom 
         Height          =   315
         Left            =   780
         TabIndex        =   4
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   38216
      End
      Begin MSComCtl2.DTPicker dtpTo 
         Height          =   315
         Left            =   2730
         TabIndex        =   6
         Top             =   180
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   38216
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   210
         Width           =   675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   2220
         TabIndex        =   5
         Top             =   210
         Width           =   435
      End
   End
   Begin MSComctlLib.ProgressBar PROGBAR 
      Height          =   405
      Left            =   8460
      TabIndex        =   14
      Top             =   450
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picReport 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2685
      ScaleWidth      =   4005
      TabIndex        =   29
      Top             =   0
      Width           =   4005
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   1830
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "AR ADHOC REPORT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   38
         Top             =   1410
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "AR AGING REPORT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   31
         Top             =   990
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "AR SCHEDULE REPORT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         MaskColor       =   &H0080FFFF&
         TabIndex        =   32
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Process AR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         MaskColor       =   &H0080FFFF&
         TabIndex        =   45
         Top             =   120
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtprocess 
         Height          =   375
         Left            =   300
         TabIndex        =   46
         Top             =   120
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   38216
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2685
      Left            =   90
      ScaleHeight     =   2685
      ScaleWidth      =   4005
      TabIndex        =   22
      Top             =   0
      Width           =   4005
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   360
         ScaleHeight     =   465
         ScaleWidth      =   285
         TabIndex        =   37
         Top             =   840
         Width           =   285
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Group by Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   36
         Top             =   510
         Width           =   2565
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Group by Customer "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   930
         TabIndex        =   35
         Top             =   240
         Value           =   -1  'True
         Width           =   2565
      End
      Begin MSComCtl2.DTPicker dtpAsOF 
         Height          =   345
         Left            =   180
         TabIndex        =   24
         Top             =   2700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20643841
         CurrentDate     =   38216
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2070
         MouseIcon       =   "ARSchedReport.frx":1267
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport.frx":13B9
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Close Window"
         Top             =   1200
         Width           =   885
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1200
         MouseIcon       =   "ARSchedReport.frx":1804
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Print Report"
         Top             =   1200
         Width           =   885
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2460
         TabIndex        =   28
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Last data  generated:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   27
         Top             =   2430
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "As Of:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture4 
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
      Height          =   2745
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   3945
      TabIndex        =   19
      Top             =   30
      Visible         =   0   'False
      Width           =   3945
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   435
         Left            =   180
         TabIndex        =   20
         Top             =   420
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblRef 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1290
         TabIndex        =   34
         Top             =   930
         Width           =   2355
      End
      Begin VB.Label Label10 
         Caption         =   "Processing:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   33
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   120
         Width           =   825
      End
   End
   Begin VB.PictureBox picNolink 
      BorderStyle     =   0  'None
      Height          =   2865
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   3915
      TabIndex        =   39
      Top             =   -240
      Width           =   3915
      Begin VB.OptionButton Option8 
         Caption         =   "APJ having AR account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   52
         Top             =   1650
         Width           =   3435
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Wrong Customer Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   50
         Top             =   1410
         Width           =   3435
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CRJ with Different Account Set Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   47
         Top             =   1170
         Width           =   3435
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1890
         MouseIcon       =   "ARSchedReport.frx":1DF5
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport.frx":1F47
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Close Window"
         Top             =   2040
         Width           =   885
      End
      Begin VB.OptionButton Option3 
         Caption         =   "CRJ No Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton Option4 
         Caption         =   "CRJ Blank Invoice"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   41
         Top             =   630
         Width           =   2775
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Vendor Having AR Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   330
         TabIndex        =   40
         Top             =   900
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1020
         MouseIcon       =   "ARSchedReport.frx":2392
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport.frx":24E4
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Print Report"
         Top             =   2040
         Width           =   885
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8460
      TabIndex        =   17
      Top             =   450
      Width           =   1125
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8460
      TabIndex        =   15
      Top             =   450
      Width           =   4185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "TEMPORARY TOOL TO REFRESH CUSTOMER AR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   8460
      TabIndex        =   13
      Top             =   450
      Width           =   4935
   End
End
Attribute VB_Name = "frmAMISARSchedReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report_type                                        As String
Dim rs                                                 As New ADODB.Recordset
Dim CountDetail                                        As Integer
Dim Dealer As String
Function SetCRJVoucherNo(XXX As String, zzz As Integer) As String
    Dim rsCRJ_Journal_HD                               As ADODB.Recordset
    Set rsCRJ_Journal_HD = New ADODB.Recordset
    If zzz = 1 Then
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and InvoiceNo = '" & XXX & "'")
    Else
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = '" & XXX & "'")
    End If
    If Not rsCRJ_Journal_HD.EOF And Not rsCRJ_Journal_HD.BOF Then
        SetCRJVoucherNo = Null2String(rsCRJ_Journal_HD!VOUCHERNO)
    End If
End Function

Private Sub cboacctcode_Change()
    ReturnAccountCode cboacctcode.Text
End Sub

Private Sub cboacctcode_Click()
    ReturnAccountCode cboacctcode.Text

End Sub


Private Sub Command1_Click()
    Report_type = "SCHED"
    getlastdate
    picReport.Visible = False
    Me.Caption = "AR SCHEDULE REPORT"
    If IsDate(Label8.Caption) = True Then
        dtpAsOF.Value = CDate(Label8.Caption)
    End If
End Sub

Private Sub Command10_Click()
'APJwithAR
End Sub

Private Sub Command2_Click()
    Report_type = "AGING"
    getlastdate
    picReport.Visible = False
    Picture1.Visible = True
    Me.Caption = "AR AGING REPORT"
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    picNolink.Visible = True
    picNolink.ZOrder 0
    picReport.Visible = True

End Sub

Private Sub Command5_Click()
    picNolink.Visible = False
End Sub

Private Sub Command6_Click()
    rptAMISDueReport.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAMISDueReport.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If Option3.Value = True Then
        'ProcessCRJNolink
        rptAMISDueReport.WindowTitle = "CASH RECEIPT JOURNAL NO DETAIL  AS OF: " & dtprocess
        rptAMISDueReport.ReportTitle = "CASH RECEIPT JOURNAL NO DETAIL AS OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARCRJNODETAIL.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If
    If Option4.Value = True Then
        rptAMISDueReport.WindowTitle = "CASH RECEIPT JOURNAL BLANK INVOICE AS OF: " & dtprocess
        rptAMISDueReport.ReportTitle = "CASH RECEIPT JOURNAL BLANK INVOICE AS OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\CRJBLANKINV.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If
    If Option5.Value = True Then
        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE IN CASH DISBURSEMENT JOURNAL  AS OF: " & dtprocess
        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE IN CASH DISBURSEMENT JOURNAL AS OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARCDJ.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If
    If Option6.Value = True Then
        rptAMISDueReport.WindowTitle = "CRJ WITH DEFFERENT ACCOUNT SET UP IN SALES JOURNAL AS OF: " & dtprocess
        rptAMISDueReport.ReportTitle = "CRJ WITH DEFFERENT ACCOUNT SET UP IN SALES JOURNAL AS OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\CRJWRONGCODE.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If
     If Option7.Value = True Then
        rptAMISDueReport.WindowTitle = "WRONG CUSTOMER CODE AS OF: " & dtprocess
        rptAMISDueReport.ReportTitle = "WRONG CUSTOMER CODE AS OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\AR_WrongCustomer.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If
     If Option8.Value = True Then
        rptAMISDueReport.WindowTitle = "APJ having AR account: " & dtprocess
        rptAMISDueReport.ReportTitle = "APJ having AR account OF: " & dtprocess
        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\APJWITHAR.Rpt", "", DMIS_REPORT_Connection, 1
        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
    End If

End Sub

Private Sub Command7_Click()
    dtpAsOF = dtprocess
    picReport.Visible = False
    Picture4.Visible = True
    Picture1.Visible = False

    If ar = False Then
        picReport.Visible = True
        MsgBox "No AR Date as of :" & dtpAsOF
    End If

End Sub

Private Sub Command8_Click()
    ' kit
    'gconDMIS.Execute ("DELETE FROM AMIS_AR WHERE SYSTEMREMARK IN('CRJWCA','NOINV','NL','CDJ','CRJWAC')")
    
   'ProcessCRJNolink
   'ProcessARinCDJ
   'ProcessCRJNoInvoice
   'ProcessCRJmaliangcodesaCRJtoSJ
   'ProcessCRJWithClearingAccount
   '  MsgBox "s"
End Sub

Private Sub Command9_Click()
Transfer_SalesJournal
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    If Report_AR = "SCHED" Then
        Me.Caption = "Sched of A/R"
    Else
        Me.Caption = "A/R Aging Report"
    End If
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpFrom = Month(LOGDATE) & "/1/" & Year(LOGDATE)
    dtpTo = LOGDATE
    dtpAsOF = LOGDATE
    dtpTranDate = LOGDATE
    dtprocess = LOGDATE

    GetAcctcode
    picNolink.Visible = False
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    Screen.MousePointer = 0
    getlastdate
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub cmdCancel_Click()
    picReport.Visible = True
    Me.Caption = "AR REPORT"
End Sub

Private Sub cmdPrint_Click()
    'On Error GoTo Errorcode:

    Dim Rs_CRJTotal                                    As New ADODB.Recordset
    Dim Rs_ARAing                                      As New ADODB.Recordset

    

    rptAMISDueReport.Reset
    
    rptAMISDueReport.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAMISDueReport.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If Report_type = "SCHED" Then
        If IsDate(Label8.Caption) = True Then
            If dtpAsOF.Value > CDate(Label8.Caption) Then
                MsgBox "Information: Date is greater than Last data generated"
                Exit Sub
            End If
        End If
        If Label9.Caption = "PLEASE GENERATE AR DATA" Then
            MsgBox "Please Generate AR AGING DATA..This will generate last data generated", vbInformation, "INFO"
            'Exit Sub
        End If
        If Option1.Value = True Then
            If optAsOf.Value = True Then
                rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1
                LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & Label8
            Else
                '
            End If
        ElseIf Option2.Value = True Then
            rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
            rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1
            LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & Label8
        End If
    Else
        If Report_type = "AGING" Then
            Dim Ans                                    As String
            If Option1.Value = True Then
                Ans = MsgBox("Do you want to generate Ar report?", vbQuestion + vbYesNo, "Info")
                If Ans = vbYes Then
                    'If ar = True Then
                    rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    rptAMISDueReport.Formulas(3) = "DateBasis = '" & dtpAsOF.Value & "'"
                    PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "", DMIS_REPORT_Connection, 1
                    LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & Label8
                    'Else
                    'MsgBox "No AR as of the " & Label8
                    'End If
                Else
                    If Option1.Value = True Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "", DMIS_REPORT_Connection, 1
                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & Label8
                    ElseIf Option1.Value = True Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "", DMIS_REPORT_Connection, 1
                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & Label8

                    End If
                End If
            ElseIf Option2.Value = True Then          ' Group report by Account
                rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8
                rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8
                PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "", DMIS_REPORT_Connection, 1
                LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & Label8
            End If
        End If
    End If
    getlastdate
    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub List1_Click()

End Sub

Private Sub optAsOf_Click()
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
End Sub

Private Sub optForthePeriod_Click()
    picPeriod.Enabled = True
    dtpAsOF.Enabled = False
    dtpFrom.Enabled = True
    dtpTo.Enabled = True
End Sub

Sub GetAcctcode()
    Dim SQL                                            As String
    Dim rs                                             As New ADODB.Recordset

    SQL = "SELECT DESCRIPTION from AMIS_chartaccount where left(acctcode,5)='11-02' or left(acctcode,5)='11-03' ORDER BY DESCRIPTION"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    cboacctcode.Clear

    Do While Not rs.EOF
        cboacctcode.AddItem (rs!Description)
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
Function ReturnAccountCode(Xacct_desc As String)
    Dim SQL                                            As String
    Dim rs                                             As New ADODB.Recordset

    SQL = "select description,acctcode from AMIS_chartaccount where description = '" & Xacct_desc & "'"

    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        txtdescription = Null2String(rs!acctcode)
    End If
    Set rs = Nothing
End Function
Function ar() As Boolean
    ' Update by BTT : this process kill the AR
    Dim rsHeader                                       As New ADODB.Recordset
    Dim rsdetail                                       As New ADODB.Recordset
    Dim BALANCE                                        As Double
    Dim totalpayment                                   As Double
    Dim CRJVoucher                                     As String
    Dim Reference                                      As String
    Dim SystemRemarks                                  As String
    Dim CRJInvoiceno                                   As String
    Dim CRJInvoicetype                                 As String
    Dim AMOUNT2PAY                                     As Double
    Dim invoicedate                                    As String
    Dim THECSJ                                         As String
    Dim theHDInvoice                                   As String
    Dim rscountx                                       As ADODB.Recordset
    Dim CustomerCode
    Dim RSCOUNT_ME As New ADODB.Recordset
    Dim BILANG As Integer
    Dim Counter_check As New ADODB.Recordset
    Dim countDuplicated  As New ADODB.Recordset
    Dim Validate As New ADODB.Recordset
    Dim CounterCheck_HD As New ADODB.Recordset
    THECSJ = "CSJ"
    
    
    gconDMIS.Execute ("DELETE FROM AMIS_AR")
    gconDMIS.Execute ("Update AMIS_CRJ_DETAIL set status = 'P'")
    'ValidateDetail
    Transfer_SalesJournal
    Me.Caption = "Loading Transaction.."
    
    Dim ARNIE As New ADODB.Recordset
    'Set rsHeader = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
    '                                "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")
    
    'Set rsHeader = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
    '                                "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' and dbo.AMIS_Journal_HD.CustomerCode = 'G00041' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")
    
    'Set rsHeader = gconDMIS.Execute("Select * from AMIS_AR_HD where jdate <= " & N2Str2Null(dtpAsOF) & " and SJ_customercode='f00015'")
    Set rsHeader = gconDMIS.Execute("Select * from AMIS_AR_HD where jdate <= " & N2Str2Null(dtpAsOF) & "")
    
    
       
    Dim LNGX                                           As Long
    If rsHeader.EOF Or rsHeader.BOF Then
        ar = False
        Exit Function
    Else
        ar = True
    End If

    Picture1.Visible = False
    Picture4.Visible = True
    rsHeader.MoveFirst
    ProgressBar1.Value = 0
    ProgressBar1.Max = rsHeader.RecordCount
    
    Do While Not rsHeader.EOF
        DoEvents

        Reference = Null2String(rsHeader!jtype) + "-" + Null2String(rsHeader!VOUCHERNO)
        'Invoicedate = Null2String(rsHeader!xInvoicedate)
        invoicedate = Null2String(rsHeader!jdate)
        theHDInvoice = Null2String(rsHeader!INVOICENO)
        
        
        
        If (rsHeader!jtype) = "SJ" Then
            AMOUNT2PAY = DebitTotalAmount(Null2String(rsHeader!VOUCHERNO), Null2String(rsHeader!jtype), Null2String(rsHeader!ACCT_CODE))
        ElseIf (rsHeader!jtype) = "CCM" Then
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmnt) * (-1)    'to deduct to total payment
            'AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt)
        Else
            'AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt)
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmnt)
        End If
         
 
    
                 
                 Set RSCOUNT_ME = gconDMIS.Execute("Select count(*) from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "' and status = 'P' and jdate <=" & N2Str2Null(dtpAsOF))
                 BILANG = RSCOUNT_ME(0)
                        If BILANG > 1 Then
                            CustomerCode = ReturnCustomerCode(Null2String(rsHeader!VOUCHERNO), Null2String(rsHeader!jtype))
                            Set rsdetail = gconDMIS.Execute("Select invoiceno,invoicetype,ISNULL(invoiceamount,0) AS INVOICEAMOUNT,voucherno,jdate,customercode as CRJ_customercode ,SJ_VOUCHERNO , J_CLASS,ID from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "'  and  detail_status <> 'Y'  and Status ='P' and jdate <=" & N2Str2Null(dtpAsOF))
                        Else
                            Set rsdetail = gconDMIS.Execute("Select invoiceno,invoicetype,ISNULL(invoiceamount,0) AS INVOICEAMOUNT,voucherno,jdate,customercode as CRJ_customercode ,SJ_VOUCHERNO , J_CLASS,ID from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "' and status = 'P' and jdate <=" & N2Str2Null(dtpAsOF))
                        End If
                 
       LNGX = 0
       If Not rsdetail.EOF And Not rsdetail.BOF Then
            If Null2String(rsdetail!SJ_voucherno) <> "" Then
                Set rscountx = gconDMIS.Execute("select count(*) from amis_journal_det where jtype='SJ' AND VOUCHERNO=" & N2Str2Null(rsdetail!SJ_voucherno) & " and LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) in ('11-02', '11-03') ")
                LNGX = rscountx(0)
            End If
            
            Do While Not rsdetail.EOF
                
                If LNGX <= 1 Then
                    
                    If (Null2String(rsHeader!SJ_CustomerCode) = Null2String(rsdetail!CRJ_customercode)) Then
                        CRJVoucher = Null2String(rsdetail!VOUCHERNO)
 
                            CRJInvoiceno = Null2String(rsdetail!INVOICENO)
                            CRJInvoicetype = Null2String(rsdetail!InvoiceType)
                            totalpayment = totalpayment + rsdetail!invoiceamount
                            SystemRemarks = N2Str2Null("")
                            gconDMIS.Execute ("UPDATE AMIS_CRJ_DETAIL set STATUS ='Y' where ID='" & rsdetail!ID & "'")
                            
                            
                            
                    Else
                        'wrong Customercode
                         CRJVoucher = Null2String(rsdetail!VOUCHERNO)
                         SystemRemarks = "Wrong customer code"
                         
                        
                         Dim CRJ_details As New ADODB.Recordset
                         Dim SJ_header As New ADODB.Recordset
                         Set Validate = gconDMIS.Execute("SELECT distinct AMIS_Journal_HD.CustomerCode, " & _
                                           "AMIS_Journal_HD.VoucherNo, " & _
                                           "AMIS_Journal_HD.JType, " & _
                                           "AMIS_Journal_Det.Debit, " & _
                                           "AMIS_Journal_Det.Credit, " & _
                                           "AMIS_Journal_Det.Acct_Code , " & _
                                           "AMIS_Journal_HD.Status " & _
                                           "FROM AMIS_Journal_HD INNER JOIN AMIS_Journal_Det ON AMIS_Journal_HD.JType = AMIS_Journal_Det.JType AND AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo " & _
                                           "WHERE (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') AND (AMIS_Journal_HD.JType = 'CRJ') AND (AMIS_Journal_HD.Status = 'P') and AMIS_Journal_HD.VoucherNo='" & CRJVoucher & "'")
                         
                         If Not Validate.EOF And Not Validate.BOF Then
                            'If Validate!acct_code = Null2String(RSDetail!J_CLASS) Or Validate!acct_code = Null2String(RSDetail!J_CLASS) Then
                                Set CRJ_details = gconDMIS.Execute("select invoiceno,invoicetype,voucherno,invoiceamount from AMIS_CRJ_DETAIL where voucherno = '" & CRJVoucher & "' and invoiceno='" & rsdetail!INVOICENO & "'")
                                    
                                    'Debug.Print CRJ_details.Source
                                    Do While Not CRJ_details.EOF
                                        Set SJ_header = gconDMIS.Execute("select customercode from AMIS_journal_hd where invoiceno='" & CRJ_details!INVOICENO & "' and InvoiceType = '" & CRJ_details!InvoiceType & "' and customercode='" & (rsdetail!CRJ_customercode) & "'")
                                        If SJ_header.EOF And SJ_header.BOF Then
                                     '       Debug.Print SJ_header.Source
                                            'Set countDuplicated = gconDMIS.Execute("Select Count(*) from AMIS_AR where crjvoucherno='" & CRJVoucher & "'")
                                            Set countDuplicated = gconDMIS.Execute("Select Count(*) from AMIS_AR where crjvoucherno='" & CRJVoucher & "' and invoiceno ='" & CRJ_details!INVOICENO & "'")
                                            If Not (countDuplicated(0) >= 1) Then
                                                If Not (Null2String(CRJ_details!INVOICENO) = "" Or Len(CRJ_details!INVOICENO)) < 6 Then
                                                           gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & ("CRJ" & "-" & CRJVoucher) & _
                                                                             "','" & CRJVoucher & "','" & CRJInvoicetype & "','" & CRJ_details!INVOICENO & _
                                                                             "','" & invoicedate & "','" & rsdetail!CRJ_customercode & "','" & NumericVal(CRJ_details!invoiceamount) & "','" & NumericVal(0) & _
                                                                             "','" & NumericVal(CRJ_details!invoiceamount) * (-1) & "','" & Null2String(Validate!ACCT_CODE) & "','WrongCC')")
                                                    
                                                End If
                                            End If
                                         End If
                                        CRJ_details.MoveNext
                                        Loop
                            'End If
                         End If
                    End If
                    
                Else
                    If (Null2String(rsHeader!SJ_CustomerCode) = Null2String(rsdetail!CRJ_customercode) And Null2String(rsHeader!ACCT_CODE) = Null2String(rsdetail!j_class)) Then
                        CRJVoucher = Null2String(rsdetail!VOUCHERNO)
                        CRJInvoiceno = Null2String(rsdetail!INVOICENO)
                        CRJInvoicetype = Null2String(rsdetail!InvoiceType)
                        totalpayment = totalpayment + rsdetail!invoiceamount
                        SystemRemarks = N2Str2Null("")
                    Else
                            CRJVoucher = Null2String(rsdetail!VOUCHERNO)
                        If Null2String(rsHeader!SJ_CustomerCode) <> Null2String(rsdetail!CRJ_customercode) Then
                            SystemRemarks = "Wrong customer code"
                        End If
                        
                        If Null2String(rsHeader!ACCT_CODE) = Null2String(rsdetail!j_class) Then
                            
                            SystemRemarks = SystemRemarks & " Invalid Linking"
                        End If
                    End If
                End If
                rsdetail.MoveNext
            Loop                                      'CRJ loop
        Else
            ' if no payment
            CRJVoucher = N2Str2Null("")
            CRJInvoiceno = N2Str2Null("")
            CRJInvoicetype = N2Str2Null("")
        End If
        BALANCE = NumericVal(AMOUNT2PAY) - NumericVal(totalpayment)
        
        'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'DATE UPDATED: 05-25-2009
        'DESCRIPTION: UPDATE THE AR BALANCE WHICH IS SET TO ZERO IN AMIS_JOURNAL_HD IN ORDER TO BE NOT INCLUDED IN PROCESSING
        '             ADHOC
        
        If BALANCE = 0 Then
             gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET AR_BALANCE = 0, AR_DATEGEN = " & N2Date2Null(dtprocess) & " WHERE VOUCHERNO = " & N2Str2Null(rsHeader!VOUCHERNO) & " AND JTYPE = " & N2Str2Null(rsHeader!jtype) & " AND CUSTOMERCODE = " & N2Str2Null(rsHeader!SJ_CustomerCode) & "")
        Else
             gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET AR_BALANCE = '" & BALANCE & "', AR_DATEGEN = " & N2Date2Null(dtprocess) & " WHERE VOUCHERNO = " & N2Str2Null(rsHeader!VOUCHERNO) & " AND JTYPE = " & N2Str2Null(rsHeader!jtype) & " AND CUSTOMERCODE = " & N2Str2Null(rsHeader!SJ_CustomerCode) & "")
        End If
        'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        If rsHeader!jdate <= dtprocess.Value Then
        

            gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                              "','" & CRJVoucher & "','" & CRJInvoicetype & "','" & theHDInvoice & _
                              "','" & invoicedate & "','" & rsHeader!SJ_CustomerCode & "','" & NumericVal(AMOUNT2PAY) & "','" & NumericVal(totalpayment) & _
                              "','" & NumericVal(BALANCE) & "','" & Null2String(rsHeader!ACCT_CODE) & "','" & SystemRemarks & "')")
        End If
        'initialized
        BALANCE = 0
        totalpayment = 0
        DoEvents
        lblRef.Caption = Reference
        ProgressBar1.Value = ProgressBar1.Value + 1
        Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
        rsHeader.MoveNext
    Loop                                              'SJ loop

   'UPDATED BY: JUN---------------------------------------------------------------------------------
   'DATE UPDATED: 05-20-2009
   'DESCRIPTION: DM/CM ISSUE IT CHECK THE BALANCE FOR A CERTAIN CUSTOMER CODE IF THE BALANCE IS ZERO
   '             THE IF IT IS ZERO DELETE THE CUSTOMER REPORT IN AMIS_AR
    Call CHECK_IF_BALANCE_ZERO_PER_ACCOUNT
   'UPDATED BY: JUN---------------------------------------------------------------------------------
      
   ProcessCRJNolink
   ProcessARinCDJ
   APJwithAR
   ProcessCRJNoInvoice
   If COMPANY_CODE = "HGC" Then
        ProcessCRJmaliangcodesaCRJtoSJ
        ProcessCRJWithClearingAccount
   End If
    Me.Caption = "AR AGING REPORT"
    gconDMIS.Execute ("update AMIS_AR SET LASTUPDATED='" & dtpAsOF & "'")
    MsgBox "You can now generate AR/adhoc report..", vbInformation, "Proccess completed"
    picReport.Visible = True
    
    Set rsHeader = Nothing
    Set rsdetail = Nothing

End Function
Sub CHECK_IF_BALANCE_ZERO_PER_ACCOUNT()
    'UPDATED BY: JUN
    'DATE UPDATED: 05-20-2009
    'DESCRIPTION: DM/CM ISSUE
    '             CHECK THE CUSTOMER CODE PER ACCOUNT IF THE BALANCE IS ZERO
    '             THEN IF IT IS ZERO DELETE IT FROM THE AR REPORT BECAUSE IT NOT NEEDED
    
    Dim rsZERO              As ADODB.Recordset
    Dim rsPER_CUSTOMER      As ADODB.Recordset
    Dim xCUST_KOD           As String
    
        'Set rsZERO = gconDMIS.Execute("Select CUSTOMERCODE from AMIS_AR where customercode = 'A00069'")
        Set rsZERO = gconDMIS.Execute("Select CUSTOMERCODE from AMIS_AR GROUP BY CUSTOMERCODE")
        If Not rsZERO.EOF And Not rsZERO.BOF Then
            
            Do While Not rsZERO.EOF
                xCUST_KOD = Null2String(rsZERO!CustomerCode)
                Set rsPER_CUSTOMER = gconDMIS.Execute("SELECT ROUND(SUM(BALANCE),2) AS CUSTOMER_BALANCE FROM AMIS_AR WHERE CUSTOMERCODE = '" & xCUST_KOD & "'")
                If Not rsPER_CUSTOMER.EOF And Not rsPER_CUSTOMER.BOF Then
                    If rsPER_CUSTOMER!CUSTOMER_BALANCE = 0 Then
                         'gconDMIS.Execute ("DELETE FROM AMIS_AR WHERE CUSTOMERCODE = 'A00069'")
                         gconDMIS.Execute ("DELETE FROM AMIS_AR WHERE CUSTOMERCODE = '" & xCUST_KOD & "'")
                    ElseIf rsPER_CUSTOMER!CUSTOMER_BALANCE > 0 Or rsPER_CUSTOMER!CUSTOMER_BALANCE < 0 Then
                        ' IT IS NEEDED TO BE DISPLAY  IN THE AR REPORT
                    Else
                        ' IT IS NEED TO BE DISPLAY IN  THE AR REPORT
                    End If
                End If
                rsZERO.MoveNext
            Loop
        End If
    
    Set rsZERO = Nothing
    Set rsPER_CUSTOMER = Nothing
End Sub

Sub getlastdate()
    Set rs = gconDMIS.Execute("select max(LASTUPDATED)LASTUPDATED from AMIS_AR")
    If Not (rs.EOF Or rs.BOF) Then
        If Null2String(rs!lastupdated) = "" Then
            Label9.Caption = "PLEASE GENERATE AR DATA"
            Label8.Visible = False
        Else
            Label9.Visible = True
            Label8.Visible = True
            Label8.Caption = Null2String(rs!lastupdated)
            dtpAsOF.Value = Null2String(rs!lastupdated)
            dtprocess.Value = Null2String(rs!lastupdated)
        End If
    Else
        Label8.Visible = False
        Label8.Caption = "PLEASE GENERATE AR DATA"
        Label9.Visible = False
    End If
    Set rs = Nothing
End Sub
Function DebitTotalAmount(XVoucher As String, xJtype As String, Optional ByVal xARAcount As String) As Double
    Dim rsdetail                                       As New ADODB.Recordset
    Dim rscount                                        As ADODB.Recordset

    Set rscount = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & XVoucher & _
                                   "' AND JTYPE = '" & xJtype & "' AND (LEFT(ACCT_CODE,5)='11-02' OR LEFT(ACCT_CODE,5)='11-03')")

    If rscount.Fields(0).Value > 1 Then
        Set rsdetail = gconDMIS.Execute("SELECT SUM(DEBIT) AS DEBIT,SUM(CREDIT) AS CREDIT  FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & XVoucher & _
                                        "' AND JTYPE = '" & xJtype & "' AND ACCT_CODE='" & xARAcount & "' GROUP BY ACCT_CODE")

    Else
        Set rsdetail = gconDMIS.Execute("SELECT DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & XVoucher & _
                                        "' AND JTYPE = '" & xJtype & "' AND (LEFT(ACCT_CODE,5)='11-02' OR LEFT(ACCT_CODE,5)='11-03')")
    End If


    DebitTotalAmount = 0
    If Not (rsdetail.EOF And rsdetail.BOF) Then
        rsdetail.MoveFirst
        Do While Not rsdetail.EOF
            If rsdetail!DEBIT = 0 Then
                DebitTotalAmount = DebitTotalAmount + NumericVal(rsdetail!CREDIT) * (-1)
            Else
                DebitTotalAmount = DebitTotalAmount + NumericVal(rsdetail!DEBIT)
            End If
            rsdetail.MoveNext
        Loop
    Else
    End If
    Set rsdetail = Nothing
End Function
'Function DebitTotalAmount(XVoucher As String, xJtype As String, xARAcount) As Double
'
'
'  Dim RSDetail                                       As New ADODB.Recordset
'    Set RSDetail = gconDMIS.Execute("select SUM(DEBIT) AS DEBIT,SUM(CREDIT) AS CREDIT  from AMIS_journal_det where voucherno='" & XVoucher & _
     '                                    "' and jtype = '" & xJtype & "' and ACCT_CODE='" & xARAcount & "' GROUP BY ACCT_CODE")
'    DebitTotalAmount = 0
'
'    If Not (RSDetail.EOF And RSDetail.BOF) Then
'        RSDetail.MoveFirst
'        Do While Not RSDetail.EOF
'            If RSDetail!DEBIT = 0 Then
'                DebitTotalAmount = DebitTotalAmount + NumericVal(RSDetail!CREDIT) * (-1)
'            Else
'                DebitTotalAmount = DebitTotalAmount + NumericVal(RSDetail!DEBIT)
'            End If
'            RSDetail.MoveNext
'        Loop
'    Else
'    End If
'    Set RSDetail = Nothing
'End Function

Sub ProcessLinkButDifferentCustomer()
    'Update By BTT : to find the CRJ without link
    Dim RSCRJ                                          As New ADODB.Recordset
    Dim TheAMOUNT                                      As Double
    Dim Reference                                      As String
    Dim rsCRJ_Detail                                   As New ADODB.Recordset
        
    Dim SQL As String
    SQL = " SELECT HD.VOUCHERNO AS VOUCHERNO,HD.JTYPE AS JTYPE,HD.CUSTOMERCODE AS CCODE,HD.JDATE AS JDATE,AVG(HD.INVOICEAMT) AS INVOICEAMT, SUM(DBO.AMIS_JOURNAL_DET.DEBIT) AS DEBIT,SUM(DBO.AMIS_JOURNAL_DET.CREDIT) AS CREDIT, HD.INVOICENO AS ORNUM, " & _
            " AMIS_JOURNAL_DET.ACCT_CODE AS ACCT_CODE,HD.STATUS  FROM AMIS_JOURNAL_HD HD INNER JOIN DBO.AMIS_JOURNAL_DET  ON HD.VOUCHERNO = DBO.AMIS_JOURNAL_DET.VOUCHERNO AND HD.JTYPE = DBO.AMIS_JOURNAL_DET.JTYPE" & _
            " WHERE HD.JDATE < = '" & dtprocess.Value & "'  AND HD.STATUS = 'P' AND" & _
            " (HD.JTYPE = 'CRJ') AND (LEFT(DBO.AMIS_JOURNAL_DET.ACCT_CODE, 5) = '11-02' OR LEFT(DBO.AMIS_JOURNAL_DET.ACCT_CODE, 5) = '21-07')" & _
            " GROUP BY HD.VOUCHERNO ,HD.JTYPE ,HD.CUSTOMERCODE ,HD.JDATE ,HD.INVOICENO, ACCT_CODE ,HD.STATUS"
Set RSCRJ = gconDMIS.Execute(SQL)


    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CRJDCN'")
    gconDMIS.Execute ("delete from AMIS_AR where SYSTEMREMARK='CRJDCN'")
    Picture1.Visible = False
    Picture4.Visible = True
    Me.Caption = "Loading CRJ No detail.."
    RSCRJ.MoveFirst
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCRJ.RecordCount
    
    If Not (RSCRJ.EOF And RSCRJ.BOF) Then
        Do While Not RSCRJ.EOF
            Reference = "CRJ" + "-" + Null2String(RSCRJ!VOUCHERNO)
            If RSCRJ!DEBIT = 0 Then
                TheAMOUNT = NumericVal(RSCRJ!CREDIT)
            Else
                TheAMOUNT = NumericVal(RSCRJ!DEBIT)
            End If
            
            Set rsCRJ_Detail = gconDMIS.Execute("SELECT VoucherNo FROM AMIS_CRJ_DETAIL WHERE VoucherNo ='" & Null2String(RSCRJ!VOUCHERNO) & "'")
            If rsCRJ_Detail.EOF Or rsCRJ_Detail.BOF Then
                If RSCRJ!jdate <= dtprocess.Value Then
                    
                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values(" & _
                            " NULL,'" & Reference & "'," & N2Str2Null(RSCRJ!InvoiceType) & "," & N2Str2Null(RSCRJ!INVOICENO) & "," & N2Str2Null("") & _
                                      "'," & N2Str2Null(RSCRJ!jdate) & ",'" & RSCRJ!CCODE & "','" & NumericVal(0) & "','" & NumericVal(RSCRJ!InvoiceAmt) & _
                                      "','" & ((TheAMOUNT) * (-1)) & "','" & Null2String(RSCRJ!ACCT_CODE) & "','NL')")

                End If
            End If
            DoEvents
            lblRef.Caption = RSCRJ!VOUCHERNO
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
            RSCRJ.MoveNext
        Loop
    End If
    Set RSCRJ = Nothing
    Set rsCRJ_Detail = Nothing
End Sub


Sub ProcessCRJNolink()
    'Update By BTT : to find the CRJ without link
    Dim RSCRJ                                          As New ADODB.Recordset
    Dim TheAMOUNT                                      As Double
    Dim Reference                                      As String
    Dim rsCRJ_Detail                                   As New ADODB.Recordset
    'Set RSCRJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS VOUCHERNO,dbo.AMIS_Journal_HD.JType as Jtype,dbo.AMIS_Journal_HD.CustomerCode as CCode,dbo.AMIS_Journal_HD.JDate as Jdate,dbo.AMIS_Journal_HD.InvoiceAmt as InvoiceAmt, dbo.AMIS_Journal_Det.Debit as Debit,dbo.AMIS_Journal_Det.credit as Credit,dbo.AMIS_Journal_HD.InvoiceNo as ORnum, dbo.AMIS_Journal_Det.Acct_Code as Acct_code,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE dbo.AMIS_Journal_HD.JDate < = '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND (dbo.AMIS_Journal_HD.JType = 'CRJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '21-07') order by voucherno")
    
    'Set RSCRJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS VOUCHERNO,dbo.AMIS_Journal_HD.JType as Jtype,dbo.AMIS_Journal_HD.CustomerCode as CCode,dbo.AMIS_Journal_HD.JDate as Jdate,dbo.AMIS_Journal_HD.InvoiceAmt as InvoiceAmt, dbo.AMIS_Journal_Det.Debit as Debit,dbo.AMIS_Journal_Det.credit as Credit,dbo.AMIS_Journal_HD.InvoiceNo as ORnum, dbo.AMIS_Journal_Det.Acct_Code as Acct_code,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE dbo.AMIS_Journal_HD.JDate < = '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND (dbo.AMIS_Journal_HD.JType = 'CRJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '21-07') order by voucherno")
    
    Dim SQL As String
    SQL = " SELECT HD.VOUCHERNO AS VOUCHERNO,HD.JTYPE AS JTYPE,HD.CUSTOMERCODE AS CCODE,HD.JDATE AS JDATE,AVG(HD.INVOICEAMT) AS INVOICEAMT, SUM(DBO.AMIS_JOURNAL_DET.DEBIT) AS DEBIT,SUM(DBO.AMIS_JOURNAL_DET.CREDIT) AS CREDIT, HD.INVOICENO AS ORNUM, " & _
          " AMIS_JOURNAL_DET.ACCT_CODE AS ACCT_CODE,HD.STATUS FROM AMIS_JOURNAL_HD HD INNER JOIN DBO.AMIS_JOURNAL_DET  ON HD.VOUCHERNO = DBO.AMIS_JOURNAL_DET.VOUCHERNO AND HD.JTYPE = DBO.AMIS_JOURNAL_DET.JTYPE" & _
          " WHERE HD.JDATE < = '" & dtprocess.Value & "' AND HD.STATUS = 'P' AND" & _
          " (HD.JTYPE = 'CRJ') AND (LEFT(DBO.AMIS_JOURNAL_DET.ACCT_CODE, 5) = '11-02' OR LEFT(DBO.AMIS_JOURNAL_DET.ACCT_CODE, 5) = '21-07') AND ((HD.AR_BALANCE IS NULL OR HD.AR_BALANCE <> 0) AND (HD.AR_DATEGEN IS NULL OR HD.AR_DATEGEN <= '" & dtprocess & "'))" & _
          " GROUP BY HD.VOUCHERNO ,HD.JTYPE ,HD.CUSTOMERCODE ,HD.JDATE ,HD.INVOICENO, ACCT_CODE ,HD.STATUS"
    Set RSCRJ = gconDMIS.Execute(SQL)
    
    
    
    
 

    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='NL'")
    gconDMIS.Execute ("delete from AMIS_AR where SYSTEMREMARK='NL'")
    Picture1.Visible = False
    Picture4.Visible = True
    Me.Caption = "Loading CRJ No detail.."
    RSCRJ.MoveFirst
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCRJ.RecordCount
    
    If Not (RSCRJ.EOF And RSCRJ.BOF) Then
        Do While Not RSCRJ.EOF
            'UPDATED BY: JUN
            Dim rsCHECK_BALANCE As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & Null2String(RSCRJ!VOUCHERNO) & "' AND JTYPE = '" & RSCRJ!jtype & "' ")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                    If rsCHECK_BALANCE!AR_BALANCE <> 0 Then
                        Reference = "CRJ" + "-" + Null2String(RSCRJ!VOUCHERNO)
                        If RSCRJ!DEBIT = 0 Then
                            TheAMOUNT = NumericVal(RSCRJ!CREDIT)
                        Else
                            TheAMOUNT = NumericVal(RSCRJ!DEBIT)
                        End If
                        
                        Set rsCRJ_Detail = gconDMIS.Execute("SELECT VoucherNo FROM AMIS_CRJ_DETAIL WHERE VoucherNo ='" & Null2String(RSCRJ!VOUCHERNO) & "'")
                        If rsCRJ_Detail.EOF Or rsCRJ_Detail.BOF Then
                            If RSCRJ!jdate <= dtprocess.Value Then
                            
                            
                                gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCRJ!CCODE & _
                                                  "'," & N2Str2Null(RSCRJ!VOUCHERNO) & "," & N2Str2Null(RSCRJ!InvoiceAmt) & _
                                                  "," & TheAMOUNT & "," & N2Str2Null(RSCRJ!ACCT_CODE) & "," & N2Str2Null(RSCRJ!jdate) & ",'NL')")
                
                                gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                                  "'," & N2Str2Null(RSCRJ!VOUCHERNO) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                                  "'," & N2Str2Null(RSCRJ!jdate) & ",'" & RSCRJ!CCODE & "','" & NumericVal(0) & "','" & NumericVal(RSCRJ!InvoiceAmt) & _
                                                  "','" & ((TheAMOUNT) * (-1)) & "','" & Null2String(RSCRJ!ACCT_CODE) & "','NL')")
                            End If
                        End If
                        DoEvents
                        lblRef.Caption = RSCRJ!VOUCHERNO
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                        RSCRJ.MoveNext
                    Else
                        DoEvents
                        lblRef.Caption = RSCRJ!VOUCHERNO
                        ProgressBar1.Value = ProgressBar1.Value + 1
                        Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                        RSCRJ.MoveNext
                    End If
            End If
        Loop
    End If
    Set RSCRJ = Nothing
    Set rsCRJ_Detail = Nothing
End Sub

Sub ProcessARinCDJ()
    'Update By BTT : to find the AR in the CDJ
    On Error Resume Next
    Dim amount                                         As Double
    Dim Reference                                      As String
    Dim RSCDJ                                          As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CDJ'")
    Set RSCDJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.CREDIT as CREDITAmount,dbo.AMIS_Journal_Det.Debit as DebitAmount, dbo.AMIS_Journal_HD.VoucherNo as voucherno, dbo.AMIS_Journal_HD.VendorCode as VCode,dbo.AMIS_Journal_HD.JDate as jdate,dbo.AMIS_Journal_HD.status,dbo.AMIS_Journal_HD.AR_DATEGEN FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.jtype = dbo.AMIS_Journal_Det.jtype  WHERE (dbo.AMIS_Journal_HD.JType = 'CDJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') and dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))")
    Me.Caption = "Loading CDJ Having AR.."
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCDJ.RecordCount
    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        Do While Not RSCDJ.EOF
            'UPDATED BY: JUN
            Dim rsCHECK_BALANCE As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & RSCDJ!VOUCHERNO & "' AND JTYPE = '" & RSCDJ!jtype & "'")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                If rsCHECK_BALANCE!AR_BALANCE <> 0 Then
                    DoEvents
                    Reference = "CDJ" + "-" + Null2String(RSCDJ!VOUCHERNO)
                    If RSCDJ!debitAmount = 0 Then
                        amount = NumericVal(RSCDJ!creditamount) * (-1)    ' Bawas ni
                    Else
                        amount = NumericVal(RSCDJ!debitAmount)
                    End If
                    If RSCDJ!jdate <= dtprocess.Value Then
                        gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCDJ!VCode & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & "," & N2Str2Null(amount) & _
                                          "," & N2Str2Null("") & "," & N2Str2Null(RSCDJ!ACCT_CODE) & "," & N2Str2Null(RSCDJ!jdate) & ",'CDJ')")
        
                        gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                          "'," & N2Str2Null(RSCDJ!jdate) & ",'" & RSCDJ!VCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                          "','" & NumericVal(amount) & "','" & Null2String(RSCDJ!ACCT_CODE) & "','CDJ')")
                    End If
                    DoEvents
                    lblRef.Caption = RSCDJ!VOUCHERNO
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RSCDJ.MoveNext
                Else
                    DoEvents
                    lblRef.Caption = RSCDJ!VOUCHERNO
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RSCDJ.MoveNext
                End If
            End If
        Loop
    End If
    Me.Caption = "Done"
    Set RSCDJ = Nothing
End Sub

Sub ProcessCRJNoInvoice()
    'Update By BTT : to find the AR no invoice
    Dim Reference                                      As String
    Dim amount                                         As Double
    Dim rs                                             As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='NOINV'")
    Set rs = gconDMIS.Execute("SELECT dbo.AMIS_CRJDETAIL_TOTAL.INVOICEAMOUNT, dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO, dbo.AMIS_CRJDETAIL_TOTAL.INVOICETYPE, " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.INVOICEDATE, dbo.AMIS_CRJDETAIL_TOTAL.JDate, dbo.AMIS_CRJDETAIL_TOTAL.JType, " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo, dbo.AMIS_CRJDETAIL_TOTAL.CustomerCode, dbo.AMIS_CRJDETAIL_TOTAL.Status, " & _
                              "dbo.AMIS_Journal_Det.Acct_Code  " & _
                              "FROM dbo.AMIS_CRJDETAIL_TOTAL INNER JOIN " & _
                              "dbo.AMIS_Journal_Det ON dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.jtype = dbo.AMIS_Journal_Det.jtype " & _
                              "WHERE (dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO IS NULL or dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO = 'INTRO' ) AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR " & _
                              "(LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') and  dbo.AMIS_CRJDETAIL_TOTAL.STATUS ='P' and dbo.AMIS_CRJDETAIL_TOTAL.JDate <='" & dtprocess.Value & "'")
                              
                              
'    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJDETAIL_TOTAL.INVOICEAMOUNT, dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO, dbo.AMIS_CRJDETAIL_TOTAL.INVOICETYPE, " & _
'                              "dbo.AMIS_CRJDETAIL_TOTAL.INVOICEDATE, dbo.AMIS_CRJDETAIL_TOTAL.JDate, dbo.AMIS_CRJDETAIL_TOTAL.JType, " & _
'                              "dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo, dbo.AMIS_CRJDETAIL_TOTAL.CustomerCode, dbo.AMIS_CRJDETAIL_TOTAL.Status, " & _
'                              "dbo.AMIS_Journal_Det.Acct_Code  " & _
'                              "FROM dbo.AMIS_CRJDETAIL_TOTAL INNER JOIN " & _
'                              "dbo.AMIS_Journal_Det ON dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND " & _
'                              "dbo.AMIS_CRJDETAIL_TOTAL.jtype = dbo.AMIS_Journal_Det.jtype " & _
'                              "WHERE (dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO IS NULL or dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO = 'INTRO' ) AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) in('11-02','11-03'))  " & _
'                              " and  dbo.AMIS_CRJDETAIL_TOTAL.STATUS ='P' and dbo.AMIS_CRJDETAIL_TOTAL.JDate <='" & dtprocess.Value & "'")
                              
    If Not rs.EOF And Not rs.BOF Then
        Me.Caption = "Loading CRJ with blank invoice.."
        ProgressBar1.Value = 0
        ProgressBar1.Max = rs.RecordCount
        If Not (rs.EOF And rs.BOF) Then
            Do While Not rs.EOF
                amount = NumericVal(rs!invoiceamount)
                Reference = "CRJ-" + Null2String(rs!VOUCHERNO)
                If (rs!jdate) <= dtprocess.Value Then
                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                      "'," & N2Str2Null(rs!VOUCHERNO) & ",'XXX','XXX'," & N2Str2Null(rs!jdate) & ",'" & rs!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                      "','" & NumericVal(amount) * (-1) & "','" & Null2String(rs!ACCT_CODE) & "','NOINV')")
                End If
                DoEvents
                lblRef.Caption = rs!VOUCHERNO
                ProgressBar1.Value = ProgressBar1.Value + 1
                Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                rs.MoveNext
            Loop
        End If
    End If
    Set rs = Nothing
End Sub
Sub ProcessCRJmaliangcodesaCRJtoSJ()
    Dim Reference                                      As String
    Dim test As New ADODB.Recordset
    Dim amount                                         As Double
    'Update By BTT : to find the AR na ang CRJ nya iba ang ang account sa SJ pero tama ang link
    Dim rs                                             As New ADODB.Recordset
    Me.Caption = "Validating Ad hoc data.."
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CRJWAC'")
    Set rs = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
                              "dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice, " & _
                              "dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode, dbo.AMIS_Journal_HD.Status, " & _
                              "dbo.AMIS_VW_AR_DET.Acct_Code as CR_ACCCODE, AMIS_VW_AR_DET_1.Acct_Code AS SJ_ACTCODE,AMIS_Journal_HD_1.CustomerCode AS CRJ_COUSTOMERCODE, dbo.AMIS_VW_AR_DET.Debit, dbo.AMIS_VW_AR_DET.Credit " & _
                              "FROM dbo.AMIS_CRJ_Detail INNER JOIN " & _
                              "dbo.AMIS_Journal_HD ON dbo.AMIS_CRJ_Detail.CR_TYPE = dbo.AMIS_Journal_HD.JType AND " & _
                              "dbo.AMIS_CRJ_Detail.VoucherNo = dbo.AMIS_Journal_HD.VoucherNo INNER JOIN " & _
                              "dbo.AMIS_VW_AR_DET ON dbo.AMIS_Journal_HD.JType = dbo.AMIS_VW_AR_DET.JType AND " & _
                              "dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_VW_AR_DET.VoucherNo INNER JOIN " & _
                              "dbo.AMIS_Journal_HD AS AMIS_Journal_HD_1 ON dbo.AMIS_CRJ_Detail.INVOICENO = AMIS_Journal_HD_1.InvoiceNo AND " & _
                              "dbo.AMIS_CRJ_Detail.INVOICETYPE = AMIS_Journal_HD_1.InvoiceType left outer JOIN " & _
                              "dbo.AMIS_VW_AR_DET AS AMIS_VW_AR_DET_1 ON AMIS_Journal_HD_1.JType = AMIS_VW_AR_DET_1.JType AND " & _
                              "AMIS_Journal_HD_1.VoucherNo = AMIS_VW_AR_DET_1.VoucherNo where dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.Status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))")
    
    
    
    'Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
    '                          "dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice, " & _
    '                          "dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode, dbo.AMIS_Journal_HD.Status, " & _
    '                          "dbo.AMIS_VW_AR_DET.Acct_Code as CR_ACCCODE, AMIS_VW_AR_DET_1.Acct_Code AS SJ_ACTCODE,AMIS_Journal_HD_1.CustomerCode AS CRJ_COUSTOMERCODE " & _
    '                          "FROM dbo.AMIS_CRJ_Detail INNER JOIN " & _
    '                          "dbo.AMIS_Journal_HD ON dbo.AMIS_CRJ_Detail.CR_TYPE = dbo.AMIS_Journal_HD.JType AND " & _
    '                          "dbo.AMIS_CRJ_Detail.VoucherNo = dbo.AMIS_Journal_HD.VoucherNo INNER JOIN " & _
    '                          "dbo.AMIS_VW_AR_DET ON dbo.AMIS_Journal_HD.JType = dbo.AMIS_VW_AR_DET.JType AND " & _
    '                          "dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_VW_AR_DET.VoucherNo INNER JOIN " & _
    '                          "dbo.AMIS_Journal_HD AS AMIS_Journal_HD_1 ON dbo.AMIS_CRJ_Detail.INVOICENO = AMIS_Journal_HD_1.InvoiceNo AND " & _
    '                          "dbo.AMIS_CRJ_Detail.INVOICETYPE = AMIS_Journal_HD_1.InvoiceType left outer JOIN " & _
    '                          "dbo.AMIS_VW_AR_DET AS AMIS_VW_AR_DET_1 ON AMIS_Journal_HD_1.JType = AMIS_VW_AR_DET_1.JType AND " & _
    '                          "AMIS_Journal_HD_1.VoucherNo = AMIS_VW_AR_DET_1.VoucherNo where dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "'")
    
    If Not (rs.EOF And rs.BOF) Then
     ProgressBar1.Value = 0
     ProgressBar1.Max = rs.RecordCount
        Do While Not rs.EOF
        Dim rsCheckBalance As ADODB.Recordset
        'UPDATED BY: JUN
        Set rsCheckBalance = gconDMIS.Execute("SELECT AR_BALANCE from AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & Null2String(rs!VOUCHERNO) & "' AND JTYPE = '" & Null2String(rs!jtype) & "'")
        If Not rsCheckBalance.EOF And Not rsCheckBalance.BOF Then
           
                If rsCheckBalance!AR_BALANCE <> 0 Then
                            Reference = "CRJ-" + Null2String(rs!VOUCHERNO)
                            If NumericVal(rs!DEBIT) > 0 Then
                                amount = NumericVal(rs!invoiceamount)
                            Else
                                amount = NumericVal(rs!invoiceamount) * (-1)
                            End If
                            
                            
                            Set test = gconDMIS.Execute("SELECT   AMIS_Journal_HD.VoucherNo " & _
                                       "From AMIS_Journal_HD Inner Join AMIS_Journal_Det ON AMIS_Journal_HD.JType = AMIS_Journal_Det.JType AND " & _
                                       "AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo " & _
                                       "Where (AMIS_Journal_HD.JType = 'CRJ') AND" & _
                                       "(LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') and AMIS_Journal_HD.VoucherNo = '" & rs!VOUCHERNO & "' AND ((AMIS_Journal_HD.AR_BALANCE IS NULL OR AMIS_Journal_HD.AR_BALANCE <> 0) AND (AMIS_Journal_HD.AR_DATEGEN IS NULL OR AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))" & _
                                       "GROUP BY AMIS_Journal_HD.VoucherNo " & _
                                       "HAVING COUNT(*) > 1  ")
                 
                            If (test.EOF And test.BOF) Then
                             
                            
                                If rs!jdate <= dtprocess.Value Then
                                
                
                                
                                    If (rs!CR_ACCCODE) <> Null2String(rs!SJ_ACTCODE) And rs!CustomerCode = rs!CRJ_COUSTOMERCODE Then
                                  
                                        gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                                      "'," & N2Str2Null(rs!VOUCHERNO) & ",'" & rs!InvoiceType & "','" & rs!INVOICENO & "'," & N2Str2Null(rs!jdate) & ",'" & rs!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                                      "','" & NumericVal(amount) & "','" & Null2String(rs!CR_ACCCODE) & "','CRJWAC')")
                                    End If
                                End If
                            End If
                            'lblRef.Caption = RSCDJ!voucherno
                             DoEvents
                            ProgressBar1.Value = ProgressBar1.Value + 1
                            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                            rs.MoveNext
                Else
                     DoEvents
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                     rs.MoveNext
                End If
        End If
        Loop
    End If
    Set rs = Nothing
    Set rsCheckBalance = Nothing
    'Temp
    gconDMIS.Execute ("update AMIS_AR set balance = 0 where CRJvoucherno = '004335' and invoicetype = 'VI' and customercode = 'P00172'")
End Sub
Sub ProcessCRJWithClearingAccount()
    Dim test As New ADODB.Recordset
    Dim Reference                                      As String
    Dim amount                                         As Double
    ' CRJWCA : CRJ - with clrearing account and tama ang link
    Me.Caption = "Validating Ad hoc Data.."
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CRJWCA'")
    Dim rs                                             As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
                              "dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice, " & _
                              "dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode, dbo.AMIS_Journal_HD.Status,dbo.AMIS_Journal_HD.AR_DATEGEN, " & _
                              "AMIS_Journal_HD_1.CustomerCode AS CRJ_COUSTOMERCODE, dbo.AMIS_VW_AR_DET.Acct_Code AS CRJ_ACCTCODE, " & _
                              "AMIS_VW_AR_DET_1.Acct_Code AS SJ_ACCTCODE " & _
                              "FROM dbo.AMIS_CRJ_Detail INNER JOIN " & _
                              "dbo.AMIS_Journal_HD ON dbo.AMIS_CRJ_Detail.CR_TYPE = dbo.AMIS_Journal_HD.JType AND " & _
                              "dbo.AMIS_CRJ_Detail.VoucherNo = dbo.AMIS_Journal_HD.VoucherNo INNER JOIN " & _
                              "dbo.AMIS_Journal_HD AS AMIS_Journal_HD_1 ON dbo.AMIS_CRJ_Detail.INVOICENO = AMIS_Journal_HD_1.InvoiceNo AND " & _
                              "dbo.AMIS_CRJ_Detail.INVOICETYPE = AMIS_Journal_HD_1.InvoiceType INNER JOIN " & _
                              "dbo.AMIS_VW_AR_DET AS AMIS_VW_AR_DET_1 ON AMIS_Journal_HD_1.JType = AMIS_VW_AR_DET_1.JType AND " & _
                              "AMIS_Journal_HD_1.VoucherNo = AMIS_VW_AR_DET_1.VoucherNo LEFT OUTER JOIN " & _
                              "dbo.AMIS_VW_AR_DET ON dbo.AMIS_Journal_HD.JType = dbo.AMIS_VW_AR_DET.JType AND " & _
                              "dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_VW_AR_DET.VoucherNo " & _
                              "WHERE (dbo.AMIS_Journal_HD.JDate) <= '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.Status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))")
    If Not rs.EOF And Not rs.BOF Then
     ProgressBar1.Value = 0
     ProgressBar1.Max = rs.RecordCount
        Do While Not rs.EOF
        'UPDATED BY: JUN
        Dim rsCHECK_AR_BALANCE As ADODB.Recordset
        Set rsCHECK_AR_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & rs!VOUCHERNO & "' AND JTYPE = '" & rs!jtype & "'")
            If Not rsCHECK_AR_BALANCE.EOF And Not rsCHECK_AR_BALANCE.BOF Then
                If rsCHECK_AR_BALANCE!AR_BALANCE <> 0 Then
                   Reference = "CRJ-" + Null2String(rs!VOUCHERNO)
                   amount = NumericVal(rs!invoiceamount)
                   
                   
                   Set test = gconDMIS.Execute(" SELECT   AMIS_Journal_HD.VoucherNo " & _
                   "From AMIS_Journal_HD Inner Join AMIS_Journal_Det ON AMIS_Journal_HD.JType = AMIS_Journal_Det.JType AND " & _
                   "AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo " & _
                   "Where (AMIS_Journal_HD.JType = 'CRJ')AND AMIS_Journal_HD.AR_BALANCE <> 0 AND " & _
                   "(LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') and AMIS_Journal_HD.VoucherNo = '" & rs!VOUCHERNO & "'" & _
                   "GROUP BY AMIS_Journal_HD.VoucherNo " & _
                   "HAVING COUNT(*) > 1  ")
                   
                  
                   If Not (test.EOF And test.BOF) Then
        
                   
                   Else
                       If rs!jdate <= dtprocess.Value Then
                       If Null2String(rs!CRJ_ACCTCODE) <> Null2String(rs!SJ_ACCTCODE) And rs!CustomerCode = rs!CRJ_COUSTOMERCODE Then
                           gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                             "'," & N2Str2Null(rs!VOUCHERNO) & ",'" & rs!InvoiceType & "','" & rs!INVOICENO & "'," & N2Str2Null(rs!jdate) & ",'" & rs!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                             "','" & NumericVal(amount) * (1) & "','" & Null2String(rs!SJ_ACCTCODE) & "','CRJWCA')")
                       End If
                   End If
                   End If
                   DoEvents
                   ProgressBar1.Value = ProgressBar1.Value + 1
                   Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                   rs.MoveNext
                Else
                   DoEvents
                   ProgressBar1.Value = ProgressBar1.Value + 1
                   Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                   rs.MoveNext
                End If
            End If
        Loop
    End If
    Set rs = Nothing
End Sub
Function ReturnCustomerCode(TheVoucherno As String, theJtype As String) As String
    Dim rs As New ADODB.Recordset
    Dim GetDealer As New ADODB.Recordset
    Set rs = gconDMIS.Execute("Select customercode from AMIS_journal_hd where voucherno='" & TheVoucherno & "' and jtype = '" & theJtype & "'")
    If Not rs.EOF And Not rs.BOF Then
        ReturnCustomerCode = Null2String(rs!CustomerCode)
        'Set GetDealer = gconDMIS.Execute("Select Acctname from all_customer_table where cuscde='" & ReturnCustomerCode & "'")
        'If Not GetDealer.EOF And Not GetDealer.BOF Then
        '    Dealer = Null2String(GetDealer!Acctname)
        'End If
        Else
        ReturnCustomerCode = N2Str2Null("")
    End If
    Set rs = Nothing
End Function
Sub Transfer_SalesJournal()
    Dim Bernard As New ADODB.Recordset
    Dim CounterCheck_F As New ADODB.Recordset
    Dim xVoucherNo As String
    Dim xJdate As String
    Dim xStatus As String
    Dim xJtype As String
    Dim xSj_customercode As String
    Dim xInvoicedate As String
    Dim xinvoiceType As String
    Dim xInvoiceAmnt As String
    Dim xamounttopay As Double
    Dim xacct_code As String
    Dim xamountPaid As Double
    Dim xInvoiceNo As String
    Dim xDebit As Double
    Dim THECSJ As String
    THECSJ = "CSJ"
    gconDMIS.Execute ("Delete from AMIS_AR_HD")
    Set Bernard = gconDMIS.Execute("SELECT DISTINCT AMIS_Journal_HD.VoucherNo,AMIS_Journal_HD.jdate, AMIS_Journal_HD.Status, AMIS_Journal_HD.JType,AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, AMIS_Journal_HD.InvoiceType, AMIS_Journal_HD.InvoiceNo,AMIS_Journal_HD.InvoiceDate as XInvoiceDate , AMIS_Journal_HD.InvoiceAmt, AMIS_Journal_HD.AmountToPay, AMIS_Journal_HD.AmountPaid,AMIS_Journal_Det.Acct_Code AS acct_code, AMIS_Journal_Det.Acct_Name, AMIS_Journal_Det.Debit AS Detdebit FROM AMIS_Journal_HD LEFT OUTER JOIN AMIS_Journal_Det ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND AMIS_Journal_HD.JType = AMIS_Journal_Det.JType WHERE (LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (AMIS_Journal_HD.JType = 'SJ' OR AMIS_Journal_HD.JType = 'COB' OR AMIS_Journal_HD.JType = 'CCM' OR AMIS_Journal_HD.JType = '" & THECSJ & _
                                    "') and AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and AMIS_Journal_HD.status ='P' AND (AMIS_Journal_HD.AR_BALANCE <> 0 OR AMIS_Journal_HD.AR_BALANCE IS NULL)  ORDER BY AMIS_Journal_HD.VoucherNo")
 

    
    'Set Bernard = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
    '                                "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' and dbo.AMIS_Journal_HD.CustomerCode = 'A00069' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")
    Me.Caption = "Validating Transaction.."
    If Not Bernard.EOF And Not Bernard.BOF Then
     ProgressBar1.Value = 0
     ProgressBar1.Max = Bernard.RecordCount
     Do While Not Bernard.EOF
            xVoucherNo = N2Str2Null(Bernard!VOUCHERNO)
            xJdate = N2Date2Null(Bernard!jdate)
            xStatus = N2Str2Null(Bernard!Status)
            xJtype = N2Str2Null(Bernard!jtype)
            xSj_customercode = N2Str2Null(Bernard!SJ_CustomerCode)
            xinvoiceType = N2Str2Null(Bernard!InvoiceType)
            xInvoiceNo = N2Str2Null(Bernard!INVOICENO)
            xInvoiceAmnt = N2Str2Null(Bernard!InvoiceAmt)
            xInvoicedate = N2Str2Null(Bernard!xInvoicedate)
            xamounttopay = N2Str2Zero(Bernard!amounttopay)
            xamountPaid = N2Str2Zero(Bernard!AmountPaid)
            xacct_code = N2Str2Null(Bernard!ACCT_CODE)
            xDebit = N2Str2Zero(Bernard!detdebit)
                     
                     Set CounterCheck_F = gconDMIS.Execute("Select count(*) from AMIS_AR_HD where voucherno = " & xVoucherNo & " and jtype =" & xJtype & "")
                     If Not CounterCheck_F(0) = 1 Then
                        If Bernard!jdate <= dtprocess.Value Then
                        gconDMIS.Execute ("Insert into AMIS_AR_HD(Voucherno,jdate,status,jtype,Sj_customercode,invoicetype,invoiceno,invoicedate,invoiceamnt,amounttopay,amountpaid,acct_code,debit) values(" & xVoucherNo & _
                                      "," & xJdate & "," & xStatus & "," & xJtype & _
                                      "," & xSj_customercode & "," & xinvoiceType & "," & xInvoiceNo & "," & xInvoicedate & "," & xInvoiceAmnt & _
                                      "," & xamounttopay & "," & xamountPaid & "," & xacct_code & "," & xDebit & ")")
                                      
                        
                     
                     End If
                     End If
                     DoEvents
                     lblRef.Caption = xVoucherNo
                     ProgressBar1.Value = ProgressBar1.Value + 1
                     Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                     Bernard.MoveNext
                 Loop
    End If
    Set Bernard = Nothing
End Sub
Sub ValidateDetail()
    Dim RSXX As New ADODB.Recordset
     
    Set RSXX = gconDMIS.Execute("SELECT AMIS_CRJ_Detail.CR_TYPE, AMIS_CRJ_Detail.VoucherNo,AMIS_CRJ_Detail.invoiceno,AMIS_CRJ_Detail.invoicetype  " & _
                                "FROM AMIS_CRJ_Detail INNER JOIN " & _
                                "AMIS_Journal_HD ON AMIS_CRJ_Detail.CR_TYPE = AMIS_Journal_HD.JType AND " & _
                                "AMIS_CRJ_Detail.VoucherNo = AMIS_Journal_HD.VoucherNo INNER JOIN " & _
                                "AMIS_Journal_Det ON AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo AND " & _
                                "AMIS_Journal_HD.jtype = AMIS_Journal_Det.jtype " & _
                                "WHERE dbo.AMIS_Journal_Det.Acct_Code = '11-01019-00'")
                                
   If Not (RSXX.EOF And RSXX.BOF) Then
        Do While Not RSXX.EOF
            gconDMIS.Execute ("UPDATE AMIS_CRJ_DETAIL SET STATUS ='X' where voucherno=" & RSXX!VOUCHERNO & " and invoiceno='" & RSXX!INVOICENO & "' and invoicetype='" & RSXX!InvoiceType & "'")
            RSXX.MoveNext
        Loop
   End If
Set RSXX = Nothing
End Sub
Sub APJwithAR()
    'Update By BTT : to find the AR in the CDJ
    On Error Resume Next
    Dim amount                                         As Double
    Dim Reference                                      As String
    Dim RSCDJ                                          As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='APJ'")
    
    Set RSCDJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.CREDIT as CREDITAmount,dbo.AMIS_Journal_Det.Debit as DebitAmount, dbo.AMIS_Journal_HD.VoucherNo as voucherno, dbo.AMIS_Journal_HD.VendorCode as VCode,dbo.AMIS_Journal_HD.JDate as jdate,dbo.AMIS_Journal_HD.status,dbo.AMIS_Journal_HD.AR_DATEGEN FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.jtype = dbo.AMIS_Journal_Det.jtype  WHERE (dbo.AMIS_Journal_HD.JType = 'APJ') AND dbo.AMIS_Journal_Det.Acct_Code = '11-02000-00' and dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "')) ")
    
    Me.Caption = "Loading AP Having AR.."
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCDJ.RecordCount
    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        Do While Not RSCDJ.EOF
            'UPDATED BY : JUN
            Dim rsCHECK_BALANCE As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & RSCDJ!VOUCHERNO & "' AND JTYPE = '" & RSCDJ!jtype & "'")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                If RSCDJ!AR_BALANCE <> 0 Then
                    DoEvents
                    Reference = "CDJ" + "-" + Null2String(RSCDJ!VOUCHERNO)
                    If RSCDJ!debitAmount = 0 Then
                        amount = NumericVal(RSCDJ!creditamount) * (-1)    ' Bawas ni
                    Else
                        amount = NumericVal(RSCDJ!debitAmount)
                    End If
                    If RSCDJ!jdate <= dtprocess.Value Then
                        gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCDJ!VCode & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & "," & N2Str2Null(amount) & _
                                          "," & N2Str2Null("") & "," & N2Str2Null(RSCDJ!ACCT_CODE) & "," & N2Str2Null(RSCDJ!jdate) & ",'APJ')")
        
                        'gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                        '                  "'," & N2Str2Null(RSCDJ!voucherno) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                        '                  "'," & N2Str2Null(RSCDJ!jdate) & ",'" & RSCDJ!VCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                        '                  "','" & NumericVal(amount) & "','" & Null2String(RSCDJ!ACCT_CODE) & "','CDJ')")
                        End If
                    DoEvents
                    lblRef.Caption = RSCDJ!VOUCHERNO
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RSCDJ.MoveNext
                Else
                    DoEvents
                    lblRef.Caption = RSCDJ!VOUCHERNO
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RSCDJ.MoveNext
                End If
            End If
        Loop
    End If
    Me.Caption = "Done"
    Set RSCDJ = Nothing
End Sub

Private Sub picReport_Click()
    ProcessCRJmaliangcodesaCRJtoSJ
    ProcessCRJWithClearingAccount
End Sub


