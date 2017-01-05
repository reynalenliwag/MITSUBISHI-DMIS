VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form ARSchedReport_Old 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GENERATE AR REPORT"
   ClientHeight    =   3210
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   5955
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
   Icon            =   "ARSchedReport_Old.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   915
      Left            =   4440
      TabIndex        =   50
      Top             =   1080
      Width           =   855
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
      Left            =   6960
      TabIndex        =   18
      Top             =   1620
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
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1710
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
      Left            =   7290
      TabIndex        =   7
      Top             =   2340
      Width           =   4965
      Begin VB.CommandButton cmdCheck 
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   4035
         MouseIcon       =   "ARSchedReport_Old.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport_Old.frx":0F94
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
         Picture         =   "ARSchedReport_Old.frx":122F
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "ARSchedReport_Old.frx":124B
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
         Format          =   51576833
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
      Left            =   7200
      TabIndex        =   1
      Top             =   60
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
      Left            =   7560
      TabIndex        =   0
      Top             =   450
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3585
   End
   Begin Crystal.CrystalReport rptAMISDueReport 
      Left            =   3870
      Top             =   4560
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
      Left            =   7200
      TabIndex        =   2
      Top             =   660
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
         Format          =   51576833
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
         Format          =   51576833
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
      Left            =   7200
      TabIndex        =   14
      Top             =   660
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   30
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   28
      Top             =   60
      Width           =   4125
      Begin VB.CommandButton Command3 
         Caption         =   "EXIT"
         Height          =   435
         Left            =   240
         TabIndex        =   29
         Top             =   2070
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "AR ADHOC REPORT"
         Height          =   435
         Left            =   240
         TabIndex        =   37
         Top             =   1650
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "AR AGING REPORT"
         Height          =   435
         Left            =   240
         TabIndex        =   30
         Top             =   1230
         Width           =   3495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "AR SCHEDULE REPORT"
         Height          =   405
         Left            =   240
         MaskColor       =   &H0080FFFF&
         TabIndex        =   31
         Top             =   840
         Width           =   3495
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF80&
         Caption         =   "Process AR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2250
         MaskColor       =   &H0080FFFF&
         TabIndex        =   44
         Top             =   270
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker dtprocess 
         Height          =   375
         Left            =   300
         TabIndex        =   45
         Top             =   270
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   51576833
         CurrentDate     =   38216
      End
   End
   Begin MSComCtl2.DTPicker dtpAsOF 
      Height          =   405
      Left            =   6450
      TabIndex        =   49
      Top             =   1080
      Width           =   2475
      _ExtentX        =   4366
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
      Format          =   51576833
      CurrentDate     =   38216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   30
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   22
      Top             =   60
      Width           =   4125
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   -240
         ScaleHeight     =   465
         ScaleWidth      =   285
         TabIndex        =   36
         Top             =   0
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   2565
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
         MouseIcon       =   "ARSchedReport_Old.frx":1267
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport_Old.frx":13B9
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
         MouseIcon       =   "ARSchedReport_Old.frx":1804
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport_Old.frx":1956
         Style           =   1  'Graphical
         TabIndex        =   24
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture4 
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
      Height          =   3105
      Left            =   30
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   4125
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
         TabIndex        =   33
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
         TabIndex        =   32
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
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   30
      ScaleHeight     =   3075
      ScaleWidth      =   4095
      TabIndex        =   38
      Top             =   60
      Width           =   4125
      Begin VB.OptionButton opt_CustomerSJCRJ 
         Caption         =   "Different Customer Information In SJ and CRJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   48
         Top             =   1620
         Width           =   3825
      End
      Begin VB.OptionButton opt_SJWithTwoAccounts 
         Caption         =   "Sales Journal with Two or More AR Accounts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   4155
      End
      Begin VB.OptionButton Option6 
         Caption         =   "CRJ with Defferent Account Set Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1035
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
         Left            =   1680
         MouseIcon       =   "ARSchedReport_Old.frx":1DF5
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport_Old.frx":1F47
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Close Window"
         Top             =   2100
         Width           =   885
      End
      Begin VB.OptionButton Option3 
         Caption         =   "CRJ No Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   210
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.OptionButton Option4 
         Caption         =   "CRJ Blank Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Vendor Having AR Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   765
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
         Left            =   810
         MouseIcon       =   "ARSchedReport_Old.frx":2392
         MousePointer    =   99  'Custom
         Picture         =   "ARSchedReport_Old.frx":24E4
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Print Report"
         Top             =   2100
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
      Left            =   7830
      TabIndex        =   17
      Top             =   780
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
      Left            =   7200
      TabIndex        =   15
      Top             =   660
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
      Left            =   7200
      TabIndex        =   13
      Top             =   660
      Width           =   4935
   End
End
Attribute VB_Name = "ARSchedReport_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report_type                                        As String
Dim RS                                                 As New ADODB.Recordset
Dim CountDetail                                        As Integer
Function SetCRJVoucherNo(XXX As String, zzz As Integer) As String
    Dim rsCRJ_Journal_HD                               As ADODB.Recordset
    Set rsCRJ_Journal_HD = New ADODB.Recordset
    If zzz = 1 Then
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and InvoiceNo = '" & XXX & "'")
    Else
        Set rsCRJ_Journal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD where Jtype = 'CRJ' and LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = '" & XXX & "'")
    End If
    If Not rsCRJ_Journal_HD.EOF And Not rsCRJ_Journal_HD.BOF Then
        SetCRJVoucherNo = Null2String(rsCRJ_Journal_HD!voucherno)
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
    picReport.Visible = True
    picNolink.ZOrder 0
End Sub

Private Sub Command5_Click()
    picNolink.Visible = False
End Sub

Private Sub Command6_Click()
    rptAMISDueReport.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAMISDueReport.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    If opt_CustomerSJCRJ.Value = True Then
        DifferentCustomerSJCRJ
        Exit Sub
    End If
    If opt_SJWithTwoAccounts.Value = True Then
        TwoOrMoreARAccounts
        Exit Sub
    End If
    If Option3.Value = True Then
        '
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

End Sub
Sub DifferentCustomerSJCRJ()

    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim i                                              As Long
    Screen.MousePointer = 11
    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If
    On Error Resume Next
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
     wsXL.Name = "DifferentCustomers"
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT  SJVOUCHERNO,CRJVOUCHERNO FROM AMIS_AR WHERE SYSTEMREMARK='WRONG CUSTOMER CODE'")
    i = 1
    wsXL.Cells(1, 1) = "SJ VOUCHER#"
    wsXL.Cells(1, 2) = "SJ CUSTOMER CODE"
    wsXL.Cells(1, 3) = "SJ CUSTOMER NAME"

    wsXL.Cells(1, 4) = "CRJ VOUCHER#"
    wsXL.Cells(1, 5) = "CRJ CUSTOMER CODE"
    wsXL.Cells(1, 6) = "CRJ CUSTOMER NAME"

    Dim SJ_CUSTNAME                                    As String
    Dim SJ_CUSTCODE                                    As String

    Dim CRJ_CUSTNAME                                   As String
    Dim CRJ_CUSTCODE                                   As String

    Dim RSCUST                                         As ADODB.Recordset

    If Not RS.EOF Or Not RS.BOF Then
        While Not RS.EOF
            SJ_CUSTCODE = "": SJ_CUSTNAME = ""
            CRJ_CUSTCODE = "": CRJ_CUSTNAME = ""


            If Left(RS!SJVoucherno, 2) = "SJ" Then
                Set RSCUST = gconDMIS.Execute("SELECT CUSTOMERCODE ,ACCTNAME CUSTOMERNAME  FROM AMIS_JOURNAL_HD  INNER JOIN ALL_CUSTOMER_TABLE ON CUSCDE=CUSTOMERCODE  WHERE JTYPE='SJ' AND VOUCHERNO=" & N2Str2Null(Replace(RS!SJVoucherno, "SJ-", "")))
                If Not RSCUST.EOF Or Not RSCUST.BOF Then
                    SJ_CUSTCODE = Null2String(RSCUST!CustomerCode)
                    SJ_CUSTNAME = Null2String(RSCUST!CustomerName)
                End If
            Else
                Set RSCUST = gconDMIS.Execute("SELECT CUSTOMERCODE ,ACCTNAME CUSTOMERNAME  FROM AMIS_JOURNAL_HD  INNER JOIN ALL_CUSTOMER_TABLE ON CUSCDE=CUSTOMERCODE WHERE JTYPE='COB' AND VOUCHERNO=" & N2Str2Null(Replace(RS!SJVoucherno, "COB-", "")))
                If Not RSCUST.EOF Or Not RSCUST.BOF Then
                    SJ_CUSTCODE = Null2String(RSCUST!CustomerCode)
                    SJ_CUSTNAME = Null2String(RSCUST!CustomerName)
                End If
            End If


            Set RSCUST = gconDMIS.Execute("SELECT CUSTOMERCODE ,ACCTNAME CUSTOMERNAME  FROM AMIS_JOURNAL_HD  INNER JOIN ALL_CUSTOMER_TABLE ON CUSCDE=CUSTOMERCODE WHERE JTYPE='CRJ' AND VOUCHERNO=" & N2Str2Null(RS!CRJVoucherno))
            If Not RSCUST.EOF Or Not RSCUST.BOF Then
                CRJ_CUSTCODE = Null2String(RSCUST!CustomerCode)
                CRJ_CUSTNAME = Null2String(RSCUST!CustomerName)
            End If

            If SJ_CUSTCODE <> CRJ_CUSTCODE Then
                i = i + 1
                wsXL.Cells(i, 1) = RS!SJVoucherno
                wsXL.Cells(i, 2) = SJ_CUSTCODE
                wsXL.Cells(i, 3) = SJ_CUSTNAME
                wsXL.Cells(i, 4) = RS!CRJVoucherno
                wsXL.Cells(i, 5) = CRJ_CUSTCODE
                wsXL.Cells(i, 6) = CRJ_CUSTNAME
            End If

            RS.MoveNext
            DoEvents
        Wend

        wsXL.Columns(1).AutoFit
        wsXL.Columns(2).AutoFit
        wsXL.Columns(3).AutoFit
        wsXL.Columns(4).AutoFit
        wsXL.Columns(5).AutoFit
        wsXL.Columns(6).AutoFit

    End If
    objXL.Visible = True
    Set wbXL = Nothing
    Set wsXL = Nothing
    Screen.MousePointer = 0

End Sub

Sub TwoOrMoreARAccounts()

    Dim objXL                                          As New Excel.Application
    Dim wbXL                                           As New Excel.Workbook
    Dim wsXL                                           As New Excel.Worksheet
    Dim i                                              As Long
    Screen.MousePointer = 11
    If Not IsObject(objXL) Then
        MsgBox "You need Microsoft Excel to use this function", _
               vbExclamation, "Print to Excel"
        Exit Sub
    End If
    On Error Resume Next
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
    '    wsXL.Name = "Different Customer In SJ and CRJ"
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT VOUCHERNO ,ACCT_CODE,ACCT_NAME,SUM(DEBIT) as DEBIT, SUM(CREDIT) CREDIT FROM AMIS_JOURNAL_DET WHERE JTYPE='SJ' AND LEFT(ACCT_CODE,5)  IN('11-02','11-03') GROUP BY VOUCHERNO ,ACCT_CODE,ACCT_NAME HAVING COUNT(ACCT_CODE)>1")
    i = 1
    wsXL.Cells(1, 1) = "SJ VOUCHER#"
    wsXL.Cells(1, 2) = "ACCOUNT CODE"
    wsXL.Cells(1, 3) = "ACCOUNT NAME"
    wsXL.Cells(1, 4) = "DEBIT"
    wsXL.Cells(1, 5) = "CREDIT"
   
    Dim SJ_CUSTNAME                                    As String
    Dim SJ_CUSTCODE                                    As String
    Dim CRJ_CUSTNAME                                   As String
    Dim CRJ_CUSTCODE                                   As String

    Dim RSCUST                                         As ADODB.Recordset

    If Not RS.EOF Or Not RS.BOF Then
        While Not RS.EOF
                   i = i + 1
                wsXL.Cells(i, 1) = Null2String(RS!voucherno)
                wsXL.Cells(i, 2) = Null2String(RS!acct_code)
                wsXL.Cells(i, 3) = Null2String(RS!acct_Name)
                wsXL.Cells(i, 4) = N2Str2Zero(RS!DEBIT)
                wsXL.Cells(i, 5) = N2Str2Zero(RS!CREDIT)
                
           

            RS.MoveNext
            DoEvents
        Wend

        wsXL.Columns(1).AutoFit
        wsXL.Columns(2).AutoFit
        wsXL.Columns(3).AutoFit
        wsXL.Columns(4).AutoFit
        wsXL.Columns(5).AutoFit
        wsXL.Columns(6).AutoFit

    End If
    objXL.Visible = True
    Set wbXL = Nothing
    Set wsXL = Nothing
    Screen.MousePointer = 0

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
    'ProcessCRJNolink
    'ProcessARinCDJ
    'ProcessCRJNoInvoice
    'ProcessCRJWithClearingAccount
    ProcessCRJmaliangcodesaCRJtoSJ
    
    
    
    MsgBox "tapos"
    
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
    'DtCrjDate = LOGDATE
    dtprocess = LOGDATE

    'gconDMIS.Execute (" IF NOT EXISTS (SELECT * FROM SYSCOLUMNS WHERE ID=OBJECT_ID('AMIS_AR') AND NAME='LASTUPDATED')" & _
     '            " ALTER TABLE AMIS_AR  ADD LASTUPDATED SMALLDATETIME NULL")
    Dim RS                                             As ADODB.Recordset
    getlastdate

    GetAcctcode
    picNolink.Visible = False
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    Screen.MousePointer = 0

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


    rptAMISDueReport.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAMISDueReport.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If Report_type = "SCHED" Then
        If IsDate(Label8.Caption) = True Then
            If dtpAsOF.Value > CDate(Label8.Caption) Then
                MsgBox "Information:Date is greater than Last data generated"
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
                LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & dtpAsOF
            Else
                '
            End If
        ElseIf Option2.Value = True Then
            rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
            rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
            PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1
            LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & dtpAsOF
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
                    PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "", DMIS_REPORT_Connection, 1
                    LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
                    'Else
                    'MsgBox "No AR as of the " & dtpAsOF
                    'End If
                Else
                    If Option1.Value = True Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "", DMIS_REPORT_Connection, 1
                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
                    ElseIf Option1.Value = True Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & Label8.Caption
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "", DMIS_REPORT_Connection, 1
                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF

                    End If
                End If
            ElseIf Option2.Value = True Then          ' Group report by Account
                rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "", DMIS_REPORT_Connection, 1
                LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
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
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT DESCRIPTION from AMIS_chartaccount where left(acctcode,5)='11-02' or left(acctcode,5)='11-03' ORDER BY DESCRIPTION"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cboacctcode.Clear

    Do While Not RS.EOF
        cboacctcode.AddItem (RS!Description)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub
Function ReturnAccountCode(Xacct_desc As String)
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "select description,acctcode from AMIS_chartaccount where description = '" & Xacct_desc & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtdescription = Null2String(RS!AcctCode)
    End If
    Set RS = Nothing
End Function
Function ar() As Boolean
    ' Update by BTT : this process kill the AR
    Dim rsHeader                                       As New ADODB.Recordset
    Dim RSDetail                                       As New ADODB.Recordset
    Dim BALANCE                                        As Double
    Dim totalpayment                                   As Double
    Dim CRJVoucher                                     As String
    Dim Reference                                      As String
    Dim SystemRemarks                                  As String
    Dim CRJInvoiceno                                   As String
    Dim CRJInvoicetype                                 As String
    Dim AMOUNT2PAY                                     As Double
    Dim Invoicedate                                    As String
    Dim THECSJ                                         As String
    Dim theHDInvoice                                   As String
    THECSJ = "CSJ"
    Me.Caption = "Loading SJ/COB transaction.."
    gconDMIS.Execute ("DELETE FROM AMIS_AR")
    Dim ARNIE                                          As New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
                                    "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")

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
        Reference = Null2String(rsHeader!jtype) + "-" + Null2String(rsHeader!voucherno)
        Invoicedate = Null2String(rsHeader!xInvoicedate)
        theHDInvoice = Null2String(rsHeader!InvoiceNo)
        If (rsHeader!jtype) = "SJ" Then
            AMOUNT2PAY = DebitTotalAmount(Null2String(rsHeader!voucherno), Null2String(rsHeader!jtype))
        ElseIf (rsHeader!jtype) = "CCM" Then                        ' Credit memo
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt) * (-1)    'to deduct to total payment
        Else
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt)
        End If
        Set RSDetail = gconDMIS.Execute("Select invoiceno,invoicetype,ISNULL(invoiceamount,0) AS INVOICEAMOUNT,voucherno,jdate,customercode as CRJ_customercode from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!InvoiceNo & "' and status = 'P' and jdate <= " & N2Str2Null(dtpAsOF) & "")
        If Not RSDetail.EOF And Not RSDetail.BOF Then
            RSDetail.MoveFirst
            Do While Not RSDetail.EOF
                If (Null2String(rsHeader!SJ_CustomerCode) = Null2String(RSDetail!CRJ_customercode)) Then
                    CRJVoucher = Null2String(RSDetail!voucherno)
                    CRJInvoiceno = Null2String(RSDetail!InvoiceNo)
                    CRJInvoicetype = Null2String(RSDetail!InvoiceType)
                    totalpayment = totalpayment + RSDetail!invoiceAmount
                    SystemRemarks = "NULL"
                Else
                    CRJVoucher = Null2String(RSDetail!voucherno)
                    SystemRemarks = "Wrong customer code"
                End If
                RSDetail.MoveNext
            Loop                                      'CRJ loop
        Else
            ' if no payment
            CRJVoucher = N2Str2Null("")
            CRJInvoiceno = N2Str2Null("")
            CRJInvoicetype = N2Str2Null("")
        End If
        BALANCE = NumericVal(AMOUNT2PAY) - NumericVal(totalpayment)
        If rsHeader!jdate <= dtprocess.Value Then
            gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                              "','" & CRJVoucher & "','" & CRJInvoicetype & "','" & theHDInvoice & _
                              "','" & Invoicedate & "','" & rsHeader!SJ_CustomerCode & "','" & NumericVal(AMOUNT2PAY) & "','" & NumericVal(totalpayment) & _
                              "','" & NumericVal(BALANCE) & "','" & Null2String(rsHeader!acct_code) & "','" & SystemRemarks & "')")
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

    'ProcessCRJNolink
    'ProcessARinCDJ
    'ProcessCRJNoInvoice
    'ProcessCRJmaliangcodesaCRJtoSJ
    'ProcessCRJWithClearingAccount
    Me.Caption = "AR AGING REPORT"
    gconDMIS.Execute ("update AMIS_AR SET LASTUPDATED='" & dtpAsOF & "'")
    MsgBox "You can now generate AR/adhoc report..", vbInformation, "Proccess completed"
    picReport.Visible = True
    Set rsHeader = Nothing
    Set RSDetail = Nothing

End Function
Sub getlastdate()
    Set RS = gconDMIS.Execute("select max(LASTUPDATED)LASTUPDATED from AMIS_AR")
    If Not (RS.EOF Or RS.BOF) Then
        If Null2String(RS!LASTUPDATED) = "" Then
            Label9.Caption = "PLEASE GENERATE AR DATA"
            Label8.Visible = False
        Else
            Label9.Visible = True
            Label8.Visible = True
            Label8.Caption = Null2String(RS!LASTUPDATED)
        End If
    Else
        Label8.Visible = False
        Label8.Caption = "PLEASE GENERATE AR DATA"
        Label9.Visible = False
    End If
    Set RS = Nothing
End Sub
Function DebitTotalAmount(XVoucher As String, xJtype As String) As Double
    Dim RSDetail                                       As New ADODB.Recordset
    Set RSDetail = gconDMIS.Execute("select debit,credit from AMIS_journal_det where voucherno='" & XVoucher & _
                                    "' and jtype = '" & xJtype & "' and (left(acct_code,5)='11-02' or left(acct_code,5)='11-03')")
    DebitTotalAmount = 0
    If Not (RSDetail.EOF And RSDetail.BOF) Then
        RSDetail.MoveFirst
        Do While Not RSDetail.EOF
            If RSDetail!DEBIT = 0 Then
                DebitTotalAmount = DebitTotalAmount + NumericVal(RSDetail!CREDIT) * (-1)
            Else
                DebitTotalAmount = DebitTotalAmount + NumericVal(RSDetail!DEBIT)
            End If
            RSDetail.MoveNext
        Loop
    Else
    End If
    Set RSDetail = Nothing
End Function
Sub ProcessCRJNolink()
    'Update By BTT : to find the CRJ without link
    Dim RSCRJ                                          As New ADODB.Recordset
    Dim TheAMOUNT                                      As Double
    Dim Reference                                      As String
    Dim rsCRJ_Detail                                   As New ADODB.Recordset
    Set RSCRJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS VOUCHERNO,dbo.AMIS_Journal_HD.JType as Jtype,dbo.AMIS_Journal_HD.CustomerCode as CCode,dbo.AMIS_Journal_HD.JDate as Jdate,dbo.AMIS_Journal_HD.InvoiceAmt as InvoiceAmt, dbo.AMIS_Journal_Det.Debit as Debit,dbo.AMIS_Journal_Det.credit as Credit,dbo.AMIS_Journal_HD.InvoiceNo as ORnum, dbo.AMIS_Journal_Det.Acct_Code as Acct_code,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE dbo.AMIS_Journal_HD.JDate < = '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND (dbo.AMIS_Journal_HD.JType = 'CRJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '21-07') order by voucherno")
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
            Reference = "CRJ" + "-" + Null2String(RSCRJ!voucherno)
            If RSCRJ!DEBIT = 0 Then
                TheAMOUNT = NumericVal(RSCRJ!CREDIT)
            Else
                TheAMOUNT = NumericVal(RSCRJ!DEBIT)
            End If
            Set rsCRJ_Detail = gconDMIS.Execute("SELECT VoucherNo FROM AMIS_CRJ_DETAIL WHERE VoucherNo ='" & Null2String(RSCRJ!voucherno) & "'")
            If rsCRJ_Detail.EOF And rsCRJ_Detail.BOF Then
                If RSCRJ!jdate <= dtprocess.Value Then
                    gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCRJ!CCode & _
                                      "'," & N2Str2Null(RSCRJ!voucherno) & "," & N2Str2Null(RSCRJ!InvoiceAmt) & _
                                      "," & TheAMOUNT & "," & N2Str2Null(RSCRJ!acct_code) & "," & N2Str2Null(RSCRJ!jdate) & ",'NL')")

                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                      "'," & N2Str2Null(RSCRJ!voucherno) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                      "'," & N2Str2Null(RSCRJ!jdate) & ",'" & RSCRJ!CCode & "','" & NumericVal(0) & "','" & NumericVal(RSCRJ!InvoiceAmt) & _
                                      "','" & ((TheAMOUNT) * (-1)) & "','" & Null2String(RSCRJ!acct_code) & "','NL')")

                End If
            End If
            DoEvents
            lblRef.Caption = RSCRJ!voucherno
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
            RSCRJ.MoveNext
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
    Set RSCDJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.CREDIT as CREDITAmount,dbo.AMIS_Journal_Det.Debit as DebitAmount, dbo.AMIS_Journal_HD.VoucherNo as voucherno, dbo.AMIS_Journal_HD.VendorCode as VCode,dbo.AMIS_Journal_HD.JDate as jdate,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.jtype = dbo.AMIS_Journal_Det.jtype  WHERE (dbo.AMIS_Journal_HD.JType = 'CDJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') and dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P'")
    Me.Caption = "Loading CDJ Having AR.."
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCDJ.RecordCount
    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        Do While Not RSCDJ.EOF
            DoEvents
            Reference = "CDJ" + "-" + Null2String(RSCDJ!voucherno)
            If RSCDJ!debitAmount = 0 Then
                amount = NumericVal(RSCDJ!creditamount) * (-1)    ' Bawas ni
            Else
                amount = NumericVal(RSCDJ!debitAmount)
            End If
            If RSCDJ!jdate <= dtprocess.Value Then
                gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCDJ!VCode & _
                                  "'," & N2Str2Null(RSCDJ!voucherno) & "," & N2Str2Null(amount) & _
                                  "," & N2Str2Null("") & "," & N2Str2Null(RSCDJ!acct_code) & "," & N2Str2Null(RSCDJ!jdate) & ",'CDJ')")

                gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                  "'," & N2Str2Null(RSCDJ!voucherno) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                  "'," & N2Str2Null(RSCDJ!jdate) & ",'" & RSCDJ!VCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                  "','" & NumericVal(amount) & "','" & Null2String(RSCDJ!acct_code) & "','CDJ')")
            End If
            DoEvents
            lblRef.Caption = RSCDJ!voucherno
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
            RSCDJ.MoveNext
        Loop
    End If
    Me.Caption = "Done"
    Set RSCDJ = Nothing
End Sub

Sub ProcessCRJNoInvoice()
    'Update By BTT : to find the AR no invoice
    Dim Reference                                      As String
    Dim amount                                         As Double
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJDETAIL_TOTAL.INVOICEAMOUNT, dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO, dbo.AMIS_CRJDETAIL_TOTAL.INVOICETYPE, " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.INVOICEDATE, dbo.AMIS_CRJDETAIL_TOTAL.JDate, dbo.AMIS_CRJDETAIL_TOTAL.JType, " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo, dbo.AMIS_CRJDETAIL_TOTAL.CustomerCode, dbo.AMIS_CRJDETAIL_TOTAL.Status, " & _
                              "dbo.AMIS_Journal_Det.Acct_Code  " & _
                              "FROM dbo.AMIS_CRJDETAIL_TOTAL INNER JOIN " & _
                              "dbo.AMIS_Journal_Det ON dbo.AMIS_CRJDETAIL_TOTAL.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND " & _
                              "dbo.AMIS_CRJDETAIL_TOTAL.jtype = dbo.AMIS_Journal_Det.jtype " & _
                              "WHERE (dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO IS NULL or dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO = 'INTRO' ) AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') OR " & _
                              "(LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') and  dbo.AMIS_CRJDETAIL_TOTAL.STATUS ='P' and dbo.AMIS_CRJDETAIL_TOTAL.JDate <='" & dtprocess.Value & "'")

    Me.Caption = "Loading CRJ with blank invoice.."
    If Not (RS.EOF And RS.BOF) Then
    ProgressBar1.Value = 0
    ProgressBar1.Max = RS.RecordCount
        Do While Not RS.EOF
            amount = NumericVal(RS!invoiceAmount)
            Reference = "CRJ-" + Null2String(RS!voucherno)
            If (RS!jdate) <= dtprocess.Value Then
                gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                  "'," & N2Str2Null(RS!voucherno) & ",'XXX','XXX'," & N2Str2Null(RS!jdate) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                  "','" & NumericVal(amount) * (-1) & "','" & Null2String(RS!acct_code) & "','NOINV')")
            End If
            DoEvents
            lblRef.Caption = RS!voucherno
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub
Sub ProcessCRJmaliangcodesaCRJtoSJ()
    Dim xcount As New ADODB.Recordset
    Dim Reference                                      As String
    Dim amount                                         As Double
    'Update By BTT : to find the AR na ang CRJ nya iba ang ang account sa SJ pero tama ang link
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
                              "dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice, " & _
                              "dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode, dbo.AMIS_Journal_HD.Status, " & _
                              "dbo.AMIS_VW_AR_DET.Acct_Code as CR_ACCCODE, AMIS_VW_AR_DET_1.Acct_Code AS SJ_ACTCODE,AMIS_Journal_HD_1.CustomerCode AS CRJ_COUSTOMERCODE " & _
                              "FROM dbo.AMIS_CRJ_Detail INNER JOIN " & _
                              "dbo.AMIS_Journal_HD ON dbo.AMIS_CRJ_Detail.CR_TYPE = dbo.AMIS_Journal_HD.JType AND " & _
                              "dbo.AMIS_CRJ_Detail.VoucherNo = dbo.AMIS_Journal_HD.VoucherNo INNER JOIN " & _
                              "dbo.AMIS_VW_AR_DET ON dbo.AMIS_Journal_HD.JType = dbo.AMIS_VW_AR_DET.JType AND " & _
                              "dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_VW_AR_DET.VoucherNo INNER JOIN " & _
                              "dbo.AMIS_Journal_HD AS AMIS_Journal_HD_1 ON dbo.AMIS_CRJ_Detail.INVOICENO = AMIS_Journal_HD_1.InvoiceNo AND " & _
                              "dbo.AMIS_CRJ_Detail.INVOICETYPE = AMIS_Journal_HD_1.InvoiceType left outer JOIN " & _
                              "dbo.AMIS_VW_AR_DET AS AMIS_VW_AR_DET_1 ON AMIS_Journal_HD_1.JType = AMIS_VW_AR_DET_1.JType AND " & _
                              "AMIS_Journal_HD_1.VoucherNo = AMIS_VW_AR_DET_1.VoucherNo where dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "'")

    If Not (RS.EOF And RS.BOF) Then
        Do While Not RS.EOF
            Reference = "CRJ-" + Null2String(RS!voucherno)
            amount = NumericVal(RS!invoiceAmount)
            If Reference = "CRJ-000192" Then Stop
            'If Reference = "CRJ-003340" Then Stop
            'And rs!CustomerCode = rs!CRJ_COUSTOMERCODE
            If RS!jdate <= dtprocess.Value Then
            'Set xcount = "SELECT * from "
                If (RS!CR_ACCCODE) <> Null2String(RS!SJ_ACTCODE) And RS!CustomerCode = RS!CRJ_COUSTOMERCODE Then
                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                      "'," & N2Str2Null(RS!voucherno) & ",'" & RS!InvoiceType & "','" & RS!InvoiceNo & "'," & N2Str2Null(RS!jdate) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                      "','" & NumericVal(amount) * (-1) & "','" & Null2String(RS!CR_ACCCODE) & "','CRJWAC')")
                End If
            End If

            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub
Sub ProcessCRJWithClearingAccount()
    Dim Reference                                      As String
    Dim amount                                         As Double
    ' CRJWCA : CRJ - with clrearing account and tama ang link
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
                              "dbo.AMIS_Journal_HD.JDate, dbo.AMIS_Journal_HD.JType, dbo.AMIS_CRJ_Detail.VoucherNo, dbo.AMIS_Journal_HD.InvoiceNo AS HDInvoice, " & _
                              "dbo.AMIS_CRJ_Detail.J_Class, dbo.AMIS_CRJ_Detail.SJ_voucherno, dbo.AMIS_Journal_HD.CustomerCode, dbo.AMIS_Journal_HD.Status, " & _
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
                              "WHERE (dbo.AMIS_Journal_HD.JDate) <= '" & dtprocess.Value & "'")
    If Not RS.EOF And Not RS.BOF Then
        Do While Not RS.EOF
            Reference = "CRJ-" + Null2String(RS!voucherno)
            amount = NumericVal(RS!invoiceAmount)
            If RS!jdate <= dtprocess.Value Then
                If Null2String(RS!CRJ_ACCTCODE) <> Null2String(RS!SJ_ACCTCODE) And RS!CustomerCode = RS!CRJ_COUSTOMERCODE Then
                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                      "'," & N2Str2Null(RS!voucherno) & ",'" & RS!InvoiceType & "','" & RS!InvoiceNo & "'," & N2Str2Null(RS!jdate) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                      "','" & NumericVal(amount) & "','" & Null2String(RS!SJ_ACCTCODE) & "','CRJWCA')")
                End If
            End If
            RS.MoveNext
        Loop
    End If
    Set RS = Nothing
End Sub

Private Sub picNolink_Click()
    'ProcessCRJWithClearingAccount
End Sub
