VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmNEW_ARSchedReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GENERATE AR REPORT"
   ClientHeight    =   3510
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   4740
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
   Icon            =   "NEW_ARSchedReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   4740
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   495
      Left            =   2670
      TabIndex        =   50
      Top             =   5730
      Width           =   585
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   1425
      Left            =   9450
      TabIndex        =   48
      Top             =   4710
      Width           =   2805
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   1275
      Left            =   3540
      TabIndex        =   47
      Top             =   4290
      Width           =   915
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
      Left            =   5010
      TabIndex        =   1
      Top             =   4770
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
      Left            =   5040
      TabIndex        =   0
      Top             =   5160
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   3585
   End
   Begin Crystal.CrystalReport rptAMISDueReport 
      Left            =   90
      Top             =   3090
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
      Height          =   2115
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
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
         Format          =   326434817
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
         Format          =   326434817
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
      Left            =   5100
      TabIndex        =   14
      Top             =   5370
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
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
      Left            =   30
      ScaleHeight     =   2745
      ScaleWidth      =   3945
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   3945
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   90
         Top             =   330
      End
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
      BorderStyle     =   0  'None
      Height          =   2955
      Left            =   4050
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   38
      Top             =   6150
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
         TabIndex        =   51
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
         TabIndex        =   49
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
         TabIndex        =   46
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
         MouseIcon       =   "NEW_ARSchedReport.frx":6852
         MousePointer    =   99  'Custom
         Picture         =   "NEW_ARSchedReport.frx":69A4
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Close Window"
         Top             =   1920
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         MouseIcon       =   "NEW_ARSchedReport.frx":6DEF
         MousePointer    =   99  'Custom
         Picture         =   "NEW_ARSchedReport.frx":6F41
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Print Report"
         Top             =   1920
         Width           =   885
      End
   End
   Begin VB.PictureBox picAR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   7920
      ScaleHeight     =   1125
      ScaleWidth      =   3405
      TabIndex        =   52
      Top             =   4950
      Visible         =   0   'False
      Width           =   3435
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   405
         Left            =   60
         TabIndex        =   53
         Top             =   450
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   60
         Picture         =   "NEW_ARSchedReport.frx":73E0
         ScaleHeight     =   405
         ScaleWidth      =   465
         TabIndex        =   57
         Top             =   30
         Width           =   465
      End
      Begin VB.Label labPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   54
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   56
         Top             =   30
         Width           =   3285
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "VoucehrNo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         TabIndex        =   55
         Top             =   870
         Width           =   2475
      End
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
         MouseIcon       =   "NEW_ARSchedReport.frx":77C1
         MousePointer    =   99  'Custom
         Picture         =   "NEW_ARSchedReport.frx":7913
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
         Picture         =   "NEW_ARSchedReport.frx":7BAE
         BarPicture      =   "NEW_ARSchedReport.frx":7BCA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Format          =   124125185
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
   Begin VB.PictureBox picReport 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   3555
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   4725
      TabIndex        =   28
      Top             =   0
      Width           =   4725
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
         Left            =   600
         TabIndex        =   29
         Top             =   2730
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
         Left            =   600
         TabIndex        =   30
         Top             =   2310
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
         Left            =   600
         MaskColor       =   &H0080FFFF&
         TabIndex        =   31
         Top             =   1920
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtprocess 
         Height          =   375
         Left            =   390
         TabIndex        =   45
         Top             =   4770
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
         Format          =   124125185
         CurrentDate     =   38216
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
         Left            =   2130
         MaskColor       =   &H0080FFFF&
         TabIndex        =   44
         Top             =   4290
         Width           =   1515
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
         TabIndex        =   37
         Top             =   4380
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpAsOF 
         Height          =   345
         Left            =   2010
         TabIndex        =   66
         Top             =   1320
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
         Format          =   124190721
         CurrentDate     =   38216
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "As of:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   67
         Top             =   1410
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   120
         Picture         =   "NEW_ARSchedReport.frx":7BE6
         Top             =   30
         Width           =   720
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   30
         Left            =   30
         TabIndex        =   61
         Top             =   750
         Width           =   9525
         _Version        =   655364
         _ExtentX        =   16801
         _ExtentY        =   53
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.02
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2700
         TabIndex        =   58
         Top             =   270
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "System Computed A/R Aging and Schedule Report."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   960
         TabIndex        =   59
         Top             =   30
         Width           =   3405
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   765
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   9525
         _Version        =   655364
         _ExtentX        =   16801
         _ExtentY        =   1349
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   4725
      TabIndex        =   22
      Top             =   0
      Width           =   4725
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   90
         ScaleHeight     =   465
         ScaleWidth      =   285
         TabIndex        =   36
         Top             =   810
         Width           =   285
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Customer "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1230
         Width           =   2565
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   2400
         MouseIcon       =   "NEW_ARSchedReport.frx":84DC
         MousePointer    =   99  'Custom
         Picture         =   "NEW_ARSchedReport.frx":862E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Close Window"
         Top             =   2460
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
         Left            =   1530
         MouseIcon       =   "NEW_ARSchedReport.frx":8A79
         MousePointer    =   99  'Custom
         Picture         =   "NEW_ARSchedReport.frx":8BCB
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Print Report"
         Top             =   2460
         Width           =   885
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Account"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1110
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1800
         Width           =   2565
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "NEW_ARSchedReport.frx":906A
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0C0FF&
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
         Height          =   315
         Left            =   4410
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   675
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
         Top             =   3600
         Visible         =   0   'False
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
         Height          =   225
         Left            =   -360
         TabIndex        =   25
         Top             =   3930
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Label13"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   2820
         TabIndex        =   64
         Top             =   270
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "System Computed A/R Aging and Schedule Report."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   960
         TabIndex        =   63
         Top             =   30
         Width           =   3495
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   30
         Left            =   30
         TabIndex        =   62
         Top             =   750
         Width           =   9525
         _Version        =   655364
         _ExtentX        =   16801
         _ExtentY        =   53
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.02
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   765
         Left            =   -30
         TabIndex        =   65
         Top             =   0
         Width           =   9525
         _Version        =   655364
         _ExtentX        =   16801
         _ExtentY        =   1349
         _StockProps     =   14
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         ForeColor       =   -2147483630
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6885
      Left            =   -30
      TabIndex        =   68
      Top             =   0
      Width           =   15225
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.PictureBox picDetailedSum 
      Height          =   1245
      Left            =   510
      ScaleHeight     =   1185
      ScaleWidth      =   3735
      TabIndex        =   74
      Top             =   1110
      Visible         =   0   'False
      Width           =   3795
      Begin VB.OptionButton optDetailed 
         Caption         =   "Detailed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   90
         Value           =   -1  'True
         Width           =   2565
      End
      Begin VB.OptionButton optSummary 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   660
         Width           =   2565
      End
   End
   Begin VB.PictureBox picByAccount 
      Height          =   1275
      Left            =   30
      ScaleHeight     =   1215
      ScaleWidth      =   4605
      TabIndex        =   69
      Top             =   1110
      Width           =   4665
      Begin VB.ComboBox cboCOBAcctName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1110
         TabIndex        =   70
         Top             =   720
         Width           =   3525
      End
      Begin RichTextLib.RichTextBox txtCOBAcctNo 
         Height          =   315
         Left            =   1110
         TabIndex        =   71
         Top             =   270
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         ReadOnly        =   -1  'True
         TextRTF         =   $"NEW_ARSchedReport.frx":9960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Acct. Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   60
         TabIndex        =   73
         Top             =   810
         Width           =   1035
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Acct. Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   72
         Top             =   360
         Width           =   990
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
Attribute VB_Name = "frmNEW_ARSchedReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Report_Type                                             As String
Dim RS                                                      As New ADODB.Recordset
Dim CountDetail                                             As Integer
Dim Dealer                                                  As String
Dim xREMARKS                                                As String
Dim xSJ_CustomerCode                                        As String
Dim xCRJ_CODE                                               As String
Dim xxx_credit                                              As Double

Function SetCRJVoucherNo(XXX As String, zzz As Integer) As String
    Dim rsCRJ_Journal_HD                                    As ADODB.Recordset
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

Private Sub cboCOBAcctName_Click()
    txtCOBAcctNo.Text = Setacctcode(cboCOBAcctName.Text)
    picDetailedSum.Visible = True
    picDetailedSum.ZOrder 0
    picByAccount.Visible = False
End Sub

Private Sub cboCOBAcctName_GotFocus()
    VBComBoBoxDroppedDown cboCOBAcctName
End Sub

Private Sub Command1_Click()
    Report_Type = "SCHED"
    'getlastdate
    picReport.Visible = False
    Option1.Value = False
    Option2.Value = False
    Picture1.Visible = True
    Me.Caption = "AR SCHEDULE REPORT"
    '    If IsDate(lblDate.Caption) = True Then
    '        dtpAsOF.Value = CDate(lblDate.Caption)
    '    End If
End Sub

Private Sub Command10_Click()
'APJwithAR
End Sub

Private Sub Command2_Click()
    Report_Type = "AGING"
    'getlastdate
    picReport.Visible = False
    Option1.Value = False
    Option2.Value = False
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

'ORIGINAL CODE: COMMENTED BY JHUN 07162009--------------------------------------
'    dtpAsOF = dtprocess
'    picReport.Visible = False
'    Picture4.Visible = True
'    Picture1.Visible = False
'
'    If ar = False Then
'        picReport.Visible = True
'        MsgBox "No AR Date as of :" & dtpAsOf
'    End if
'ORIGINAL CODE: COMMENTED BY JHUN 07162009--------------------------------------

    Command7.Enabled = False
    Command3.Enabled = False
    picAR.Visible = True
    picAR.ZOrder 0
    TRANS_SLS_JOURNAL
    AR_COMPUTE
    ADVANCE_CRJ
    MsgBox "Processing AR completed...", vbInformation + vbOKOnly, "INFORMATION"
    picAR.Visible = False
    picAR.ZOrder 1

    Command7.Enabled = True
    Command3.Enabled = True
End Sub
Sub AR_IN_AP_JOURNAL()

End Sub


Sub ADVANCE_CRJ()
    Dim rsADVANCE_CRJ                                       As ADODB.Recordset
    Dim rsADVANCE_CRJ2                                      As ADODB.Recordset


    Dim RSHD                                                As ADODB.Recordset
    Dim rsCODE                                              As ADODB.Recordset


    Dim jSJ_VOUCHERNO                                       As String
    Dim jINVOICETYPE                                        As String
    Dim jINVOICENO                                          As String
    Dim jCustomerCode                                       As String
    Dim jCUSTOMERNAME                                       As String
    Dim jAMOUNT_TOPAY                                       As Double
    Dim jAMOUNT_PAID                                        As Double
    Dim jBALANCE                                            As Double
    Dim jAccount_code                                       As String
    Dim jSYSTEM_REMARKS                                     As String
    Dim jinvoicedate                                        As String
    Dim jLASTUPDATED                                        As String


    Set rsADVANCE_CRJ = New ADODB.Recordset
    rsADVANCE_CRJ.Open "SELECT X.CUST_CODE,X.CRJ_VOUCHERNO,X.INV,ROUND(SUM(INV_AMOUNT),2) AS INVOICEAMOUNT,X.JDATE,X.ACCT_CODE FROM " & _
                       "(SELECT HD.CREDIT,CODE.CUSTOMERCODE AS CUST_CODE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV, HD.JDATE AS JDATE, RTRIM(LTRIM(CRJ.CR_TYPE)) + '-' + CRJ.VOUCHERNO AS CRJ_VOUCHERNO, CRJ.INVOICEAMOUNT AS INV_AMOUNT, HD.ACCT_CODE AS ACCT_CODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET HD " & _
                       "ON  CRJ.VOUCHERNO = HD.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD CODE ON CODE.VOUCHERNO = CRJ.VOUCHERNO  WHERE HD.JTYPE = 'CRJ' AND HD.JDATE <= '" & dtprocess & "' AND HD.STATUS = 'P' AND LEFT(HD.ACCT_CODE,5) IN('11-02','11-03') AND CODE.JTYPE = 'CRJ' " & _
                       ") X WHERE X.INV IN (SELECT INVOICETYPE + '-' + INVOICENO AS HD_INV FROM AMIS_JOURNAL_HD WHERE JTYPE IN('SJ','COB') AND JDATE > '" & dtprocess & "') GROUP BY X.CUST_CODE,X.INV,X.CRJ_VOUCHERNO,X.JDATE,X.ACCT_CODE", gconDMIS, adOpenKeyset

    'FOR ACCOUNT CODE DEBUGGING ONLY
    'rsADVANCE_CRJ.Open "SELECT X.CUST_CODE,X.CRJ_VOUCHERNO,X.INV,ROUND(SUM(INV_AMOUNT),2) AS INVOICEAMOUNT,X.JDATE,X.ACCT_CODE FROM " & _
     "(SELECT HD.CREDIT,CODE.CUSTOMERCODE AS CUST_CODE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV, HD.JDATE AS JDATE, RTRIM(LTRIM(CRJ.CR_TYPE)) + '-' + CRJ.VOUCHERNO AS CRJ_VOUCHERNO, CRJ.INVOICEAMOUNT AS INV_AMOUNT, HD.ACCT_CODE AS ACCT_CODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET HD " & _
     "ON  CRJ.VOUCHERNO = HD.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD CODE ON CODE.VOUCHERNO = CRJ.VOUCHERNO  WHERE HD.JTYPE = 'CRJ' AND HD.JDATE <= '" & dtprocess & "' AND HD.STATUS = 'P' AND HD.ACCT_CODE = '11-02023-00' AND CODE.JTYPE = 'CRJ' " & _
     ") X WHERE X.INV IN (SELECT INVOICETYPE + '-' + INVOICENO AS HD_INV FROM AMIS_JOURNAL_HD WHERE JTYPE IN('SJ','COB') AND JDATE > '" & dtprocess & "') GROUP BY X.CUST_CODE,X.INV,X.CRJ_VOUCHERNO,X.JDATE,X.ACCT_CODE", gconDMIS, adOpenKeyset

    If Not rsADVANCE_CRJ.EOF And Not rsADVANCE_CRJ.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsADVANCE_CRJ.RecordCount
        Label11.Caption = "Processing CRJ... Please Wait.."

        Do While Not rsADVANCE_CRJ.EOF

            If Null2String(rsADVANCE_CRJ!ACCT_CODE) = "11-02002-00" Then
                'DESCRIPTION: THIS IS FOR CRJ TO CRJ - PAYMENT IS ADVANCE TO THE AR ACCOUNTS
                Set rsADVANCE_CRJ2 = New ADODB.Recordset
                rsADVANCE_CRJ2.Open "SELECT X.CUST_CODE,X.CRJ_VOUCHERNO,X.INV,ROUND(SUM(INV_AMOUNT),2) AS INVOICEAMOUNT,X.JDATE,X.ACCT_CODE FROM " & _
                                    "(SELECT HD.CREDIT,CODE.CUSTOMERCODE AS CUST_CODE,'CRJ' + '-' + CRJ.VOUCHERNO AS INV, HD.JDATE AS JDATE, RTRIM(LTRIM(CRJ.CR_TYPE)) + '-' + CRJ.VOUCHERNO AS CRJ_VOUCHERNO, CRJ.INVOICEAMOUNT AS INV_AMOUNT, HD.ACCT_CODE AS ACCT_CODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET HD " & _
                                    "ON  CRJ.VOUCHERNO = HD.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD CODE ON CODE.VOUCHERNO = CRJ.VOUCHERNO  WHERE HD.JTYPE = 'CRJ' AND HD.JDATE <= '" & dtprocess & "' AND HD.STATUS = 'P' AND HD.ACCT_CODE = '11-02002-00' AND CODE.JTYPE = 'CRJ' " & _
                                    ") X WHERE X.INV IN (SELECT JTYPE + '-' + VOUCHERNO AS HD_INV FROM AMIS_JOURNAL_HD WHERE JTYPE IN('CRJ') AND JDATE >'" & dtprocess & "') GROUP BY X.CUST_CODE,X.INV,X.CRJ_VOUCHERNO,X.JDATE,X.ACCT_CODE", gconDMIS, adOpenKeyset

                If Not rsADVANCE_CRJ2.EOF And Not rsADVANCE_CRJ2.BOF Then

                    jSJ_VOUCHERNO = N2Str2Null(rsADVANCE_CRJ!CRJ_VOUCHERNO)
                    jINVOICETYPE = N2Str2Null(Left(rsADVANCE_CRJ!INV, 2))
                    jINVOICENO = N2Str2Null(Right(rsADVANCE_CRJ!INV, 6))
                    jCustomerCode = N2Str2Null(rsADVANCE_CRJ!CUST_CODE)
                    jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(rsADVANCE_CRJ!CUST_CODE))
                    jAMOUNT_TOPAY = 0
                    jAMOUNT_PAID = NumericVal(rsADVANCE_CRJ!invoiceamount)
                    jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                    jAccount_code = N2Str2Null(rsADVANCE_CRJ!ACCT_CODE)
                    jSYSTEM_REMARKS = N2Str2Null("u")
                    jinvoicedate = N2Date2Null(rsADVANCE_CRJ!JDATE)
                    jLASTUPDATED = N2Date2Null(LOGDATE)

                    gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                     "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"

                    If CHECK_DUPLICATE(jINVOICENO, jINVOICETYPE, jSJ_VOUCHERNO) = False Then
                        'THIS IS FOR AR PAYMENT DETAIL
                        gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                                         "VALUES(" & N2Str2Null(jINVOICENO) & ", " & N2Str2Null(jINVOICETYPE) & ", " & NumericVal(jAMOUNT_PAID) & ", " & N2Str2Null(jCustomerCode) & "," & N2Str2Null(jAccount_code) & "," & N2Date2Null(rsADVANCE_CRJ!JDATE) & ",'4'," & jSJ_VOUCHERNO & ")"
                    End If


                End If
            Else
                Dim rsINVOICE_CUSCDE                        As ADODB.Recordset
                'DESCRIPTION: THIS IS FOR DUPLICATE INVOICENO AND INVOICETYPE JTYPE IS (COB AND SJ) WHICH IS WITHIN THE PERIOD OF PROCESS
                Set rsINVOICE_CUSCDE = New ADODB.Recordset
                rsINVOICE_CUSCDE.Open "SELECT * FROM AMIS_AR WHERE INVOICENO = " & N2Str2Null(Right(rsADVANCE_CRJ!INV, 6)) & " AND INVOICETYPE = " & N2Str2Null(Left(rsADVANCE_CRJ!INV, 2)) & " AND CUSTOMERCODE = " & N2Str2Null(rsADVANCE_CRJ!CUST_CODE) & " AND SJVOUCHERNO = '" & rsADVANCE_CRJ!CRJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
                If Not rsINVOICE_CUSCDE.EOF And Not rsINVOICE_CUSCDE.BOF Then
                Else
                    If LTrim(RTrim(Right(rsADVANCE_CRJ!INV, 6))) = "INT RO" Then
                    Else
                        jSJ_VOUCHERNO = N2Str2Null(rsADVANCE_CRJ!CRJ_VOUCHERNO)
                        jINVOICETYPE = N2Str2Null(Left(rsADVANCE_CRJ!INV, 2))
                        jINVOICENO = N2Str2Null(Right(rsADVANCE_CRJ!INV, 6))
                        jCustomerCode = N2Str2Null(rsADVANCE_CRJ!CUST_CODE)
                        jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(rsADVANCE_CRJ!CUST_CODE))
                        jAMOUNT_TOPAY = 0
                        jAMOUNT_PAID = NumericVal(rsADVANCE_CRJ!invoiceamount)
                        jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                        jAccount_code = N2Str2Null(rsADVANCE_CRJ!ACCT_CODE)
                        jSYSTEM_REMARKS = N2Str2Null("x")
                        jinvoicedate = N2Date2Null(rsADVANCE_CRJ!JDATE)
                        jLASTUPDATED = N2Date2Null(LOGDATE)

                        gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                         "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"

                        If CHECK_DUPLICATE(jINVOICENO, jINVOICETYPE, jSJ_VOUCHERNO) = False Then
                            'THIS IS FOR AR PAYMENT DETAIL
                            gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                                             "VALUES(" & N2Str2Null(jINVOICENO) & ", " & N2Str2Null(jINVOICETYPE) & ", " & NumericVal(jAMOUNT_PAID) & ", " & N2Str2Null(jCustomerCode) & "," & N2Str2Null(jAccount_code) & "," & N2Date2Null(rsADVANCE_CRJ!JDATE) & ",'5'," & jSJ_VOUCHERNO & ")"
                        End If
                    End If

                End If
            End If

            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = Null2String(rsADVANCE_CRJ!CRJ_VOUCHERNO)
            DoEvents
            rsADVANCE_CRJ.MoveNext
        Loop
    End If

    'DESCRIPTION: THIS IS FOR PAYMENT FOR CREDIT CARD WHICH HAS NO AR CREDIT CARD RECEIVABLE
    Dim rsAR_CCARD                                          As ADODB.Recordset
    Set rsAR_CCARD = New ADODB.Recordset
    rsAR_CCARD.Open "SELECT X.VOUCHERNO,X.JDATE,RTRIM(LTRIM(X.INV_TYPE)) AS INV_TYPE,X.I_TYPE,X.I_NO,X.INV_AMOUNT,X.C_CODE,X.ACCT_CODE,X.JTYPE FROM " & _
                    "(SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV_TYPE,CRJ.INVOICENO AS I_NO,CRJ.INVOICETYPE AS I_TYPE,CRJ.INVOICEAMOUNT AS INV_AMOUNT, DET.JDATE AS JDATE, CRJ.VOUCHERNO AS VOUCHERNO, HD.CUSTOMERCODE AS C_CODE, DET.ACCT_CODE AS ACCT_CODE,DET.JTYPE AS JTYPE  FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO  AND CRJ.CR_TYPE = DET.JTYPE " & _
                    "INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO  AND CRJ.CR_TYPE = HD.JTYPE WHERE DET.ACCT_CODE = '11-02002-00' AND DET.CREDIT <> 0 AND DET.STATUS = 'P' AND DET.JDATE <= '" & dtprocess & "' AND DET.JTYPE IN ('CRJ') " & _
                    ") X WHERE INV_TYPE NOT IN (SELECT  RTRIM(LTRIM(CRJ.INVOICETYPE)) + '-' + RTRIM(LTRIM(CRJ.INVOICENO)) AS INV_TYPE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO WHERE DET.ACCT_CODE = '11-02002-00' AND DET.DEBIT <> 0 AND DET.STATUS = 'P' AND DET.JDATE <= '" & dtprocess & "' AND DET.JTYPE = 'CRJ')", gconDMIS, adOpenKeyset

    If Not rsAR_CCARD.EOF And Not rsAR_CCARD.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsAR_CCARD.RecordCount
        Label11.Caption = "Validating CRJ... Please Wait.."
        Do While Not rsAR_CCARD.EOF
            'DESCRIPTION: CHECK THE INVOICE IF ITS IN COB
            Dim rsIS_IN_COB                                 As ADODB.Recordset
            Set rsIS_IN_COB = New ADODB.Recordset
            rsIS_IN_COB.Open "SELECT * FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                             "WHERE DET.ACCT_CODE = '11-02002-00' AND HD.INVOICENO = '" & rsAR_CCARD!I_NO & "' AND HD.INVOICETYPE = '" & rsAR_CCARD!I_TYPE & "' AND HD.JTYPE = 'COB' AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
            If Not rsIS_IN_COB.EOF And Not rsIS_IN_COB.BOF Then
                'FOUND
            Else
                'DESCRIPTION: CHECK IF THE INVOICE HAS COMPLETELY NO LINK
                Dim rsCOMP_NOLINK                           As ADODB.Recordset
                Set rsCOMP_NOLINK = New ADODB.Recordset
                rsCOMP_NOLINK.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & rsAR_CCARD!I_NO & "' AND INVOICETYPE = '" & rsAR_CCARD!I_TYPE & "' AND JTYPE IN('SJ','COB') AND STATUS = 'P'", gconDMIS, adOpenKeyset
                If Not rsCOMP_NOLINK.EOF And Not rsCOMP_NOLINK.BOF Then
                    jSJ_VOUCHERNO = N2Str2Null(rsAR_CCARD!VOUCHERNO)
                    jINVOICETYPE = N2Str2Null(rsAR_CCARD!I_TYPE)
                    jINVOICENO = N2Str2Null(rsAR_CCARD!I_NO)
                    jCustomerCode = N2Str2Null(rsAR_CCARD!C_CODE)
                    jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(rsAR_CCARD!C_CODE))
                    jAMOUNT_TOPAY = 0
                    jAMOUNT_PAID = NumericVal(rsAR_CCARD!INV_AMOUNT)
                    jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                    jAccount_code = N2Str2Null(rsAR_CCARD!ACCT_CODE)
                    jSYSTEM_REMARKS = N2Str2Null("w")
                    jinvoicedate = N2Date2Null(rsAR_CCARD!JDATE)
                    jLASTUPDATED = N2Date2Null(LOGDATE)

                    gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                     "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"

                    If CHECK_DUPLICATE(jINVOICENO, jINVOICETYPE, jSJ_VOUCHERNO) = False Then
                        'THIS IS FOR AR PAYMENT DETAIL
                        gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                                         "VALUES(" & N2Str2Null(jINVOICENO) & ", " & N2Str2Null(jINVOICETYPE) & ", " & NumericVal(jAMOUNT_PAID) & ", " & N2Str2Null(jCustomerCode) & "," & N2Str2Null(jAccount_code) & "," & N2Date2Null(rsAR_CCARD!JDATE) & ",'6'," & jSJ_VOUCHERNO & ")"
                    End If
                Else
                    'THIS IS ALREADY INSERT IN AR_COMPUTE SUB ROUTINE INVALID INVOICE
                End If
            End If
            Set rsIS_IN_COB = Nothing
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = Null2String(rsAR_CCARD!VOUCHERNO)
            DoEvents
            rsAR_CCARD.MoveNext
        Loop
    End If


    'THIS IS FOR CRJ PAYMENT WHICH DETAILS AR ACCOUNTING ENTRY IS IN DEBIT SIDE NOT IN CREDIT
    'FOR REFERENCE SEE VOUCHERNO: CRJ-005932
    Dim WRONG_DETAILS                                       As ADODB.Recordset
    Set WRONG_DETAILS = New ADODB.Recordset
    WRONG_DETAILS.Open "SELECT DET.JDATE AS JDATE,DET.ACCT_CODE AS Acct_Code,CRJ.VOUCHERNO AS CRJ_VOUCHERNO,HD.CUSTOMERCODE AS CUST_CODE,CRJ.INVOICEAMOUNT AS INV_AMT,CRJ.INVOICETYPE AS INV_TYPE,CRJ.INVOICENO AS INV_NO FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_CRJ_DETAIL CRJ ON DET.VOUCHERNO = CRJ.VOUCHERNO AND DET.JTYPE = CRJ.CR_TYPE INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE  WHERE DET.ACCT_CODE <> '11-02002-00' AND DET.DEBIT <> 0 AND DET.JDATE <= '" & dtprocess & "' AND LEFT(DET.ACCT_CODE,5)IN('11-02','11-03') AND DET.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not WRONG_DETAILS.EOF And Not WRONG_DETAILS.BOF Then
        jAMOUNT_TOPAY = NumericVal(Abs(WRONG_DETAILS!INV_AMT))
        jAMOUNT_PAID = 0
        jINVOICETYPE = N2Str2Null(WRONG_DETAILS!INV_TYPE)
        jINVOICENO = N2Str2Null(WRONG_DETAILS!INV_NO)
        jSJ_VOUCHERNO = N2Str2Null(WRONG_DETAILS!CRJ_VOUCHERNO)
        jCustomerCode = N2Str2Null(WRONG_DETAILS!CUST_CODE)
        jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(WRONG_DETAILS!CUST_CODE))
        jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
        jAccount_code = N2Str2Null(WRONG_DETAILS!ACCT_CODE)
        jSYSTEM_REMARKS = N2Str2Null("y")
        jinvoicedate = N2Date2Null(WRONG_DETAILS!JDATE)
        jLASTUPDATED = N2Date2Null(LOGDATE)

        gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                         "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
        If CHECK_DUPLICATE(jINVOICENO, jINVOICETYPE, jSJ_VOUCHERNO) = False Then
            'THIS IS FOR AR PAYMENT DETAIL
            gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                             "VALUES(" & N2Str2Null(jINVOICENO) & ", " & N2Str2Null(jINVOICETYPE) & ", " & NumericVal(jAMOUNT_PAID) & ", " & N2Str2Null(jCustomerCode) & "," & N2Str2Null(jAccount_code) & "," & N2Date2Null(WRONG_DETAILS!JDATE) & ",'7'," & jSJ_VOUCHERNO & ")"
        End If
    End If
    Set WRONG_DETAILS = Nothing

    'THIS IS FOR ACCOUNT RECEIVABLE WITH JTYPE IN ('SJ','COB') WHERE AR AMOUNT IS IN CREDIT SIDE
    Dim rsSJ_CREDIT                                         As ADODB.Recordset
    Set rsSJ_CREDIT = New ADODB.Recordset
    rsSJ_CREDIT.Open "SELECT HD.JDATE AS JDATE,DET.ACCT_CODE AS ACCT_CODE,HD.CUSTOMERCODE AS CUSCODE,HD.VOUCHERNO AS HD_VOUCHERNO,HD.INVOICENO AS INV_NO,HD.INVOICETYPE AS INV_TYPE,DET.DEBIT,DET.CREDIT AS CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO  " & _
                     "AND HD.JTYPE = DET.JTYPE WHERE LEFT(DET.ACCT_CODE,5)IN('11-02','11-03') AND HD.JTYPE IN ('SJ','COB') AND DET.CREDIT <> 0 AND HD.JDATE <= '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsSJ_CREDIT.EOF And Not rsSJ_CREDIT.BOF Then
        jAMOUNT_TOPAY = 0
        jAMOUNT_PAID = NumericVal(rsSJ_CREDIT!Credit)
        jINVOICETYPE = N2Str2Null(rsSJ_CREDIT!INV_TYPE)
        jINVOICENO = N2Str2Null(rsSJ_CREDIT!INV_NO)
        jSJ_VOUCHERNO = N2Str2Null(rsSJ_CREDIT!HD_VOUCHERNO)
        jCustomerCode = N2Str2Null(rsSJ_CREDIT!CUSCODE)
        jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(rsSJ_CREDIT!CUSCODE))
        jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
        jAccount_code = N2Str2Null(rsSJ_CREDIT!ACCT_CODE)
        jSYSTEM_REMARKS = N2Str2Null("z")
        jinvoicedate = N2Date2Null(rsSJ_CREDIT!JDATE)
        jLASTUPDATED = N2Date2Null(LOGDATE)

        gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                         "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
        If CHECK_DUPLICATE(jINVOICENO, jINVOICETYPE, jSJ_VOUCHERNO) = False Then
            'THIS IS FOR AR PAYMENT DETAIL
            gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO)" & _
                             "VALUES(" & N2Str2Null(jINVOICENO) & ", " & N2Str2Null(jINVOICETYPE) & ", " & NumericVal(jAMOUNT_PAID) & ", " & N2Str2Null(jCustomerCode) & "," & N2Str2Null(jAccount_code) & "," & N2Date2Null(rsSJ_CREDIT!JDATE) & ",'8'," & jSJ_VOUCHERNO & ")"
        End If
    End If

    Set rsSJ_CREDIT = Nothing


    'THIS IS FOR CRJ PAYMENT WHICH IS INVOICE NO AND INVOICE TYPE IS IN SJ BUT THE ACCT_CODE IS NOT EQUAL
    Dim rsCODE_NOLINK                                       As ADODB.Recordset
    Dim rsdetail                                            As ADODB.Recordset
    Dim rsDUP_INVOICE                                       As ADODB.Recordset
    Dim OLD_VOUCHER                                         As String


    Set rsCODE_NOLINK = New ADODB.Recordset

    rsCODE_NOLINK.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM " & _
                       "(SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO,CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, " & _
                       "CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE " & _
                       "INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE " & _
                       "WHERE LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' AND DET.STATUS = 'P' AND DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0 " & _
                       ") X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE X.ACCT_CODE <> DET.ACCT_CODE AND HD.CUSTOMERCODE = X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') AND (DET.DEBIT <> 0 OR DET.DEBIT IS NULL)) ", gconDMIS, adOpenKeyset

    If Not rsCODE_NOLINK.EOF And Not rsCODE_NOLINK.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsCODE_NOLINK.RecordCount
        Label11.Caption = "Validating CRJ... Please Wait.."
        Do While Not rsCODE_NOLINK.EOF

            'If rsCODE_NOLINK!DET_VOUCHERNO = "000209" Then Stop

            If OLD_VOUCHER = rsCODE_NOLINK!DET_VOUCHERNO Then
            Else
                Set rsDUP_INVOICE = New ADODB.Recordset
                'DESCRIPTION: THIS IS FOR REVALIDATION FOR DUPLICATE INVOICENO 1 ACCT CODE IS EQUAL AND THE OTHER ONE IS NOT PRONE TO WARRANTY
                rsDUP_INVOICE.Open "SELECT X.xINVOICE, X.xACCT_CODE,X.xCUSTOMERCODE FROM ( " & _
                                   "SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS xINVOICE ,DET.ACCT_CODE AS xACCT_CODE, HD.CUSTOMERCODE AS xCUSTOMERCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.CR_TYPE = HD.JTYPE AND CRJ.VOUCHERNO = HD.VOUCHERNO WHERE CRJ.VOUCHERNO = " & N2Str2Null(rsCODE_NOLINK!DET_VOUCHERNO) & " AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') " & _
                                   ")X WHERE X.xINVOICE IN(SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAl_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE  DET.ACCT_CODE = X.xACCT_CODE AND HD.CUSTOMERCODE = X.xCUSTOMERCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03'))", gconDMIS, adOpenKeyset
                If Not rsDUP_INVOICE.EOF And Not rsDUP_INVOICE.BOF Then
                    Dim rsMULTIPLE_ACCT                     As ADODB.Recordset
                    Set rsMULTIPLE_ACCT = New ADODB.Recordset
                    rsMULTIPLE_ACCT.Open "SELECT COUNT(DISTINCT ACCT_CODE) AS CODE_COUNT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & rsCODE_NOLINK!DET_VOUCHERNO & "' AND JTYPE = 'CRJ' AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset
                    If Not rsMULTIPLE_ACCT.EOF And Not rsMULTIPLE_ACCT.BOF Then
                        If rsMULTIPLE_ACCT!CODE_COUNT > 1 Then
                            Set rsdetail = New ADODB.Recordset
                            rsdetail.Open "SELECT HD.JDATE,DET.CREDIT,DET.ACCT_CODE FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON DET.VOUCHERNO = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE  WHERE HD.VOUCHERNO = " & N2Str2Null(rsCODE_NOLINK!DET_VOUCHERNO) & " AND HD.JTYPE = 'CRJ' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
                            If Not rsdetail.EOF And Not rsdetail.BOF Then
                                Do While Not rsdetail.EOF
                                    jAMOUNT_TOPAY = 0
                                    jAMOUNT_PAID = NumericVal(rsdetail!Credit)
                                    jINVOICETYPE = N2Str2Null("")
                                    jINVOICENO = N2Str2Null("")
                                    jSJ_VOUCHERNO = N2Str2Null("CRJ" & "-" & rsCODE_NOLINK!DET_VOUCHERNO)
                                    jCustomerCode = N2Str2Null("XXXXXX")
                                    jCUSTOMERNAME = N2Str2Null("XXXXXX")
                                    jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                                    jAccount_code = N2Str2Null(rsdetail!ACCT_CODE)
                                    jSYSTEM_REMARKS = N2Str2Null("INVALID CODE")
                                    jinvoicedate = N2Date2Null(rsdetail!JDATE)
                                    jLASTUPDATED = N2Date2Null(LOGDATE)

                                    If Null2String(rsdetail!ACCT_CODE) <> "11-02002-00" Then
                                        gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                                         "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
                                    End If
                                    rsdetail.MoveNext
                                Loop
                            End If
                        Else
                        End If
                    End If
                    Set rsMULTIPLE_ACCT = Nothing


                Else
                    Set rsdetail = New ADODB.Recordset
                    rsdetail.Open "SELECT HD.JDATE,DET.CREDIT,DET.ACCT_CODE FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON DET.VOUCHERNO = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE  WHERE HD.VOUCHERNO = " & N2Str2Null(rsCODE_NOLINK!DET_VOUCHERNO) & " AND HD.JTYPE = 'CRJ' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
                    If Not rsdetail.EOF And Not rsdetail.BOF Then
                        Do While Not rsdetail.EOF
                            jAMOUNT_TOPAY = 0
                            jAMOUNT_PAID = NumericVal(rsdetail!Credit)
                            jINVOICETYPE = N2Str2Null("")
                            jINVOICENO = N2Str2Null("")
                            jSJ_VOUCHERNO = N2Str2Null("CRJ" & "-" & rsCODE_NOLINK!DET_VOUCHERNO)
                            jCustomerCode = N2Str2Null("XXXXXX")
                            jCUSTOMERNAME = N2Str2Null("XXXXXX")
                            jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                            jAccount_code = N2Str2Null(rsdetail!ACCT_CODE)
                            jSYSTEM_REMARKS = N2Str2Null("INVALID CODE")
                            jinvoicedate = N2Date2Null(rsdetail!JDATE)
                            jLASTUPDATED = N2Date2Null(LOGDATE)

                            If Null2String(rsdetail!ACCT_CODE) <> "11-02002-00" Then
                                gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                                 "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
                            End If
                            rsdetail.MoveNext
                        Loop
                    End If
                End If
            End If
            OLD_VOUCHER = Null2String(rsCODE_NOLINK!DET_VOUCHERNO)
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = Null2String(rsCODE_NOLINK!DET_VOUCHERNO)
            DoEvents
            rsCODE_NOLINK.MoveNext
        Loop
    Else
        'DO NOTHING
    End If
    Set rsCODE_NOLINK = Nothing


    'DESCRIPTION: THIS IS FOR CRJ PAYMENT WHICH IS INVOICE NO AND INVOICE TYPE IS IN SJ BUT THE ACCT_CODE IS NOT EQUAL AND CUSTOMER CODE IS NOT EQUAL
    Dim rsWRONG_ENTRY                                       As ADODB.Recordset
    Set rsWRONG_ENTRY = New ADODB.Recordset
    rsWRONG_ENTRY.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM ( " & _
                       "SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO, " & _
                       "CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE FROM AMIS_CRJ_DETAIL CRJ " & _
                       "INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE INNER JOIN AMIS_JOURNAL_HD HD " & _
                       "ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE WHERE LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' " & _
                       "AND DET.STATUS = 'P' AND DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0) X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO FROM " & _
                       "AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                       "WHERE X.ACCT_CODE <> DET.ACCT_CODE AND HD.CUSTOMERCODE <> X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') AND (DET.DEBIT <> 0 OR DET.DEBIT IS NULL) )", gconDMIS, adOpenKeyset
    If Not rsWRONG_ENTRY.EOF And Not rsWRONG_ENTRY.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsWRONG_ENTRY.RecordCount
        Label11.Caption = "Validating CRJ... Please Wait.."
        Do While Not rsWRONG_ENTRY.EOF

            'If rsWRONG_ENTRY!DET_VOUCHERNO = "005501" Then Stop

            Dim rsDUP_INVOICE2                              As ADODB.Recordset
            Set rsDUP_INVOICE2 = New ADODB.Recordset
            'DESCRIPTION: THIS IS FOR REVALIDATION FOR DUPLICATE INVOICENO 1 ACCT CODE IS EQUAL AND THE OTHER ONE IS NOT PRONE TO WARRANTY
            rsDUP_INVOICE2.Open "SELECT X.xINVOICE, X.xACCT_CODE,X.xCUSTOMERCODE FROM ( " & _
                                "SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS xINVOICE ,DET.ACCT_CODE AS xACCT_CODE, HD.CUSTOMERCODE AS xCUSTOMERCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.CR_TYPE = HD.JTYPE AND CRJ.VOUCHERNO = HD.VOUCHERNO WHERE CRJ.VOUCHERNO = " & N2Str2Null(rsWRONG_ENTRY!DET_VOUCHERNO) & " AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') " & _
                                ")X WHERE X.xINVOICE IN(SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAl_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE  DET.ACCT_CODE = X.xACCT_CODE AND HD.CUSTOMERCODE = X.xCUSTOMERCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03'))", gconDMIS, adOpenKeyset
            If Not rsDUP_INVOICE2.EOF And Not rsDUP_INVOICE2.BOF Then
            Else
                Dim rsdetail2                               As ADODB.Recordset
                Set rsdetail2 = New ADODB.Recordset
                rsdetail2.Open "SELECT HD.JDATE,DET.CREDIT,DET.ACCT_CODE FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON DET.VOUCHERNO = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE  WHERE HD.VOUCHERNO = " & N2Str2Null(rsWRONG_ENTRY!DET_VOUCHERNO) & " AND HD.JTYPE = 'CRJ' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
                jAMOUNT_TOPAY = 0
                jAMOUNT_PAID = NumericVal(rsdetail2!Credit)
                jINVOICETYPE = N2Str2Null("")
                jINVOICENO = N2Str2Null("")
                jSJ_VOUCHERNO = N2Str2Null("CRJ" & "-" & rsWRONG_ENTRY!DET_VOUCHERNO)
                jCustomerCode = N2Str2Null("XXXXXX")
                jCUSTOMERNAME = N2Str2Null("XXXXXX")
                jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                jAccount_code = N2Str2Null(rsdetail2!ACCT_CODE)
                jSYSTEM_REMARKS = N2Str2Null("INV CSCODE AND ACCTCODE")
                jinvoicedate = N2Date2Null(rsdetail2!JDATE)
                jLASTUPDATED = N2Date2Null(LOGDATE)

                If Null2String(rsdetail2!ACCT_CODE) <> "11-02002-00" Then
                    gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                     "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
                End If
                Set rsdetail2 = Nothing
            End If
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = Null2String(rsWRONG_ENTRY!DET_VOUCHERNO)
            DoEvents
            rsWRONG_ENTRY.MoveNext
        Loop
    End If


    'DESCRIPTION: THIS IS FOR CRJ HAS A CORRECT LINK IN SJ BUT THE HEADER STATUS = 'N'
    Dim rsNOT_POSTED                                        As ADODB.Recordset
    Set rsNOT_POSTED = New ADODB.Recordset
    rsNOT_POSTED.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM ( " & _
                      "SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE, " & _
                      "CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO, " & _
                      "CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, " & _
                      "CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE " & _
                      "FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON " & _
                      "CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE INNER JOIN " & _
                      "AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' AND DET.STATUS = 'P' AND " & _
                      "DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0 ) X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE X.ACCT_CODE = DET.ACCT_CODE AND HD.CUSTOMERCODE = X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') " & _
                      "AND (DET.DEBIT <> 0 OR DET.DEBIT IS NULL) AND HD.STATUS NOT IN('C','P'))", gconDMIS, adOpenKeyset
    If Not rsNOT_POSTED.EOF And Not rsNOT_POSTED.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsNOT_POSTED.RecordCount
        Label11.Caption = "Validating CRJ... Please Wait.."
        Do While Not rsNOT_POSTED.EOF
            Dim rsdetail3                                   As ADODB.Recordset
            Set rsdetail3 = New ADODB.Recordset
            'rsdetail3.Open "SELECT HD.JDATE,DET.CREDIT,DET.ACCT_CODE FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD ON DET.VOUCHERNO = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE  WHERE HD.VOUCHERNO = " & N2Str2Null(rsNOT_POSTED!DET_VOUCHERNO) & " AND HD.JTYPE = 'CRJ' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
            rsdetail3.Open "SELECT HD.CUSTOMERCODE,CRJ.INVOICETYPE,CRJ.INVOICENO,HD.JDATE,DET.CREDIT,DET.ACCT_CODE " & _
                           "FROM AMIS_JOURNAL_DET DET INNER JOIN AMIS_JOURNAL_HD HD " & _
                           "ON DET.VOUCHERNO = HD.VOUCHERNO AND DET.JTYPE = HD.JTYPE " & _
                           "INNER JOIN AMIS_CRJ_DETAIL CRJ ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO " & _
                           "WHERE HD.VOUCHERNO = " & N2Str2Null(rsNOT_POSTED!DET_VOUCHERNO) & " AND HD.JTYPE = 'CRJ' AND LEFT(DET.ACCT_CODE,5) IN('11-02','11-03') AND HD.STATUS = 'P' AND DET.CREDIT <> 0 ", gconDMIS, adOpenKeyset
            If Not rsdetail3.EOF And Not rsdetail3.BOF Then
                Do While Not rsdetail3.EOF
                    'DIM CHECK IF ALREADY INSERTED OR COMPUTED
                    Dim rsCHECK_INV                         As ADODB.Recordset
                    Set rsCHECK_INV = New ADODB.Recordset
                    rsCHECK_INV.Open "SELECT * FROM AMIS_AR WHERE INVOICENO = " & N2Str2Null(rsdetail3!INVOICENO) & " AND INVOICETYPE = " & N2Str2Null(rsdetail3!InvoiceType) & " AND CUSTOMERCODE = " & N2Str2Null(rsdetail3!CustomerCode) & " AND ACCOUNT_CODE = " & N2Str2Null(rsdetail3!ACCT_CODE) & "", gconDMIS, adOpenKeyset
                    If Not rsCHECK_INV.EOF And Not rsCHECK_INV.BOF Then
                    Else
                        jAMOUNT_TOPAY = 0
                        jAMOUNT_PAID = NumericVal(rsdetail3!Credit)
                        jINVOICETYPE = N2Str2Null(rsdetail3!InvoiceType)
                        jINVOICENO = N2Str2Null(rsdetail3!INVOICENO)
                        jSJ_VOUCHERNO = N2Str2Null("CRJ" & "-" & rsNOT_POSTED!DET_VOUCHERNO)
                        jCustomerCode = Null2String(rsdetail3!CustomerCode)
                        jCUSTOMERNAME = N2Str2Null(GET_CUST_NAME(RTrim(LTrim(jCustomerCode))))
                        jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                        jAccount_code = N2Str2Null(rsdetail3!ACCT_CODE)
                        jSYSTEM_REMARKS = N2Str2Null("SJ NOT POSTED")
                        jinvoicedate = N2Date2Null(rsdetail3!JDATE)
                        jLASTUPDATED = N2Date2Null(LOGDATE)

                        If Null2String(rsdetail3!ACCT_CODE) <> "11-02002-00" Then
                            gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                             "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & ",'" & jCustomerCode & "'," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
                        End If
                    End If
                    rsdetail3.MoveNext
                Loop
            End If
            Set rsdetail3 = Nothing
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = Null2String(rsNOT_POSTED!DET_VOUCHERNO)
            DoEvents
            rsNOT_POSTED.MoveNext
        Loop
    End If
    Set rsNOT_POSTED = Nothing

    Set rsCOMP_NOLINK = Nothing
    Set rsIS_IN_COB = Nothing
    Set rsAR_CCARD = Nothing
    Set rsADVANCE_CRJ = Nothing
    Set RSHD = Nothing
    Set rsCODE = Nothing
End Sub

Sub TRANS_SLS_JOURNAL()
'DESCRIPTION: TRANSFER THE SALES JOURNAL FROM AMIS_JOURNAL_HD TO AMIS_AR_HD
    Dim rsTRANS_SLS_JOURNAL                                 As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xSTATUS                                             As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xInvoiceType                                        As String
    Dim xInvoiceNo                                          As String
    Dim xInvoicedate                                        As String
    Dim xINVOICE_AMT                                        As Double
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xACCNT_NAME                                         As String
    Dim xdebit                                              As Double

    gconDMIS.Execute "TRUNCATE TABLE AMIS_AR_HD"

    Set rsTRANS_SLS_JOURNAL = New ADODB.Recordset


    rsTRANS_SLS_JOURNAL.Open "SELECT HD_DET.INVOICENO AS CDJ_NO,HD.VENDORCODE AS VEN_CODE,HD.VoucherNo as HD_VOUCHERNO,HD.jdate AS HD_JDATE,HD.Status AS HD_STATUS, HD.JType AS HD_JTYPE,HD.CustomerCode AS HD_CUST_CODE,HD.InvoiceType AS HD_INV_TYPE, " & _
                             "HD.InvoiceNo AS HD_INV_NO,HD.InvoiceDate AS HD_INV_DATE,HD.InvoiceAmt AS HD_INV_AMT,HD.AmountToPay AS HD_AMT_TO_PAY,HD.AmountPaid AS HD_AMT_PAID,HD_DET.Acct_Code AS DET_ACCT_CODE, " & _
                             "HD_DET.Acct_Name AS DET_ACCT_NAME,HD_DET.Debit AS DET_DEBIT FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_Journal_Det HD_DET ON HD.VoucherNo = HD_DET.VoucherNo AND HD.JType = HD_DET.JType " & _
                             "WHERE LEFT(HD_DET.Acct_Code,5) IN ('11-02','11-03') AND HD.JType IN('SJ','CDJ','COB','CRJ','APJ','CCM','GJ','VCJ','VDJ') AND HD.jdate <= " & N2Str2Null(dtprocess) & " AND HD.status ='P' AND (AR_BALANCE <> 0 OR AR_BALANCE IS NULL) ORDER BY HD.VoucherNo", gconDMIS, adOpenKeyset

    'FOR DEBUGGING PURPOSE ONLY
    'rsTRANS_SLS_JOURNAL.Open "SELECT HD_DET.INVOICENO AS CDJ_NO,HD.VENDORCODE AS VEN_CODE,HD.VoucherNo as HD_VOUCHERNO,HD.jdate AS HD_JDATE,HD.Status AS HD_STATUS, HD.JType AS HD_JTYPE,HD.CustomerCode AS HD_CUST_CODE,HD.InvoiceType AS HD_INV_TYPE, " & _
     "HD.InvoiceNo AS HD_INV_NO,HD.InvoiceDate AS HD_INV_DATE,HD.InvoiceAmt AS HD_INV_AMT,HD.AmountToPay AS HD_AMT_TO_PAY,HD.AmountPaid AS HD_AMT_PAID,HD_DET.Acct_Code AS DET_ACCT_CODE, " & _
     "HD_DET.Acct_Name AS DET_ACCT_NAME,HD_DET.Debit AS DET_DEBIT FROM AMIS_Journal_HD HD LEFT OUTER JOIN AMIS_Journal_Det HD_DET ON HD.VoucherNo = HD_DET.VoucherNo AND HD.JType = HD_DET.JType " & _
     "WHERE HD_DET.Acct_Code = '11-02017-00' AND HD.JType IN('SJ','CDJ','COB','CRJ','APJ','CCM','GJ','VCJ','VDJ') AND HD.jdate <= " & N2Str2Null(dtprocess) & " AND HD.status ='P' AND (AR_BALANCE <> 0 OR AR_BALANCE IS NULL) ORDER BY HD.VoucherNo", gconDMIS, adOpenKeyset

    If rsTRANS_SLS_JOURNAL.RecordCount = 0 Then Exit Sub

    If Not rsTRANS_SLS_JOURNAL.EOF And Not rsTRANS_SLS_JOURNAL.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsTRANS_SLS_JOURNAL.RecordCount
        Label11.Caption = "Processing SLS Journal....."

        Do While Not rsTRANS_SLS_JOURNAL.EOF
            'FOR DEBUGGING PURPOSES
            'If rsTRANS_SLS_JOURNAL!HD_VOUCHERNO = "000199" And RTrim(LTrim(rsTRANS_SLS_JOURNAL!HD_JTYPE)) = "COB" Then Stop


            xVOUCHERNO = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_VOUCHERNO)
            xJdate = N2Date2Null(rsTRANS_SLS_JOURNAL!HD_JDATE)
            xSTATUS = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_STATUS)
            xJType = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_JTYPE)

            If Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "CDJ" Or Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "APJ" Then
                XCustomerCode = N2Str2Null(rsTRANS_SLS_JOURNAL!VEN_CODE)
                xInvoiceNo = N2Str2Null(rsTRANS_SLS_JOURNAL!CDJ_NO)
            Else
                XCustomerCode = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_CUST_CODE)
                xInvoiceNo = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_INV_NO)
            End If

            xInvoiceType = N2Str2Null(rsTRANS_SLS_JOURNAL!HD_INV_TYPE)
            xInvoicedate = N2Date2Null(rsTRANS_SLS_JOURNAL!HD_INV_DATE)
            xINVOICE_AMT = NumericVal(rsTRANS_SLS_JOURNAL!HD_INV_AMT)

            If Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) = "COB" Then
                xAMOUNT_TO_PAY = NumericVal(rsTRANS_SLS_JOURNAL!HD_INV_AMT)
            Else
                xAMOUNT_TO_PAY = NumericVal(rsTRANS_SLS_JOURNAL!HD_AMT_TO_PAY)
            End If

            xAMOUNT_PAID = NumericVal(rsTRANS_SLS_JOURNAL!HD_AMT_PAID)
            xACCT_CODE = N2Str2Null(rsTRANS_SLS_JOURNAL!DET_ACCT_CODE)
            xACCNT_NAME = N2Str2Null(rsTRANS_SLS_JOURNAL!DET_ACCT_NAME)
            xdebit = NumericVal(rsTRANS_SLS_JOURNAL!DET_DEBIT)

            gconDMIS.Execute "Insert into AMIS_AR_HD(VoucherNo,Jdate,Status,JType,SJ_CustomerCode,InvoiceType,InvoiceNo,InvoiceDate,InvoiceAmnt,AmountToPay,Acct_code,Debit)" & _
                             "VALUES(" & xVOUCHERNO & "," & xJdate & "," & xSTATUS & "," & xJType & "," & XCustomerCode & "," & xInvoiceType & "," & xInvoiceNo & "," & xInvoicedate & "," & xINVOICE_AMT & "," & xAMOUNT_TO_PAY & "," & xACCT_CODE & "," & xdebit & ")"

            Label12.Caption = Null2String(rsTRANS_SLS_JOURNAL!HD_JTYPE) & "-" & Null2String(rsTRANS_SLS_JOURNAL!HD_VOUCHERNO)
            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            DoEvents
            rsTRANS_SLS_JOURNAL.MoveNext
        Loop
    End If
    Set rsTRANS_SLS_JOURNAL = Nothing
End Sub

Sub AR_COMPUTE()
'DESCRIPTION: COMPUTE THE AR FROM SJ BEING DEDUCTED BY THE CRJ PAYMENT AND COMPUTE IF THERE IS AN ADJUSTMENT
    Dim rsAR_COMPUTE                                        As ADODB.Recordset
    Dim xSJ_VOUCHERNO                                       As String
    Dim xCRJ_VOUCHERNO                                      As String
    Dim xInvoiceType                                        As String
    Dim xInvoiceNo                                          As String
    Dim xAMOUNT_TOPAY                                       As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xBALANCE                                            As Double
    Dim xACCT_CODE                                          As String
    Dim xSYSTEM_REMARKS                                     As String
    Dim xInvoicedate                                        As String
    Dim xLASTUPDATED                                        As String
    Dim xACCT_NAME                                          As String
    Dim xCHECKER                                            As Integer

    xCHECKER = 0

    gconDMIS.Execute "TRUNCATE TABLE AMIS_AR"
    gconDMIS.Execute "TRUNCATE TABLE AMIS_DETAIL"

    Set rsAR_COMPUTE = New ADODB.Recordset
    rsAR_COMPUTE.Open "SELECT DISTINCT VOUCHERNO,JTYPE,INVOICETYPE,INVOICENO,SJ_CUSTOMERCODE,AMOUNTTOPAY,ACCT_CODE,INVOICEDATE,DEBIT,JDATE FROM AMIS_AR_HD ORDER BY VOUCHERNO ASC", gconDMIS, adOpenKeyset

    'FOR DEBUGGING PURPOSES ONLY
    'rsAR_COMPUTE.Open "SELECT DISTINCT VOUCHERNO,JTYPE,INVOICETYPE,INVOICENO,SJ_CUSTOMERCODE,AMOUNTTOPAY,ACCT_CODE,INVOICEDATE,JDATE FROM AMIS_AR_HD WHERE VOUCHERNO = '000170' AND JTYPE = 'COB' ORDER BY VOUCHERNO ASC", gconDMIS, adOpenKeyset

    If rsAR_COMPUTE.RecordCount = 0 Then Exit Sub

    If Not rsAR_COMPUTE.EOF And Not rsAR_COMPUTE.BOF Then
        ProgressBar2.Value = 0
        ProgressBar2.Max = rsAR_COMPUTE.RecordCount
        Label11.Caption = "Processing AR... Please Wait.."
        Do While Not rsAR_COMPUTE.EOF
            xSJ_VOUCHERNO = Null2String(LTrim(RTrim(rsAR_COMPUTE!JTYPE))) & "-" & Null2String(LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)))

            'If xSJ_VOUCHERNO = "CDJ-000918" Then Stop

            Dim rsJNO                                       As ADODB.Recordset
            Set rsJNO = New ADODB.Recordset
            rsJNO.Open "Select VOUCHERNO,JTYPE,JNO,DEBIT,CREDIT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!JTYPE & "' AND (CUSTOMERCODE = '" & rsAR_COMPUTE!SJ_CustomerCode & "' OR VENDORCODE = '" & rsAR_COMPUTE!SJ_CustomerCode & "') AND " & _
                       "JNO IN(SELECT JNO FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!JTYPE & "')", gconDMIS, adOpenKeyset

            If Not rsJNO.EOF And Not rsJNO.BOF Then

                'FOR CRJ AR ACCOUNT
                xSJ_CustomerCode = Null2String(rsAR_COMPUTE!SJ_CustomerCode)

                If Null2String(rsAR_COMPUTE!JTYPE) = "CRJ" And Null2String(rsAR_COMPUTE!ACCT_CODE) = "11-02002-00" Then
                    Dim rsCRJ_AR                            As ADODB.Recordset
                    Set rsCRJ_AR = New ADODB.Recordset
                    rsCRJ_AR.Open "Select InvoiceNo,Invoicetype from AMIS_CRJ_DETAIL WHERE VOUCHERNO = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' ", gconDMIS, adOpenKeyset
                    If Not rsCRJ_AR.EOF And Not rsCRJ_AR.BOF Then
                        xInvoiceNo = Null2String(rsCRJ_AR!INVOICENO)
                        xInvoiceType = Null2String(rsCRJ_AR!InvoiceType)
                        xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(xSJ_CustomerCode)))
                        xAMOUNT_TOPAY = GET_CRJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                        If NumericVal(xAMOUNT_TOPAY) <> 0 Then
                            xAMOUNT_PAID = GET_CRJ_AR_AMT_PAID(xInvoiceNo, xInvoiceType, "C00067", Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDATE))
                        Else
                            'THIS IS CRJ WITH NO LINK
                            Dim rsNOLINK                    As ADODB.Recordset
                            Set rsNOLINK = New ADODB.Recordset
                            rsNOLINK.Open "Select INVOICENO,INVOICETYPE FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
                            If Not rsNOLINK.EOF And Not rsNOLINK.BOF Then
                            Else
                                xAMOUNT_PAID = GET_CRJ_AR_AMT_PAID(xInvoiceNo, xInvoiceType, "C00067", Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDATE))
                            End If
                            Set rsNOLINK = Nothing
                        End If
                    Else
                        xAMOUNT_TOPAY = GET_CRJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                        xAMOUNT_PAID = GET_CRJ_AR_AMT_PAID(xInvoiceNo, xInvoiceType, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDATE))
                    End If
                    Set rsCRJ_AR = Nothing
                Else
                    xInvoiceType = Null2String(rsAR_COMPUTE!InvoiceType)
                    xInvoiceNo = Null2String(rsAR_COMPUTE!INVOICENO)


                    '                                If Null2String(rsAR_COMPUTE!jtype) = "CRJ" Then
                    '                                'THIS IS FOR CRJ PAYMENT WHICH IS INVOICE NO AND INVOICE TYPE IS IN SJ BUT THE ACCT_CODE IS NOT EQUAL
                    '                                    Dim rsCODE_NOLINK As ADODB.Recordset
                    '                                    Set rsCODE_NOLINK = New ADODB.Recordset
                    '                                        rsCODE_NOLINK.Open "SELECT X.INV,X.ACCT_CODE,X.INV_AMT,X.I_NO,X.I_TYPE FROM " & _
                                                             '                                                           "(SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO,CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, CRJ.INVOICEAMOUNT AS INV_AMT FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO " & _
                                                             '                                                           "AND CRJ.CR_TYPE = DET.JTYPE  WHERE DET.ACCT_CODE = '" & rsAR_COMPUTE!Acct_Code & "' AND DET.JDATE <= '" & dtprocess & "' AND DET.STATUS = 'P' AND DET.VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND DET.JTYPE = '" & rsAR_COMPUTE!jtype & "' " & _
                                                             '                                                           ") X WHERE X.INV IN (SELECT INVOICETYPE + '-'+ INVOICENO FROM AMIS_VW_VLEDGER WHERE JDATE <= '" & dtprocess & "' AND ACCT_CODE = '" & rsAR_COMPUTE!Acct_Code & "' AND X.ACCT_CODE = ACCT_CODE ) ", gconDMIS, adOpenKeyset
                    '
                    '                                        If Not rsCODE_NOLINK.EOF And Not rsCODE_NOLINK.BOF Then
                    '                                        Else
                    '                                            xAMOUNT_TOPAY = GET_CRJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), Null2String(rsAR_COMPUTE!Acct_Code))
                    '                                            Dim rsGET_INV As ADODB.Recordset
                    '                                                Set rsGET_INV = New ADODB.Recordset
                    '                                                'rsGET_INV.Open "Select CRJ.INVOICENO,CRJ.INVOICETYPE, SUM(CRJ.INVOICEAMOUNT)AS X from Amis_Crj_detail INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO where CRJ.VOUCHERNO = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' AND CRJ.CR_TYPE = 'CRJ' AND DET.ACCT_CODE = '" & rsAR_COMPUTE!Acct_Code & "'", gconDMIS, adOpenKeyset
                    '                                                rsGET_INV.Open "Select CRJ.INVOICENO,CRJ.INVOICETYPE, SUM(CRJ.INVOICEAMOUNT)AS INV_AMOUNT from Amis_Crj_detail CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.J_CLASS = DET.ACCT_CODE where CRJ.VOUCHERNO = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' AND CRJ.CR_TYPE = 'CRJ' AND DET.ACCT_CODE ='" & rsAR_COMPUTE!Acct_Code & "' GROUP BY CRJ.INVOICENO,CRJ.INVOICETYPE", gconDMIS, adOpenKeyset
                    '                                                    If Not rsGET_INV.EOF And Not rsGET_INV.BOF Then
                    '                                                        'DUE TO WRONG FORMAT OF INVOICETYPE AND INVOICENO REVALIDATE THE INVOICENO AND INVOICETYPE IF FOUND IN HEADER
                    '                                                        xREMARKS = "INV ACCT CODE"
                    '                                                        xINVOICETYPE = Null2String(rsGET_INV!InvoiceType)
                    '                                                        xINVOICENO = Null2String(rsGET_INV!INVOICENO)
                    '
                    '                                                        Dim XINV_TYPE   As String
                    '                                                        Dim XINV_FORMAT As String
                    '                                                        Dim rsX_INVALID As ADODB.Recordset
                    '
                    '                                                        If xINVOICETYPE = "VEHICLE INVOICE" Then
                    '                                                            XINV_TYPE = "VI"
                    '                                                        ElseIf xINVOICETYPE = "SERVICE INVOICE" Then
                    '                                                            XINV_TYPE = "SI"
                    '                                                        ElseIf xINVOICETYPE = "SI" Then
                    '                                                            XINV_TYPE = "SI"
                    '                                                        ElseIf xINVOICETYPE = "VI" Then
                    '                                                            XINV_TYPE = "VI"
                    '                                                        End If
                    '
                    '                                                        If IsNumeric(xINVOICENO) = True Then
                    '                                                            XINV_FORMAT = Format(xINVOICENO, "000000")
                    '                                                            Set rsX_INVALID = New ADODB.Recordset
                    '                                                            rsX_INVALID.Open "SELECT * FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE (HD.INVOICENO = '" & Abs(xINVOICENO) & "' or HD.INVOICENO = '" & XINV_FORMAT & "') AND " & _
                                                                                 '                                                                             "(HD.INVOICETYPE = '" & xINVOICETYPE & "' OR HD.INVOICETYPE = '" & XINV_TYPE & "') AND HD.JDATE <= '" & dtprocess & "' AND DET.ACCT_CODE = '" & Null2String(rsAR_COMPUTE!Acct_Code) & "'", gconDMIS, adOpenKeyset
                    '                                                        Else
                    '                                                            XINV_FORMAT = xINVOICENO
                    '                                                            rsX_INVALID.Open "SELECT * FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.INVOICENO = '" & XINV_FORMAT & "' AND " & _
                                                                                 '                                                                             "(HD.INVOICETYPE = '" & xINVOICETYPE & "' OR HD.INVOICETYPE = '" & XINV_TYPE & "') AND HD.JDATE <= '" & dtprocess & "' AND DET.ACCT_CODE = '" & Null2String(rsAR_COMPUTE!Acct_Code) & "'", gconDMIS, adOpenKeyset
                    '                                                        End If
                    '
                    '                                                        If Not rsX_INVALID.EOF And Not rsX_INVALID.BOF Then
                    '                                                        Else
                    '                                                            xAMOUNT_PAID = NumericVal(rsGET_INV!INV_AMOUNT)
                    '                                                            xCHECKER = 1
                    '                                                        End If
                    '                                                        Set rsX_INVALID = Nothing
                    '                                                    End If
                    '                                            Set rsGET_INV = Nothing
                    '                                        End If
                    '                                    Set rsCODE_NOLINK = Nothing

                    'TEMPORARY COMMENTED -----------------------------------------------------------------------10-14-2009------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    'THIS FOR INVALID INVOICES
                    Dim rsINVLID_INVOICE                    As ADODB.Recordset
                    Dim INV_TEMP                            As String
                    Dim rsINVALID                           As ADODB.Recordset
                    Dim INV_FROMAT                          As String

                    Set rsINVLID_INVOICE = New ADODB.Recordset
                    'rsINVLID_INVOICE.Open "SELECT RTRIM(LTRIM(X.INV_NO_TYPE)),X.INV_AMT FROM " & _
                     "(SELECT DET.INVOICETYPE + '-' + DET.INVOICENO AS INV_NO_TYPE,DET.INVOICEAMOUNT AS INV_AMT FROM AMIS_AR_HD HD INNER JOIN AMIS_CRJ_DETAIL DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE  = DET.CR_TYPE  WHERE HD.VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!jtype & "'" & _
                     ") X WHERE X.INV_NO_TYPE IN(SELECT INVOICETYPE + '-' + INVOICENO  FROM AMIS_JOURNAL_HD WHERE JDATE <= '" & dtprocess & "')", gconDMIS, adOpenKeyset


                    rsINVLID_INVOICE.Open "SELECT HD.SJ_CUSTOMERCODE AS CUS_CODE, DET.INVOICETYPE AS INV_TYPE,DET.INVOICENO AS INV_NO,DET.INVOICEAMOUNT AS INV_AMT FROM AMIS_AR_HD HD INNER JOIN AMIS_CRJ_DETAIL DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE  = DET.CR_TYPE  WHERE HD.VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND JTYPE = '" & rsAR_COMPUTE!JTYPE & "'", gconDMIS, adOpenKeyset

                    If Not rsINVLID_INVOICE.EOF And Not rsINVLID_INVOICE.BOF Then
                        If Null2String(rsINVLID_INVOICE!INV_TYPE) = "VEHICLE INVOICE" Then
                            INV_TEMP = "VI"
                        ElseIf Null2String(rsINVLID_INVOICE!INV_TYPE) = "SERVICE INVOICE" Then
                            INV_TEMP = "SI"
                        ElseIf Null2String(rsINVLID_INVOICE!INV_TYPE) = "SI" Then
                            INV_TEMP = "SI"
                        ElseIf Null2String(rsINVLID_INVOICE!INV_TYPE) = "VI" Then
                            INV_TEMP = "VI"
                        End If
                        Set rsINVALID = New ADODB.Recordset

                        If IsNumeric(rsINVLID_INVOICE!INV_NO) = True Then
                            INV_FROMAT = Format(Null2String(rsINVLID_INVOICE!INV_NO), "000000")
                            rsINVALID.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE (INVOICENO = '" & Abs(rsINVLID_INVOICE!INV_NO) & "' or INVOICENO = '" & INV_FROMAT & "') AND (INVOICETYPE = '" & Null2String(rsINVLID_INVOICE!INV_TYPE) & "' OR INVOICETYPE = '" & INV_TEMP & "') AND JDATE <= '" & dtprocess & "' AND STATUS = 'P' AND CUSTOMERCODE = '" & Null2String(rsINVLID_INVOICE!CUS_CODE) & "'", gconDMIS, adOpenKeyset
                        Else
                            INV_FROMAT = Null2String(rsINVLID_INVOICE!INV_NO)
                            rsINVALID.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & rsINVLID_INVOICE!INV_NO & "' AND (INVOICETYPE = '" & Null2String(rsINVLID_INVOICE!INV_TYPE) & "' OR INVOICETYPE = '" & INV_TEMP & "') AND JDATE <= '" & dtprocess & "' AND STATUS = 'P' AND CUSTOMERCODE = '" & Null2String(rsINVLID_INVOICE!CUS_CODE) & "'", gconDMIS, adOpenKeyset
                        End If
                        If Not rsINVALID.EOF And Not rsINVALID.BOF Then
                            Dim INV_VOUCHERNO               As String
                            Dim INV_JTYPE                   As String
                            INV_VOUCHERNO = Null2String(RTrim(LTrim(rsINVALID!VOUCHERNO)))
                            INV_JTYPE = Null2String(RTrim(LTrim(rsINVALID!JTYPE)))
                        Else
                            'THIS IS ONLY FOR VALIDATING IF AMIS_JOURNAL_DET HAS NO AR ACCOUNT CODE OR GOOD AS CASH
                            Dim rsNO_AR_CODE                As ADODB.Recordset
                            Set rsNO_AR_CODE = New ADODB.Recordset
                            rsNO_AR_CODE.Open "SELECT * FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.INVOICENO = '" & INV_FROMAT & "' AND HD.INVOICETYPE = '" & INV_TEMP & "' AND HD.CUSTOMERCODE = '" & Null2String(rsINVLID_INVOICE!CUS_CODE) & "' AND LEFT(ACCT_CODE,5) = '11-02'", gconDMIS, adOpenKeyset
                            If Not rsNO_AR_CODE.EOF And Not rsNO_AR_CODE.BOF Then
                            Else
                                'VALIDATE IF THERE IS AR ACCOUNT CODE IN CRJ <> TO SJ
                                Dim rsWRONG_ENTRY           As ADODB.Recordset
                                Set rsWRONG_ENTRY = New ADODB.Recordset
                                rsWRONG_ENTRY.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM ( " & _
                                                   "SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO, " & _
                                                   "CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE FROM AMIS_CRJ_DETAIL CRJ " & _
                                                   "INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE INNER JOIN AMIS_JOURNAL_HD HD " & _
                                                   "ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE WHERE HD.VOUCHERNO = " & N2Str2Null(rsAR_COMPUTE!VOUCHERNO) & " AND LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' " & _
                                                   "AND DET.STATUS = 'P' AND DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0) X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO FROM " & _
                                                   "AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                                   "WHERE X.ACCT_CODE <> DET.ACCT_CODE AND HD.CUSTOMERCODE <> X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') AND (DET.DEBIT <> 0 OR DET.DEBIT IS NULL) )", gconDMIS, adOpenKeyset
                                If Not rsWRONG_ENTRY.EOF And Not rsWRONG_ENTRY.BOF Then
                                Else
                                    xAMOUNT_TOPAY = GET_CRJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                                    Dim rsTEMP              As ADODB.Recordset
                                    Set rsTEMP = New ADODB.Recordset
                                    rsTEMP.Open "SELECT INVOICENO,INVOICETYPE,INVOICEAMOUNT FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO = '" & rsAR_COMPUTE!VOUCHERNO & "' AND CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
                                    If Not rsTEMP.EOF And Not rsTEMP.BOF Then
                                        xAMOUNT_PAID = NumericVal(rsTEMP!invoiceamount)
                                        xInvoiceType = Null2String(rsTEMP!InvoiceType)
                                        xInvoiceNo = Null2String(rsTEMP!INVOICENO)
                                        xSJ_CustomerCode = "XXXXXX"
                                        xACCT_NAME = "XXXXXX"
                                        'xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(Null2String(rsINVLID_INVOICE!CUS_CODE))))
                                        xCHECKER = 1
                                    End If
                                    Set rsTEMP = Nothing
                                End If
                            End If
                            Set rsNO_AR_CODE = Nothing
                        End If
                    End If
                    Set rsINVLID_INVOICE = Nothing
                    'TEMPORARY COMMENTED -----------------------------------------------------------------------10-14-2009------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                End If
            End If

            If Null2String(rsAR_COMPUTE!JTYPE) = "CDJ" Then
                'xAMOUNT_TOPAY = GET_AR_CDJ_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!jtype), xSJ_CustomerCode)
                xAMOUNT_TOPAY = GET_AR_CDJ_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_PAID = GET_AR_CDJ_PAYENT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!SJ_CustomerCode), Null2String(rsAR_COMPUTE!INVOICENO))
                xACCT_NAME = GET_VEN_NAME(LTrim(RTrim(xSJ_CustomerCode)))

                If NumericVal(xAMOUNT_TOPAY) = 0 Then
                    xAMOUNT_TOPAY = GET_CDJ_DEBIT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), xSJ_CustomerCode, Null2String(rsAR_COMPUTE!INVOICENO))
                    xAMOUNT_PAID = GET_CDJ_CREDIT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                End If
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "SJ" Then
                xAMOUNT_TOPAY = GET_SJ_DEBIT_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), xInvoiceNo, xInvoiceType, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_PAID = COMP_AMT_PAID(xInvoiceNo, xInvoiceType, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDATE))
                xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(xSJ_CustomerCode)))
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "COB" Then
                xAMOUNT_TOPAY = GET_COB_AMOUNT(rsAR_COMPUTE!VOUCHERNO, rsAR_COMPUTE!JTYPE, rsAR_COMPUTE!SJ_CustomerCode)
                xAMOUNT_PAID = COMP_AMT_PAID(xInvoiceNo, xInvoiceType, xSJ_CustomerCode, Null2String(rsAR_COMPUTE!ACCT_CODE), Null2String(rsAR_COMPUTE!VOUCHERNO), Null2Date(rsAR_COMPUTE!JDATE))
                xACCT_NAME = GET_CUST_NAME(LTrim(RTrim(xSJ_CustomerCode)))
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "APJ" Then
                xAMOUNT_TOPAY = GET_APJ_AR(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_PAID = GET_APJ_CREDIT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xACCT_NAME = GET_VEN_NAME(LTrim(RTrim(xSJ_CustomerCode)))
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "CCM" Then
                xAMOUNT_PAID = GET_CCM_AMOUNT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xACCT_NAME = "XXXXXX"
                xSJ_CustomerCode = "XXXXXX"
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "GJ" Then
                xAMOUNT_TOPAY = GET_GJ_DEBIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_PAID = GET_GJ_CREDIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xACCT_NAME = "XXXXXX"
                xSJ_CustomerCode = "XXXXXX"
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "VCJ" Then
                xAMOUNT_TOPAY = GET_VCJ_DEBIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_PAID = GET_VCJ_CREDIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
            ElseIf Null2String(rsAR_COMPUTE!JTYPE) = "VDJ" Then
                xAMOUNT_PAID = GET_VDJ_CREDIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xAMOUNT_TOPAY = GET_VDJ_DEBIT_AMT(Null2String(rsAR_COMPUTE!VOUCHERNO), Null2String(rsAR_COMPUTE!JTYPE), Null2String(rsAR_COMPUTE!ACCT_CODE))
                xACCT_NAME = "XXXXXX"
                xSJ_CustomerCode = "XXXXXX"
            End If


            xBALANCE = Round(NumericVal(xAMOUNT_TOPAY) - NumericVal(xAMOUNT_PAID), 2)
            xACCT_CODE = Null2String(rsAR_COMPUTE!ACCT_CODE)
            xInvoicedate = Null2Date(rsAR_COMPUTE!invoicedate)
            xLASTUPDATED = LOGDATE

            'CHECK IF VOUCHERNO AND JTYPE IS ALREADY EXISTING IN AMIS_AR
            Dim rsVOUCHERNO_IN_AR                           As ADODB.Recordset
            Set rsVOUCHERNO_IN_AR = New ADODB.Recordset

            If Null2String(rsAR_COMPUTE!JTYPE) = "CDJ" Then
                rsVOUCHERNO_IN_AR.Open "Select * from Amis_Ar where SJVOUCHERNO = '" & xSJ_VOUCHERNO & "' and ACCOUNT_CODE = '" & xACCT_CODE & "' AND SYSTEMREMARK IS NULL", gconDMIS, adOpenKeyset
            Else
                rsVOUCHERNO_IN_AR.Open "Select * from Amis_Ar where SJVOUCHERNO = '" & xSJ_VOUCHERNO & "' and ACCOUNT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
            End If
            'rsVOUCHERNO_IN_AR.Open "Select * from Amis_Ar where SJVOUCHERNO = '" & xSJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
            If Not rsVOUCHERNO_IN_AR.EOF And Not rsVOUCHERNO_IN_AR.BOF Then
            Else
                If NumericVal(xAMOUNT_TOPAY) = 0 And NumericVal(xAMOUNT_PAID) = 0 And NumericVal(xBALANCE) = 0 Then
                    'DONT INSERT
                Else
                    gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                     "VALUES('" & xSJ_VOUCHERNO & "','" & xInvoiceType & "','" & xInvoiceNo & "','" & xSJ_CustomerCode & "','" & xACCT_NAME & "'," & xAMOUNT_TOPAY & "," & xAMOUNT_PAID & ",'" & xBALANCE & "','" & xACCT_CODE & "'," & N2Str2Null("") & ",'" & xInvoicedate & "','" & xLASTUPDATED & "')"
                    'THIS IS FOR AMIS_AR DETAIL

                    '                                        If NumericVal(xCHECKER) = 1 Then
                    '                                            If CHECK_DUPLICATE(N2Str2Null(INV_FROMAT), N2Str2Null(INV_TEMP), N2Str2Null(LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)))) = False Then
                    '                                                gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                                                                     '                                                                 "VALUES(" & N2Str2Null(INV_FROMAT) & ", " & N2Str2Null(INV_TEMP) & ", " & NumericVal(xAMOUNT_PAID) & ", " & N2Str2Null(xSJ_CustomerCode) & "," & N2Str2Null(rsAR_COMPUTE!Acct_Code) & "," & N2Date2Null(rsAR_COMPUTE!JDate) & ",'1','" & LTrim(RTrim(rsAR_COMPUTE!VOUCHERNO)) & "')"
                    '                                            End If
                    '                                        End If
                End If
            End If
            Set rsVOUCHERNO_IN_AR = Nothing
            '
            '                            If NumericVal(xBALANCE) = 0 Then
            '                                gconDMIS.Execute "Update Amis_Journal_Hd set AR_DATEGEN = '" & xLASTUPDATED & "', AR_BALANCE = '" & xBALANCE & "' where Voucherno = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' and JTYPE = '" & Null2String(rsAR_COMPUTE!jtype) & "'"
            '                            Else
            '                                gconDMIS.Execute "Update Amis_Journal_Hd set AR_DATEGEN = '" & xLASTUPDATED & "', AR_BALANCE = '" & xBALANCE & "' where Voucherno = '" & Null2String(rsAR_COMPUTE!VOUCHERNO) & "' and JTYPE = '" & Null2String(rsAR_COMPUTE!jtype) & "'"
            '                            End If
            'End If
            Set rsJNO = Nothing


            xAMOUNT_TOPAY = 0
            xAMOUNT_PAID = 0
            xBALANCE = 0
            xCHECKER = 0


            ProgressBar2.Value = ProgressBar2.Value + 1
            labPercent.Caption = Round((ProgressBar2.Value / ProgressBar2.Max) * 100, 0) & "%"
            Label12.Caption = xSJ_VOUCHERNO
            DoEvents
            rsAR_COMPUTE.MoveNext
        Loop
    End If

    Set rsAR_COMPUTE = Nothing
    'Set rsCHECK_IN_AR = Nothing
End Sub
Function GET_VDJ_CREDIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsGET_VDJ_CREDIT_AMT                                As ADODB.Recordset
    Set rsGET_VDJ_CREDIT_AMT = New ADODB.Recordset
    rsGET_VDJ_CREDIT_AMT.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_VDJ_CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_VDJ_CREDIT_AMT.EOF And Not rsGET_VDJ_CREDIT_AMT.BOF Then
        GET_VDJ_CREDIT_AMT = NumericVal(rsGET_VDJ_CREDIT_AMT!SUM_VDJ_CREDIT)
    End If
    Set rsGET_VDJ_CREDIT_AMT = Nothing
End Function

Function GET_VDJ_DEBIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsGET_VDJ_DEBIT_AMT                                 As ADODB.Recordset
    Set rsGET_VDJ_DEBIT_AMT = New ADODB.Recordset
    rsGET_VDJ_DEBIT_AMT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_VDJ_DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_VDJ_DEBIT_AMT.EOF And Not rsGET_VDJ_DEBIT_AMT.BOF Then
        GET_VDJ_DEBIT_AMT = NumericVal(rsGET_VDJ_DEBIT_AMT!SUM_VDJ_DEBIT)
    End If
    Set rsGET_VDJ_DEBIT_AMT = Nothing
End Function

Function GET_VCJ_CREDIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE VCJ CREDIT AMOUNT
    Dim rsGET_VCJ_CREDIT_AMT                                As ADODB.Recordset
    Set rsGET_VCJ_CREDIT_AMT = New ADODB.Recordset
    rsGET_VCJ_CREDIT_AMT.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_VCJ_CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_VCJ_CREDIT_AMT.EOF And Not rsGET_VCJ_CREDIT_AMT.BOF Then
        GET_VCJ_CREDIT_AMT = NumericVal(rsGET_VCJ_CREDIT_AMT!SUM_VCJ_CREDIT)
    End If
    Set rsGET_VCJ_CREDIT_AMT = Nothing
End Function
Function GET_VCJ_DEBIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE VCJ DEBIT AMOUNT
    Dim rsGET_VCJ_DEBIT_AMT                                 As ADODB.Recordset
    Set rsGET_VCJ_DEBIT_AMT = New ADODB.Recordset
    rsGET_VCJ_DEBIT_AMT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_VCJ_DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_VCJ_DEBIT_AMT.EOF And Not rsGET_VCJ_DEBIT_AMT.BOF Then
        GET_VCJ_DEBIT_AMT = NumericVal(rsGET_VCJ_DEBIT_AMT!SUM_VCJ_DEBIT)
    End If
    Set rsGET_VCJ_DEBIT_AMT = Nothing
End Function

Function GET_GJ_DEBIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'GETTING THE DEBIT AMOUNT OF GJ
    Dim rsGET_GJ_DEBIT_AMT                                  As ADODB.Recordset
    Set rsGET_GJ_DEBIT_AMT = New ADODB.Recordset
    rsGET_GJ_DEBIT_AMT.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_DEBIT_AMT.EOF And Not rsGET_GJ_DEBIT_AMT.BOF Then
        GET_GJ_DEBIT_AMT = NumericVal(rsGET_GJ_DEBIT_AMT!SUM_GJ_DEBIT)
    End If
    Set rsGET_GJ_DEBIT_AMT = Nothing
End Function
Function GET_GJ_CREDIT_AMT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'GETTING THE CREDIT AMOUNT OF GJ
    Dim rsGET_GJ_CREDIT_AMT                                 As ADODB.Recordset
    Set rsGET_GJ_CREDIT_AMT = New ADODB.Recordset
    rsGET_GJ_CREDIT_AMT.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_GJ_CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND STATUS = 'P' AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_CREDIT_AMT.EOF And Not rsGET_GJ_CREDIT_AMT.BOF Then
        GET_GJ_CREDIT_AMT = NumericVal(rsGET_GJ_CREDIT_AMT!SUM_GJ_CREDIT)
    End If
    Set rsGET_GJ_CREDIT_AMT = Nothing
End Function

Function GET_CDJ_DEBIT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xInvoiceNo As String) As Double
'DESCRIPTION: GET THE CDJ DEBIT AMOUNT
    Dim rsGET_CDJ_DEBIT                                     As ADODB.Recordset
    Dim sum_AR_ADJ_PAYMENT                                  As Double
    Dim sum_ADJ_PAYMENT                                     As Double
    sum_ADJ_PAYMENT = 0
    Set rsGET_CDJ_DEBIT = New ADODB.Recordset
    rsGET_CDJ_DEBIT.Open "Select DEBIT FROM  AMIS_JOURNAL_DET WHERE  JTYPE = 'GJ' AND RIGHT(ENTITY,6) = '" & xVENDORCODE & "' AND INVOICENO = '" & xVOUCHERNO & "' AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_DEBIT.EOF And Not rsGET_CDJ_DEBIT.BOF Then
        Do While Not rsGET_CDJ_DEBIT.EOF
            sum_AR_ADJ_PAYMENT = Round((sum_AR_ADJ_PAYMENT + NumericVal(rsGET_CDJ_DEBIT!Debit)), 2)
            rsGET_CDJ_DEBIT.MoveNext
        Loop
    End If
    GET_CDJ_DEBIT = NumericVal(sum_ADJ_PAYMENT)
    Set rsGET_CDJ_DEBIT = Nothing
End Function

Function GET_CDJ_CREDIT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRITION: GET THE CDJ CREDIT
    Dim rsGET_CDJ_CREDIT                                    As ADODB.Recordset
    Set rsGET_CDJ_CREDIT = New ADODB.Recordset
    rsGET_CDJ_CREDIT.Open "SELECT ROUND(SUM(CREDIT),2) AS CRJ_CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_CDJ_CREDIT.EOF And Not rsGET_CDJ_CREDIT.BOF Then
        GET_CDJ_CREDIT = NumericVal(rsGET_CDJ_CREDIT!CRJ_CREDIT)
    End If
    Set rsGET_CDJ_CREDIT = Nothing
End Function

Function GET_APJ_CREDIT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE APJ CREDIT AMOUNT
    Dim rsGET_APJ_CREDIT                                    As ADODB.Recordset
    Set rsGET_APJ_CREDIT = New ADODB.Recordset
    rsGET_APJ_CREDIT.Open "SELECT ROUND(SUM(CREDIT),2) AS CREDIT_APJ FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_APJ_CREDIT.EOF And Not rsGET_APJ_CREDIT.BOF Then
        GET_APJ_CREDIT = NumericVal(rsGET_APJ_CREDIT!CREDIT_APJ)
    End If
    Set rsGET_APJ_CREDIT = Nothing
End Function

Function GET_CCM_AMOUNT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE CCM AMOUNT
    Dim rsGET_CCM_AMOUNT                                    As ADODB.Recordset
    Set rsGET_CCM_AMOUNT = New ADODB.Recordset
    rsGET_CCM_AMOUNT.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_CCM FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_CCM_AMOUNT.EOF And Not rsGET_CCM_AMOUNT.BOF Then
        GET_CCM_AMOUNT = NumericVal(rsGET_CCM_AMOUNT!SUM_CCM)
    End If
    Set rsGET_CCM_AMOUNT = Nothing
End Function

Function CHECK_DUPLICATE(xInvoiceNo As String, xInvoiceType As String, xVOUCHERNO As String) As Boolean
'DESCRIPTION: THIS IS TO CHECK IF PAYMENT DETAIL IS EXISTING IN THE AMIS_DETAIL TABLE
'             AMIS_DETAIL TABLE IS USE AS A SUB REPORT IN AR SCHEDULE AGING REPORT
    Dim rsCHECK_DUPLICATE                                   As ADODB.Recordset
    Set rsCHECK_DUPLICATE = New ADODB.Recordset
    rsCHECK_DUPLICATE.Open "Select * from AMIs_detail where INVOICENO = " & xInvoiceNo & " AND INVOICETYPE = " & xInvoiceType & " AND VOUCHERNO = " & xVOUCHERNO & "", gconDMIS, adOpenKeyset
    If Not rsCHECK_DUPLICATE.EOF And Not rsCHECK_DUPLICATE.BOF Then
        CHECK_DUPLICATE = True
    Else
        CHECK_DUPLICATE = False
    End If
End Function

Function GET_APJ_AR(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCTIPION: GET THE APJ DEBIT AMOUNT
    Dim rsGET_APJ_AR                                        As ADODB.Recordset
    Set rsGET_APJ_AR = New ADODB.Recordset
    rsGET_APJ_AR.Open "Select ROUND(SUM(DEBIT),2) AS APJ_AR FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_APJ_AR.EOF And Not rsGET_APJ_AR.BOF Then
        GET_APJ_AR = NumericVal(rsGET_APJ_AR!APJ_AR)
    End If
    Set rsGET_APJ_AR = Nothing
End Function

Function GET_CRJ_AR(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE APJ CREDIT AMOUNT
    Dim rsGET_CRJ_AR                                        As ADODB.Recordset
    Set rsGET_CRJ_AR = New ADODB.Recordset
    'rsGET_CRJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS CRJ_AR FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND  JTYPE = '" & xJTYPE & "' AND STATUS = 'P' and Left(Acct_Code,5) IN('11-02','11-03') and JDATE <= '" & dtprocess & "'", gconDMIS, adOpenKeyset
    rsGET_CRJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS CRJ_AR FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND  JTYPE = '" & xJType & "' AND STATUS = 'P' and ACCT_CODE  = '" & xACCT_CODE & "' and JDATE <= '" & dtprocess & "' AND DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_CRJ_AR.EOF And Not rsGET_CRJ_AR.BOF Then
        GET_CRJ_AR = NumericVal(rsGET_CRJ_AR!CRJ_AR)
    End If
    Set rsGET_CRJ_AR = Nothing
End Function

Function GET_CRJ_AR_AMT_PAID(xInvoiceNo As String, xInvoiceType As String, xCUSCODE As String, xACCT_CODE As String, xVOUCHERNO As String, xJdate As String) As Double
    Dim rsGET_CRJ_AR_AMT_PAID                               As ADODB.Recordset
    Dim rsCUS_CODE                                          As ADODB.Recordset
    Dim rsDOUBLE_INVOICE                                    As ADODB.Recordset
    Dim sumCRJ_AMNT                                         As Double
    sumCRJ_AMNT = 0
    Set rsGET_CRJ_AR_AMT_PAID = New ADODB.Recordset
    Set rsDOUBLE_INVOICE = New ADODB.Recordset
    rsDOUBLE_INVOICE.Open "SELECT X.INV_INVTYPE FROM  ( " & _
                          "SELECT HD.INVOICETYPE + '-' + HD.INVOICENO AS INV_INVTYPE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON HD.VOUCHERNO = DET.VOUCHERNO  AND HD.JTYPE = DET.JTYPE  WHERE DET.ACCT_CODE = '11-02002-00' AND HD.INVOICENO = '" & xInvoiceNo & "' AND HD.INVOICETYPE = '" & xInvoiceType & "' AND HD.STATUS = 'P' " & _
                          ") X WHERE X.INV_INVTYPE IN (SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET " & _
                          "ON CRJ.VOUCHERNO = DET.VOUCHERNO  AND CRJ.CR_TYPE = DET.JTYPE WHERE DET.ACCT_CODE = '11-02002-00' AND DET.STATUS = 'P')", gconDMIS, adOpenKeyset
    If Not rsDOUBLE_INVOICE.EOF And Not rsDOUBLE_INVOICE.BOF Then
        rsGET_CRJ_AR_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where InvoiceType = '" & xInvoiceType & "' and InvoiceNo = '" & xInvoiceNo & "' and CR_TYPE = 'CRJ' AND SJ_VOUCHERNO = '" & xVOUCHERNO & "'", gconDMIS, adOpenKeyset
    Else
        rsGET_CRJ_AR_AMT_PAID.Open "Select VOUCHERNO,INVOICEAMOUNT,INVOICENO,INVOICETYPE from Amis_Crj_detail where INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "'", gconDMIS, adOpenKeyset
    End If

    If Not rsGET_CRJ_AR_AMT_PAID.EOF And Not rsGET_CRJ_AR_AMT_PAID.BOF Then
        Do While Not rsGET_CRJ_AR_AMT_PAID.EOF
            Set rsCUS_CODE = New ADODB.Recordset
            rsCUS_CODE.Open "SELECT HD.CUSTOMERCODE,HD.VOUCHERNO AS VOUCHERNO  FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO INNER JOIN AMIS_CRJ_DETAIL CRJ  ON CRJ.VOUCHERNO = HD.VOUCHERNO WHERE  CRJ.INVOICENO = '" & xInvoiceNo & "' AND CRJ.INVOICETYPE = '" & xInvoiceType & "' AND  HD.VOUCHERNO = '" & rsGET_CRJ_AR_AMT_PAID!VOUCHERNO & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' AND HD.JTYPE = 'CRJ' AND (HD.CUSTOMERCODE = '" & xCUSCODE & "' OR HD.CUSTOMERCODE = 'R00197') AND HD.JDATE <= '" & dtprocess & "' AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
            If Not rsCUS_CODE.EOF And Not rsCUS_CODE.BOF Then
                sumCRJ_AMNT = Round((NumericVal(sumCRJ_AMNT) + NumericVal(rsGET_CRJ_AR_AMT_PAID!invoiceamount)), 2)

                'THIS IS FOR AR PAYMENT DETAIL
                If CHECK_DUPLICATE(N2Str2Null(rsGET_CRJ_AR_AMT_PAID!INVOICENO), N2Str2Null(rsGET_CRJ_AR_AMT_PAID!InvoiceType), N2Str2Null(rsGET_CRJ_AR_AMT_PAID!VOUCHERNO)) = False Then
                    gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,JDATE,REMARKS,VOUCHERNO) " & _
                                     "VALUES(" & N2Str2Null(rsGET_CRJ_AR_AMT_PAID!INVOICENO) & ", " & N2Str2Null(rsGET_CRJ_AR_AMT_PAID!InvoiceType) & ", " & NumericVal(rsGET_CRJ_AR_AMT_PAID!invoiceamount) & ", " & N2Str2Null(xCUSCODE) & "," & N2Date2Null(xJdate) & ",'2','" & RTrim(LTrim(rsGET_CRJ_AR_AMT_PAID!VOUCHERNO)) & "')"

                    'UPDATE THE AMIS_CRJ_DETAIL WITH SJ_VOUCHERNO AND CUSTOMER CODE
                    gconDMIS.Execute "UPDATE AMIS_CRJ_DETAIL SET SJ_VOUCHERNO = '" & xVOUCHERNO & "', CUSTOMERCODE = '" & xCUSCODE & "' where INVOICENO = '" & rsGET_CRJ_AR_AMT_PAID!INVOICENO & "' AND INVOICETYPE = '" & rsGET_CRJ_AR_AMT_PAID!InvoiceType & "' AND VOUCHERNO = '" & rsGET_CRJ_AR_AMT_PAID!VOUCHERNO & "'"
                End If
            End If
            rsGET_CRJ_AR_AMT_PAID.MoveNext
        Loop
    End If
    GET_CRJ_AR_AMT_PAID = NumericVal(sumCRJ_AMNT)
    Set rsGET_CRJ_AR_AMT_PAID = Nothing
End Function

Function GET_COB_AMOUNT(xVOUCHERNO As String, xJType As String, xCUSCDE As String) As Double
    Dim rsGET_COB_AMOUNT                                    As ADODB.Recordset
    Set rsGET_COB_AMOUNT = New ADODB.Recordset
    rsGET_COB_AMOUNT.Open "Select INVOICEAMT FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND CUSTOMERCODE = '" & xCUSCDE & "'  AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsGET_COB_AMOUNT.EOF And Not rsGET_COB_AMOUNT.BOF Then
        GET_COB_AMOUNT = NumericVal(rsGET_COB_AMOUNT!InvoiceAmt)
    End If
    Set rsGET_COB_AMOUNT = Nothing
End Function

Function GET_SJ_DEBIT_AMOUNT(xVOUCHERNO As String, xInvoiceNo As String, xInvoiceType As String, xSJ_CustCode As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE  SJ DEBIT AMOUNT  FROM AMIS_JOURNAL_DET AND CHECK IF THERE IS AN ADJUSTMENT
    Dim rsDEBIT_AMOUNT                                      As ADODB.Recordset
    Dim rsSJ_ADJUSTMENT                                     As ADODB.Recordset
    Dim xSUM_SJ                                             As Double
    Dim xADJ_AMOUNT                                         As Double
    xSUM_SJ = 0
    xADJ_AMOUNT = 0

    Set rsDEBIT_AMOUNT = New ADODB.Recordset
    rsDEBIT_AMOUNT.Open "Select ROUND(SUM(Debit),2) AS DEBIT from Amis_Journal_det where VoucherNo = '" & xVOUCHERNO & "' and JTYPE = 'SJ' and ACCT_CODE = '" & xACCT_CODE & "' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsDEBIT_AMOUNT.EOF And Not rsDEBIT_AMOUNT.BOF Then
        xSUM_SJ = NumericVal(rsDEBIT_AMOUNT!Debit)
    End If

    '        Set rsSJ_ADJUSTMENT = New ADODB.Recordset
    '        rsSJ_ADJUSTMENT.Open "SELECT ADJ_AMOUNT,ADJ_TYPE from AMIS_JOURNAL_HD WHERE InvoiceNo= '" & xInvoiceNo & "' and InvoiceType = '" & xinvoiceType & "' and JTYPE = 'ADJ' and STATUS = 'P'", gconDMIS, adOpenKeyset
    '        If Not rsSJ_ADJUSTMENT.EOF And Not rsSJ_ADJUSTMENT.BOF Then
    '            Do While Not rsSJ_ADJUSTMENT.EOF
    '                If Null2String(rsSJ_ADJUSTMENT!ADJ_TYPE) = "DEBIT" Then
    '                    xADJ_AMOUNT = Round(NumericVal(xADJ_AMOUNT) + NumericVal(rsSJ_ADJUSTMENT!ADJ_AMOUNT), 2)
    '                End If
    '                rsSJ_ADJUSTMENT.MoveNext
    '            Loop
    '        End If

    Set rsSJ_ADJUSTMENT = New ADODB.Recordset
    rsSJ_ADJUSTMENT.Open "SELECT DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND RIGHT(ENTITY,6) = '" & xSJ_CustCode & "' AND Left(Acct_Code,5) IN('11-02','11-03') and ADJ_JTYPE = 'SJ' AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsSJ_ADJUSTMENT.EOF And Not rsSJ_ADJUSTMENT.BOF Then
        Do While Not rsSJ_ADJUSTMENT.EOF
            If NumericVal(rsSJ_ADJUSTMENT!Debit) <> 0 Then
                xADJ_AMOUNT = Round((NumericVal(xADJ_AMOUNT) + NumericVal(rsSJ_ADJUSTMENT!Debit)), 2)
            End If
            rsSJ_ADJUSTMENT.MoveNext
        Loop
    End If
    GET_SJ_DEBIT_AMOUNT = Round((NumericVal(xSUM_SJ) + NumericVal(xADJ_AMOUNT)), 2)
    Set rsDEBIT_AMOUNT = Nothing
End Function

Function GET_AR_CDJ_AMOUNT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xACCT_CODE As String) As Double
'DESCRIPTION: GET THE AR SCHEDULE AMOUNT FROM CASH DISBURSEMENT
    Dim rsGET_AR_CDJ_AMOUNT                                 As ADODB.Recordset
    Dim sumCDJ                                              As Double
    sumCDJ = 0
    Set rsGET_AR_CDJ_AMOUNT = New ADODB.Recordset
    'rsGET_AR_CDJ_AMOUNT.Open "Select DEBIT From Amis_journal_det where Jtype = '" & xJTYPE & "' and VoucherNo = '" & xVOUCHERNO & "' AND Left(Acct_Code,5) IN('11-02','11-03') AND STATUS = 'P' and JDATE <= '" & dtprocess & "'", gconDMIS, adOpenKeyset
    rsGET_AR_CDJ_AMOUNT.Open "Select DEBIT From Amis_journal_det where Jtype = '" & xJType & "' and VoucherNo = '" & xVOUCHERNO & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P' and JDATE <= '" & dtprocess & "' AND DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_AR_CDJ_AMOUNT.EOF And Not rsGET_AR_CDJ_AMOUNT.BOF Then
        Do While Not rsGET_AR_CDJ_AMOUNT.EOF
            sumCDJ = Round((sumCDJ + NumericVal(rsGET_AR_CDJ_AMOUNT!Debit)), 2)
            rsGET_AR_CDJ_AMOUNT.MoveNext
        Loop
    End If
    GET_AR_CDJ_AMOUNT = NumericVal(sumCDJ)
    Set rsGET_AR_CDJ_AMOUNT = Nothing

    If sumCDJ <> 0 Then
        'DESCRIPTION: THIS IS FOR BOTH CREDIT AND DEBIT SIDE HAS AN AR ENTRY
        '             THIS GET THE AR ON THE CREDIT SIDE IF IT HAS AN AR ENTRY

        Dim rsCDJ_CREDIT                                    As ADODB.Recordset
        Dim rsGET_DETAIL                                    As ADODB.Recordset
        Dim jSJ_VOUCHERNO                                   As String
        Dim jINVOICETYPE                                    As String
        Dim jINVOICENO                                      As String
        Dim jCustomerCode                                   As String
        Dim jCUSTOMERNAME                                   As String
        Dim jAMOUNT_TOPAY                                   As Double
        Dim jAMOUNT_PAID                                    As Double
        Dim jBALANCE                                        As Double
        Dim jAccount_code                                   As String
        Dim jSYSTEM_REMARKS                                 As String
        Dim jinvoicedate                                    As String
        Dim jLASTUPDATED                                    As String

        jAMOUNT_TOPAY = 0

        Set rsCDJ_CREDIT = New ADODB.Recordset
        rsCDJ_CREDIT.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_CDJ_CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P' AND JDATE <= '" & dtprocess & "'  AND CREDIT <> 0 ", gconDMIS, adOpenKeyset
        If Not rsCDJ_CREDIT.EOF And Not rsCDJ_CREDIT.BOF Then
            Set rsGET_DETAIL = New ADODB.Recordset
            rsGET_DETAIL.Open "SELECT JTYPE,VOUCHERNO,ACCT_CODE,JDATE FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "' " & _
                              "AND ACCT_CODE = '" & xACCT_CODE & "' AND STATUS = 'P' AND JDATE <= '" & dtprocess & "' AND CREDIT <> 0", gconDMIS, adOpenKeyset
            If Not rsGET_DETAIL.EOF And Not rsGET_DETAIL.BOF Then
                jSJ_VOUCHERNO = N2Str2Null(rsGET_DETAIL!JTYPE & "-" & rsGET_DETAIL!VOUCHERNO)
                jINVOICETYPE = N2Str2Null("")
                jINVOICENO = N2Str2Null("")
                jCustomerCode = N2Str2Null("XXXXXX")
                jCUSTOMERNAME = N2Str2Null("XXXXXX")
                jAMOUNT_TOPAY = 0
                jAMOUNT_PAID = NumericVal(rsCDJ_CREDIT!SUM_CDJ_CREDIT)
                jBALANCE = Round((NumericVal(jAMOUNT_TOPAY) - NumericVal(jAMOUNT_PAID)), 2)
                jAccount_code = N2Str2Null(rsGET_DETAIL!ACCT_CODE)
                jSYSTEM_REMARKS = N2Str2Null("OFFSET")
                jinvoicedate = N2Date2Null(rsGET_DETAIL!JDATE)
                jLASTUPDATED = N2Date2Null(LOGDATE)

                Dim rsCHECK_EXIST                           As ADODB.Recordset
                Set rsCHECK_EXIST = New ADODB.Recordset
                rsCHECK_EXIST.Open "SELECT * FROM AMIS_AR WHERE SJVOUCHERNO = " & jSJ_VOUCHERNO & " AND ACCOUNT_CODE = " & jAccount_code & " AND SYSTEMREMARK = 'OFFSET'", gconDMIS, adOpenKeyset
                If Not rsCHECK_EXIST.EOF And Not rsCHECK_EXIST.BOF Then
                Else
                    gconDMIS.Execute "Insert into Amis_Ar (SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                     "VALUES(" & jSJ_VOUCHERNO & "," & jINVOICETYPE & "," & jINVOICENO & "," & jCustomerCode & "," & jCUSTOMERNAME & "," & jAMOUNT_TOPAY & "," & jAMOUNT_PAID & "," & jBALANCE & "," & jAccount_code & "," & jSYSTEM_REMARKS & "," & jinvoicedate & "," & jLASTUPDATED & ")"
                End If
            End If
        End If
        Set rsCDJ_CREDIT = Nothing
    End If
End Function

Function GET_AR_CDJ_PAYENT(xVOUCHERNO As String, xJType As String, xVENDORCODE As String, xInvoiceNo As String) As Double
'DESCRIPTION: GET PAYMENT AMOUNT TO THE CORRESPONDING AR ACCOUNT SCHED AMOUNT TO PAY FOR THE CORRESPONDING DISBURSEMENT
    Dim rsGET_AR_CDJ_PAYENT                                 As ADODB.Recordset
    Dim sum_AR_ADJ_PAYMENT                                  As Double
    sum_AR_ADJ_PAYMENT = 0
    Set rsGET_AR_CDJ_PAYENT = New ADODB.Recordset
    rsGET_AR_CDJ_PAYENT.Open "Select CREDIT FROM  AMIS_JOURNAL_DET WHERE  JTYPE = 'GJ' AND RIGHT(ENTITY,6) = '" & xVENDORCODE & "' AND INVOICENO = '" & xVOUCHERNO & "' AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsGET_AR_CDJ_PAYENT.EOF And Not rsGET_AR_CDJ_PAYENT.BOF Then
        Do While Not rsGET_AR_CDJ_PAYENT.EOF
            sum_AR_ADJ_PAYMENT = Round((sum_AR_ADJ_PAYMENT + NumericVal(rsGET_AR_CDJ_PAYENT!Credit)), 2)
            rsGET_AR_CDJ_PAYENT.MoveNext
        Loop
    End If
    GET_AR_CDJ_PAYENT = NumericVal(sum_AR_ADJ_PAYMENT)
    Set rsGET_AR_CDJ_PAYENT = Nothing
End Function

Function GET_CUST_NAME(xCUS_CODE As String) As String
'DESCRIPTION: GET THE ACCTNAME FOR THE CORRESPONDING ACCOUNT ACCOUNT CODE FROM THE ALL_CUSTOMER_TABLE
    Dim rsGET_CUST_NAME                                     As ADODB.Recordset
    Set rsGET_CUST_NAME = New ADODB.Recordset
    rsGET_CUST_NAME.Open "Select ACCTNAME from ALL_CUSTOMER_TABLE where CUSCDE = '" & xCUS_CODE & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CUST_NAME.EOF And Not rsGET_CUST_NAME.BOF Then
        GET_CUST_NAME = Null2String(rsGET_CUST_NAME!AcctName)
    End If
    Set rsGET_CUST_NAME = Nothing
End Function

Function GET_VEN_NAME(xVEN_CODE As String) As String
'DESCRITPION: GET THE NAME OF VENDOR FOR THE CORRESPONDING CODE IN THE ALL_VENDOR_TABLE
    Dim rsGET_VEN_NAME                                      As ADODB.Recordset
    Set rsGET_VEN_NAME = New ADODB.Recordset
    rsGET_VEN_NAME.Open "Select NAMEOFVENDOR from All_VENDOR_TABLE WHERE CODE = '" & xVEN_CODE & "'", gconDMIS, adOpenKeyset
    If Not rsGET_VEN_NAME.EOF And Not rsGET_VEN_NAME.BOF Then
        GET_VEN_NAME = Null2String(rsGET_VEN_NAME!nameofvendor)
    End If
    Set rsGET_VEN_NAME = Nothing
End Function

Function COMP_AMT_PAID(xInvoiceNo As String, xInvoiceType As String, xSJ_CustomerCode, xACCT_CODE As String, xSJ_VOUCHERNO As String, xJdate As String) As Double
'DESCRIPTION: COMPUTE THE TOTAL AMOUNT PAID  AND CHECK IF THERE IS AN ADJUSTMENT MADE
    Dim rsCOMP_AMT_PAID                                     As ADODB.Recordset
    Dim rsGET_ADJ                                           As ADODB.Recordset
    Dim SUM_CRJ                                             As Double
    Dim xINVOICE_AMOUNT                                     As Double
    Dim SUM_ADJ                                             As Double
    Dim SJVoucherno                                         As String
    Dim INVOICETYPE_TEMP                                    As String
    Dim DOUBLE_CARD                                         As String

    SUM_ADJ = 0
    xINVOICE_AMOUNT = 0
    SUM_CRJ = 0
    xREMARKS = ""
    Set rsCOMP_AMT_PAID = New ADODB.Recordset
    DOUBLE_CARD = ""

    'THIS IS FOR CREDIT CARD TRANSACTION ONLY
    Dim rsDOUBLE_ARCREDIT                                   As ADODB.Recordset
    Set rsDOUBLE_ARCREDIT = New ADODB.Recordset
    rsDOUBLE_ARCREDIT.Open "SELECT X.INV_INVTYPE FROM  ( " & _
                           "SELECT HD.INVOICETYPE + '-' + HD.INVOICENO AS INV_INVTYPE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON HD.VOUCHERNO = DET.VOUCHERNO  AND HD.JTYPE = DET.JTYPE  WHERE DET.ACCT_CODE = '11-02002-00' AND HD.INVOICENO = '" & xInvoiceNo & "' AND HD.INVOICETYPE = '" & xInvoiceType & "' AND HD.STATUS = 'P' " & _
                           ") X WHERE X.INV_INVTYPE IN (SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET " & _
                           "ON CRJ.VOUCHERNO = DET.VOUCHERNO  AND CRJ.CR_TYPE = DET.JTYPE WHERE DET.ACCT_CODE = '11-02002-00' AND DET.STATUS = 'P')", gconDMIS, adOpenKeyset
    If Not rsDOUBLE_ARCREDIT.EOF And Not rsDOUBLE_ARCREDIT.BOF Then
        DOUBLE_CARD = "X"
    End If
    Set rsDOUBLE_ARCREDIT = Nothing


    'THIS IS FOR GETTING THE RIGTH AMOUNT FOR THE DOUBLE INVOICENO
    Dim rsDOUBLE                                            As ADODB.Recordset
    Dim xDOUBLE_INV                                         As Integer
    Set rsDOUBLE = New ADODB.Recordset
    rsDOUBLE.Open "SELECT COUNT(*) AS INV_COUNT FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND CUSTOMERCODE = '" & xSJ_CustomerCode & "' AND JDATE <= '" & dtprocess & "' GROUP BY JTYPE", gconDMIS, adOpenKeyset
    If Not rsDOUBLE.EOF And Not rsDOUBLE.BOF Then
        xDOUBLE_INV = NumericVal(rsDOUBLE!INV_COUNT)
    End If
    Set rsDOUBLE = Nothing

    'THIS IS TO SET ANOTHER INVOICETYPE DUE SERVICE INVOICE AND VEHICLE INVOICE IS DIFFERENT FROM VI AND SI
    If RTrim(LTrim(xInvoiceType)) = "VI" Then
        INVOICETYPE_TEMP = "VEHICLE INVOICE"
    ElseIf RTrim(LTrim(xInvoiceType)) = "SI" Then
        INVOICETYPE_TEMP = "SERVICE INVOICE"
    End If

    If xInvoiceNo = "INT RO" Then
        rsCOMP_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where InvoiceType = '" & xInvoiceType & "' and InvoiceNo = '" & xInvoiceNo & "' and CR_TYPE = 'CRJ' AND SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
    ElseIf xDOUBLE_INV > 1 Then
        rsCOMP_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where InvoiceType = '" & xInvoiceType & "' and InvoiceNo = '" & xInvoiceNo & "' and CR_TYPE = 'CRJ' AND SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
    ElseIf DOUBLE_CARD = "X" Then
        rsCOMP_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where InvoiceType = '" & xInvoiceType & "' and InvoiceNo = '" & xInvoiceNo & "' and CR_TYPE = 'CRJ' AND SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "'", gconDMIS, adOpenKeyset
    Else
        If IsNumeric(xInvoiceNo) = True Then
            rsCOMP_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where (InvoiceType = '" & xInvoiceType & "' OR InvoiceType = '" & INVOICETYPE_TEMP & "') and (InvoiceNo = '" & Abs(xInvoiceNo) & "' or INVOICENO = '" & xInvoiceNo & "') and CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
        Else
            rsCOMP_AMT_PAID.Open "Select SJ_VOUCHERNO,VoucherNo,Cr_Type,InvoiceAmount,InvoiceNo,InvoiceType from Amis_Crj_Detail where (InvoiceType = '" & xInvoiceType & "' OR InvoiceType = '" & INVOICETYPE_TEMP & "') and  INVOICENO = '" & xInvoiceNo & "' and CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
        End If
    End If

    If Not rsCOMP_AMT_PAID.EOF And Not rsCOMP_AMT_PAID.BOF Then
        Do While Not rsCOMP_AMT_PAID.EOF

            SJVoucherno = Null2String(LTrim(RTrim(rsCOMP_AMT_PAID!CR_type))) & "-" & Null2String(RTrim(LTrim(rsCOMP_AMT_PAID!VOUCHERNO)))

            'VALIDATE IF IT IS AR SCHEDULE
            Dim rsAR_SCHED                                  As ADODB.Recordset
            Set rsAR_SCHED = New ADODB.Recordset
            'rsAR_SCHED.Open "Select * from Amis_Journal_det where VOUCHERNO = '" & rsCOMP_AMT_PAID!VOUCHERNO & "' AND JTYPE = 'CRJ' AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset

            'UPDATE ADD IN FILTERING CREDIT <> 0 DUE TO WRONG ENTRY IN VOUCHERNO-005932
            rsAR_SCHED.Open "Select * from Amis_Journal_det where VOUCHERNO = '" & rsCOMP_AMT_PAID!VOUCHERNO & "' AND JTYPE = 'CRJ' AND ACCT_CODE = '" & xACCT_CODE & "' and CREDIT <> 0", gconDMIS, adOpenKeyset
            'FOR DEBUGGING PURPOSES ONLY
            'If rsAR_SCHED!VOUCHERNO = "000126" And rsAR_SCHED!jtype = "COB" Then Stop

            If Not rsAR_SCHED.EOF And Not rsAR_SCHED.BOF Then
                'Call VAL_SJ_CR_CODE(Null2String(rsCOMP_AMT_PAID!INVOICENO), Null2String(rsCOMP_AMT_PAID!InvoiceType), Null2String(rsCOMP_AMT_PAID!VOUCHERNO))
                Call VAL_SJ_CR_CODE(xInvoiceNo, xInvoiceType, Null2String(rsCOMP_AMT_PAID!VOUCHERNO), xACCT_CODE)
                If IS_POSTED_IN_HD(Null2String(rsCOMP_AMT_PAID!VOUCHERNO)) = True Then
                    If xREMARKS = "Wrong Code" Then
                        'Not related payment
                        Dim rsCHECK_IN_SJ                   As ADODB.Recordset
                        Set rsCHECK_IN_SJ = New ADODB.Recordset
                        rsCHECK_IN_SJ.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE INVOICENO = '" & xInvoiceNo & "' AND INVOICETYPE = '" & xInvoiceType & "' AND CUSTOMERCODE = '" & xSJ_CustomerCode & "' AND JTYPE IN('SJ','COB')", gconDMIS, adOpenKeyset
                        If Not rsCHECK_IN_SJ.EOF And Not rsCHECK_IN_SJ.BOF Then
                            'THIS PAYMENT BELONG TO OTHER CODE
                        End If
                        Set rsCHECK_IN_SJ = Nothing
                    ElseIf xREMARKS = "RILI RONG" Then
                        'THIS REALLY WRONG CODE
                        Dim xWRONG_CDE_AMT                  As Double
                        xWRONG_CDE_AMT = Round((0 - NumericVal(rsCOMP_AMT_PAID!invoiceamount)), 2)

                        Dim rsALREADY_IN_AR                 As ADODB.Recordset
                        Set rsALREADY_IN_AR = New ADODB.Recordset
                        rsALREADY_IN_AR.Open "Select * from Amis_Ar where SJVOUCHERNO = '" & SJVoucherno & "'", gconDMIS, adOpenKeyset
                        If Not rsALREADY_IN_AR.EOF And Not rsALREADY_IN_AR.BOF Then
                        Else
                            If xxx_credit <> 0 Then
                                gconDMIS.Execute "Insert Into AMIS_AR (SJVOUCHERNO,CRJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,SYSTEMREMARK,INVOICEDATE,LASTUPDATED)" & _
                                                 "VALUES('" & SJVoucherno & "','" & SJVoucherno & "','" & xInvoiceType & "','" & xInvoiceNo & "','" & xCRJ_CODE & "','" & GET_CUST_NAME(xCRJ_CODE) & "',0,'" & NumericVal(rsCOMP_AMT_PAID!invoiceamount) & "'," & xWRONG_CDE_AMT & ",'" & xACCT_CODE & "','" & xREMARKS & "'," & N2Date2Null(rsAR_SCHED!JDATE) & ",'" & LOGDATE & "')"
                            End If
                        End If
                        Set rsALREADY_IN_AR = Nothing
                    Else
                        'FOR CHECKING OF WRONG ACCOUNT CODE ENTRY
                        If WRONG_ENTRY(Null2String(RTrim(LTrim(rsCOMP_AMT_PAID!VOUCHERNO)))) = False Then
                            SUM_CRJ = Round((SUM_CRJ + NumericVal(rsCOMP_AMT_PAID!invoiceamount)), 2)
                            'FOR DEBUGGING PURPOSES
                            gconDMIS.Execute "UPDATE AMIS_AR_HD SET AMOUNTPAID = 1 where voucherno = '" & rsCOMP_AMT_PAID!VOUCHERNO & "' AND JTYPE  = 'CRJ'"

                            If CHECK_DUPLICATE(N2Str2Null(xInvoiceNo), N2Str2Null(xInvoiceType), N2Str2Null(RTrim(LTrim(rsCOMP_AMT_PAID!VOUCHERNO)))) = False Then
                                'THIS IS FOR PAYMENT DETAIL
                                gconDMIS.Execute "INSERT INTO AMIS_DETAIL (INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,REMARKS,VOUCHERNO) " & _
                                                 "VALUES(" & N2Str2Null(xInvoiceNo) & ", " & N2Str2Null(xInvoiceType) & ", " & NumericVal(rsCOMP_AMT_PAID!invoiceamount) & ", " & N2Str2Null(xSJ_CustomerCode) & "," & N2Str2Null(xACCT_CODE) & "," & N2Date2Null(xJdate) & ",'3','" & RTrim(LTrim(rsCOMP_AMT_PAID!VOUCHERNO)) & "')"

                                'UPDATE THE AMIS_CRJ_DETAIL WITH SJ_VOUCHERNO AND CUSTOMER CODE
                                gconDMIS.Execute "UPDATE AMIS_CRJ_DETAIL SET SJ_VOUCHERNO = '" & xSJ_VOUCHERNO & "', CUSTOMERCODE = '" & xSJ_CustomerCode & "' where INVOICENO = '" & rsCOMP_AMT_PAID!INVOICENO & "' AND INVOICETYPE = '" & rsCOMP_AMT_PAID!InvoiceType & "' AND VOUCHERNO = '" & rsCOMP_AMT_PAID!VOUCHERNO & "'"
                            End If
                        End If
                    End If
                Else
                    'NOT POSTED
                End If
            Else
                'THIS IS CASH TRANSACTION
            End If
            Set rsAR_SCHED = Nothing


            '                Call VAL_SJ_CR_CODE(Null2String(rsCOMP_AMT_PAID!INVOICENO), Null2String(rsCOMP_AMT_PAID!InvoiceType), Null2String(rsCOMP_AMT_PAID!VOUCHERNO))
            '                If IS_POSTED_IN_HD(Null2String(rsCOMP_AMT_PAID!VOUCHERNO)) = True Then
            '                    If xREMARKS = "Wrong Code" Then
            '                        'Not related payment
            '                    Else
            '                        SUM_CRJ = Round((SUM_CRJ + NumericVal(rsCOMP_AMT_PAID!INVOICEAMOUNT)), 2)
            '                    End If
            '                Else
            '                    'NOT POSTED
            '                End If
            rsCOMP_AMT_PAID.MoveNext
        Loop
    End If

    '        Set rsGET_ADJ = New ADODB.Recordset
    '        rsGET_ADJ.Open "Select ADJ_AMOUNT,ADJ_TYPE from Amis_Journal_hd where InvoiceNo = '" & xInvoiceNo & "' and InvoiceType = '" & xinvoiceType & "' and JType = 'ADJ' and STATUS = 'P'", gconDMIS, adOpenKeyset
    '        If Not rsGET_ADJ.EOF And Not rsGET_ADJ.BOF Then
    '            Do While Not rsGET_ADJ.EOF
    '                    If RTrim(LTrim(Null2String(rsGET_ADJ!ADJ_TYPE))) = "CREDIT" Then
    '                        SUM_ADJ = SUM_ADJ + NumericVal(rsGET_ADJ!ADJ_AMOUNT)
    '                    End If
    '                rsGET_ADJ.MoveNext
    '            Loop
    '        End If

    Set rsGET_ADJ = New ADODB.Recordset
    rsGET_ADJ.Open "Select Debit,Credit from Amis_journal_det where InvoiceNo = '" & xInvoiceNo & "' and invoicetype = '" & xInvoiceType & "' and right(Entity,6) = '" & xSJ_CustomerCode & "' AND Left(Acct_Code,5) IN('11-02','11-03') AND STATUS = 'P' and JDATE< = '" & dtprocess & "'", gconDMIS, adOpenKeyset
    If Not rsGET_ADJ.EOF And Not rsGET_ADJ.BOF Then
        Do While Not rsGET_ADJ.EOF
            If NumericVal(rsGET_ADJ!Credit) <> 0 Then
                SUM_ADJ = Round((SUM_ADJ + NumericVal(rsGET_ADJ!Credit)), 2)
            End If
            rsGET_ADJ.MoveNext
        Loop
    End If

    COMP_AMT_PAID = Round((NumericVal(SUM_CRJ) + NumericVal(SUM_ADJ)), 2)
    Set rsCOMP_AMT_PAID = Nothing
    Set rsGET_ADJ = Nothing
End Function
Function WRONG_ENTRY(xVOUCHERNO As String) As Boolean
'DESCRIPTION: IF WRONG ACCT_CODE ENTRY DO NOT ALLOW TO INCLUDE IN COMPUTATION
    Dim rsWRONG_ENTRY                                       As ADODB.Recordset
    Dim rsWRONG_ENTRY2                                      As ADODB.Recordset
    Dim rsDUP_INVOICE                                       As ADODB.Recordset

    WRONG_ENTRY = False
    Set rsWRONG_ENTRY = New ADODB.Recordset
    rsWRONG_ENTRY.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM ( " & _
                       "SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO, " & _
                       "CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE FROM AMIS_CRJ_DETAIL CRJ " & _
                       "INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE INNER JOIN AMIS_JOURNAL_HD HD " & _
                       "ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE WHERE HD.VOUCHERNO = '" & xVOUCHERNO & "' AND LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' " & _
                       "AND DET.STATUS = 'P' AND DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0) X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO FROM " & _
                       "AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                       "WHERE X.ACCT_CODE <> DET.ACCT_CODE AND HD.CUSTOMERCODE <> X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') AND DET.DEBIT <> 0 )", gconDMIS, adOpenKeyset
    If Not rsWRONG_ENTRY.EOF And Not rsWRONG_ENTRY.BOF Then
        Dim rsDUP_INVOICE2                                  As ADODB.Recordset
        Set rsDUP_INVOICE2 = New ADODB.Recordset
        'DESCRIPTION: THIS IS FOR REVALIDATION FOR DUPLICATE INVOICENO 1 ACCT CODE IS EQUAL AND THE OTHER ONE IS NOT PRONE TO WARRANTY
        rsDUP_INVOICE2.Open "SELECT X.xINVOICE, X.xACCT_CODE,X.xCUSTOMERCODE FROM ( " & _
                            "SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS xINVOICE ,DET.ACCT_CODE AS xACCT_CODE, HD.CUSTOMERCODE AS xCUSTOMERCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.CR_TYPE = HD.JTYPE AND CRJ.VOUCHERNO = HD.VOUCHERNO WHERE CRJ.VOUCHERNO = " & N2Str2Null(rsWRONG_ENTRY!DET_VOUCHERNO) & " AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') " & _
                            ")X WHERE X.xINVOICE IN(SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAl_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE  DET.ACCT_CODE = X.xACCT_CODE AND HD.CUSTOMERCODE = X.xCUSTOMERCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03'))", gconDMIS, adOpenKeyset
        If Not rsDUP_INVOICE2.EOF And Not rsDUP_INVOICE2.BOF Then
        Else
            WRONG_ENTRY = True
        End If
    End If

    Set rsWRONG_ENTRY2 = New ADODB.Recordset
    rsWRONG_ENTRY2.Open "SELECT DISTINCT X.DET_VOUCHERNO,X.HD_CUSCODE,X.INV,X.ACCT_CODE FROM " & _
                        "(SELECT DISTINCT DET.VOUCHERNO AS DET_VOUCHERNO,DET.JDATE AS DET_JDATE,CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS INV,CRJ.INVOICENO AS I_NO,CRJ.INVOICETYPE AS I_TYPE,DET.ACCT_CODE AS ACCT_CODE, " & _
                        "CRJ.INVOICEAMOUNT AS INV_AMT,HD.CUSTOMERCODE AS HD_CUSCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.VOUCHERNO = DET.VOUCHERNO AND CRJ.CR_TYPE = DET.JTYPE " & _
                        "INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE " & _
                        "WHERE HD.VOUCHERNO = '" & xVOUCHERNO & "' AND LEFT(DET.ACCT_CODE,5) = '11-02' AND DET.JDATE <= '" & dtprocess & "' AND DET.STATUS = 'P' AND DET.JTYPE = 'CRJ' AND DET.CREDIT <> 0 " & _
                        ") X WHERE X.INV IN (SELECT HD.INVOICETYPE + '-'+ HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE X.ACCT_CODE <> DET.ACCT_CODE AND HD.CUSTOMERCODE = X.HD_CUSCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02') AND DET.DEBIT <> 0 ) ", gconDMIS, adOpenKeyset
    If Not rsWRONG_ENTRY2.EOF And Not rsWRONG_ENTRY2.BOF Then
        Do While Not rsWRONG_ENTRY2.EOF
            Set rsDUP_INVOICE = New ADODB.Recordset
            'DESCRIPTION: THIS IS FOR REVALIDATION FOR DUPLICATE INVOICENO 1 ACCT CODE IS EQUAL AND THE OTHER ONE IS NOT PRONE TO WARRANTY
            rsDUP_INVOICE.Open "SELECT X.xINVOICE, X.xACCT_CODE,X.xCUSTOMERCODE FROM ( " & _
                               "SELECT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO AS xINVOICE ,DET.ACCT_CODE AS xACCT_CODE, HD.CUSTOMERCODE AS xCUSTOMERCODE FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_DET DET ON CRJ.CR_TYPE = DET.JTYPE AND CRJ.VOUCHERNO = DET.VOUCHERNO INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.CR_TYPE = HD.JTYPE AND CRJ.VOUCHERNO = HD.VOUCHERNO WHERE CRJ.VOUCHERNO = " & N2Str2Null(rsWRONG_ENTRY2!DET_VOUCHERNO) & " AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') " & _
                               ")X WHERE X.xINVOICE IN(SELECT HD.INVOICETYPE + '-' + HD.INVOICENO FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAl_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE WHERE  DET.ACCT_CODE = X.xACCT_CODE AND HD.CUSTOMERCODE = X.xCUSTOMERCODE AND LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03'))", gconDMIS, adOpenKeyset
            If Not rsDUP_INVOICE.EOF And Not rsDUP_INVOICE.BOF Then
                Dim rsMULTIPLE_ACCT                         As ADODB.Recordset
                Set rsMULTIPLE_ACCT = New ADODB.Recordset
                rsMULTIPLE_ACCT.Open "SELECT COUNT(DISTINCT ACCT_CODE) AS CODE_COUNT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & rsWRONG_ENTRY2!DET_VOUCHERNO & "' AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset
                If Not rsMULTIPLE_ACCT.EOF And Not rsMULTIPLE_ACCT.BOF Then
                    If rsMULTIPLE_ACCT!CODE_COUNT > 1 Then
                        WRONG_ENTRY = True
                    End If
                End If
            Else
                WRONG_ENTRY = True
            End If
            rsWRONG_ENTRY2.MoveNext
        Loop
    End If
    Set rsWRONG_ENTRY = Nothing
    Set rsWRONG_ENTRY2 = Nothing
End Function


Function IS_POSTED_IN_HD(xVOUCHERNO As String) As Boolean
    Dim rsIS_POSTED_IN_HD                                   As ADODB.Recordset
    Set rsIS_POSTED_IN_HD = New ADODB.Recordset
    rsIS_POSTED_IN_HD.Open "Select STATUS from Amis_Journal_hd where VoucherNo = '" & xVOUCHERNO & "' AND JTYPE = 'CRJ'", gconDMIS, adOpenKeyset
    If Not rsIS_POSTED_IN_HD.EOF And Not rsIS_POSTED_IN_HD.BOF Then
        If Null2String(rsIS_POSTED_IN_HD!Status) = "P" Then
            IS_POSTED_IN_HD = True
        Else
            IS_POSTED_IN_HD = False
        End If
    End If
    Set rsIS_POSTED_IN_HD = Nothing
End Function

Sub VAL_SJ_CR_CODE(xInvoiceNo As String, xInvoiceType As String, xVOUCHERNO As String, xACCT_CODE As String)
'DESCRIPTION: VALIDATE THE CRJ CODE AGAINST THE SJ CUSTOMER CODE
    Dim rsCRJ_CODE                                          As ADODB.Recordset
    Dim rsSJ_CODE                                           As ADODB.Recordset
    Dim rsVoucherCode                                       As ADODB.Recordset
    Dim xSJ_CODE                                            As String
    Dim xJdate                                              As String


    xSJ_CODE = ""
    xCRJ_CODE = ""

    'Set rsCRJ_CODE = New ADODB.Recordset
    'rsCRJ_CODE.Open "Select VoucherNo from Amis_Crj_Detail where InvoiceNo = '" & xInvoiceNo & "' and InvoiceType ='" & xinvoiceType & "' and CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
    'If Not rsCRJ_CODE.EOF And Not rsCRJ_CODE.BOF Then
    Set rsVoucherCode = New ADODB.Recordset

    'rsVoucherCode.Open "Select Jdate,CustomerCode from Amis_journal_hd where VoucherNo = '" & xVOUCHERNO & "' and Jtype = 'CRJ'  and JDATE <= '" & dtprocess & "'", gconDMIS, adOpenKeyset
    rsVoucherCode.Open "Select CRJ.J_Class,hd.Jdate,hd.CustomerCode from Amis_journal_hd HD Inner Join Amis_Crj_detail CRJ on HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE where hd.VoucherNo = '" & xVOUCHERNO & "' and hd.Jtype = 'CRJ'  and hd.JDATE <= '" & dtprocess & "'", gconDMIS, adOpenKeyset

    If Not rsVoucherCode.EOF And Not rsVoucherCode.BOF Then
        xCRJ_CODE = Null2String(rsVoucherCode!CustomerCode)
    Else

        'SJ customer code not found in CRJ
    End If
    'End If

    Set rsSJ_CODE = New ADODB.Recordset
    rsSJ_CODE.Open "Select CustomerCode from Amis_Journal_hd where InvoiceNo = '" & xInvoiceNo & "' and InvoiceType = '" & xInvoiceType & "' and JTYPE IN('SJ','COB')  AND CUSTOMERCODE = '" & xSJ_CustomerCode & "'", gconDMIS, adOpenKeyset
    If Not rsSJ_CODE.EOF And Not rsSJ_CODE.BOF Then
        xSJ_CODE = Null2String(rsSJ_CODE!CustomerCode)
    End If

    If RTrim(LTrim(xSJ_CODE)) <> RTrim(LTrim(xCRJ_CODE)) Then
        xREMARKS = "Wrong Code"

        If xCRJ_CODE <> "" Then
            Dim rsRILI_RONG                                 As ADODB.Recordset
            Set rsRILI_RONG = New ADODB.Recordset
            'rsRILI_RONG.Open "Select * from AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.INVOICENO = '" & xINVOICENO & "' AND HD.INVOICETYPE = '" & xINVOICETYPE & "' AND HD.CUSTOMERCODE = '" & xCRJ_CODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "' and DET.CREDIT <> 0", gconDMIS, adOpenKeyset
            rsRILI_RONG.Open "Select * from AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE  WHERE HD.INVOICENO = '" & xInvoiceNo & "' AND HD.INVOICETYPE = '" & xInvoiceType & "' AND HD.CUSTOMERCODE = '" & xCRJ_CODE & "' AND DET.ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset    'LAST UPDATED 09/07/2009 REMOVE DET.CREDIT <> 0 IN QUERY
            If Not rsRILI_RONG.EOF And Not rsRILI_RONG.BOF Then
            Else
                Dim rsDEBIT                                 As ADODB.Recordset
                Set rsDEBIT = New ADODB.Recordset
                rsDEBIT.Open "Select round(sum(CREDIT),2) AS CREDIT from amis_journal_det where VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = 'CRJ' AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
                If Not rsDEBIT.EOF And Not rsDEBIT.BOF Then
                    xxx_credit = NumericVal(rsDEBIT!Credit)
                End If
                xREMARKS = "RILI RONG"
            End If
            Set rsRILI_RONG = Nothing
        End If
    Else
        xREMARKS = ""
    End If
    Set rsCRJ_CODE = Nothing
    Set rsSJ_CODE = Nothing
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
    getlastdate
    dtpTranDate = LOGDATE
    dtprocess = LOGDATE

    GetAcctcode
    picNolink.Visible = False
    picPeriod.Enabled = False
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
    dtpAsOF.Enabled = True
    'Call GET_MAX_DATE
    Label13.Caption = "AS OF: " & GET_MAX_DATE
    Label16.Caption = "AS OF: " & GET_MAX_DATE
    Screen.MousePointer = 0

    'getlastdate ' commented by: JUN
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub cmdCancel_Click()
    picReport.Visible = True
    picByAccount.Visible = False
    picDetailedSum.Visible = False
    Me.Caption = "AR REPORT"
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    Dim Rs_CRJTotal                                         As New ADODB.Recordset
    Dim Rs_ARAing                                           As New ADODB.Recordset
    Dim rptApp                                              As CRAXDRT.Application
    Dim rptRep                                              As REPORT
    Dim crSections                                          As CRAXDRT.Sections
    Dim crSection                                           As CRAXDRT.Section
    Dim crRepObjs                                           As CRAXDRT.ReportObjects
    Dim crSubRepObj                                         As CRAXDRT.SubreportObject
    Dim crSubReport                                         As CRAXDRT.REPORT
    Dim j As Integer, k                                     As Integer
    Dim ellaine                                             As Integer
    'ACL 0452010
    'DESC: Report using crystal report viewer
    rptAMISDueReport.Reset
    rptAMISDueReport.WindowShowSearchBtn = True

    rptAMISDueReport.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAMISDueReport.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HMH") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HLP") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HAM") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HSP") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HGC") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HCI" Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "HPI") Or (cboCOBAcctName.Text = "" And Option2.Value = True And COMPANY_CODE = "DGI")) Then
        MsgBox "Please select from the list.", vbInformation, "Account Description"
        cboCOBAcctName.SetFocus
    Else
        If Report_Type = "SCHED" Then
            If IsDate(lblDate.Caption) = True Then
                If dtpAsOF.Value > CDate(lblDate.Caption) Then
                    MsgBox "Information: Date is greater than Last data generated"
                    'Exit Sub
                End If
            End If
            '        If Label9.Caption = "PLEASE GENERATE AR DATA" Then
            '            MsgBox "Please Generate AR AGING DATA..This will generate last data generated", vbInformation, "INFO"
            '            'Exit Sub
            '        End If
            If Option1.Value = True Then
                If optAsOf.Value = True Then
                    If COMPANY_CODE = "" Then
                        'If COMPANY_CODE = "HPI" Then
                        rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                        rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                        'rptAMISDueReport.Formulas(10) = "@JDATE = '" & dtpAsOF.Value & "'"
                        'PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1

                        'Dim rsScheduleReport As New ADODB.Recordset

                        '                Dim cmd As New ADODB.Command
                        '                Dim asOF As String
                        '                Dim xCounter As Integer
                        '                Dim xlApplication As Excel.Application
                        '                Dim xlWorkbook As Excel.Workbook
                        '                Dim xlWorksheet As Excel.Worksheet
                        '                Set xlApplication = CreateObject("Excel.Application")
                        '                Set xlWorkbook = xlApplication.Workbooks.Open(AMIS_REPORT_PATH & "DueReports\" & "ARScheduleReport.xlt")
                        '                Set xlWorksheet = xlWorkbook.Worksheets(1)
                        '                xCounter = 2
                        '                With cmd
                        '                    Set .ActiveConnection = gconDMIS
                        '                        .CommandType = adCmdStoredProc
                        '                        .CommandText = "USP_SCHEDULEREPORT"
                        '                End With
                        '                With cmd.Parameters
                        '                    .Append cmd.CreateParameter("@JDATE", adDate, adParamInput, 8, dtpAsOF)
                        '                End With
                    Else
                        Me.WindowState = vbMaximized
                        CRViewer1.Height = Me.Height - 800
                        CRViewer1.Width = Me.Width
                        CRViewer1.ZOrder 0
                        Set rptApp = New CRAXDRT.Application
                        Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt", 1)
                        rptRep.DiscardSavedData
                        rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
                        rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
                        rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                        Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                        Set crSections = rptRep.Sections
                        For ellaine = 1 To crSections.Count
                            Set crSection = crSections.Item(ellaine)
                            Set crRepObjs = crSection.ReportObjects
                            For j = 1 To crRepObjs.Count
                                If crRepObjs.Item(j).Kind = crSubreportObject Then
                                    Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                    'For k = 1 To crSubReport.ParameterFields.Count
                                    If ellaine = 7 Then
                                        Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                        'Call crSubReport.ParameterFields(6).AddCurrentValue(CDate(Day(dtpAsOF) & "-" & Month(dtpAsOF) & "-" & Year(dtpAsOF)))
                                        Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                    Else
                                        Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                        'Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(Day(dtpAsOF) & "-" & Month(dtpAsOF) & "-" & Year(dtpAsOF)))
                                        Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                    End If
                                    'Next
                                End If
                            Next
                        Next
                        With CRViewer1
                            .ReportSource = rptRep
                            .DisplayGroupTree = False
                            .DisplayTabs = False
                            .DisplayToolbar = True
                            .ViewReport
                        End With

                        Set rptApp = Nothing
                        Set rptRep = Nothing
                        '                Set rsScheduleReport = cmd.Execute
                        '                Do While Not rsScheduleReport.EOF
                        '                    xlWorksheet.Cells(xCounter, 1) = Null2String(rsScheduleReport!SJVoucherno)
                        '                    xlWorksheet.Cells(xCounter, 2) = Null2String(rsScheduleReport!invoicedate)
                        '                    xlWorksheet.Cells(xCounter, 3) = Null2String(rsScheduleReport!INVOICENO)
                        '                    xlWorksheet.Cells(xCounter, 4) = Null2String(rsScheduleReport!InvoiceType)
                        '                    xlWorksheet.Cells(xCounter, 5) = Null2String(rsScheduleReport!AR_TOPAY)
                        '                    xlWorksheet.Cells(xCounter, 6) = Null2String(rsScheduleReport!AMOUNT_PAID)
                        '                    xlWorksheet.Cells(xCounter, 7) = Null2String(rsScheduleReport!BALANCE)
                        '                    xCounter = xCounter + 1
                        '                    rsScheduleReport.MoveNext
                        '                    DoEvents
                        '                Loop
                        '                'asOF = cmd.Parameters("@JDATE")
                        '                xlApplication.Visible = True
                        '                Set xlApplication = Nothing
                        '                Set rsScheduleReport = Nothing

                        'rptAMISDueReport.StoredProcParam(0) = "@JDATE = CDATE('" & dtpAsOF & "')"
                        'rptAMISDueReport.ReportFileName = AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt"
                        'rptAMISDueReport.ParameterFields(0) = "JDATE;" & "DATE(2010,1,31)" & ";false"
                        'rptAMISDueReport.SelectionFormula = "{@JDATE} = cdate('" & dtpAsOF & "')"
                        'rptAMISDueReport.ParameterFields(0) = "@JDATE = CDATE('" & dtpAsOF & "')"

                        'rptAMISDueReport.PrintReport
                        'rptAMISDueReport.ReplaceSelectionFormula dtpAsOF
                        'PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARScheduleReport.Rpt", "{@JDATE} = cdate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                        'rptAMISDueReport.Action = 1
                        'rptAMISDueReport.ParameterFields(0) = "@JDATE"
                        LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & dtpAsOF
                    End If
                Else
                    '
                End If
            ElseIf Option2.Value = True Then
                If COMPANY_CODE = "" Then
                    'If COMPANY_CODE = "HPI" Then
                    rptAMISDueReport.WindowTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                    rptAMISDueReport.ReportTitle = "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF
                    'rptAMISDueReport.Formulas(10) = "JDATE = '" & dtpAsOF.Value & "'"
                    'PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                    PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", "", DMIS_REPORT_Connection, 1
                Else
                    Me.WindowState = vbMaximized
                    CRViewer1.Height = Me.Height - 800
                    CRViewer1.Width = Me.Width
                    CRViewer1.ZOrder 0
                    Set rptApp = New CRAXDRT.Application

                    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "DGI" Then
                        If cboCOBAcctName.Text = "ALL" Then
                            Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", 1)
                        Else
                            Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\GroupARScheduleReportAccount.Rpt", 1)
                        End If
                    Else
                        Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\GroupARScheduleReport.Rpt", 1)
                    End If

                    rptRep.DiscardSavedData
                    rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
                    rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
                    rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "SCHEDULE OF ACCOUNTS RECEIVABLE AS OF: " & dtpAsOF

                    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "DGI" Then
                        If cboCOBAcctName.Text = "ALL" Then
                            Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                            Set crSections = rptRep.Sections
                            For ellaine = 1 To crSections.Count
                                Set crSection = crSections.Item(ellaine)
                                Set crRepObjs = crSection.ReportObjects
                                For j = 1 To crRepObjs.Count
                                    If crRepObjs.Item(j).Kind = crSubreportObject Then
                                        Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                        If ellaine = 7 Then
                                            Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                            Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                        Else
                                            Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                            Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                        End If
                                    End If
                                Next
                            Next
                        Else
                            Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                            Call rptRep.ParameterFields(5).AddCurrentValue(txtCOBAcctNo.Text)
                            Set crSections = rptRep.Sections
                            For ellaine = 1 To crSections.Count
                                Set crSection = crSections.Item(ellaine)
                                Set crRepObjs = crSection.ReportObjects
                                For j = 1 To crRepObjs.Count
                                    If crRepObjs.Item(j).Kind = crSubreportObject Then
                                        Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                        If ellaine = 7 Then
                                            Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                            Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                            Call crSubReport.ParameterFields(6).ClearCurrentValueAndRange
                                            Call crSubReport.ParameterFields(6).AddCurrentValue(txtCOBAcctNo.Text)
                                            '                                Else
                                            '                                    Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                            '                                    Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    Else
                        Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                        Set crSections = rptRep.Sections
                        For ellaine = 1 To crSections.Count
                            Set crSection = crSections.Item(ellaine)
                            Set crRepObjs = crSection.ReportObjects
                            For j = 1 To crRepObjs.Count
                                If crRepObjs.Item(j).Kind = crSubreportObject Then
                                    Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                    If ellaine = 7 Then
                                        Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                        Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                    Else
                                        Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                        Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                    End If
                                End If
                            Next
                        Next
                    End If
                    With CRViewer1
                        .ReportSource = rptRep
                        .DisplayGroupTree = False
                        .DisplayTabs = False
                        .DisplayToolbar = True
                        .ViewReport
                    End With

                    Set rptApp = Nothing
                    Set rptRep = Nothing
                End If
                LogAudit "V", "SCHEDULE OF ACCOUNTS RECEIVABLE", "As of: " & dtpAsOF
            End If
        Else
            If Report_Type = "AGING" Then
                Dim Ans                                     As String
                If Option1.Value = True Then
                    '                Ans = MsgBox("Do you want to generate Ar report?", vbQuestion + vbYesNo, "Info")
                    '                If Ans = vbYes Then
                    'If ar = True Then
                    If COMPANY_CODE = "" Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                        'rptAMISDueReport.Formulas(10) = "JDATE = '" & dtpAsOF.Value & "'"
                        'PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "", DMIS_REPORT_Connection, 1
                    Else
                        Me.WindowState = vbMaximized
                        CRViewer1.Height = Me.Height - 800
                        CRViewer1.Width = Me.Width
                        CRViewer1.ZOrder 0
                        Set rptApp = New CRAXDRT.Application
                        If optDetailed = True Then
                            Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", 1)
                        Else
                            Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTBYCUSTOMER.Rpt", 1)
                        End If
                        
                        rptRep.DiscardSavedData
                        rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
                        rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
                        rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF

                        Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                        Set crSections = rptRep.Sections
                        For ellaine = 1 To crSections.Count
                            Set crSection = crSections.Item(ellaine)
                            Set crRepObjs = crSection.ReportObjects
                            For j = 1 To crRepObjs.Count
                                If crRepObjs.Item(j).Kind = crSubreportObject Then
                                    Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                    If ellaine = 7 Then
                                        'Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                        'Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                    Else
                                        Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                        Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                    End If
                                End If
                            Next
                        Next
                        With CRViewer1
                            .ReportSource = rptRep
                            .DisplayGroupTree = False
                            .DisplayTabs = False
                            .DisplayToolbar = True
                            .ViewReport
                        End With

                        Set rptApp = Nothing
                        Set rptRep = Nothing
                    End If
                    LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & lblDate
                    'Else
                    'MsgBox "No AR as of the " & lblDate
                    'End If
                    '                Else
                    '                    If Option1.Value = True Then
                    '                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    '                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    '                        rptAMISDueReport.Formulas(10) = "JDATE = '" & dtpAsOF.Value & "'"
                    '                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORT.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                    '                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
                    '                    ElseIf Option1.Value = True Then
                    '                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    '                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                    '                        rptAMISDueReport.Formulas(10) = "JDATE = '" & dtpAsOF.Value & "'"
                    '                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                    '                        LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
                    '
                    '                    End If
                    '                End If
                ElseIf Option2.Value = True Then      ' Group report by Account
                    If COMPANY_CODE = "" Then
                        'If COMPANY_CODE = "HPI" Then
                        rptAMISDueReport.WindowTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                        rptAMISDueReport.ReportTitle = "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF
                        'rptAMISDueReport.Formulas(10) = "JDATE = '" & dtpAsOF.Value & "'"
                        'PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "{AMIS_AR.JDATE} <= CDate('" & dtpAsOF & "')", DMIS_REPORT_Connection, 1
                        PrintSQLReport rptAMISDueReport, AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", "", DMIS_REPORT_Connection, 1
                    Else
                        Me.WindowState = vbMaximized
                        CRViewer1.Height = Me.Height - 800
                        CRViewer1.Width = Me.Width
                        CRViewer1.ZOrder 0
                        Set rptApp = New CRAXDRT.Application

                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "DGI" Then
                            If cboCOBAcctName.Text = "ALL" Then
                                If optDetailed = True Then
                                    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", 1)
                                Else
                                    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUPBYCUSTOMER.Rpt", 1)
                                End If
                            Else
                                If optDetailed = True Then
                                    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUPACCOUNT.Rpt", 1)
                                Else
                                    Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUPACCOUNTBYCUSTOMER.Rpt", 1)
                                End If
                            End If
                        Else
                            Set rptRep = rptApp.OpenReport(AMIS_REPORT_PATH & "DueReports\ARAGINGREPORTGROUP.Rpt", 1)
                        End If

                        rptRep.DiscardSavedData
                        rptRep.ParameterFields.GetItemByName("CompanyName").AddCurrentValue COMPANY_NAME
                        rptRep.ParameterFields.GetItemByName("CompanyAddress").AddCurrentValue COMPANY_ADDRESS
                        rptRep.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "ACCOUNTS RECEIVABLE AGING REPORT AS OF: " & dtpAsOF

                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "DGI" Then
                            If cboCOBAcctName = "ALL" Then
                                Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                                Set crSections = rptRep.Sections
                                For ellaine = 1 To crSections.Count
                                    Set crSection = crSections.Item(ellaine)
                                    Set crRepObjs = crSection.ReportObjects
                                    For j = 1 To crRepObjs.Count
                                        If crRepObjs.Item(j).Kind = crSubreportObject Then
                                            Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                            If ellaine = 7 Then
                                                'Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                                'Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                            Else
                                                Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                                Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                            End If
                                        End If
                                    Next
                                Next
                            Else
                                Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                                Call rptRep.ParameterFields(5).AddCurrentValue(txtCOBAcctNo.Text)
                                Set crSections = rptRep.Sections
                                For ellaine = 1 To crSections.Count
                                    Set crSection = crSections.Item(ellaine)
                                    Set crRepObjs = crSection.ReportObjects
                                    For j = 1 To crRepObjs.Count
                                        If crRepObjs.Item(j).Kind = crSubreportObject Then
                                            Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                            If ellaine = 7 Then
                                                'Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                                'Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                            Else
                                                Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                                Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                            End If
                                        End If
                                    Next
                                Next
                            End If
                        Else
                            Call rptRep.ParameterFields(4).AddCurrentValue(CDate(dtpAsOF))
                            Set crSections = rptRep.Sections
                            For ellaine = 1 To crSections.Count
                                Set crSection = crSections.Item(ellaine)
                                Set crRepObjs = crSection.ReportObjects
                                For j = 1 To crRepObjs.Count
                                    If crRepObjs.Item(j).Kind = crSubreportObject Then
                                        Set crSubReport = rptRep.OpenSubreport(crRepObjs.Item(j).SubreportName)
                                        If ellaine = 7 Then
                                            'Call crSubReport.ParameterFields(5).ClearCurrentValueAndRange
                                            'Call crSubReport.ParameterFields(5).AddCurrentValue(CDate(dtpAsOF))
                                        Else
                                            Call crSubReport.ParameterFields(1).ClearCurrentValueAndRange
                                            Call crSubReport.ParameterFields(1).AddCurrentValue(CDate(dtpAsOF))
                                        End If
                                    End If
                                Next
                            Next
                        End If
                        With CRViewer1
                            .ReportSource = rptRep
                            .DisplayGroupTree = False
                            .DisplayTabs = False
                            .DisplayToolbar = True
                            .ViewReport
                        End With

                        Set rptApp = Nothing
                        Set rptRep = Nothing
                    End If
                    LogAudit "V", "ACCOUNTS RECEIVABLE AGING REPORT", "As of: " & dtpAsOF
                End If
            End If
        End If
    End If
    'getlastdate
    Exit Sub

ErrorCode:
    ShowVBError
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
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    SQL = "SELECT DESCRIPTION from AMIS_chartaccount where left(acctcode,5)='11-02' or left(acctcode,5)='11-03' ORDER BY DESCRIPTION"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    cboacctcode.Clear

    Do While Not RS.EOF
        cboacctcode.AddItem (RS!DESCRIPTION)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub
Function ReturnAccountCode(Xacct_desc As String)
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

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
    Dim rsHeader                                            As New ADODB.Recordset
    Dim rsdetail                                            As New ADODB.Recordset
    Dim BALANCE                                             As Double
    Dim totalpayment                                        As Double
    Dim CRJVoucher                                          As String
    Dim Reference                                           As String
    Dim SystemRemarks                                       As String
    Dim CRJInvoiceno                                        As String
    Dim CRJInvoicetype                                      As String
    Dim AMOUNT2PAY                                          As Double
    Dim invoicedate                                         As String
    Dim THECSJ                                              As String
    Dim theHDInvoice                                        As String
    Dim rscountx                                            As ADODB.Recordset
    Dim CustomerCode
    Dim RSCOUNT_ME                                          As New ADODB.Recordset
    Dim BILANG                                              As Integer
    Dim Counter_check                                       As New ADODB.Recordset
    Dim countDuplicated                                     As New ADODB.Recordset
    Dim Validate                                            As New ADODB.Recordset
    Dim CounterCheck_HD                                     As New ADODB.Recordset
    THECSJ = "CSJ"

    ProcessCRJmaliangcodesaCRJtoSJ

    gconDMIS.Execute ("DELETE FROM AMIS_AR")
    gconDMIS.Execute ("Update AMIS_CRJ_DETAIL set status = 'P'")
    'ValidateDetail
    Transfer_SalesJournal
    Me.Caption = "Loading Transaction.."

    Dim ARNIE                                               As New ADODB.Recordset
    'Set rsHeader = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
     '                                "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")

    'Set rsHeader = gconDMIS.Execute("SELECT DISTINCT dbo.AMIS_Journal_HD.VoucherNo,dbo.AMIS_Journal_HD.jdate, dbo.AMIS_Journal_HD.Status, dbo.AMIS_Journal_HD.JType,dbo.AMIS_Journal_HD.CustomerCode AS SJ_CustomerCode, dbo.AMIS_Journal_HD.InvoiceType, dbo.AMIS_Journal_HD.InvoiceNo,dbo.AMIS_Journal_HD.InvoiceDate as XInvoiceDate , dbo.AMIS_Journal_HD.InvoiceAmt, dbo.AMIS_Journal_HD.AmountToPay, dbo.AMIS_Journal_HD.AmountPaid,dbo.AMIS_Journal_Det.Acct_Code AS acct_code, dbo.AMIS_Journal_Det.Acct_Name, dbo.AMIS_Journal_Det.Debit AS Detdebit FROM dbo.AMIS_Journal_HD LEFT OUTER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-03') AND (dbo.AMIS_Journal_HD.JType = 'SJ' OR dbo.AMIS_Journal_HD.JType = 'COB' OR dbo.AMIS_Journal_HD.JType = 'CCM' OR dbo.AMIS_Journal_HD.JType = '" & THECSJ & _
     '                                "') and dbo.AMIS_Journal_HD.jdate <= " & N2Str2Null(dtpAsOF) & " and dbo.AMIS_Journal_HD.status ='P' and dbo.AMIS_Journal_HD.CustomerCode = 'G00041' ORDER BY dbo.AMIS_Journal_HD.VoucherNo")

    'Set rsHeader = gconDMIS.Execute("Select * from AMIS_AR_HD where jdate <= " & N2Str2Null(dtpAsOF) & " and SJ_customercode='f00015'")
    Set rsHeader = gconDMIS.Execute("Select * from AMIS_AR_HD where jdate <= " & N2Str2Null(dtpAsOF) & "")



    Dim LNGX                                                As Long
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

        Reference = Null2String(rsHeader!JTYPE) + "-" + Null2String(rsHeader!VOUCHERNO)
        'Invoicedate = Null2String(rsHeader!xInvoicedate)
        invoicedate = Null2String(rsHeader!JDATE)
        theHDInvoice = Null2String(rsHeader!INVOICENO)



        If (rsHeader!JTYPE) = "SJ" Then
            AMOUNT2PAY = DebitTotalAmount(Null2String(rsHeader!VOUCHERNO), Null2String(rsHeader!JTYPE), Null2String(rsHeader!ACCT_CODE))
        ElseIf (rsHeader!JTYPE) = "CCM" Then
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmnt) * (-1)    'to deduct to total payment
            'AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt)
        Else
            'AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmt)
            AMOUNT2PAY = N2Str2Zero(rsHeader!InvoiceAmnt)
        End If

        '     If Reference = "SJ-004397" Then Stop ' Yaon dgde
        '     If Reference = "SJ-004399" Then Stop
        'If Reference = "SJ-004605" Then Stop


        Set RSCOUNT_ME = gconDMIS.Execute("Select count(*) from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "' and status = 'P' and jdate <=" & N2Str2Null(dtpAsOF))
        BILANG = RSCOUNT_ME(0)
        If BILANG > 1 Then
            CustomerCode = ReturnCustomerCode(Null2String(rsHeader!VOUCHERNO), Null2String(rsHeader!JTYPE))
            Set rsdetail = gconDMIS.Execute("Select invoiceno,invoicetype,ISNULL(invoiceamount,0) AS INVOICEAMOUNT,voucherno,jdate,customercode as CRJ_customercode ,SJ_VOUCHERNO , J_CLASS,ID from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "'  and  detail_status <> 'Y'  and Status ='P' and jdate <=" & N2Str2Null(dtpAsOF))
        Else
            Set rsdetail = gconDMIS.Execute("Select invoiceno,invoicetype,ISNULL(invoiceamount,0) AS INVOICEAMOUNT,voucherno,jdate,customercode as CRJ_customercode ,SJ_VOUCHERNO , J_CLASS,ID from AMIS_crjdetail_total where invoicetype='" & Null2String(rsHeader!InvoiceType) & "' and invoiceno='" & rsHeader!INVOICENO & "' and status = 'P' and jdate <=" & N2Str2Null(dtpAsOF))
        End If

        LNGX = 0
        If Not rsdetail.EOF And Not rsdetail.BOF Then
            If Null2String(rsdetail!SJ_VOUCHERNO) <> "" Then
                Set rscountx = gconDMIS.Execute("select count(*) from amis_journal_det where jtype='SJ' AND VOUCHERNO=" & N2Str2Null(rsdetail!SJ_VOUCHERNO) & " and LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) in ('11-02', '11-03') ")
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


                        Dim CRJ_details                     As New ADODB.Recordset
                        Dim SJ_header                       As New ADODB.Recordset
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
                            'If CRJ_details!voucherno = "004249" Then Stop
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
                    '
                Else
                    If (Null2String(rsHeader!SJ_CustomerCode) = Null2String(rsdetail!CRJ_customercode) And Null2String(rsHeader!ACCT_CODE) = Null2String(rsdetail!J_CLASS)) Then
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

                        If Null2String(rsHeader!ACCT_CODE) = Null2String(rsdetail!J_CLASS) Then

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
        'If BALANCE <> 0 And (rsHeader!jtype) = "CCM" Then
        If BALANCE = 0 Then
            gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET AR_BALANCE = 0, AR_DATEGEN = " & N2Date2Null(dtprocess) & " WHERE VOUCHERNO = " & N2Str2Null(rsHeader!VOUCHERNO) & " AND JTYPE = " & N2Str2Null(rsHeader!JTYPE) & " AND CUSTOMERCODE = " & N2Str2Null(rsHeader!SJ_CustomerCode) & "")
        Else
            '             gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET AR_BALANCE = '" & ToDoubleNumber(BALANCE) & "', AR_DATEGEN = " & N2Date2Null(dtprocess) & " WHERE VOUCHERNO = " & N2Str2Null(rsHeader!VOUCHERNO) & " AND JTYPE = " & N2Str2Null(rsHeader!jtype) & " AND CUSTOMERCODE = " & N2Str2Null(rsHeader!SJ_CustomerCode) & "")
        End If
        'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        If rsHeader!JDATE <= dtprocess.Value Then


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
    'gconDMIS.Execute ("update AMIS_AR SET LASTUPDATED='" & dtpAsOF & "'")
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

    Dim rsZERO                                              As ADODB.Recordset
    Dim rsPER_CUSTOMER                                      As ADODB.Recordset
    Dim xCUST_KOD                                           As String

    'Set rsZERO = gconDMIS.Execute("Select CUSTOMERCODE from AMIS_AR where customercode = 'A00069'")
    Set rsZERO = gconDMIS.Execute("Select CUSTOMERCODE from AMIS_AR GROUP BY CUSTOMERCODE")
    If Not rsZERO.EOF And Not rsZERO.BOF Then
        'If rsZERO!CustomerCode = "A00069" Then
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
    Set RS = gconDMIS.Execute("SELECT MAX(JDATE) AS JDATE FROM (SELECT MAX(JDATE)AS JDATE FROM AMIS_AR UNION SELECT MAX(JDATE) AS JDATE FROM AMIS_DETAIL)T")
    If Not (RS.EOF Or RS.BOF) Then
        If Null2String(RS!JDATE) = "" Then
            Label9.Caption = "PLEASE GENERATE AR DATA"
            lblDate.Visible = False
        Else
            Label9.Visible = True
            lblDate.Visible = True
            lblDate.Caption = Null2String(RS!JDATE)
            dtpAsOF.Value = Null2String(RS!JDATE)
            dtprocess.Value = Null2String(RS!JDATE)
        End If
    Else
        lblDate.Visible = False
        lblDate.Caption = "PLEASE GENERATE AR DATA"
        Label9.Visible = False
    End If
    Set RS = Nothing
End Sub
Function DebitTotalAmount(xVOUCHER As String, xJType As String, Optional ByVal xARAcount As String) As Double
    Dim rsdetail                                            As New ADODB.Recordset
    Dim rscount                                             As ADODB.Recordset

    Set rscount = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & xVOUCHER & _
                                   "' AND JTYPE = '" & xJType & "' AND (LEFT(ACCT_CODE,5)='11-02' OR LEFT(ACCT_CODE,5)='11-03')")

    If rscount.Fields(0).Value > 1 Then
        Set rsdetail = gconDMIS.Execute("SELECT SUM(DEBIT) AS DEBIT,SUM(CREDIT) AS CREDIT  FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & xVOUCHER & _
                                        "' AND JTYPE = '" & xJType & "' AND ACCT_CODE='" & xARAcount & "' GROUP BY ACCT_CODE")

    Else
        Set rsdetail = gconDMIS.Execute("SELECT DEBIT,CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & xVOUCHER & _
                                        "' AND JTYPE = '" & xJType & "' AND (LEFT(ACCT_CODE,5)='11-02' OR LEFT(ACCT_CODE,5)='11-03')")
    End If


    DebitTotalAmount = 0
    If Not (rsdetail.EOF And rsdetail.BOF) Then
        rsdetail.MoveFirst
        Do While Not rsdetail.EOF
            If rsdetail!Debit = 0 Then
                DebitTotalAmount = DebitTotalAmount + NumericVal(rsdetail!Credit) * (-1)
            Else
                DebitTotalAmount = DebitTotalAmount + NumericVal(rsdetail!Debit)
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
    Dim RSCRJ                                               As New ADODB.Recordset
    Dim TheAMOUNT                                           As Double
    Dim Reference                                           As String
    Dim rsCRJ_Detail                                        As New ADODB.Recordset

    Dim SQL                                                 As String
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
            If RSCRJ!Debit = 0 Then
                TheAMOUNT = NumericVal(RSCRJ!Credit)
            Else
                TheAMOUNT = NumericVal(RSCRJ!Debit)
            End If

            Set rsCRJ_Detail = gconDMIS.Execute("SELECT VoucherNo FROM AMIS_CRJ_DETAIL WHERE VoucherNo ='" & Null2String(RSCRJ!VOUCHERNO) & "'")
            If rsCRJ_Detail.EOF Or rsCRJ_Detail.BOF Then
                If RSCRJ!JDATE <= dtprocess.Value Then

                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values(" & _
                                      " NULL,'" & Reference & "'," & N2Str2Null(RSCRJ!InvoiceType) & "," & N2Str2Null(RSCRJ!INVOICENO) & "," & N2Str2Null("") & _
                                      "'," & N2Str2Null(RSCRJ!JDATE) & ",'" & RSCRJ!CCODE & "','" & NumericVal(0) & "','" & NumericVal(RSCRJ!InvoiceAmt) & _
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
    Dim RSCRJ                                               As New ADODB.Recordset
    Dim TheAMOUNT                                           As Double
    Dim Reference                                           As String
    Dim rsCRJ_Detail                                        As New ADODB.Recordset
    'Set RSCRJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS VOUCHERNO,dbo.AMIS_Journal_HD.JType as Jtype,dbo.AMIS_Journal_HD.CustomerCode as CCode,dbo.AMIS_Journal_HD.JDate as Jdate,dbo.AMIS_Journal_HD.InvoiceAmt as InvoiceAmt, dbo.AMIS_Journal_Det.Debit as Debit,dbo.AMIS_Journal_Det.credit as Credit,dbo.AMIS_Journal_HD.InvoiceNo as ORnum, dbo.AMIS_Journal_Det.Acct_Code as Acct_code,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE dbo.AMIS_Journal_HD.JDate < = '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND (dbo.AMIS_Journal_HD.JType = 'CRJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '21-07') order by voucherno")

    'Set RSCRJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_HD.VoucherNo AS VOUCHERNO,dbo.AMIS_Journal_HD.JType as Jtype,dbo.AMIS_Journal_HD.CustomerCode as CCode,dbo.AMIS_Journal_HD.JDate as Jdate,dbo.AMIS_Journal_HD.InvoiceAmt as InvoiceAmt, dbo.AMIS_Journal_Det.Debit as Debit,dbo.AMIS_Journal_Det.credit as Credit,dbo.AMIS_Journal_HD.InvoiceNo as ORnum, dbo.AMIS_Journal_Det.Acct_Code as Acct_code,dbo.AMIS_Journal_HD.status FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.JType = dbo.AMIS_Journal_Det.JType WHERE dbo.AMIS_Journal_HD.JDate < = '" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND (dbo.AMIS_Journal_HD.JType = 'CRJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02' OR LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '21-07') order by voucherno")

    Dim SQL                                                 As String
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
            Dim rsCHECK_BALANCE                             As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & Null2String(RSCRJ!VOUCHERNO) & "' AND JTYPE = '" & RSCRJ!JTYPE & "' ")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                If rsCHECK_BALANCE!AR_BALANCE <> 0 Then
                    Reference = "CRJ" + "-" + Null2String(RSCRJ!VOUCHERNO)
                    If RSCRJ!Debit = 0 Then
                        TheAMOUNT = NumericVal(RSCRJ!Credit)
                    Else
                        TheAMOUNT = NumericVal(RSCRJ!Debit)
                    End If

                    Set rsCRJ_Detail = gconDMIS.Execute("SELECT VoucherNo FROM AMIS_CRJ_DETAIL WHERE VoucherNo ='" & Null2String(RSCRJ!VOUCHERNO) & "'")
                    If rsCRJ_Detail.EOF Or rsCRJ_Detail.BOF Then
                        If RSCRJ!JDATE <= dtprocess.Value Then
                            If RSCRJ!VOUCHERNO = "001277" Then
                                '
                            End If

                            gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCRJ!CCODE & _
                                              "'," & N2Str2Null(RSCRJ!VOUCHERNO) & "," & N2Str2Null(RSCRJ!InvoiceAmt) & _
                                              "," & TheAMOUNT & "," & N2Str2Null(RSCRJ!ACCT_CODE) & "," & N2Str2Null(RSCRJ!JDATE) & ",'NL')")

                            gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                              "'," & N2Str2Null(RSCRJ!VOUCHERNO) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                              "'," & N2Str2Null(RSCRJ!JDATE) & ",'" & RSCRJ!CCODE & "','" & NumericVal(0) & "','" & NumericVal(RSCRJ!InvoiceAmt) & _
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
    Dim amount                                              As Double
    Dim Reference                                           As String
    Dim RSCDJ                                               As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CDJ'")
    Set RSCDJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.CREDIT as CREDITAmount,dbo.AMIS_Journal_Det.Debit as DebitAmount, dbo.AMIS_Journal_HD.VoucherNo as voucherno, dbo.AMIS_Journal_HD.VendorCode as VCode,dbo.AMIS_Journal_HD.JDate as jdate,dbo.AMIS_Journal_HD.status,dbo.AMIS_Journal_HD.AR_DATEGEN FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.jtype = dbo.AMIS_Journal_Det.jtype  WHERE (dbo.AMIS_Journal_HD.JType = 'CDJ') AND (LEFT(dbo.AMIS_Journal_Det.Acct_Code, 5) = '11-02') and dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))")
    Me.Caption = "Loading CDJ Having AR.."
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCDJ.RecordCount
    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        Do While Not RSCDJ.EOF
            'UPDATED BY: JUN
            Dim rsCHECK_BALANCE                             As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & RSCDJ!VOUCHERNO & "' AND JTYPE = '" & RSCDJ!JTYPE & "'")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                If rsCHECK_BALANCE!AR_BALANCE <> 0 Then
                    DoEvents
                    Reference = "CDJ" + "-" + Null2String(RSCDJ!VOUCHERNO)
                    If RSCDJ!debitAmount = 0 Then
                        amount = NumericVal(RSCDJ!creditamount) * (-1)    ' Bawas ni
                    Else
                        amount = NumericVal(RSCDJ!debitAmount)
                    End If
                    If RSCDJ!JDATE <= dtprocess.Value Then
                        gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCDJ!VCode & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & "," & N2Str2Null(amount) & _
                                          "," & N2Str2Null("") & "," & N2Str2Null(RSCDJ!ACCT_CODE) & "," & N2Str2Null(RSCDJ!JDATE) & ",'CDJ')")

                        gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & ",'" & N2Str2Null("") & "','" & N2Str2Null("") & _
                                          "'," & N2Str2Null(RSCDJ!JDATE) & ",'" & RSCDJ!VCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
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
    Dim Reference                                           As String
    Dim amount                                              As Double
    Dim RS                                                  As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='NOINV'")
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJDETAIL_TOTAL.INVOICEAMOUNT, dbo.AMIS_CRJDETAIL_TOTAL.INVOICENO, dbo.AMIS_CRJDETAIL_TOTAL.INVOICETYPE, " & _
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

    If Not RS.EOF And Not RS.BOF Then
        Me.Caption = "Loading CRJ with blank invoice.."
        ProgressBar1.Value = 0
        ProgressBar1.Max = RS.RecordCount
        If Not (RS.EOF And RS.BOF) Then
            Do While Not RS.EOF
                amount = NumericVal(RS!invoiceamount)
                Reference = "CRJ-" + Null2String(RS!VOUCHERNO)
                If (RS!JDATE) <= dtprocess.Value Then
                    gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                      "'," & N2Str2Null(RS!VOUCHERNO) & ",'XXX','XXX'," & N2Str2Null(RS!JDATE) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                      "','" & NumericVal(amount) * (-1) & "','" & Null2String(RS!ACCT_CODE) & "','NOINV')")
                End If
                DoEvents
                lblRef.Caption = RS!VOUCHERNO
                ProgressBar1.Value = ProgressBar1.Value + 1
                Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                RS.MoveNext
            Loop
        End If
    End If
    Set RS = Nothing
End Sub
Sub ProcessCRJmaliangcodesaCRJtoSJ()
    Dim Reference                                           As String
    Dim test                                                As New ADODB.Recordset
    Dim amount                                              As Double
    'Update By BTT : to find the AR na ang CRJ nya iba ang ang account sa SJ pero tama ang link
    Dim RS                                                  As New ADODB.Recordset
    Me.Caption = "Validating Ad hoc data.."
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CRJWAC'")
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
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

    If Not (RS.EOF And RS.BOF) Then
        ProgressBar1.Value = 0
        ProgressBar1.Max = RS.RecordCount
        Do While Not RS.EOF
            Dim rsCheckBalance                              As ADODB.Recordset
            'UPDATED BY: JUN
            Set rsCheckBalance = gconDMIS.Execute("SELECT AR_BALANCE from AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & Null2String(RS!VOUCHERNO) & "' AND JTYPE = '" & Null2String(RS!JTYPE) & "'")
            If Not rsCheckBalance.EOF And Not rsCheckBalance.BOF Then
                'If rsCheckBalance!AR_BALANCE <> 0 Then
                If rsCheckBalance!AR_BALANCE <> 0 Then
                    Reference = "CRJ-" + Null2String(RS!VOUCHERNO)
                    If NumericVal(RS!Debit) > 0 Then
                        amount = NumericVal(RS!invoiceamount)
                    Else
                        amount = NumericVal(RS!invoiceamount) * (-1)
                    End If


                    Set test = gconDMIS.Execute("SELECT   AMIS_Journal_HD.VoucherNo " & _
                                                "From AMIS_Journal_HD Inner Join AMIS_Journal_Det ON AMIS_Journal_HD.JType = AMIS_Journal_Det.JType AND " & _
                                                "AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo " & _
                                                "Where (AMIS_Journal_HD.JType = 'CRJ') AND" & _
                                                "(LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') and AMIS_Journal_HD.VoucherNo = '" & RS!VOUCHERNO & "' AND ((AMIS_Journal_HD.AR_BALANCE IS NULL OR AMIS_Journal_HD.AR_BALANCE <> 0) AND (AMIS_Journal_HD.AR_DATEGEN IS NULL OR AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "'))" & _
                                                "GROUP BY AMIS_Journal_HD.VoucherNo " & _
                                                "HAVING COUNT(*) > 1  ")

                    If (test.EOF And test.BOF) Then


                        If RS!JDATE <= dtprocess.Value Then

                            '                If amount > 2330 And amount < 2337 Then
                            '
                            '                End If

                            If (RS!CR_ACCCODE) <> Null2String(RS!SJ_ACTCODE) And RS!CustomerCode = RS!CRJ_COUSTOMERCODE Then

                                gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                                  "'," & N2Str2Null(RS!VOUCHERNO) & ",'" & RS!InvoiceType & "','" & RS!INVOICENO & "'," & N2Str2Null(RS!JDATE) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                                  "','" & NumericVal(amount) & "','" & Null2String(RS!CR_ACCCODE) & "','CRJWAC')")
                            End If
                        End If
                    End If
                    'lblRef.Caption = RSCDJ!voucherno
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RS.MoveNext
                Else
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RS.MoveNext
                End If
            End If
        Loop
    End If
    Set RS = Nothing
    Set rsCheckBalance = Nothing
    'Temp
    'gconDMIS.Execute ("update AMIS_AR set balance = 0 where CRJvoucherno = '004335' and invoicetype = 'VI' and customercode = 'P00172'")
End Sub
Sub ProcessCRJWithClearingAccount()
    Dim test                                                As New ADODB.Recordset
    Dim Reference                                           As String
    Dim amount                                              As Double
    ' CRJWCA : CRJ - with clrearing account and tama ang link
    Me.Caption = "Validating Ad hoc Data.."
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='CRJWCA'")
    Dim RS                                                  As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT dbo.AMIS_CRJ_Detail.INVOICEAMOUNT, dbo.AMIS_CRJ_Detail.INVOICENO, dbo.AMIS_CRJ_Detail.INVOICETYPE, dbo.AMIS_CRJ_Detail.INVOICEDATE, " & _
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
    If Not RS.EOF And Not RS.BOF Then
        ProgressBar1.Value = 0
        ProgressBar1.Max = RS.RecordCount
        Do While Not RS.EOF
            'UPDATED BY: JUN
            Dim rsCHECK_AR_BALANCE                          As ADODB.Recordset
            Set rsCHECK_AR_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & RS!VOUCHERNO & "' AND JTYPE = '" & RS!JTYPE & "'")
            If Not rsCHECK_AR_BALANCE.EOF And Not rsCHECK_AR_BALANCE.BOF Then
                If rsCHECK_AR_BALANCE!AR_BALANCE <> 0 Then
                    Reference = "CRJ-" + Null2String(RS!VOUCHERNO)
                    amount = NumericVal(RS!invoiceamount)

                    Set test = gconDMIS.Execute(" SELECT   AMIS_Journal_HD.VoucherNo " & _
                                                "From AMIS_Journal_HD Inner Join AMIS_Journal_Det ON AMIS_Journal_HD.JType = AMIS_Journal_Det.JType AND " & _
                                                "AMIS_Journal_HD.VoucherNo = AMIS_Journal_Det.VoucherNo " & _
                                                "Where (AMIS_Journal_HD.JType = 'CRJ')AND AMIS_Journal_HD.AR_BALANCE <> 0 AND " & _
                                                "(LEFT(AMIS_Journal_Det.Acct_Code, 5) = '11-02') and AMIS_Journal_HD.VoucherNo = '" & RS!VOUCHERNO & "'" & _
                                                "GROUP BY AMIS_Journal_HD.VoucherNo " & _
                                                "HAVING COUNT(*) > 1  ")


                    If Not (test.EOF And test.BOF) Then


                    Else
                        If RS!JDATE <= dtprocess.Value Then
                            If Null2String(RS!CRJ_ACCTCODE) <> Null2String(RS!SJ_ACCTCODE) And RS!CustomerCode = RS!CRJ_COUSTOMERCODE Then
                                gconDMIS.Execute ("insert into amis_ar(sjvoucherno,crjvoucherno,invoicetype,invoiceno,invoicedate,customercode,amount_topay,amount_paid,balance,Account_code,SystemRemark) values('" & Reference & _
                                                  "'," & N2Str2Null(RS!VOUCHERNO) & ",'" & RS!InvoiceType & "','" & RS!INVOICENO & "'," & N2Str2Null(RS!JDATE) & ",'" & RS!CustomerCode & "','" & NumericVal(0) & "','" & NumericVal(amount) & _
                                                  "','" & NumericVal(amount) * (1) & "','" & Null2String(RS!SJ_ACCTCODE) & "','CRJWCA')")
                            End If
                        End If
                    End If
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RS.MoveNext
                Else
                    DoEvents
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
                    RS.MoveNext
                End If
            End If
        Loop
    End If
    Set RS = Nothing
End Sub
Function ReturnCustomerCode(TheVoucherno As String, theJtype As String) As String
    Dim RS                                                  As New ADODB.Recordset
    Dim GetDealer                                           As New ADODB.Recordset
    Set RS = gconDMIS.Execute("Select customercode from AMIS_journal_hd where voucherno='" & TheVoucherno & "' and jtype = '" & theJtype & "'")
    If Not RS.EOF And Not RS.BOF Then
        ReturnCustomerCode = Null2String(RS!CustomerCode)
        'Set GetDealer = gconDMIS.Execute("Select Acctname from all_customer_table where cuscde='" & ReturnCustomerCode & "'")
        'If Not GetDealer.EOF And Not GetDealer.BOF Then
        '    Dealer = Null2String(GetDealer!Acctname)
        'End If
    Else
        ReturnCustomerCode = N2Str2Null("")
    End If
    Set RS = Nothing
End Function
Sub Transfer_SalesJournal()
    Dim Bernard                                             As New ADODB.Recordset
    Dim CounterCheck_F                                      As New ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xSTATUS                                             As String
    Dim xJType                                              As String
    Dim xSJ_CustomerCode                                    As String
    Dim xInvoicedate                                        As String
    Dim xInvoiceType                                        As String
    Dim xInvoiceAmnt                                        As String
    Dim xamounttopay                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xAMOUNTPAID                                         As Double
    Dim xInvoiceNo                                          As String
    Dim xdebit                                              As Double
    Dim THECSJ                                              As String
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
            xVOUCHERNO = N2Str2Null(Bernard!VOUCHERNO)
            xJdate = N2Date2Null(Bernard!JDATE)
            xSTATUS = N2Str2Null(Bernard!Status)
            xJType = N2Str2Null(Bernard!JTYPE)
            xSJ_CustomerCode = N2Str2Null(Bernard!SJ_CustomerCode)
            xInvoiceType = N2Str2Null(Bernard!InvoiceType)
            xInvoiceNo = N2Str2Null(Bernard!INVOICENO)
            xInvoiceAmnt = N2Str2Null(Bernard!InvoiceAmt)
            xInvoicedate = N2Str2Null(Bernard!xInvoicedate)
            xamounttopay = N2Str2Zero(Bernard!AMOUNTTOPAY)
            xAMOUNTPAID = N2Str2Zero(Bernard!AMOUNTPAID)
            xACCT_CODE = N2Str2Null(Bernard!ACCT_CODE)
            xdebit = N2Str2Zero(Bernard!detdebit)

            Set CounterCheck_F = gconDMIS.Execute("Select count(*) from AMIS_AR_HD where voucherno = " & xVOUCHERNO & " and jtype =" & xJType & "")
            If Not CounterCheck_F(0) = 1 Then
                If Bernard!JDATE <= dtprocess.Value Then
                    gconDMIS.Execute ("Insert into AMIS_AR_HD(Voucherno,jdate,status,jtype,Sj_customercode,invoicetype,invoiceno,invoicedate,invoiceamnt,amounttopay,amountpaid,acct_code,debit) values(" & xVOUCHERNO & _
                                      "," & xJdate & "," & xSTATUS & "," & xJType & _
                                      "," & xSJ_CustomerCode & "," & xInvoiceType & "," & xInvoiceNo & "," & xInvoicedate & "," & xInvoiceAmnt & _
                                      "," & xamounttopay & "," & xAMOUNTPAID & "," & xACCT_CODE & "," & xdebit & ")")



                End If
            End If
            DoEvents
            lblRef.Caption = xVOUCHERNO
            ProgressBar1.Value = ProgressBar1.Value + 1
            Label22(0).Caption = Round((ProgressBar1.Value / ProgressBar1.Max * 100), 0) & "%"
            Bernard.MoveNext
        Loop
    End If
    Set Bernard = Nothing
End Sub
Sub ValidateDetail()
    Dim RSXX                                                As New ADODB.Recordset

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
    Dim amount                                              As Double
    Dim Reference                                           As String
    Dim RSCDJ                                               As New ADODB.Recordset
    gconDMIS.Execute ("delete from AMIS_CRJ_nodetail where proccess_type='APJ'")

    Set RSCDJ = gconDMIS.Execute("SELECT dbo.AMIS_Journal_Det.Acct_Code as acct_code, dbo.AMIS_Journal_Det.CREDIT as CREDITAmount,dbo.AMIS_Journal_Det.Debit as DebitAmount, dbo.AMIS_Journal_HD.VoucherNo as voucherno, dbo.AMIS_Journal_HD.VendorCode as VCode,dbo.AMIS_Journal_HD.JDate as jdate,dbo.AMIS_Journal_HD.status,dbo.AMIS_Journal_HD.AR_DATEGEN FROM dbo.AMIS_Journal_HD INNER JOIN dbo.AMIS_Journal_Det ON dbo.AMIS_Journal_HD.VoucherNo = dbo.AMIS_Journal_Det.VoucherNo AND dbo.AMIS_Journal_HD.jtype = dbo.AMIS_Journal_Det.jtype  WHERE (dbo.AMIS_Journal_HD.JType = 'APJ') AND dbo.AMIS_Journal_Det.Acct_Code = '11-02000-00' and dbo.AMIS_Journal_HD.JDate <='" & dtprocess.Value & "' and dbo.AMIS_Journal_HD.status = 'P' AND ((dbo.AMIS_Journal_HD.AR_BALANCE IS NULL OR dbo.AMIS_Journal_HD.AR_BALANCE <> 0) AND (dbo.AMIS_Journal_HD.AR_DATEGEN IS NULL OR dbo.AMIS_Journal_HD.AR_DATEGEN <= '" & dtprocess & "')) ")

    Me.Caption = "Loading AP Having AR.."
    ProgressBar1.Value = 0
    ProgressBar1.Max = RSCDJ.RecordCount
    If Not (RSCDJ.EOF And RSCDJ.BOF) Then
        Do While Not RSCDJ.EOF
            'UPDATED BY : JUN
            Dim rsCHECK_BALANCE                             As ADODB.Recordset
            Set rsCHECK_BALANCE = gconDMIS.Execute("SELECT AR_BALANCE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & RSCDJ!VOUCHERNO & "' AND JTYPE = '" & RSCDJ!JTYPE & "'")
            If Not rsCHECK_BALANCE.EOF And Not rsCHECK_BALANCE.BOF Then
                If RSCDJ!AR_BALANCE <> 0 Then
                    DoEvents
                    Reference = "CDJ" + "-" + Null2String(RSCDJ!VOUCHERNO)
                    If RSCDJ!debitAmount = 0 Then
                        amount = NumericVal(RSCDJ!creditamount) * (-1)    ' Bawas ni
                    Else
                        amount = NumericVal(RSCDJ!debitAmount)
                    End If
                    If RSCDJ!JDATE <= dtprocess.Value Then
                        gconDMIS.Execute ("INSERT INTO AMIS_CRJ_NoDetail(CUSTOMERCODE,CRJ_VOUCHERNO,ORAMOUNT,ORNUM,ACC_CODE,INVOICEDATE,Proccess_type) VALUES('" & RSCDJ!VCode & _
                                          "'," & N2Str2Null(RSCDJ!VOUCHERNO) & "," & N2Str2Null(amount) & _
                                          "," & N2Str2Null("") & "," & N2Str2Null(RSCDJ!ACCT_CODE) & "," & N2Str2Null(RSCDJ!JDATE) & ",'APJ')")

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

Private Sub Option1_Click()
    'frmARReportCustomer.Show
    picDetailedSum.Visible = True
    picDetailedSum.ZOrder 0
End Sub

Private Sub Option2_Click()
    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "DGI" Then
        picByAccount.Visible = True
        picByAccount.ZOrder 0
        Dim rsChartOfAccounts                               As ADODB.Recordset
        Set rsChartOfAccounts = New ADODB.Recordset
        rsChartOfAccounts.Open "SELECT ACCTCODE,DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) IN ('11-02','11-03') AND IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenKeyset
        If Not rsChartOfAccounts.EOF And Not rsChartOfAccounts.BOF Then
            cboCOBAcctName.AddItem "ALL"
            Do While Not rsChartOfAccounts.EOF
                cboCOBAcctName.AddItem rsChartOfAccounts!DESCRIPTION
                rsChartOfAccounts.MoveNext
            Loop
        End If
        Set rsChartOfAccounts = Nothing
    End If
End Sub

Private Sub picReport_Click()
'    ProcessCRJmaliangcodesaCRJtoSJ
'    ProcessCRJWithClearingAccount
End Sub

Private Sub Timer1_Timer()
    If Label11.Caption <> "" Then
        If Picture3.Visible = True Then
            Picture3.Visible = False
        Else
            Picture3.Visible = True
        End If
    End If
End Sub

Function GET_MAX_DATE() As String
    Dim rsGET_MAX_DATE                                      As ADODB.Recordset
    Set rsGET_MAX_DATE = New ADODB.Recordset
    rsGET_MAX_DATE.Open "SELECT * FROM (SELECT MAX(JDATE) AS JDATE FROM AMIS_AR)T WHERE JDATE IS NOT NULL ", gconDMIS, adOpenKeyset
    If Not rsGET_MAX_DATE.EOF And Not rsGET_MAX_DATE.BOF Then
        GET_MAX_DATE = Null2Date(rsGET_MAX_DATE!JDATE)
        'dtprocess.Value = GET_MAX_DATE
        Command1.Enabled = True
        Command2.Enabled = True
    Else
        dtpAsOF.Value = LOGDATE
        Command1.Enabled = False
        Command2.Enabled = False
        MessagePop InfoFriend, "Info", "No such Record!"
    End If
    Set rsGET_MAX_DATE = Nothing
End Function

Function Setacctcode(xDescription As String) As String
    Dim rsDescription                                       As ADODB.Recordset
    Set rsDescription = New ADODB.Recordset
    rsDescription.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE LEFT(ACCTCODE,5) IN ('11-02','11-03') AND IS_SCHEDULE_ACCNT=1 AND DESCRIPTION = '" & xDescription & "'", gconDMIS, adOpenKeyset
    If Not rsDescription.EOF And Not rsDescription.BOF Then
        Setacctcode = Null2String(rsDescription!AcctCode)
    End If
    Set rsDescription = Nothing
End Function







