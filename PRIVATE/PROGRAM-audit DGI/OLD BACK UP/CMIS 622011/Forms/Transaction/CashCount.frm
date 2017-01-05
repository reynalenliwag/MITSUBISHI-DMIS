VERSION 5.00
Begin VB.Form frmCMISCashCount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Count Entry"
   ClientHeight    =   6000
   ClientLeft      =   180
   ClientTop       =   795
   ClientWidth     =   8220
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CashCount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8220
   Begin VB.PictureBox picCashCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3360
      Picture         =   "CashCount.frx":08CA
      ScaleHeight     =   4665
      ScaleWidth      =   945
      TabIndex        =   77
      Top             =   840
      Width           =   945
      Begin VB.TextBox txtISANGLIBO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   0
         Text            =   "0"
         Top             =   0
         Width           =   885
      End
      Begin VB.TextBox txtLIMANGDAAN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   885
      End
      Begin VB.TextBox txtDALAWANGDAAN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   2
         Text            =   "0"
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtISANGDAAN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   3
         Text            =   "0"
         Top             =   1080
         Width           =   885
      End
      Begin VB.TextBox txtSINGKWENTA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   4
         Text            =   "0"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox txtBENTE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   5
         Text            =   "0"
         Top             =   1800
         Width           =   885
      End
      Begin VB.TextBox txtSAMPU 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   6
         Text            =   "0"
         Top             =   2160
         Width           =   885
      End
      Begin VB.TextBox txtLIMANGPISO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   7
         Text            =   "0"
         Top             =   2520
         Width           =   885
      End
      Begin VB.TextBox txtPISO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   8
         Text            =   "0"
         Top             =   2880
         Width           =   885
      End
      Begin VB.TextBox txtBENTESINKO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   9
         Text            =   "0"
         Top             =   3210
         Width           =   885
      End
      Begin VB.TextBox txtDYES 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   10
         Text            =   "0"
         Top             =   3600
         Width           =   885
      End
      Begin VB.TextBox txtSINKO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   11
         Text            =   "0"
         Top             =   3960
         Width           =   885
      End
      Begin VB.TextBox txtSENTIMO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   30
         TabIndex        =   12
         Text            =   "0"
         Top             =   4320
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   825
      Left            =   90
      ScaleHeight     =   825
      ScaleWidth      =   2835
      TabIndex        =   92
      Top             =   1860
      Width           =   2835
      Begin VB.TextBox txtCHANGE_FUND 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   945
         TabIndex        =   93
         Text            =   "0.00"
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Change Fund"
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
         TabIndex        =   94
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   6630
      Picture         =   "CashCount.frx":3606
      ScaleHeight     =   4665
      ScaleWidth      =   1515
      TabIndex        =   78
      Top             =   840
      Width           =   1515
      Begin VB.TextBox txt01total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   91
         Text            =   "0.00"
         Top             =   4320
         Width           =   1485
      End
      Begin VB.TextBox txt05total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   90
         Text            =   "0.00"
         Top             =   3960
         Width           =   1485
      End
      Begin VB.TextBox txt010total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   89
         Text            =   "0.00"
         Top             =   3600
         Width           =   1485
      End
      Begin VB.TextBox txt25total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   88
         Text            =   "0.00"
         Top             =   3240
         Width           =   1485
      End
      Begin VB.TextBox txt1total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   87
         Text            =   "0.00"
         Top             =   2880
         Width           =   1485
      End
      Begin VB.TextBox txt5total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   86
         Text            =   "0.00"
         Top             =   2520
         Width           =   1485
      End
      Begin VB.TextBox txt10total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   85
         Text            =   "0.00"
         Top             =   2160
         Width           =   1485
      End
      Begin VB.TextBox txt20total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   84
         Text            =   "0.00"
         Top             =   1800
         Width           =   1485
      End
      Begin VB.TextBox txt50total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   83
         Text            =   "0.00"
         Top             =   1440
         Width           =   1485
      End
      Begin VB.TextBox txt100total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   82
         Text            =   "0.00"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txt200total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   81
         Text            =   "0.00"
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txt500total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   80
         Text            =   "0.00"
         Top             =   360
         Width           =   1485
      End
      Begin VB.TextBox txt1000total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   330
         Left            =   0
         TabIndex        =   79
         Text            =   "0.00"
         Top             =   0
         Width           =   1485
      End
   End
   Begin VB.CommandButton cmdF4 
      Caption         =   "F4 - Update Cash Count"
      Height          =   345
      Left            =   90
      TabIndex        =   76
      ToolTipText     =   "Update Cash Count"
      Top             =   5580
      Width           =   2865
   End
   Begin VB.CommandButton cmdF10 
      Caption         =   "F10 - Print"
      Height          =   345
      Left            =   3000
      TabIndex        =   75
      ToolTipText     =   "Print"
      Top             =   5580
      Width           =   1425
   End
   Begin VB.CommandButton cmdF11 
      Caption         =   "F11 - Calculator"
      Height          =   345
      Left            =   4470
      TabIndex        =   74
      ToolTipText     =   "View Calculator"
      Top             =   5580
      Width           =   1905
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<<"
      Height          =   345
      Left            =   6420
      TabIndex        =   73
      ToolTipText     =   "Previous"
      Top             =   5580
      Width           =   825
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>"
      Height          =   345
      Left            =   7290
      TabIndex        =   72
      ToolTipText     =   "Next"
      Top             =   5580
      Width           =   825
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   2835
      TabIndex        =   26
      Top             =   3660
      Width           =   2835
      Begin VB.TextBox txtShortBy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   930
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Short by"
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
         TabIndex        =   28
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   825
      Left            =   90
      ScaleHeight     =   825
      ScaleWidth      =   2835
      TabIndex        =   23
      Top             =   2760
      Width           =   2835
      Begin VB.TextBox txtCashCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   930
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Count Total"
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
         TabIndex        =   25
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   795
      Left            =   90
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   20
      Top             =   4590
      Width           =   2835
      Begin VB.TextBox txtOverBy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   960
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label43 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Over by"
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
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   795
      Left            =   90
      ScaleHeight     =   795
      ScaleWidth      =   2835
      TabIndex        =   17
      Top             =   90
      Width           =   2835
      Begin VB.TextBox txtCUTDATE 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   930
         TabIndex        =   18
         Top             =   330
         Width           =   1755
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
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
         TabIndex        =   19
         Top             =   30
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   825
      Left            =   90
      ScaleHeight     =   825
      ScaleWidth      =   2835
      TabIndex        =   14
      Top             =   960
      Width           =   2835
      Begin VB.TextBox txtCashOnHand 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   930
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash on Hand"
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
         TabIndex        =   16
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.Label Label49 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   71
      Top             =   4770
      Width           =   255
   End
   Begin VB.Label Label48 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   70
      Top             =   4830
      Width           =   225
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "0.05"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   69
      Top             =   4830
      Width           =   555
   End
   Begin VB.Label Label46 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   68
      Top             =   4410
      Width           =   255
   End
   Begin VB.Label Label45 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   67
      Top             =   4470
      Width           =   225
   End
   Begin VB.Label Label44 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "0.10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   66
      Top             =   4470
      Width           =   555
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "1000.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5190
      TabIndex        =   65
      Top             =   870
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   64
      Top             =   870
      Width           =   225
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   63
      Top             =   810
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "500.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5340
      TabIndex        =   62
      Top             =   1230
      Width           =   705
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   61
      Top             =   1230
      Width           =   225
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   60
      Top             =   1170
      Width           =   255
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "200.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5370
      TabIndex        =   59
      Top             =   1590
      Width           =   675
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   58
      Top             =   1590
      Width           =   225
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   57
      Top             =   1530
      Width           =   255
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "100.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5340
      TabIndex        =   56
      Top             =   1950
      Width           =   705
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   55
      Top             =   1950
      Width           =   225
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   54
      Top             =   1890
      Width           =   255
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "50.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   53
      Top             =   2310
      Width           =   555
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   52
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   51
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "20.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   50
      Top             =   2670
      Width           =   555
   End
   Begin VB.Label Label20 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   49
      Top             =   2670
      Width           =   225
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   48
      Top             =   2610
      Width           =   255
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "10.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   47
      Top             =   3030
      Width           =   555
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   46
      Top             =   3030
      Width           =   225
   End
   Begin VB.Label Label24 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   45
      Top             =   2970
      Width           =   255
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "5.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   44
      Top             =   3390
      Width           =   555
   End
   Begin VB.Label Label26 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   43
      Top             =   3390
      Width           =   225
   End
   Begin VB.Label Label27 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   42
      Top             =   3330
      Width           =   255
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   41
      Top             =   3750
      Width           =   555
   End
   Begin VB.Label Label29 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   40
      Top             =   3750
      Width           =   225
   End
   Begin VB.Label Label30 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   39
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "0.25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   38
      Top             =   4110
      Width           =   555
   End
   Begin VB.Label Label32 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   37
      Top             =   4110
      Width           =   225
   End
   Begin VB.Label Label33 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   36
      Top             =   4050
      Width           =   255
   End
   Begin VB.Label Label39 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6300
      TabIndex        =   35
      Top             =   5130
      Width           =   255
   End
   Begin VB.Label Label40 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4650
      TabIndex        =   34
      Top             =   5190
      Width           =   225
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "0.01"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   5490
      TabIndex        =   33
      Top             =   5190
      Width           =   555
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash Count Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3060
      TabIndex        =   32
      Top             =   90
      Width           =   5055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Pc(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3060
      TabIndex        =   31
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4770
      TabIndex        =   30
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Total per Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6450
      TabIndex        =   29
      Top             =   450
      Width           =   1665
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   510
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmCMISCashCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCASH                                                            As ADODB.Recordset
Dim rsOFF_HD                                                          As ADODB.Recordset

Sub rsRefresh()
    Set rsCASH = New ADODB.Recordset
    Set rsCASH = gconDMIS.Execute("Select * from CMIS_Cash order by cutdate asc")
End Sub

Sub KIM_SUCKS()
    txtCashCount.Text = ToDoubleNumber(NumericVal(txt1000total.Text) + NumericVal(txt500total.Text) + NumericVal(txt200total.Text) + NumericVal(txt100total.Text) + NumericVal(txt50total.Text) + NumericVal(txt20total.Text) + NumericVal(txt10total.Text) + NumericVal(txt5total.Text) + NumericVal(txt1total.Text) + NumericVal(txt25total.Text) + NumericVal(txt010total.Text) + NumericVal(txt05total.Text) + NumericVal(txt01total.Text)):    'MyFirstLove
End Sub

Sub StoreMemvars()
    Dim ADVANCES_COLL                                                 As Double
    If Not rsCASH.EOF And Not rsCASH.BOF Then
        LABID.Caption = rsCASH!Id
        txtCutDate.Text = Null2String(rsCASH!CUTDATE)
        txtISANGLIBO.Text = N2Str2Zero(rsCASH!ISANGLIBO)
        txtLIMANGDAAN.Text = N2Str2Zero(rsCASH!LIMANGDAAN)
        txtDALAWANGDAAN.Text = N2Str2Zero(rsCASH!DALAWANGDAAN)
        txtISANGDAAN.Text = N2Str2Zero(rsCASH!ISANGDAAN)
        txtSINGKWENTA.Text = N2Str2Zero(rsCASH!SINGKWENTA)
        txtBENTE.Text = N2Str2Zero(rsCASH!BENTE)
        txtSAMPU.Text = N2Str2Zero(rsCASH!SAMPU)
        txtLIMANGPISO.Text = N2Str2Zero(rsCASH!LIMANGPISO)
        txtPISO.Text = N2Str2Zero(rsCASH!PISO)
        txtBENTESINKO.Text = N2Str2Zero(rsCASH!BENTESINKO)
        txtDYES.Text = N2Str2Zero(rsCASH!DYES)
        txtSINKO.Text = N2Str2Zero(rsCASH!SINKO)
        txtSENTIMO.Text = N2Str2Zero(rsCASH!SENTIMO)
        txtCHANGE_FUND.Text = ToDoubleNumber(CHANGE_FUND)
        Set rsOFF_HD = New ADODB.Recordset
        Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & rsCASH!CUTDATE & "'")
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            txtCashOnHand.Text = N2Str2Zero(rsOFF_HD!CASH)
            If N2Str2Zero(rsOFF_HD!LTO) < (N2Str2Zero(rsOFF_HD!LTO_EXP) + N2Str2Zero(rsOFF_HD!LTO_ADV) + N2Str2Zero(rsOFF_HD!LTO_REPL)) Then
                ADVANCES_COLL = (N2Str2Zero(rsOFF_HD!LTO_EXP) + N2Str2Zero(rsOFF_HD!LTO_ADV) + N2Str2Zero(rsOFF_HD!LTO_REPL)) - N2Str2Zero(rsOFF_HD!LTO)
            Else
                ADVANCES_COLL = 0
            End If
            If N2Str2Zero(rsOFF_HD!FUND) < (N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES) + N2Str2Zero(rsOFF_HD!REPLENISH)) Then
                ADVANCES_COLL = ((N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES) + N2Str2Zero(rsOFF_HD!REPLENISH)) - N2Str2Zero(rsOFF_HD!FUND)) + NumericVal(ADVANCES_COLL)
            End If
            txtCashOnHand.Text = ToDoubleNumber(NumericVal(txtCashOnHand.Text) - ADVANCES_COLL)
        End If
        MyFirstLove
    End If
End Sub

Sub MyFirstLove()
    If NumericVal(txtCashOnHand.Text) > NumericVal(txtCashCount.Text) Then
        txtShortBy.Text = ToDoubleNumber((NumericVal(txtCashOnHand.Text)) - NumericVal(txtCashCount.Text))
        txtOverBy.Text = "0.00"
    Else
        txtShortBy.Text = "0.00"
        txtOverBy.Text = ToDoubleNumber(NumericVal(txtCashCount.Text) - (NumericVal(txtCashOnHand.Text)))
    End If
End Sub

Private Sub cmdF10_Click()
    LogAudit "V", "CASH COUNT ENTRY", "CUT OFF DATE: " & txtCutDate
End Sub

Private Sub cmdF4_Click()
    If Function_Access(LOGID, "Acess_Edit", "TRANSACTION CASHIER CASH COUNT") = False Then Exit Sub
    cmdF4.Enabled = False
    picCashCount.Enabled = True
    On Error Resume Next
    txtISANGLIBO.SetFocus
    LogAudit "E", "CASH COUNT ENTRY"
End Sub

Private Sub cmdNext_Click()
    rsCASH.MoveNext
    If rsCASH.EOF Then
        rsCASH.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrev_Click()
    rsCASH.MovePrevious
    If rsCASH.BOF Then
        rsCASH.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If cmdF4.Enabled = True Then cmdF4.Value = True
        Case vbKeyF10
        Case vbKeyF11
            Shell "calc.exe"
        Case vbKeyEscape
            If picCashCount.Enabled = True Then
                picCashCount.Enabled = False
                cmdF4.Enabled = True
                StoreMemvars
            Else
                Unload Me
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    If Not rsCASH.EOF And Not rsCASH.BOF Then
        rsCASH.MoveLast
    End If
    StoreMemvars
    KIM_SUCKS
    Screen.MousePointer = 0
End Sub

Private Sub txtBENTE_Change()
    txt20total.Text = ToDoubleNumber(NumericVal(txtBENTE.Text) * 20): KIM_SUCKS
End Sub

Private Sub txtBENTE_GotFocus()
    If NumericVal(txtBENTE.Text) = 0 Then txtBENTE.Text = ""
End Sub

Private Sub txtBENTE_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtBENTE_LostFocus()
    If txtBENTE.Text = "" Then txtBENTE.Text = "0"
    MyFirstLove
End Sub

Private Sub txtBENTESINKO_Change()
    txt25total.Text = ToDoubleNumber(NumericVal(txtBENTESINKO.Text) * 0.25): KIM_SUCKS
End Sub

Private Sub txtBENTESINKO_GotFocus()
    If NumericVal(txtBENTESINKO.Text) = 0 Then txtBENTESINKO.Text = ""
End Sub

Private Sub txtBENTESINKO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtBENTESINKO_LostFocus()
    If txtBENTESINKO.Text = "" Then txtBENTESINKO.Text = "0"
    MyFirstLove
End Sub

Private Sub txtDALAWANGDAAN_Change()
    txt200total.Text = ToDoubleNumber(NumericVal(txtDALAWANGDAAN.Text) * 200): KIM_SUCKS
End Sub

Private Sub txtDALAWANGDAAN_GotFocus()
    If NumericVal(txtDALAWANGDAAN.Text) = 0 Then txtDALAWANGDAAN.Text = ""
End Sub

Private Sub txtDALAWANGDAAN_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtDALAWANGDAAN_LostFocus()
    If txtDALAWANGDAAN.Text = "" Then txtDALAWANGDAAN.Text = "0"
    MyFirstLove
End Sub

Private Sub txtDYES_Change()
    txt010total.Text = ToDoubleNumber(NumericVal(txtDYES.Text) * 0.1)
End Sub

Private Sub txtDYES_GotFocus()
    If NumericVal(txtDYES.Text) = 0 Then txtDYES.Text = ""
End Sub

Private Sub txtDYES_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtDYES_LostFocus()
    If txtDYES.Text = "" Then txtDYES.Text = "0"
    MyFirstLove
End Sub

Private Sub txtISANGDAAN_Change()
    txt100total.Text = ToDoubleNumber(NumericVal(txtISANGDAAN.Text) * 100): KIM_SUCKS
End Sub

Private Sub txtISANGDAAN_GotFocus()
    If NumericVal(txtISANGDAAN.Text) = 0 Then txtISANGDAAN.Text = ""
End Sub

Private Sub txtISANGDAAN_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtISANGDAAN_LostFocus()
    If txtISANGDAAN.Text = "" Then txtISANGDAAN.Text = "0"
    MyFirstLove
End Sub

Private Sub txtISANGLIBO_Change()
    txt1000total.Text = ToDoubleNumber(NumericVal(txtISANGLIBO.Text) * 1000): KIM_SUCKS
End Sub

Private Sub txtISANGLIBO_GotFocus()
    If NumericVal(txtISANGLIBO.Text) = 0 Then txtISANGLIBO.Text = ""
End Sub

Private Sub txtISANGLIBO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtISANGLIBO_LostFocus()
    If txtISANGLIBO.Text = "" Then txtISANGLIBO.Text = "0"
    MyFirstLove
End Sub

Private Sub txtLIMANGDAAN_Change()
    txt500total.Text = ToDoubleNumber(NumericVal(txtLIMANGDAAN.Text) * 500): KIM_SUCKS
End Sub

Private Sub txtLIMANGDAAN_GotFocus()
    If NumericVal(txtLIMANGDAAN.Text) = 0 Then txtLIMANGDAAN.Text = ""
End Sub

Private Sub txtLIMANGDAAN_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLIMANGDAAN_LostFocus()
    If txtLIMANGDAAN.Text = "" Then txtLIMANGDAAN.Text = "0"
    MyFirstLove
End Sub

Private Sub txtLIMANGPISO_Change()
    txt5total.Text = ToDoubleNumber(NumericVal(txtLIMANGPISO.Text) * 5): KIM_SUCKS
End Sub

Private Sub txtLIMANGPISO_GotFocus()
    If NumericVal(txtLIMANGPISO.Text) = 0 Then txtLIMANGPISO.Text = ""
End Sub

Private Sub txtLIMANGPISO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLIMANGPISO_LostFocus()
    If txtLIMANGPISO.Text = "" Then txtLIMANGPISO.Text = "0"
    MyFirstLove
End Sub

Private Sub txtPISO_Change()
    txt1total.Text = ToDoubleNumber(NumericVal(txtPISO.Text) * 1): KIM_SUCKS
End Sub

Private Sub txtPISO_GotFocus()
    If NumericVal(txtPISO.Text) = 0 Then txtPISO.Text = ""
End Sub

Private Sub txtPISO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPISO_LostFocus()
    If txtPISO.Text = "" Then txtPISO.Text = "0"
    MyFirstLove
End Sub

Private Sub txtSAMPU_Change()
    txt10total.Text = ToDoubleNumber(NumericVal(txtSAMPU.Text) * 10): KIM_SUCKS
End Sub

Private Sub txtSAMPU_GotFocus()
    If NumericVal(txtSAMPU.Text) = 0 Then txtSAMPU.Text = ""
End Sub

Private Sub txtSAMPU_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSAMPU_LostFocus()
    If txtSAMPU.Text = "" Then txtSAMPU.Text = "0"
    MyFirstLove
End Sub

Private Sub txtSENTIMO_Change()
    txt01total.Text = ToDoubleNumber(NumericVal(txtSENTIMO.Text) * 0.01): KIM_SUCKS
End Sub

Private Sub txtSENTIMO_GotFocus()
    If NumericVal(txtSENTIMO.Text) = 0 Then txtSENTIMO.Text = ""
End Sub

Private Sub txtSENTIMO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If MsgBox("Save Cash Count?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            picCashCount.Enabled = False
            gconDMIS.Execute ("update CMIS_Cash Set " & _
                            " CUTDATE = " & N2Str2Null(txtCutDate.Text) & "," & _
                            " ISANGLIBO = " & NumericVal(txtISANGLIBO.Text) & "," & _
                            " LIMANGDAAN = " & NumericVal(txtLIMANGDAAN.Text) & "," & _
                            " DALAWANGDAAN = " & NumericVal(txtDALAWANGDAAN.Text) & "," & _
                            " ISANGDAAN = " & NumericVal(txtISANGDAAN.Text) & "," & _
                            " SINGKWENTA = " & NumericVal(txtSINGKWENTA.Text) & "," & _
                            " BENTE = " & NumericVal(txtBENTE.Text) & "," & _
                            " SAMPU = " & NumericVal(txtSAMPU.Text) & "," & _
                            " LIMANGPISO = " & NumericVal(txtLIMANGPISO.Text) & "," & _
                            " PISO = " & NumericVal(txtPISO.Text) & "," & _
                            " BENTESINKO = " & NumericVal(txtBENTESINKO.Text) & "," & _
                            " DYES = " & NumericVal(txtDYES.Text) & "," & _
                            " SINKO = " & NumericVal(txtSINKO.Text) & "," & _
                            " SENTIMO = " & NumericVal(txtSENTIMO.Text) & _
                            " where id = " & LABID.Caption)

            LogAudit "E", "CASH COUNT ENTRY", "CUT OFF DATE: " & txtCutDate
            cmdF4.Enabled = True
            rsRefresh
            rsCASH.Find "id = " & LABID.Caption
            StoreMemvars
        End If
    End If
End Sub

Private Sub txtSENTIMO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSENTIMO_LostFocus()
    If txtSENTIMO.Text = "" Then txtSENTIMO.Text = "0"
    MyFirstLove
End Sub

Private Sub txtSINGKWENTA_Change()
    txt50total.Text = ToDoubleNumber(NumericVal(txtSINGKWENTA.Text) * 50): KIM_SUCKS
End Sub

Private Sub txtSINGKWENTA_GotFocus()
    If NumericVal(txtSINGKWENTA.Text) = 0 Then txtSINGKWENTA.Text = ""
End Sub

Private Sub txtSINGKWENTA_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSINGKWENTA_LostFocus()
    If txtSINGKWENTA.Text = "" Then txtSINGKWENTA.Text = "0"
    MyFirstLove
End Sub

Private Sub txtSINKO_Change()
    txt05total.Text = ToDoubleNumber(NumericVal(txtSINKO.Text) * 0.05)
End Sub

Private Sub txtSINKO_GotFocus()
    If NumericVal(txtSINKO.Text) = 0 Then txtSINKO.Text = ""
End Sub

Private Sub txtSINKO_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtSINKO_LostFocus()
    If txtSINKO.Text = "" Then txtSINKO.Text = "0"
    MyFirstLove
End Sub

