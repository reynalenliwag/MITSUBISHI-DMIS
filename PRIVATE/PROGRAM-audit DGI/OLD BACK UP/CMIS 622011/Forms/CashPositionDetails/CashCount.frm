VERSION 5.00
Begin VB.Form frmCASHPOSITIONCashCount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Count Entry"
   ClientHeight    =   5700
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   8310
   Enabled         =   0   'False
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CashCount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8310
   Begin VB.PictureBox picCashCount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   3330
      Picture         =   "CashCount.frx":08CA
      ScaleHeight     =   4665
      ScaleWidth      =   945
      TabIndex        =   14
      Top             =   810
      Width           =   945
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
         TabIndex        =   27
         Text            =   "0"
         Top             =   4320
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
         TabIndex        =   26
         Text            =   "0"
         Top             =   3960
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
         TabIndex        =   25
         Text            =   "0"
         Top             =   3600
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
         TabIndex        =   24
         Text            =   "0"
         Top             =   3240
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
         TabIndex        =   23
         Text            =   "0"
         Top             =   2880
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
         TabIndex        =   22
         Text            =   "0"
         Top             =   2520
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
         TabIndex        =   21
         Text            =   "0"
         Top             =   2160
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
         TabIndex        =   20
         Text            =   "0"
         Top             =   1800
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
         TabIndex        =   19
         Text            =   "0"
         Top             =   1440
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
         TabIndex        =   18
         Text            =   "0"
         Top             =   1080
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
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
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
         TabIndex        =   16
         Text            =   "0"
         Top             =   360
         Width           =   885
      End
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
         TabIndex        =   15
         Text            =   "0"
         Top             =   0
         Width           =   885
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2835
      TabIndex        =   40
      Top             =   1140
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
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   450
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
         TabIndex        =   42
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2835
      TabIndex        =   37
      Top             =   60
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
         TabIndex        =   38
         Top             =   450
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
         TabIndex        =   39
         Top             =   60
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2835
      TabIndex        =   34
      Top             =   4500
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
         Height          =   360
         Left            =   960
         TabIndex        =   35
         Text            =   "0.00"
         Top             =   450
         Width           =   1725
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
         TabIndex        =   36
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2835
      TabIndex        =   31
      Top             =   2250
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
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   450
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
         TabIndex        =   33
         Top             =   60
         Width           =   1875
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   945
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   2835
      TabIndex        =   28
      Top             =   3390
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
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   450
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
         TabIndex        =   30
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
      Left            =   6600
      Picture         =   "CashCount.frx":3606
      ScaleHeight     =   4665
      ScaleWidth      =   1515
      TabIndex        =   0
      Top             =   810
      Width           =   1515
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
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   0
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
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   360
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
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   720
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
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   1080
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
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1440
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
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1800
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
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   2160
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
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   2520
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
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   2880
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
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3240
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
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   3600
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
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   3960
         Width           =   1485
      End
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
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   4320
         Width           =   1485
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   1050
      TabIndex        =   86
      Top             =   480
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   6420
      TabIndex        =   85
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4740
      TabIndex        =   84
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3030
      TabIndex        =   83
      Top             =   420
      Width           =   1665
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
      Left            =   3030
      TabIndex        =   82
      Top             =   60
      Width           =   5055
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
      Left            =   5460
      TabIndex        =   81
      Top             =   5160
      Width           =   555
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
      Left            =   4620
      TabIndex        =   80
      Top             =   5160
      Width           =   225
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
      Left            =   6270
      TabIndex        =   79
      Top             =   5100
      Width           =   255
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
      Left            =   6270
      TabIndex        =   78
      Top             =   4020
      Width           =   255
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
      Left            =   4620
      TabIndex        =   77
      Top             =   4080
      Width           =   225
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
      Left            =   5460
      TabIndex        =   76
      Top             =   4080
      Width           =   555
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
      Left            =   6270
      TabIndex        =   75
      Top             =   3660
      Width           =   255
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
      Left            =   4620
      TabIndex        =   74
      Top             =   3720
      Width           =   225
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
      Left            =   5460
      TabIndex        =   73
      Top             =   3720
      Width           =   555
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
      Left            =   6270
      TabIndex        =   72
      Top             =   3300
      Width           =   255
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
      Left            =   4620
      TabIndex        =   71
      Top             =   3360
      Width           =   225
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
      Left            =   5460
      TabIndex        =   70
      Top             =   3360
      Width           =   555
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
      Left            =   6270
      TabIndex        =   69
      Top             =   2940
      Width           =   255
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
      Left            =   4620
      TabIndex        =   68
      Top             =   3000
      Width           =   225
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
      Left            =   5460
      TabIndex        =   67
      Top             =   3000
      Width           =   555
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
      Left            =   6270
      TabIndex        =   66
      Top             =   2580
      Width           =   255
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
      Left            =   4620
      TabIndex        =   65
      Top             =   2640
      Width           =   225
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
      Left            =   5460
      TabIndex        =   64
      Top             =   2640
      Width           =   555
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
      Left            =   6270
      TabIndex        =   63
      Top             =   2220
      Width           =   255
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
      Left            =   4620
      TabIndex        =   62
      Top             =   2280
      Width           =   225
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
      Left            =   5460
      TabIndex        =   61
      Top             =   2280
      Width           =   555
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
      Left            =   6270
      TabIndex        =   60
      Top             =   1860
      Width           =   255
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
      Left            =   4620
      TabIndex        =   59
      Top             =   1920
      Width           =   225
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
      Left            =   5310
      TabIndex        =   58
      Top             =   1920
      Width           =   705
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
      Left            =   6270
      TabIndex        =   57
      Top             =   1500
      Width           =   255
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
      Left            =   4620
      TabIndex        =   56
      Top             =   1560
      Width           =   225
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
      Left            =   5340
      TabIndex        =   55
      Top             =   1560
      Width           =   675
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
      Left            =   6270
      TabIndex        =   54
      Top             =   1140
      Width           =   255
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
      Left            =   4620
      TabIndex        =   53
      Top             =   1200
      Width           =   225
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
      Left            =   5310
      TabIndex        =   52
      Top             =   1200
      Width           =   705
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
      Left            =   6270
      TabIndex        =   51
      Top             =   780
      Width           =   255
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
      Left            =   4620
      TabIndex        =   50
      Top             =   840
      Width           =   225
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
      Left            =   5160
      TabIndex        =   49
      Top             =   840
      Width           =   855
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
      Left            =   5460
      TabIndex        =   48
      Top             =   4440
      Width           =   555
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
      Left            =   4620
      TabIndex        =   47
      Top             =   4440
      Width           =   225
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
      Left            =   6270
      TabIndex        =   46
      Top             =   4380
      Width           =   255
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
      Left            =   5460
      TabIndex        =   45
      Top             =   4800
      Width           =   555
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
      Left            =   4620
      TabIndex        =   44
      Top             =   4800
      Width           =   225
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
      Left            =   6270
      TabIndex        =   43
      Top             =   4740
      Width           =   255
   End
End
Attribute VB_Name = "frmCASHPOSITIONCashCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCASH                                                            As ADODB.Recordset
Dim rsOFF_HD                                                          As ADODB.Recordset
Dim VPettyCAFromCollection, VTotalAdvances                            As Double
Attribute VTotalAdvances.VB_VarUserMemId = 1073938434

Sub KIM_SUCKS()
    txtCashCount.Text = ToDoubleNumber(NumericVal(txt1000total.Text) + NumericVal(txt500total.Text) + NumericVal(txt200total.Text) + NumericVal(txt100total.Text) + NumericVal(txt50total.Text) + NumericVal(txt20total.Text) + NumericVal(txt10total.Text) + NumericVal(txt5total.Text) + NumericVal(txt1total.Text) + NumericVal(txt25total.Text) + NumericVal(txt010total.Text) + NumericVal(txt05total.Text) + NumericVal(txt01total.Text)): MyFirstLove
End Sub

Sub StoreMemvars()
    Set rsCASH = New ADODB.Recordset
    Set rsCASH = gconDMIS.Execute("Select * from CMIS_Cash Where CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "'")
    If Not rsCASH.EOF And Not rsCASH.BOF Then
        labid.Caption = rsCASH!Id
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
        Set rsOFF_HD = New ADODB.Recordset
        Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & CASHPOSITION_CUTOFF_DATE & "'")
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
            'txtCashOnHand.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CASH))
            If N2Str2Zero(rsOFF_HD!FUND) < N2Str2Zero(rsOFF_HD!REPLENISH) + N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES) Then
                VPettyCAFromCollection = NumericVal(N2Str2Zero(rsOFF_HD!REPLENISH) + N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES)) - N2Str2Zero(rsOFF_HD!FUND)
            Else
                VPettyCAFromCollection = "0.00"
            End If
            If N2Str2Zero(rsOFF_HD!LTO) < (N2Str2Zero(rsOFF_HD!LTO_EXP) + N2Str2Zero(rsOFF_HD!LTO_ADV) + N2Str2Zero(rsOFF_HD!LTO_REPL)) Then
                VTotalAdvances = (N2Str2Zero(rsOFF_HD!LTO_EXP) + N2Str2Zero(rsOFF_HD!LTO_ADV) + N2Str2Zero(rsOFF_HD!LTO_REPL)) - N2Str2Zero(rsOFF_HD!LTO)
            Else
                VTotalAdvances = 0
            End If
            If N2Str2Zero(rsOFF_HD!FUND) < N2Str2Zero(rsOFF_HD!REPLENISH) + N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES) Then
                VTotalAdvances = ToDoubleNumber(((N2Str2Zero(rsOFF_HD!REPLENISH) + N2Str2Zero(rsOFF_HD!EXPENSE) + N2Str2Zero(rsOFF_HD!ADVANCES)) - N2Str2Zero(rsOFF_HD!FUND)) + NumericVal(VTotalAdvances))
            End If
            txtCashOnHand.Text = ToDoubleNumber(N2Str2Zero(rsOFF_HD!CASH) - NumericVal(VTotalAdvances))

        End If
        KIM_SUCKS
    End If
End Sub

Sub MyFirstLove()
    If NumericVal(txtCashOnHand.Text) >= NumericVal(txtCashCount.Text) Then
        txtShortBy.Text = ToDoubleNumber(NumericVal(txtCashOnHand.Text) - NumericVal(txtCashCount.Text))
        txtOverBy.Text = "0.00"
    Else
        txtShortBy.Text = "0.00"
        txtOverBy.Text = ToDoubleNumber(NumericVal(txtCashCount.Text) - NumericVal(txtCashOnHand.Text))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub txtBENTE_Change()
    txt20total.Text = ToDoubleNumber(NumericVal(txtBENTE.Text) * 20)
End Sub

Private Sub txtBENTESINKO_Change()
    txt25total.Text = ToDoubleNumber(NumericVal(txtBENTESINKO.Text) * 0.25)
End Sub

Private Sub txtDALAWANGDAAN_Change()
    txt200total.Text = ToDoubleNumber(NumericVal(txtDALAWANGDAAN.Text) * 200)
End Sub

Private Sub txtDYES_Change()
    txt010total.Text = ToDoubleNumber(NumericVal(txtDYES.Text) * 0.1)
End Sub

Private Sub txtISANGDAAN_Change()
    txt100total.Text = ToDoubleNumber(NumericVal(txtISANGDAAN.Text) * 100)
End Sub

Private Sub txtISANGLIBO_Change()
    txt1000total.Text = ToDoubleNumber(NumericVal(txtISANGLIBO.Text) * 1000)
End Sub

Private Sub txtLIMANGDAAN_Change()
    txt500total.Text = ToDoubleNumber(NumericVal(txtLIMANGDAAN.Text) * 500)
End Sub

Private Sub txtLIMANGPISO_Change()
    txt5total.Text = ToDoubleNumber(NumericVal(txtLIMANGPISO.Text) * 5)
End Sub

Private Sub txtPISO_Change()
    txt1total.Text = ToDoubleNumber(NumericVal(txtPISO.Text) * 1)
End Sub

Private Sub txtSAMPU_Change()
    txt10total.Text = ToDoubleNumber(NumericVal(txtSAMPU.Text) * 10)
End Sub

Private Sub txtSENTIMO_Change()
    txt01total.Text = ToDoubleNumber(NumericVal(txtSENTIMO.Text) * 0.01)
End Sub

Private Sub txtSINGKWENTA_Change()
    txt50total.Text = ToDoubleNumber(NumericVal(txtSINGKWENTA.Text) * 50)
End Sub

Private Sub txtSINKO_Change()
    txt05total.Text = ToDoubleNumber(NumericVal(txtSINKO.Text) * 0.05)
End Sub

