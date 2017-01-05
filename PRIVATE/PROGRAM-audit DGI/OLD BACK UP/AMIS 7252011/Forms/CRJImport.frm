VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmCRJImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Receipts Import Process"
   ClientHeight    =   7935
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14025
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CRJImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   14025
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3270
      ScaleHeight     =   465
      ScaleWidth      =   4905
      TabIndex        =   22
      Top             =   7500
      Width           =   4905
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "- Not Yet Imported"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2250
         TabIndex        =   28
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   27
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "- Imported"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4140
         TabIndex        =   26
         Top             =   90
         Width           =   720
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   3750
         MouseIcon       =   "CRJImport.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "- No Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   24
         Top             =   90
         Width           =   1200
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   30
         MouseIcon       =   "CRJImport.frx":091E
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   90
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   3195
      TabIndex        =   16
      Top             =   6810
      Width           =   3195
      Begin VB.ComboBox cboYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   420
         Width           =   1005
      End
      Begin VB.ComboBox cboMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "CRJImport.frx":0C28
         Left            =   780
         List            =   "CRJImport.frx":0C2A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   30
         Width           =   1455
      End
      Begin VB.CommandButton cmdShowImp 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2250
         Picture         =   "CRJImport.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   30
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   9390
      ScaleHeight     =   765
      ScaleWidth      =   4515
      TabIndex        =   13
      Top             =   6150
      Width           =   4575
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Paid Invoices If not yet imported, will be automatically imported in Sales Journal since its already paid."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1125
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   4395
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   13230
      MouseIcon       =   "CRJImport.frx":1CAE
      MousePointer    =   99  'Custom
      Picture         =   "CRJImport.frx":1E00
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit Window"
      Top             =   7080
      Width           =   720
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Import"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   12525
      MouseIcon       =   "CRJImport.frx":2166
      MousePointer    =   99  'Custom
      Picture         =   "CRJImport.frx":22B8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   7080
      Width           =   720
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Un-Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   6210
      Value           =   -1  'True
      Width           =   4155
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Deposited Official Receipts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   6540
      Width           =   4155
   End
   Begin VB.CommandButton cmdShowTrans 
      Caption         =   "Show Transactions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3660
      MouseIcon       =   "CRJImport.frx":2553
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Import of SALES"
      Top             =   120
      Width           =   2010
   End
   Begin VB.CommandButton cmdClearJournals 
      BackColor       =   &H0080FF80&
      Caption         =   "Clear Selected Date"
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
      Height          =   375
      Left            =   11970
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   150
      Width           =   1935
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   4740
      TabIndex        =   6
      Top             =   6450
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      Picture         =   "CRJImport.frx":26A5
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "CRJImport.frx":26C1
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
      Left            =   1860
      TabIndex        =   7
      Top             =   120
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
      Format          =   48168961
      CurrentDate     =   40063
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   3240
      ScaleHeight     =   765
      ScaleWidth      =   5925
      TabIndex        =   29
      Top             =   6840
      Width           =   5925
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   60
         MouseIcon       =   "CRJImport.frx":26DD
         MousePointer    =   99  'Custom
         TabIndex        =   60
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   420
         MouseIcon       =   "CRJImport.frx":29E7
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   780
         MouseIcon       =   "CRJImport.frx":2CF1
         MousePointer    =   99  'Custom
         TabIndex        =   58
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1140
         MouseIcon       =   "CRJImport.frx":2FFB
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   1500
         MouseIcon       =   "CRJImport.frx":3305
         MousePointer    =   99  'Custom
         TabIndex        =   56
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":360F
         MousePointer    =   99  'Custom
         TabIndex        =   55
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2220
         MouseIcon       =   "CRJImport.frx":3919
         MousePointer    =   99  'Custom
         TabIndex        =   54
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2580
         MouseIcon       =   "CRJImport.frx":3C23
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   2940
         MouseIcon       =   "CRJImport.frx":3F2D
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   3300
         MouseIcon       =   "CRJImport.frx":4237
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   3660
         MouseIcon       =   "CRJImport.frx":4541
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4020
         MouseIcon       =   "CRJImport.frx":484B
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   4380
         MouseIcon       =   "CRJImport.frx":4B55
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   4740
         MouseIcon       =   "CRJImport.frx":4E5F
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   5100
         MouseIcon       =   "CRJImport.frx":5169
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   60
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   60
         MouseIcon       =   "CRJImport.frx":5473
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   420
         MouseIcon       =   "CRJImport.frx":577D
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   780
         MouseIcon       =   "CRJImport.frx":5A87
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   1140
         MouseIcon       =   "CRJImport.frx":5D91
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   1500
         MouseIcon       =   "CRJImport.frx":609B
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   1860
         MouseIcon       =   "CRJImport.frx":63A5
         MousePointer    =   99  'Custom
         TabIndex        =   40
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   2220
         MouseIcon       =   "CRJImport.frx":66AF
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   2580
         MouseIcon       =   "CRJImport.frx":69B9
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   2940
         MouseIcon       =   "CRJImport.frx":6CC3
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   3300
         MouseIcon       =   "CRJImport.frx":6FCD
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   3660
         MouseIcon       =   "CRJImport.frx":72D7
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   4020
         MouseIcon       =   "CRJImport.frx":75E1
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   4380
         MouseIcon       =   "CRJImport.frx":78EB
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   390
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   4740
         MouseIcon       =   "CRJImport.frx":7BF5
         MousePointer    =   99  'Custom
         TabIndex        =   32
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   5100
         MouseIcon       =   "CRJImport.frx":7EFF
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   5460
         MouseIcon       =   "CRJImport.frx":8209
         MousePointer    =   99  'Custom
         TabIndex        =   30
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4965
      Left            =   90
      TabIndex        =   61
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4965
      Left            =   4710
      TabIndex        =   62
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4965
      Left            =   9330
      TabIndex        =   63
      Top             =   1170
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PMIS/CSMS/SMIS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   9330
      TabIndex        =   15
      Top             =   660
      Width           =   4575
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
      Left            =   4770
      TabIndex        =   12
      Top             =   6180
      Width           =   5835
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DEPOSITED OR'S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   4710
      TabIndex        =   10
      Top             =   645
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNDEPOSITED OR'S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   90
      TabIndex        =   9
      Top             =   660
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported                        Invoices can be                            Imported"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   795
      Left            =   9390
      TabIndex        =   8
      Top             =   7020
      Width           =   3375
   End
End
Attribute VB_Name = "frmCRJImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TransactionID                                      As String
Dim rsCSMIOS_REPOR                                     As ADODB.Recordset
Dim rsCSMIOS_SUBLET                                    As ADODB.Recordset
Dim rsCSMIOS_PMS                                       As ADODB.Recordset
Dim rsCSMIOS_LABOR                                     As ADODB.Recordset
Dim rsCSMIOS_PARTS                                     As ADODB.Recordset
Dim rsCSMIOS_MATERIALS                                 As ADODB.Recordset
Dim rsSMIS_PURCHAGREE                                  As ADODB.Recordset
Dim rsCSMIOS_ACCESSORIES                               As New ADODB.Recordset
Dim rsCSMIOS_TINSMITH                                  As New ADODB.Recordset

'Discount
Dim CSMIOS_PMS_DISCOUNT                                As Double

'Warranty
Dim WARRANTY_CSMIOS_PARTS_COST                         As Double
Dim WARRANTY_CSMIOS_MATERIALS_COST                     As Double
Dim WARRANTY_CSMIOS_ACCESSORIES_COST                   As Double
Dim WARRANTY_JNO                                       As String
Dim WARRANTY_VOUCHERNO                                 As String
Dim WARRANTY_ItemCnt                                   As Integer
Dim WARRANTY_J_JITEMNO                                 As String
Dim WARRANTY_DIRECT_EXPENSE_LABOR_COST                 As Double
Dim WARRANTY_J_AMOUNTTOPAY                             As Double
Dim WARRANTY_J_INVOICEAMT                              As Double
Dim WARRANTY_J_BALANCE                                 As Double
Dim WARRANTY_J_AMOUNTPAID                              As Double

'Accessoreis

Dim CSMIOS_ACCESSORIES_DISCOUNT                        As Double

Dim CSMIOS_REP_OR                                      As String
Dim CSMIOS_ACCT_NO                                     As String
Dim CSMIOS_PARTICIPAT                                  As String
Dim CSMIOS_PLATE_NO                                    As String
Dim CSMIOS_NIYM                                        As String
Dim CSMIOS_TERM                                        As String
Dim CSMIOS_DTE_REL                                     As String
Dim CSMIOS_INVOICE                                     As String
Dim CSMIOS_VAT_EXEMPT                                  As Boolean
Dim CSMIOS_RO_AMOUNT                                   As Double

Dim CSMIOS_LABOR                                       As Double
Dim CSMIOS_PARTS                                       As Double
Dim CSMIOS_MATERIALS                                   As Double
Dim CSMIOS_ACCESSORIES                                 As Double

Dim CSMIOS_LABOR_COST                                  As Double
Dim CSMIOS_PARTS_COST                                  As Double
Dim CSMIOS_MATERIALS_COST                              As Double
Dim CSMIOS_ACCESSORIES_COST                            As Double

Dim CSMIOS_PMS_COST                                    As Double

'Dim CSMIOS_RO_AMOUNT                               As Double

Dim CSMIOS_TINSPAINT                                   As Double
Dim CSMIOS_SUBLET                                      As Double
Dim CSMIOS_AIRCON                                      As Double

Dim CSMIOS_TINSPAINT_DISCOUNT                          As Double
Dim CSMIOS_SUBLET_DISCOUNT                             As Double

Dim CSMIOS_LABOR_DISCOUNT                              As Double
Dim CSMIOS_PARTS_DISCOUNT                              As Double
Dim CSMIOS_MATERIALS_DISCOUNT                          As Double

Dim WARRANTY_DIRECT_EXPENSE_LABOR                      As Double
Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS                 As Double
Dim WARRANTY_DIRECT_EXPENSE_GOL                        As Double

Dim COMPANY_DIRECT_EXPENSE_LABOR                       As Double
Dim COMPANY_DIRECT_EXPENSE_SPAREPARTS                  As Double
Dim COMPANY_DIRECT_EXPENSE_GOL                         As Double

Dim SALES_DIRECT_EXPENSE_LABOR                         As Double
Dim SALES_DIRECT_EXPENSE_SPAREPARTS                    As Double
Dim SALES_DIRECT_EXPENSE_GOL                           As Double

Dim INSURANCE_DIRECT_EXPENSE_LABOR                     As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS                As Double
Dim INSURANCE_DIRECT_EXPENSE_GOL                       As Double
Dim INSURANCE_DIRECT_EXPENSE_ACCESSORIES               As Double
Dim CSMIOS_PMS                                         As Double

Dim TOTAL_INSURANCE_AMOUNT                             As Double
Dim TOTAL_DISCOUNT_AMOUNT                              As Double
Dim CSMIOS_SUBLET_COST                                 As Double
Dim ALL_DEBIT, ALL_CREDIT                              As Double
Attribute ALL_CREDIT.VB_VarUserMemId = 1073938500
Dim CSMIOS_TINSPAINT_COST                              As Double
Attribute CSMIOS_TINSPAINT_COST.VB_VarUserMemId = 1073938502

'Internal
Dim INTERNAL_LABOR_AMT                                 As Double
Attribute INTERNAL_LABOR_AMT.VB_VarUserMemId = 1073938503
Dim INTERNAL_PARTS_AMT                                 As Double
Attribute INTERNAL_PARTS_AMT.VB_VarUserMemId = 1073938504
Dim INTERNAL_MATERIALS_AMT                             As Double
Attribute INTERNAL_MATERIALS_AMT.VB_VarUserMemId = 1073938505
Dim INTERNAL_LABOR_COST                                As Double
Attribute INTERNAL_LABOR_COST.VB_VarUserMemId = 1073938506
Dim INTERNAL_PARTS_COST                                As Double
Attribute INTERNAL_PARTS_COST.VB_VarUserMemId = 1073938507
Dim INTERNAL_MATERIALS_COST                            As Double
Attribute INTERNAL_MATERIALS_COST.VB_VarUserMemId = 1073938508

Dim J_ACCT_CODE                                        As String
Attribute J_ACCT_CODE.VB_VarUserMemId = 1073938509
Dim J_ACCT_NAME                                        As String
Attribute J_ACCT_NAME.VB_VarUserMemId = 1073938510
Dim J_GROSS                                            As Double
Attribute J_GROSS.VB_VarUserMemId = 1073938511
Dim J_TAX                                              As Double
Attribute J_TAX.VB_VarUserMemId = 1073938512
Dim J_NET                                              As Double
Attribute J_NET.VB_VarUserMemId = 1073938513

Dim J_DEBIT                                            As Double
Attribute J_DEBIT.VB_VarUserMemId = 1073938514
Dim J_CREDIT                                           As Double
Attribute J_CREDIT.VB_VarUserMemId = 1073938515

Dim TOTAL_DEBIT                                        As Double
Attribute TOTAL_DEBIT.VB_VarUserMemId = 1073938516
Dim TOTAL_CREDIT                                       As Double
Attribute TOTAL_CREDIT.VB_VarUserMemId = 1073938517
'Dim CSMIOS_VAT_EXEMPT                              As Boolean

Dim ItemCnt                                            As Integer
Attribute ItemCnt.VB_VarUserMemId = 1073938518
Dim CSMS_ACCCOST                                       As Double
Attribute CSMS_ACCCOST.VB_VarUserMemId = 1073938519
Dim rsINTERNAL_RO_DET                                  As ADODB.Recordset
Attribute rsINTERNAL_RO_DET.VB_VarUserMemId = 1073938520

Dim J_JDATE                                            As String
Attribute J_JDATE.VB_VarUserMemId = 1073938521
Dim J_VOUCHERNO                                        As String
Attribute J_VOUCHERNO.VB_VarUserMemId = 1073938522
Dim J_JTYPE                                            As String
Attribute J_JTYPE.VB_VarUserMemId = 1073938523
Dim J_JNO                                              As String
Attribute J_JNO.VB_VarUserMemId = 1073938524
Dim J_REMARKS                                          As String
Attribute J_REMARKS.VB_VarUserMemId = 1073938525
Dim J_VENDORCODE                                       As String
Attribute J_VENDORCODE.VB_VarUserMemId = 1073938526
Dim J_CUSTOMERCODE                                     As String
Attribute J_CUSTOMERCODE.VB_VarUserMemId = 1073938527
Dim J_OUTBALANCE                                       As Double
Attribute J_OUTBALANCE.VB_VarUserMemId = 1073938528
Dim J_AMOUNTTOPAY                                      As Double
Attribute J_AMOUNTTOPAY.VB_VarUserMemId = 1073938529
Dim J_INVOICEAMT                                       As Double
Attribute J_INVOICEAMT.VB_VarUserMemId = 1073938530
Dim J_BALANCE                                          As Double
Attribute J_BALANCE.VB_VarUserMemId = 1073938531
Dim J_AMOUNTPAID                                       As Double
Attribute J_AMOUNTPAID.VB_VarUserMemId = 1073938532
Dim J_CHECKNO                                          As String
Attribute J_CHECKNO.VB_VarUserMemId = 1073938533
Dim J_INVOICEDATE                                      As String
Attribute J_INVOICEDATE.VB_VarUserMemId = 1073938534
Dim J_DUEDATE                                          As String
Attribute J_DUEDATE.VB_VarUserMemId = 1073938535
Dim J_PAYTYPE                                          As String
Attribute J_PAYTYPE.VB_VarUserMemId = 1073938536
Dim J_INVOICETYPE                                      As String
Attribute J_INVOICETYPE.VB_VarUserMemId = 1073938537
Dim J_INVOICENO                                        As String
Attribute J_INVOICENO.VB_VarUserMemId = 1073938538
Dim J_CHECKDATE                                        As String
Attribute J_CHECKDATE.VB_VarUserMemId = 1073938539
Dim J_BANKCODE                                         As String
Attribute J_BANKCODE.VB_VarUserMemId = 1073938540
Dim J_REFNO                                            As String
Attribute J_REFNO.VB_VarUserMemId = 1073938541
Dim J_REFDATE                                          As String
Attribute J_REFDATE.VB_VarUserMemId = 1073938542
Dim J_TERMS                                            As String
Attribute J_TERMS.VB_VarUserMemId = 1073938543
Dim J_DEALER                                           As String
Attribute J_DEALER.VB_VarUserMemId = 1073938544
Dim J_PAIDSTATUS                                       As String
Attribute J_PAIDSTATUS.VB_VarUserMemId = 1073938545
Dim J_RECEIVESTATUS                                    As String
Attribute J_RECEIVESTATUS.VB_VarUserMemId = 1073938546

'Dim J_ACCT_CODE, J_ACCT_NAME                       As String
'Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
Dim J_STATUS                                           As String
Attribute J_STATUS.VB_VarUserMemId = 1073938547
Dim J_JITEMNO                                          As String
Attribute J_JITEMNO.VB_VarUserMemId = 1073938548
Dim rsJournal_HDDup                                    As ADODB.Recordset
Attribute rsJournal_HDDup.VB_VarUserMemId = 1073938549
Dim LIM                                                As Integer
Attribute LIM.VB_VarUserMemId = 1073938550
Dim TRANSACTIONDATE                                    As String
Attribute TRANSACTIONDATE.VB_VarUserMemId = 1073938551
Dim Indx                                               As Integer
Attribute Indx.VB_VarUserMemId = 1073938552
Dim xTranDate                                          As String
Attribute xTranDate.VB_VarUserMemId = 1073938553
Dim xCUT_OFF_DATE                                      As String
Attribute xCUT_OFF_DATE.VB_VarUserMemId = 1073938554

Function SetOTHChartCodes(XXX As String) As String
    Dim rsSBOOK_CHARTCODES                             As ADODB.Recordset
    Set rsSBOOK_CHARTCODES = New ADODB.Recordset
    Set rsSBOOK_CHARTCODES = gconDMIS.Execute("Select * from CMIS_SBOOK where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOK_CHARTCODES.EOF And Not rsSBOOK_CHARTCODES.BOF Then
        SetOTHChartCodes = Null2String(rsSBOOK_CHARTCODES!CHARTCODES)
    End If
    Set rsSBOOK_CHARTCODES = Nothing
End Function

Function ReturnSITerm(XXX As String) As String
    Dim rsREPOR_INVOICE                                As ADODB.Recordset
    Set rsREPOR_INVOICE = New ADODB.Recordset
    Set rsREPOR_INVOICE = gconDMIS.Execute("Select TERM from CSMS_Repor Where INVOICE = '" & XXX & "'")
    If Not rsREPOR_INVOICE.EOF And Not rsREPOR_INVOICE.BOF Then
        ReturnSITerm = Null2String(rsREPOR_INVOICE!TERM)
    End If
    Set rsREPOR_INVOICE = Nothing
End Function

Function ReturnTranType(XXX As String, YYY As String) As String
    Dim rsPMIS_Invoice                                 As ADODB.Recordset
    Set rsPMIS_Invoice = New ADODB.Recordset
    rsPMIS_Invoice.Open "SELECT TRANTYPE FROM PMIS_VW_ISS_HISTORY WHERE TRANNO = '" & XXX & "' AND TYPE='" & YYY & "' AND TRANTYPE<>'RIV'", gconDMIS, adOpenKeyset
    If Not rsPMIS_Invoice.EOF And Not rsPMIS_Invoice.BOF Then
        ReturnTranType = Null2String(rsPMIS_Invoice!TranType)
    End If
    Set rsPMIS_Invoice = Nothing
End Function

Function SetTransaction(XXX As Variant) As String
    Dim rsSBOOKTransaction                             As ADODB.Recordset
    Set rsSBOOKTransaction = New ADODB.Recordset
    Set rsSBOOKTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'A' and CODE = '" & XXX & "'")
    If Not rsSBOOKTransaction.EOF And Not rsSBOOKTransaction.BOF Then
        SetTransaction = Null2String(rsSBOOKTransaction!DESCNAME)
    End If
    Set rsSBOOKTransaction = Nothing
End Function

Function SetOtherTransaction(XXX As Variant) As String
    Dim rsSBOOKOtherTransaction                        As ADODB.Recordset
    Set rsSBOOKOtherTransaction = New ADODB.Recordset
    Set rsSBOOKOtherTransaction = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOKOtherTransaction.EOF And Not rsSBOOKOtherTransaction.BOF Then
        SetOtherTransaction = Null2String(rsSBOOKOtherTransaction!DESCNAME)
    End If
    Set rsSBOOKOtherTransaction = Nothing
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
End Function

Function GetCRJVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'CRJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetCRJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetCRJVoucherNo = "000001"
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'DRJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function GetSJVoucherNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetSJVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetSJVoucherNo = "000001"
    End If
End Function

'Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
'    Dim rsCheckSJ_Journal_HD                      As ADODB.Recordset
'    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
'    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
'    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
'        CheckSJExisting = True
'    Else
'        CheckSJExisting = False
'    End If
'    Set rsCheckSJ_Journal_HD = Nothing
'End Function

Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String, Optional VarTranType As String) As Boolean
    Dim rsCheckSJ_Journal_HD                           As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    If VarTranType = "" Then
        Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND PayType = " & N2Str2Null(VarTranType) & " AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckSJExisting = True
    Else
        CheckSJExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function

Function CheckRefNoExisting(VarInvoiceType As String, VarInvoiceNo As String) As Boolean
    Dim rsCheckSJ_Journal_HD                           As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND RefNo = " & N2Str2Null(VarInvoiceNo))
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckRefNoExisting = True
    Else
        CheckRefNoExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function

Function CheckCRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckCRJ_Journal_HD                          As ADODB.Recordset
    Set rsCheckCRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckCRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'CRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckCRJ_Journal_HD.EOF And Not rsCheckCRJ_Journal_HD.BOF Then
        CheckCRJExisting = True
    Else
        CheckCRJExisting = False
    End If
    Set rsCheckCRJ_Journal_HD = Nothing
End Function

Function CheckDRJExisting(VarInvoiceNo As String, VarVAT As Variant) As Boolean
    Dim rsCheckDRJ_Journal_HD                          As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo))
    Else
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExisting = True
    Else
        CheckDRJExisting = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function CheckDRJExistingM(VarInvoiceNo As String, VarVAT As Variant, VarAmount As Double) As Boolean
    Dim rsCheckDRJ_Journal_HD                          As ADODB.Recordset
    Set rsCheckDRJ_Journal_HD = New ADODB.Recordset
    If VarVAT = 0 Then
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND LEFT(InvoiceNo,2) = 'NV' AND RIGHT(InvoiceNo,6) = " & N2Str2Null(VarInvoiceNo) & " AND INVOICEAMT =" & NumericVal(VarAmount) & "")
    Else
        Set rsCheckDRJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'DRJ' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo) & " And INVOICEAMT = " & NumericVal(VarAmount) & "")
    End If
    If Not rsCheckDRJ_Journal_HD.EOF And Not rsCheckDRJ_Journal_HD.BOF Then
        CheckDRJExistingM = True
    Else
        CheckDRJExistingM = False
    End If
    Set rsCheckDRJ_Journal_HD = Nothing
End Function

Function ReturnAR_AccountCode(XXX As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AR' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAR_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnClearing_AccountCode(XXX As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'CLEARING' AND TRANTYPE1 = '" & Trim(XXX) & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnClearing_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAccountCode(XXX As String, Optional YYY As String)
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If YYY = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "' AND TRANTYPE3 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnDeposit_AccountCode(XXX As String)
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'DEPOSIT' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDeposit_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetCustomerName(VVV As Variant) As String
    Dim rsCustomer2                                    As ADODB.Recordset
    Set rsCustomer2 = New ADODB.Recordset
    rsCustomer2.Open "Select CustCode,acctname from ALL_CUSTMASTER_AMIS where CustCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer2.EOF And Not rsCustomer2.BOF Then
        SetCustomerName = UCase(Null2String(rsCustomer2!AcctName))
    Else
        SetCustomerName = ""
    End If
End Function

Function ReturnSales_AccountCode(INVTYPE As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            If INVTYPE = "SALES" Then
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & Right(OTHERTYPE, 5) & "%'")
            Else
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%'")
            End If
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnSales_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnCostofSales(INVTYPE As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            If INVTYPE = "SALES" Then
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & Right(OTHERTYPE, 5) & "%'")
            Else
                Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%'")
            End If
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'COST OF SALES' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 like '%" & OTHERTYPE & "%' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnCostofSales = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInventory(INVTYPE As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    Else
        If NEXTTYPE = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "' AND TRANTYPE4 = '" & NEXTTYPE & "'")
        End If
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInventory = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnOutputTax()
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnOutputTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnDiscount_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If OTHERTYPE = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'DISCOUNT' AND TRANTYPE2 = '" & INVTYPE & "' AND TRANTYPE1 = '" & OTHERTYPE & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDiscount_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function PosibleDoubleInternal(XXX As String) As Boolean
    Dim RS                                             As New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT * FROM AMIS_journal_hd where refno='" & XXX & "'")
    If Not RS.EOF And Not RS.BOF Then
        PosibleDoubleInternal = True
    Else
        PosibleDoubleInternal = False
    End If
    Set RS = Nothing
End Function

Function ReturnInternalAccountCode(XXX As String)
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select ChartCodes from CMIS_SBOOK where BOOK = 'S' and CODE = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInternalAccountCode = Null2String(rsChartAccount!CHARTCODES)
    End If
    Set rsChartAccount = Nothing
End Function

Function GetJNo() As String
    Dim rsJournal_HD                                   As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(JNo AS int) AS MAX_JNO from AMIS_Journal_HD Order by MAX_JNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetJNo = Format(NumericVal(rsJournal_HD!MAX_JNO) + 1, "000000")
    Else
        GetJNo = "000001"
    End If
End Function

Function ReturnPlateNo(XXX As String) As String
' Update By BTT - 07092008
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT Plate_no from CSMS_REPOR WHERE REP_OR='" & XXX & "' "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        ReturnPlateNo = Null2String(RS!PLATE_NO)
    Else
        ReturnPlateNo = ""
    End If
    Set RS = Nothing
End Function

Function ReturnCodeSellingDealer(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT Selling_dealer from CSMS_Cusveh WHERE plate_no='" & XXX & "' "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        ReturnCodeSellingDealer = Null2String(RS!Selling_Dealer)
    Else
        ReturnCodeSellingDealer = ""
    End If
    Set RS = Nothing
End Function

Function SetSellingDealerName(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT DealerName from CSMS_SellingDealer WHERE DealerCOde='" & XXX & "' "

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        SetSellingDealerName = Null2String(RS!DealerName)
    Else
        SetSellingDealerName = ""
    End If
    Set RS = Nothing
End Function

Function ReturnCode(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim MARK                                           As String
    'This will return the code of the Selling Dealer

    MARK = (Replace(XXX, " ", ""))

    SQL = "SELECT Custcode, replace(custname,' ','') from ALL_Custmaster_AMIS where REPLACE(custname,' ','') like '%" & MARK & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        ReturnCode = Null2String(RS!CUSTCODE)
    Else
        ReturnCode = ""
    End If
    Set RS = Nothing
End Function

Function CheckIfPMS_Ik_to_5k(XXX As String) As Boolean
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "select Status1 from CSMS_ro_det where jobtype='PMS' and livil='1' and wcode='W' and Rep_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If Null2String(RS!status1) = "Y" Then
            CheckIfPMS_Ik_to_5k = True
        Else
            CheckIfPMS_Ik_to_5k = False
        End If
    End If
    Set RS = Nothing
End Function

Function ReturnShopSuplies(XXX As String) As Integer
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT DETAMT,detprc FROM CSMS_RO_DET WHERE LIVIL = '3' AND DETCDE = 'SVCMAT0068' and REP_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        ReturnShopSuplies = NumericVal(RS!DETPRC)
    Else
        ReturnShopSuplies = 0
    End If
    Set RS = Nothing
End Function

Function ReturnDeferredOutPutTax()
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'DEFERRED OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDeferredOutPutTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetRONiym(XXX As String)
    Dim rsRO_Niym                                      As ADODB.Recordset
    Set rsRO_Niym = New ADODB.Recordset
    Set rsRO_Niym = gconDMIS.Execute("Select NIYM from CSMS_REPOR WHERE INVOICE = '" & XXX & "'")
    If Not rsRO_Niym.EOF And Not rsRO_Niym.BOF Then
        SetRONiym = Null2String(rsRO_Niym!Niym)
    End If
End Function

Function ImportUnDeposit() As Boolean
'HEADER
    On Error GoTo ErrorCode
    Dim J_JDATE                                        As String
    Dim J_VOUCHERNO                                    As String
    Dim J_JTYPE                                        As String
    Dim J_JNO                                          As String
    Dim J_REMARKS                                      As String
    Dim J_VENDORCODE                                   As String
    Dim J_CUSTOMERCODE                                 As String
    Dim J_CUSTOMERCODE2                                As String
    Dim J_CHECKNO                                      As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE            As String
    Dim J_INVOICETYPE, J_INVOICENO                     As String
    Dim J_CHECKDATE, J_BANKCODE                        As String
    Dim J_REFNO, J_REFDATE                             As String
    Dim J_TERMS, J_DEALER                              As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                  As String
    Dim J_REFERENCENO                                  As String
    Dim J_ENTITY                                       As String
    Dim J_OUTBALANCE                                   As Double
    Dim J_AMOUNTTOPAY                                  As Double
    Dim J_INVOICEAMT                                   As Double
    Dim J_BALANCE                                      As Double
    Dim J_AMOUNTPAID                                   As Double
    Dim J_BANKCHARGES                                  As Double
    'DETAIL
    Dim J_ACCT_CODE                                    As String
    Dim J_ACCT_NAME                                    As String
    Dim J_STATUS                                       As String
    Dim J_JITEMNO                                      As String
    Dim J_JITEMNO_2                                    As String
    Dim J_ALLENTITY                                    As String
    Dim J_INVOICENUM                                   As String
    Dim J_DEBIT                                        As Double
    Dim J_CREDIT                                       As Double
    Dim J_TAX                                          As Double
    Dim J_GROSS                                        As Double
    Dim J_NET                                          As Double

    Dim rsJournal_HDDup                                As ADODB.Recordset
    Dim CMIS_OR_NUM                                    As String
    Dim CMIS_OR_DATE                                   As String
    Dim CMIS_OR_AMT                                    As String
    Dim CMIS_DISCOUNT                                  As String
    Dim CMIS_TAX                                       As String
    Dim CMIS_CUSCDE                                    As String
    Dim CMIS_CUSNAME                                   As String
    Dim CMIS_DEPOSIT                                   As String
    Dim CMIS_BANKCODE                                  As String
    Dim CMIS_BANK                                      As String
    Dim CMIS_CARDBANK                                  As String
    Dim CMIS_TSEKE                                     As String
    Dim CMIS_CHECKDATE                                 As String
    Dim CMIS_STATUS                                    As String
    Dim CMIS_TYPE_PAYMENT                              As String
    Dim CMIS_DT_TRANTYPE                               As String
    Dim CMIS_DT_REFERENCE                              As String
    Dim CMIS_DT_CUSCDE                                 As String
    Dim CMIS_DT_DESCRIPT                               As String
    Dim CMIS_DT_REFERENCENO                            As String
    Dim CMIS_DT_DOCDTE                                 As String
    Dim CMIS_DT_PAIDFOR                                As String
    Dim CMIS_ENTITYCODE                                As String
    Dim CMIS_ALLENTITY                                 As String
    Dim CMIS_CASHAMOUNT                                As Double
    Dim CMIS_CHKAMOUNT                                 As Double
    Dim CMIS_CARDAMOUNT                                As Double
    Dim TOTAL_DEBIT                                    As Double
    Dim TOTAL_CREDIT                                   As Double
    Dim CMIS_DT_AMOUNT                                 As Double
    Dim CMIS_DT_PAYMENT                                As Double
    Dim CMIS_DT_DISCOUNT                               As Double
    Dim CMIS_DT_TAX                                    As Double
    Dim CMIS_IS_VAT                                    As Boolean
    Dim i                                              As Long

    Dim rsOFF_HD                                       As ADODB.Recordset
    Dim rsOFF_DT                                       As ADODB.Recordset
    Dim rsSJ_DATA                                      As ADODB.Recordset
    Dim rsCheckJournal_HD                              As ADODB.Recordset

    Dim PV_MRRNO                                       As String
    Dim PV_INVNO                                       As String
    Dim PV_PRODNO                                      As String
    Dim J_JVOUCHERNO                                   As String
    Dim PV_STATUS, PV_ITEMNO                           As String
    Dim PV_AMOUNT                                      As Double
    Dim SJ_PV_ITEMNO                                   As Integer
    Dim GridImport                                     As Integer
    Dim OIinv                                     As String
    '    Dim J_ITEMCOUNT                               As Integer

    Dim rsCreditCard                                   As ADODB.Recordset
    Dim BANKCHARGES                                    As Double
    Dim EWT                                            As Double
    Dim TOTALCHARGES                                   As Double
    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 1 AND OR_DATE > '" & GetCUTOFF_DATE & "' AND  OR_DATE <= '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 0 AND OR_DATE > '" & GetCUTOFF_DATE & "' AND  OR_DATE <= '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                CMIS_BANKCODE = Null2String(rsOFF_HD!bankcode)
                CMIS_BANK = Null2String(rsOFF_HD!Bank)
                CMIS_CARDBANK = Null2String(rsOFF_HD!cardbnkcde)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)
                CMIS_BANKCODE = Null2String(rsOFF_HD!bankcode)
                If COMPANY_CODE = "HMH" Then
                    If Null2String(rsOFF_HD!oiiden) = "" Then
                        CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                    Else
                        CMIS_ENTITYCODE = Null2String(rsOFF_HD!oiiden)
                    End If
                Else
                    CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                End If
                'CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                CMIS_ALLENTITY = CMIS_ENTITYCODE + CMIS_CUSCDE
                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetCRJVoucherNo())
                J_JTYPE = "'CRJ'"
                J_ITEMCOUNT = 0
                '                J_JITEMNO_2 = 0
                'INSERTED SEPTEMBER 8, 2007
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        J_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_INVOICENUM = Null2String(rsOFF_DT!INVOICENO)
                        OIinv = Null2String(rsOFF_DT!TranType)

                        If Null2String(rsOFF_DT!TranType) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)
                J_DEPOSIT = 0
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CMIS_OR_AMT), 2)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 6))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 6))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_ENTITY = N2Str2Null(CMIS_ENTITYCODE)
                J_ALLENTITY = N2Str2Null(CMIS_ALLENTITY)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND

                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    '                    J_JITEMNO = "'0001'"
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                    Else
                        J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    'If COMPANY_CODE = "HGC" Then
                    'DESCRIPTION: Checking for Customer Deposits and ADD upon full payment of the invoice amount
                    If J_INVOICENUM <> "" Then
                        Dim rsCheckDeposit             As ADODB.Recordset
                        Set rsCheckDeposit = New ADODB.Recordset
                        'updated by al for customer deposit...03/24/2011
                        rsCheckDeposit.Open "Select * from CMIS_Deposits where InvoiceNo = '" & J_INVOICENUM & "' and cuscde = " & J_CUSTOMERCODE & " And InvoiceType = '" & OIinv & "'", gconDMIS, adOpenForwardOnly
                        If Not rsCheckDeposit.EOF And Not rsCheckDeposit.BOF Then
                            '                            J_JITEMNO_2 = 1
                            '                            J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                            '                            J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR))))
                            J_DEBIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                            J_DEPOSIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                        Set rsCheckDeposit = Nothing
                    End If
                End If

                If CMIS_TYPE_PAYMENT = "3" Then
                    '                    J_JITEMNO = "'0001'"
                    J_ITEMCOUNT = 0
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                    ElseIf COMPANY_CODE = "DGI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND", CMIS_CARDBANK))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND", CMIS_CARDBANK)))
                    End If
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE & "," & J_ALLENTITY & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                End If

                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 1 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 0 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst: SJ_PV_ITEMNO = 0
                    Do While Not rsOFF_DT.EOF
                        CMIS_DT_TRANTYPE = Null2String(rsOFF_DT!TranType)
                        CMIS_DT_REFERENCE = Null2String(rsOFF_DT!INVOICENO)
                        CMIS_DT_CUSCDE = N2Str2Null(rsOFF_DT!CUSCDE)
                        CMIS_DT_DESCRIPT = Null2String(rsOFF_DT!DESCRIPT)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!amount)
                        CMIS_DT_DOCDTE = Null2String(rsOFF_DT!DOCDTE)
                        CMIS_DT_PAYMENT = N2Str2Zero(rsOFF_DT!payment)
                        CMIS_DT_DISCOUNT = N2Str2Zero(rsOFF_DT!DISCOUNT)
                        CMIS_DT_TAX = N2Str2Zero(rsOFF_DT!tax)
                        CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)

                        'DESCRIPTION: For credit card detail
                        CMIS_DT_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_JVOUCHERNO = J_VOUCHERNO
                        SJ_PV_ITEMNO = SJ_PV_ITEMNO + 1
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                        PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                        PV_AMOUNT = CMIS_DT_PAYMENT   ' AMOUNT
                        PV_STATUS = "'N'"
                        J_BANKCHARGES = 0
                        PV_INVDATE = N2Date2Null(rsOFF_DT!ORDATE)
                        'CREDITCARD
                        '                        If COMPANY_CODE = "HGC" Then
                        '                            PV_INVDATE = GetCustomerORDate(CMIS_DT_REFERENCE)
                        '                        Else

                        If UCase(CMIS_DT_PAIDFOR) = "412S" Or UCase(CMIS_DT_PAIDFOR) = "412P" Or UCase(CMIS_DT_PAIDFOR) = "412V" Then
                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                            "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE)" & _
                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & ",NULL, " & PV_ITEMNO & _
                                            ", 'CI', '" & CMIS_OR_NUM & "', " & PV_INVDATE & ", " & CMIS_DT_PAYMENT & _
                                            ", " & PV_STATUS & "," & CMIS_DT_CUSCDE & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If


                        '                        End If
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        Set rsSJ_DATA = New ADODB.Recordset
                        '---al FOR GJ COUNTRY
                        If COMPANY_CODE = "HMH" Then
                            If CMIS_DT_TRANTYPE = "OI" Then
                                'Set rsSJ_DATA = gconDMIS.Execute("Select * from amis_cv_detail Where jtype = 'OI' and pv_voucherno = " & N2Str2Null(CMIS_DT_REFERENCE))
                                Set rsSJ_DATA = gconDMIS.Execute("Select docdate,j_class,vendorcode,(left(cv_jtype,3)+''+left(voucherno,6)) as voucherno from amis_cv_detail Where jtype = 'OI' and pv_voucherno = '" & CMIS_DT_REFERENCE & "' Union select jdate as docdate,j_class,entitycode as vendorcode,(left(jtype,3)+''+left(voucherno,6)) as voucherno from amis_pv_detail Where invoicetype = 'OI' and inv_no = '" & CMIS_DT_REFERENCE & "'")
                            Else
                                Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                            End If
                        Else
                            Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        End If
                        'Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            J_JVOUCHERNO = J_VOUCHERNO
                            PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                            PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                            '---al FOR GJ COUNTRY
                            If CMIS_DT_TRANTYPE = "OI" Then
                                PV_PRODNO = N2Date2Null(rsSJ_DATA!docdate)
                            Else
                                PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)    ' DATE
                            End If
                            'PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)    ' DATE
                            If CMIS_DT_TAX = 0 Then
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT), 2)    ' AMOUNT
                            Else
                                PV_AMOUNT = Round((CMIS_DT_PAYMENT + J_DEPOSIT + CMIS_DT_TAX), 2)
                            End If
                            PV_STATUS = "'N'"

                            'COMMENTED BY: JUN 10072009 ----------ORIGINNAL CODE-------------------------------------------------------------------------------------------------------------------
                            '                            'DESCRIPTION: Modify database structure - added fields: CusCde - Insert J_Class and Customer Code upon importing (,J_Class,CusCde) NOT YET
                            '                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                         '                                             "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                         '                                           " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                         '                                             ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                         '                                             ", " & PV_STATUS & ")"
                            '                            gconDMIS.Execute SQL_STATEMENT
                            'COMMENTED BY: JUN 10072009 ----------ORIGINNAL CODE-------------------------------------------------------------------------------------------------------------------

                            'UPDATED BY: JUN----------------------------------------------------------------------------------------------------------------
                            'DATE UPDATED: 10012009
                            'DESCRIPTION: ADDED FIELD IN SAVING IN AMIS_CRJ_DETAIL FOR AR PURPOSE FIELD NAME: SJ_VOUCHERNO,CUSTOMER_CODE
                            If CMIS_DT_TRANTYPE = "OI" And COMPANY_CODE = "HMH" And Now >= "03/22/2011" Then
                                SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                            "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE)" & _
                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & "," & N2Str2Null(rsSJ_DATA!VOUCHERNO) & ", " & PV_ITEMNO & _
                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                            ", " & PV_STATUS & "," & N2Str2Null(rsSJ_DATA!VendorCode) & ")"
                            Else
                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                            "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE)" & _
                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & "," & N2Str2Null(rsSJ_DATA!VOUCHERNO) & ", " & PV_ITEMNO & _
                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                            ", " & PV_STATUS & "," & N2Str2Null(rsSJ_DATA!CustomerCode) & ")"
                            End If
                            gconDMIS.Execute SQL_STATEMENT
                            'UPDATED BY: JUN----------------------------------------------------------------------------------------------------------------

                            'If COMPANY_CODE = "HGC" Then
                            'DESCRIPTION: AR Credit Card detail upon payment of Credit Card Company
                            '                            If J_INVOICENUM <> "" Then
                            '                                Dim rsCheckDetail As ADODB.Recordset
                            '                                Set rsCheckDetail = New ADODB.Recordset
                            '                                rsCheckDetail.Open "SELECT ReferenceNo FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                            '                                If Not rsCheckDetail.EOF And Not rsCheckDetail.BOF Then
                            '                                    Dim rsCheckDetail2 As ADODB.Recordset
                            '                                    Set rsCheckDetail2 = New ADODB.Recordset
                            '                                    rsCheckDetail2.Open "SELECT * FROM CMIS_OFF_DT WHERE REFERENCENO=" & N2Str2Null(rsCheckDetail!ReferenceNo) & " AND OR_NUM NOT IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                            '                                    If Not rsCheckDetail2.EOF And Not rsCheckDetail2.BOF Then
                            '                                        Dim rsCheckVoucher As ADODB.Recordset
                            '                                        Set rsCheckVoucher = New ADODB.Recordset
                            '                                        rsCheckVoucher.Open "SELECT VOUCHERNO,InvoiceAmt from AMIS_JOURNAL_HD WHERE InvoiceNo =" & N2Str2Null(rsCheckDetail2!OR_NUM) & "", gconDMIS, adOpenForwardOnly
                            '                                        If Not rsCheckVoucher.EOF And Not rsCheckVoucher.BOF Then
                            '                                            PV_AMOUNT = rsCheckVoucher!InvoiceAmt / 0.97
                            '                                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                                         '                                                            "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                                         '                                                            " values (" & N2Str2Null(rsCheckVoucher!VOUCHERNO) & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                                         '                                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                                         '                                                            ", " & PV_STATUS & ")"
                            '                                            gconDMIS.Execute SQL_STATEMENT
                            '                                        End If
                            '                                        Set rsCheckVoucher = Nothing
                            '                                    End If
                            '                                    Set rsCheckDetail2 = Nothing
                            '                                End If
                            '                                Set rsCheckDetail = Nothing
                            '                            End If
                            'End If

                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(PV_MRRNO)

                            Set rsCheckJournal_HD = New ADODB.Recordset
                            Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
                            If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'Y' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"

                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                                Else
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'N' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
                                    gconDMIS.Execute SQL_STATEMENT

                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                End If
                            Else
                                Set rsCheckJournal_HD = New ADODB.Recordset
                                Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
                                If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                    If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then

                                        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                         " ReceiveStatus = 'Y' " & "," & _
                                                         " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                         " Balance = Balance - " & PV_AMOUNT & _
                                                         " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                    Else
                                        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                        " ReceiveStatus = 'N' " & "," & _
                                                        " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                        " Balance = Balance - " & PV_AMOUNT & _
                                                        " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
                                        gconDMIS.Execute SQL_STATEMENT

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                                    End If
                                End If
                            End If
                        Else
                            'If COMPANY_CODE = "HGC" Then
                            'DESCRIPTION: Insert detail for A/R CREDIT CARD / Credit Card payment receive
                            If CheckIfBank(CMIS_CUSCDE) = True Then
                                '                                Dim rsCreditCard          As ADODB.Recordset
                                '                                Dim BANKCHARGES           As Double
                                '                                Dim EWT                   As Double
                                '                                Dim TOTALCHARGES          As Double
                                Set rsCreditCard = New ADODB.Recordset
                                rsCreditCard.Open "SELECT * FROM CMIS_CARDBANK WHERE CUSCDE ='" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                                If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                                    BANKCHARGES = NumericVal(rsCreditCard!BANKCHARGES) / 100
                                    EWT = NumericVal(rsCreditCard!EWT) / 100
                                    TOTALCHARGES = 1 - (BANKCHARGES + EWT)
                                End If

                                '                            Dim rsCheckCreditCardSJ As ADODB.Recordset
                                '                            Set rsCheckCreditCardSJ = New ADODB.Recordset
                                '                            'rsCheckCreditCardSJ.Open "select * from CMIS_OFF_DT where ReferenceNo = " & N2Str2Null(CMIS_DT_REFERENCENO) & "  and TRANTYPE <>'OTH'", gconDMIS, adOpenForwardOnly
                                '                            rsCheckCreditCardSJ.Open "select * from CMIS_OFF_HD where ReferenceNo = " & N2Str2Null(CMIS_DT_REFERENCENO) & " ", gconDMIS, adOpenForwardOnly
                                '                            If Not rsCheckCreditCardSJ.EOF And Not rsCheckCreditCardSJ.BOF Then
                                '                                PV_INVNO = N2Str2Null(rsCheckCreditCardSJ!OR_NUM)
                                PV_MRRNO = "'CI'"
                                'PV_MRRNO = "'" & rsCheckCreditCardSJ!TRANTYPE & "'"
                                'CREDITCARD
                                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "DGI" Then
                                    PV_INVDATE = GetCustomerORDate(CMIS_DT_REFERENCE)
                                    PV_AMOUNT = Round(GetCustomerORAmount(CMIS_DT_REFERENCE), 2)
                                Else
                                    PV_AMOUNT = Round((NumericVal(PV_AMOUNT) / TOTALCHARGES), 2)
                                End If

                                '                                Dim rsDT_Amount As ADODB.Recordset
                                '                                Set rsDT_Amount = New ADODB.Recordset
                                '                                rsDT_Amount.Open "SELECT SUM(PAYMENT) AS PAYMENT FROM CMIS_OFF_DT WHERE OR_NUM=" & PV_INVNO & " AND PAIDNA=1", gconDMIS, adOpenForwardOnly
                                '                                If Not rsDT_Amount.EOF And Not rsDT_Amount.BOF Then
                                '                                    PV_AMOUNT = (NumericVal(rsDT_Amount!payment))
                                '                                End If
                                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD"))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                                End If
                                SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE,J_CLASS)" & _
                                                " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_INVDATE & ", " & PV_AMOUNT & _
                                                ", " & PV_STATUS & "," & CMIS_DT_CUSCDE & "," & J_ACCT_CODE & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                '                            End If
                            End If
                            'End If
                        End If

                        '                        J_JITEMNO = "'0002'"
                        '                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        '                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"

                        'RO  - SERVICE REPAIR ORDER
                        If CMIS_DT_TRANTYPE = "RO" Or CMIS_DT_TRANTYPE = "SI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                                If ReturnSITerm(CMIS_DT_REFERENCE) = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "DGI" Then
                                If ReturnSITerm(CMIS_DT_REFERENCE) = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            ElseIf COMPANY_CODE = "HCI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            End If
                        End If
                        'CSH - PARTS CASH INVOICE
                        'CHG - PARTS CHARGE INVOICE
                        'If CMIS_DT_TRANTYPE = "CSH" Or CMIS_DT_TRANTYPE = "CHG" Then
                        If CMIS_DT_TRANTYPE = "PI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                                If ReturnTranType(CMIS_DT_REFERENCE, "P") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                            ElseIf COMPANY_CODE = "HCC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                            'J_ACCT_CODE = COA_AR_TRADE_PARTS
                            'J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_PARTS))
                        End If
                        'VI  - VEHICLE INVOICE
                        If CMIS_DT_TRANTYPE = "VI" Then
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                            End If
                        End If
                        'EST - SERVICE ESTIMATE
                        If CMIS_DT_TRANTYPE = "EST" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDeposit_AccountCode("SERVICE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeposit_AccountCode("SERVICE")))
                        End If
                        'USER DEFINE INVOICE
                        If CMIS_DT_TRANTYPE = "UI" Then
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                            End If
                        End If
                        '* IOC - INTER OFFICE COLLECTION
                        'BRA - BRANCH PAYMENT
                        'If CMIS_DT_TRANTYPE = "BRA" Then
                        '    J_ACCT_CODE = COA_BRANCH_LEGASPI
                        '    J_ACCT_NAME = N2Str2Null(Setacctname(COA_BRANCH_LEGASPI))
                        'End If
                        'WAR - A/R WARRANTY CLAIMS
                        'INV - INVENTORIES - GAS,OIL,LUBS
                        'CRD - CREDIT CARD PAYMENT
                        'OTH
                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then    ' BTT
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("ACCESSORIES")))
                            End If
                        ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HCC" Or COMPANY_CODE = "DGI" Then
                            If CMIS_DT_TRANTYPE = "AI" Then
                                If ReturnTranType(CMIS_DT_REFERENCE, "A") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                ElseIf COMPANY_CODE = "HCC" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            End If
                        End If

                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "DGI" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                If ReturnTranType(CMIS_DT_REFERENCE, "A") = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            End If
                        ElseIf COMPANY_CODE = "HCI" Then
                            If CMIS_DT_TRANTYPE = "MI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                        End If

                        If CMIS_DT_TRANTYPE = "OTH" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            'OTHER TRANSACTION
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                        End If
                        
                        If CMIS_DT_TRANTYPE = "OI" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            'OTHER TRANSACTION
                            J_ACCT_CODE = N2Str2Null(rsSJ_DATA!J_CLASS)
                            J_ACCT_NAME = N2Str2Null(Setacctname(rsSJ_DATA!J_CLASS))
                        End If

                        'If COMPANY_CODE = "HGC" Then
                        'Description: For Customer Deposit
                        If J_INVOICENUM <> "" Then
                            Dim rsCheckDeposit2        As ADODB.Recordset
                            Set rsCheckDeposit2 = New ADODB.Recordset
                            rsCheckDeposit2.Open "Select * from CMIS_Deposits where InvoiceNo = '" & J_INVOICENUM & "' and invoicetype = '" & OIinv & "'", gconDMIS, adOpenForwardOnly
                            If Not rsCheckDeposit2.EOF And Not rsCheckDeposit2.BOF Then
                                '                                J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                                '                                J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                                '                                J_ITEMCOUNT = 0
                                '                                J_ITEMCOUNT = J_ITEMCOUNT + 1
                                '                                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                J_DEPOSIT = Round(NumericVal(rsCheckDeposit2!amount), 2)
                            End If
                            Set rsCheckDeposit2 = Nothing
                        End If
                        'End If

                        '                        If InvoiceType = "OTH" And Left(J_ACCT_CODE, 5) = "21-07" Then
                        '                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                     '                                            "(VoucherNo,Jdate,SJ_VOUCHERNO,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status,CUSTOMERCODE)" & _
                                                     '                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & "," & N2Str2Null(rsSJ_DATA!VOUCHERNO) & ", " & PV_ITEMNO & _
                                                     '                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                     '                                            ", " & PV_STATUS & "," & N2Str2Null(rsSJ_DATA!CustomerCode) & ")"
                        '                            gconDMIS.Execute SQL_STATEMENT
                        '                        End If

                        'DESCRIPTION FOR A/R CREDIT CARD TRANSACTION
                        'J_ACCT_CODE = N2Str2Null("11-02002-00")
                        '                        If COMPANY_CODE = "HGC" Then

                        J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                        J_TAX = 0

                        If CMIS_DT_TRANTYPE = "OTH" Then
                            J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                        Else
                            J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(J_DEPOSIT), 2)
                        End If
                        J_DEBIT = 0

                        '                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        '                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))

                        If J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD")) Or J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND")) Then
                            '                            If CheckIfBank(CMIS_CUSCDE) = True Then
                            '                                    Set rsCreditCard = New ADODB.Recordset
                            '                                    rsCreditCard.Open "SELECT * FROM CMIS_CARDBANK WHERE CUSCDE ='" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                            '                                    If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                            '                                        BANKCHARGES = NumericVal(rsCreditCard!BANKCHARGES) / 100
                            '                                        EWT = NumericVal(rsCreditCard!EWT) / 100
                            '                                        TOTALCHARGES = 1 - (BANKCHARGES + EWT)
                            '                                    End If
                            '                            End If
                            '                            If CheckIfBank(CMIS_CUSCDE) = True Then
                            'CREDITCARD
                            '===updated by arjr&arnold
                            'all none credit transaction will not enter this condition
                            If CheckIfBank(CMIS_CUSCDE) = True Then
                                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "DGI" Then
                                    J_CREDIT = Round(GetCustomerORAmount(CMIS_DT_REFERENCE), 2)
                                Else
                                    J_CREDIT = Round((NumericVal(J_NET) / TOTALCHARGES), 2)
                                End If
                            End If
                        Else

                            J_CREDIT = Round(NumericVal(J_NET), 2) + Round(NumericVal(CMIS_DT_TAX), 2)
                        End If
                        '                        Else
                        '                            J_CREDIT = Round(NumericVal(J_NET), 2)
                        '                        End If

                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)

                        If CheckIfBank(CMIS_CUSCDE) = False Then
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If

                        If CheckIfBank(CMIS_CUSCDE) = False Then
                            '                        J_JITEMNO = "'0003'"
                            If CMIS_DT_TAX <> 0 Then
                                J_ITEMCOUNT = J_ITEMCOUNT + 1
                                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                                J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                                J_TAX = 0
                                J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(J_DEPOSIT), 2)
                                J_DEBIT = Round(NumericVal(CMIS_DT_TAX), 2)
                                J_CREDIT = 0

                                J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDITABLE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDITABLE")))

                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                            End If
                        End If
                        'If COMPANY_CODE = "HGC" Then
                        'DESCRIPTION: Insert in AMIS_Reference to be used in Customer Ledger for Customer and Credit Card Company
                        SQL_STATEMENT = "insert into AMIS_REFERENCE (VoucherNo,Jtype,ReferenceNo,JDate) values (" & J_VOUCHERNO & "," & J_JTYPE & "," & J_CUSTOMERCODE2 & "," & J_JDATE & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        'End If

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                        rsOFF_DT.MoveNext
                    Loop
                End If

                If CheckIfBank(CMIS_CUSCDE) = True Then
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & TOTAL_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    'DESCRIPTION: Bank Charges for credit card payment receive BANK CHARGES
                    Set rsCreditCard = New ADODB.Recordset
                    rsCreditCard.Open "SELECT * FROM CMIS_CARDBANK WHERE CUSCDE ='" & CMIS_CUSCDE & "'", gconDMIS, adOpenForwardOnly
                    If Not rsCreditCard.EOF And Not rsCreditCard.BOF Then
                        BANKCHARGES = NumericVal(rsCreditCard!BANKCHARGES) / 100
                        EWT = NumericVal(rsCreditCard!EWT) / 100
                        TOTALCHARGES = 1 - (BANKCHARGES + EWT)
                    End If

                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"

                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("BANK CHARGES"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("BANK CHARGES")))
                    J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = Round(((NumericVal(CMIS_CASHAMOUNT) / TOTALCHARGES) * NumericVal(BANKCHARGES)), 2)
                        J_BANKCHARGES = J_DEBIT
                        'CMIS_CASHAMOUNT
                    Else
                        J_DEBIT = Round(((NumericVal(CMIS_CHKAMOUNT) / TOTALCHARGES) * NumericVal(BANKCHARGES)), 2)
                        J_BANKCHARGES = J_DEBIT
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    'DESCRIPTION: Creditable Withholding Tax for credit card payment receive CREDITABLE WITHHOLDING TAX
                    '                            J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                    '                            J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    If COMPANY_CODE = "DGI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDITABLE"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDITABLE")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CREDITABLE"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CREDITABLE")))
                    End If
                    '"11-07000-00"
                    If CMIS_CASHAMOUNT > 0 Then
                        J_DEBIT = Round(((NumericVal(CMIS_CASHAMOUNT) / TOTALCHARGES) * NumericVal(EWT)), 2)
                        'J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                    Else
                        J_DEBIT = Round(((NumericVal(CMIS_CHKAMOUNT) / TOTALCHARGES) * NumericVal(EWT)), 2)
                        'J_DEBIT = Round(CreditCardPayment(J_VOUCHERNO, J_JTYPE, J_ACCT_CODE) - (NumericVal(CMIS_CHKAMOUNT) + J_BANKCHARGES), 2)
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "A", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)
                End If

                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,ReferenceNo,Bank,Entity_Class)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & "," & J_CUSTOMERCODE & "," & N2Str2Null(CMIS_CARDBANK) & "," & J_ENTITY & ")"

                gconDMIS.Execute SQL_STATEMENT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Screen.MousePointer = 0

    ImportUnDeposit = True
    Exit Function
ErrorCode:
    ImportUnDeposit = False
End Function

Function ImportDeposited() As Boolean
'HEADER
    On Error GoTo ErrorCode

    Dim J_JDATE                                        As String
    Dim J_VOUCHERNO                                    As String
    Dim J_JTYPE                                        As String
    Dim J_JNO                                          As String
    Dim J_REMARKS                                      As String
    Dim J_VENDORCODE                                   As String
    Dim J_CUSTOMERCODE                                 As String
    Dim J_OUTBALANCE                                   As Double
    Dim J_AMOUNTTOPAY                                  As Double
    Dim J_INVOICEAMT                                   As Double
    Dim J_BALANCE                                      As Double
    Dim J_AMOUNTPAID                                   As Double
    Dim J_CHECKNO                                      As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE            As String
    Dim J_INVOICETYPE, J_INVOICENO                     As String
    Dim J_CHECKDATE, J_BANKCODE                        As String
    Dim J_REFNO, J_REFDATE                             As String
    Dim J_TERMS, J_DEALER                              As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                  As String

    'DETAIL
    Dim J_ACCT_CODE, J_ACCT_NAME                       As String
    Dim J_DEBIT                                        As Double
    Dim J_CREDIT                                       As Double
    Dim J_TAX                                          As Double
    Dim J_GROSS                                        As Double
    Dim J_NET                                          As Double
    Dim J_STATUS                                       As String
    Dim J_JITEMNO                                      As String

    Dim rsJournal_HDDup                                As ADODB.Recordset

    Dim CMIS_OR_NUM                                    As String
    Dim CMIS_OR_DATE                                   As String
    Dim CMIS_OR_AMT                                    As String
    Dim CMIS_DISCOUNT                                  As String
    Dim CMIS_TAX                                       As String
    Dim CMIS_CASHAMOUNT                                As Double
    Dim CMIS_CHKAMOUNT                                 As Double
    Dim CMIS_CARDAMOUNT                                As Double
    Dim CMIS_CUSCDE                                    As String
    Dim CMIS_CUSNAME                                   As String
    Dim CMIS_DEPOSIT                                   As String
    Dim CMIS_BANKCODE                                  As String
    Dim CMIS_TSEKE                                     As String
    Dim CMIS_CHECKDATE                                 As String
    Dim CMIS_STATUS                                    As String
    Dim CMIS_TYPE_PAYMENT                              As String

    Dim CMIS_IS_VAT                                    As Boolean
    Dim CMIS_BANK_DEPOSITED                            As String

    Dim TOTAL_DEBIT, TOTAL_CREDIT                      As Double

    Dim rsOFF_HD                                       As ADODB.Recordset
    Dim rsOFF_DT                                       As ADODB.Recordset
    Dim i                                              As Long

    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
        COA_CASH_ON_HAND = Null2String(ReturnAccountCode("CASH ON HAND"))
        COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = Null2String(ReturnAccountCode("CARD"))
    Else
        '        COA_CASH_ON_HAND = "11-01000-00"
        '        COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = "11-02002-00"
        COA_CASH_ON_HAND = Null2String(ReturnAccountCode("CASH ON HAND"))
        COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD = Null2String(ReturnAccountCode("CARD ON HAND"))
    End If
    Dim GridImport                                     As Integer
    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If COMPANY_CODE = M_COMPANY_CODE Then
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 or DEPOSIT2 = 1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND OR_AMT = '" & NumericVal(Grid2.Cell(GridImport, 4).Text) & "' AND VAT = 1 Order by OR_NUM ASC")
                Else
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 or DEPOSIT2 = 1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 0 Order by OR_NUM ASC")
                End If
            Else
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT >= '" & GetCUTOFF_DATE & "' AND DATDEPOSIT <= '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 1 Order by OR_NUM ASC")
                Else
                    Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT >= '" & GetCUTOFF_DATE & "' AND DATDEPOSIT <= '" & CDate(dtpTranDate) & "' AND OR_NUM = '" & Grid2.Cell(GridImport, 3).Text & "' AND VAT = 0 Order by OR_NUM ASC")
                End If
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!DATDEPOSIT)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                If COMPANY_CODE = M_COMPANY_CODE Then
                    CMIS_DEPOSIT1 = Null2String(rsOFF_HD!DEPOSIT1)
                    CMIS_DEPOSIT2 = Null2String(rsOFF_HD!DEPOSIT2)
                Else
                    CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT)
                End If
                CMIS_BANKCODE = Null2String(rsOFF_HD!deposit_to)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT = Null2String(rsOFF_HD!TOF)

                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                    If CMIS_BANKCODE = "AUB" Then
                        CMIS_BANK_DEPOSITED = "'11-01007-00'"
                    Else
                        CMIS_BANK_DEPOSITED = Null2String(rsOFF_HD!BankAccountNo)
                    End If
                Else
                    CMIS_BANK_DEPOSITED = Null2String(rsOFF_HD!BankAccountNo)
                End If

                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0



                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'DRJ'"

                'INSERTED SEPTEMBER 8, 2007
                Set rsOFF_DT = New ADODB.Recordset
                If Grid2.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        If Null2String(rsOFF_DT!TranType) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = NumericVal(CMIS_OR_AMT)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 6))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 6))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                If CMIS_TYPE_PAYMENT = "1" Or CMIS_TYPE_PAYMENT = "2" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        If CMIS_TYPE_PAYMENT = "1" Then
                            J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    Else
                        If CMIS_CASHAMOUNT > 0 Then
                            J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", " DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_CASH_ON_HAND)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_CASH_ON_HAND))
                    If COMPANY_CODE = M_COMPANY_CODE Then
                        If CMIS_TYPE_PAYMENT = "1" Then
                            J_CREDIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_CREDIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    Else
                        If CMIS_CASHAMOUNT > 0 Then
                            J_CREDIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                        Else
                            J_CREDIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                    End If
                    J_DEBIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                End If
                If CMIS_TYPE_PAYMENT = "3" Then
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(CMIS_BANK_DEPOSITED)
                    J_ACCT_NAME = N2Str2Null(Setacctname(CMIS_BANK_DEPOSITED))
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD)
                    J_ACCT_NAME = N2Str2Null(Setacctname(COA_ACCOUNTS_RECEIVABLE_CREDIT_CARD))
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "DEPOSITED JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                End If

                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "DEPOSITED RECEIPTS JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

            End If
            Grid2.Cell(GridImport, 1).Text = 1
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    '=========================================================================================================

    Screen.MousePointer = 0
    '=========================================================================================================

    ImportDeposited = True
    Exit Function
ErrorCode:
    ImportDeposited = False
End Function

Sub ShowUnImportedPaidInvoices(VarTranType As String, VarTranno As String)
    Screen.MousePointer = 11
    Dim InvoiceType, InvoiceTypeCode                   As String
    Dim IS_Exist                                       As Byte
    LIM = 0

    If VarTranType = "PI" Or VarTranType = "MI" Or VarTranType = "AI" Then
        Set rsPMIOS_ORD_HD = New ADODB.Recordset
        Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where TYPE = '" & Left(VarTranType, 1) & "' AND (TranType = 'CSH' OR  TranType = 'CHG') and STATUS = 'P' AND tranno = '" & VarTranno & "' AND TRANDATE > '" & GetCUTOFF_DATE & "' AND TRANDATE <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by Tranno ASC")
        If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
            rsPMIOS_ORD_HD.MoveFirst:
            Do While Not rsPMIOS_ORD_HD.EOF
                LIM = LIM + 1
                If Null2String(rsPMIOS_ORD_HD!Type) = "P" Then
                    InvoiceType = "Parts"
                    InvoiceTypeCode = "PI"
                ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "A" Then
                    InvoiceType = "Accessories"
                    InvoiceTypeCode = "AI"
                ElseIf Null2String(rsPMIOS_ORD_HD!Type) = "M" Then
                    InvoiceType = "Materials"
                    InvoiceTypeCode = "MI"
                Else
                    InvoiceType = "Unknown"
                    InvoiceTypeCode = ""
                End If
                If CheckSJExisting(InvoiceTypeCode, Null2String(rsPMIOS_ORD_HD!TRANNO), Null2String(rsPMIOS_ORD_HD!TranType)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & UCase(InvoiceType) & Chr(9) & Null2String(rsPMIOS_ORD_HD!TranType) & "-" & Null2String(rsPMIOS_ORD_HD!TRANNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPMIOS_ORD_HD!NETINVAMT)) & Chr(9) & Null2String(rsPMIOS_ORD_HD!CUSTNAME) & Chr(9) & UCase(InvoiceType)
                rsPMIOS_ORD_HD.MoveNext
                DoEvents
            Loop
        End If
    End If
    If VarTranType = "SI" Then
        'PURELY INTERNAL
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = '" & VarTranno & "' AND DTE_COMP > '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckRefNoExisting("SI", Null2String(rsCSMIOS_REPOR!REP_OR)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice = '" & VarTranno & "' AND DTE_COMP >= '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!INVOICE)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If

        'PURELY WARRANTY
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        'UPDATE BY: JUN | DATE UPDATED: 01/14/2010 |DESCRIPTION: DO NOT ALLOWED TO IMPORT TRANSACTION LESS THAT CUTOFF DATE
        'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice = '" & VarTranno & "' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice = '" & VarTranno & "' AND DTE_COMP > '" & GetCUTOFF_DATE & "' AND DTE_COMP <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        'UPDATE BY: JUN
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                LIM = LIM + 1
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!INVOICE)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!INVOICE) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!INVOICE)) & Chr(9) & "SERVICE"
                rsCSMIOS_REPOR.MoveNext
                DoEvents
            Loop
        End If

    End If
    If VarTranType = "VI" Then
        Set rsSMIS_PURCHAGREE = New ADODB.Recordset
        Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & VarTranno & "' AND DATERELEASED > '" & GetCUTOFF_DATE & "' AND CONVERT(VARCHAR,DATERELEASED,101) <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
        If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
            rsSMIS_PURCHAGREE.MoveFirst:
            Do While Not rsSMIS_PURCHAGREE.EOF
                LIM = LIM + 1
                If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsSMIS_PURCHAGREE!IGNKEY_NO) & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!Code)) & Chr(9) & "VEHICLE"
                rsSMIS_PURCHAGREE.MoveNext
                DoEvents
            Loop
        End If
    End If
    Screen.MousePointer = 0
End Sub

Sub InitGrid3()
    With Grid3
        .Rows = 1
        .Cols = 7
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Inv. Type"
        .Cell(0, 3).Text = "Inv. No."
        .Cell(0, 4).Text = "Inv. Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 80
        .Column(4).Width = 80
        .Column(5).Width = 200
        .Column(6).Width = 50

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With
End Sub

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "OR Type"
        .Cell(0, 3).Text = "OR No."
        .Cell(0, 4).Text = "OR Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True

    End With

    With Grid2
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "OR Type"
        .Cell(0, 3).Text = "OR No."
        .Cell(0, 4).Text = "OR Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200

        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral

        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
    End With
    InitGrid3
End Sub

Function ImportPMISSales() As Boolean
    On Error GoTo ErrorCode
    Dim PMIOS_TRANTYPE                                 As String
    Dim PMIOS_TRANNO                                   As String
    Dim PMIOS_TRANDATE                                 As String
    Dim PMIOS_cuscde                                   As String
    Dim PMIOS_AcctName                                 As String
    Dim PMIOS_TTLINVAMT                                As Double
    Dim PMIOS_DS_AMT1                                  As Double
    Dim PMIOS_NETINVAMT                                As Double
    Dim PMIOS_NETCOST                                  As Double
    Dim PMIOS_PAY_CLASS                                As String
    Dim PMIOS_TYPE                                     As String
    Dim PMIS_TYPE                                      As String
    Dim TOTAL_DEBIT, TOTAL_CREDIT                      As Double
    Dim i                                              As Long

    i = 0

    Dim GridImport                                     As Integer
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Set rsPMIOS_ORD_HD = New ADODB.Recordset
            If UCase(Grid3.Cell(GridImport, 2).Text) = UCase("Accessories") Then
                'Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate > '" & GetCUTOFF_DATE & "' AND trandate <= '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                PMIS_TYPE = "AI"
            ElseIf UCase(Grid3.Cell(GridImport, 2).Text) = UCase("Materials") Then
                Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'M' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate > '" & GetCUTOFF_DATE & "' AND trandate <= '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                PMIS_TYPE = "MI"
            ElseIf UCase(Grid3.Cell(GridImport, 2).Text) = UCase("Parts") Then
                Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'P' AND TranType = '" & Left(Grid3.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid3.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate > '" & GetCUTOFF_DATE & "' AND trandate <= '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                PMIS_TYPE = "PI"
            Else
                GoTo NextGrid
            End If

            If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
                If CheckSJExisting(PMIS_TYPE, Null2String(rsPMIOS_ORD_HD!TRANNO), Null2String(rsPMIOS_ORD_HD!TranType)) = True Then
                    GoTo DuplicateInvoiceNo
                Else

                    PMIOS_TRANTYPE = Null2String(rsPMIOS_ORD_HD!TranType)
                    PMIOS_TRANNO = Null2String(rsPMIOS_ORD_HD!TRANNO)

                    PMIOS_TRANDATE = Null2String(rsPMIOS_ORD_HD!trandate)
                    PMIOS_cuscde = Null2String(rsPMIOS_ORD_HD!CUSTCODE)
                    'PMIOS_AcctName = Null2String(rsPMIOS_ORD_HD!AcctName)
                    PMIOS_AcctName = SetCustomerName(rsPMIOS_ORD_HD!CUSTCODE)
                    PMIOS_TTLINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!TTLINVAMT), 2)
                    PMIOS_NETINVAMT = Round(N2Str2Zero(rsPMIOS_ORD_HD!NETINVAMT), 2)
                    PMIOS_DS_AMT1 = Round(N2Str2Zero(rsPMIOS_ORD_HD!DS_AMT1), 2)



                    If COMPANY_CODE = "HAS" Then
                        ' Upate By BTT - 07082008: NET OF DISCOUNT
                        PMIOS_NETINVAMT = Round(N2Str2Zero(PMIOS_TTLINVAMT), 2) - PMIOS_DS_AMT1
                    End If

                    'PMIOS_NETCOST = Round(N2Str2Zero(rsPMIOS_ORD_HD!NETCOST), 2)
                    PMIOS_NETCOST = ReturnCostOFSALES_PARTS(PMIOS_TRANNO, PMIOS_TRANTYPE)


                    PMIOS_TYPE = Null2String(rsPMIOS_ORD_HD!Type)
                    PMIOS_PAY_CLASS = Mid(Null2String(rsPMIOS_ORD_HD!REFPISNO), 5, 1)

                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'HEADER
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If

                    J_JDATE = N2Date2Null(PMIOS_TRANDATE)
                    J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(PMIOS_cuscde)

                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0

                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = Round(PMIOS_NETINVAMT, 2)
                    J_BALANCE = Round(PMIOS_NETINVAMT, 2)
                    J_AMOUNTPAID = 0

                    J_STATUS = "'N'"

                    J_INVOICEDATE = N2Date2Null(PMIOS_TRANDATE)
                    J_INVOICENO = N2Str2Null(PMIOS_TRANNO)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(PMIOS_TRANDATE)
                    J_PAYTYPE = N2Str2Null(PMIOS_TRANTYPE)
                    If PMIOS_TYPE = "P" Then J_INVOICETYPE = "'PI'"
                    If PMIOS_TYPE = "A" Then J_INVOICETYPE = "'AI'"
                    If PMIOS_TYPE = "M" Then J_INVOICETYPE = "'MI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = "NULL"
                    J_REFDATE = N2Date2Null(PMIOS_TRANDATE)
                    J_TERMS = "NULL"
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"

                    'If CheckSJExisting(J_INVOICETYPE, PMIOS_TRANNO) = True Then
                    J_JITEMNO = "'0001'"
                    '---AL
                    If COMPANY_CODE = "HCC" Then
                        If PMIOS_TYPE = "P" Or PMIOS_TYPE = "A" Or PMIOS_TYPE = "M" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HCC" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                                End If
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))

                            Else
                                'INTERNAL
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                            End If
                        End If
                        J_GROSS = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                        J_TAX = Round(NumericVal(Round((PMIOS_TTLINVAMT / 1.12), 2) * 0.12), 2)
                        J_NET = Round(NumericVal(PMIOS_TTLINVAMT) - Round(NumericVal(J_TAX), 2), 2)
                        J_DEBIT = 0
                        J_CREDIT = NumericVal(J_NET)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                        'COST OF SALES
                        J_JITEMNO = "'0002'"
                        If PMIOS_TYPE = "P" Or PMIOS_TYPE = "A" Or PMIOS_TYPE = "M" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                'WARRANTY

                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                            Else
                                'INTERNAL
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "INTERNAL")))
                            End If
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(NumericVal(PMIOS_NETCOST), 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO

                        'INVENTORY
                        J_JITEMNO = "'0003'"
                        If PMIOS_TYPE = "P" Or PMIOS_TYPE = "A" Or PMIOS_TYPE = "M" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                        End If
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(PMIOS_NETCOST), 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO


                    Else
                        If PMIOS_TYPE = "P" Then
                            If COMPANY_CODE = "HBK" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                            ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "COUNTER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "COUNTER")))
                            Else
                                If PMIOS_PAY_CLASS = "C" Then
                                    'CUSTOMER PAID
                                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HCC" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "COUNTER"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "COUNTER")))
                                    End If
                                ElseIf PMIOS_PAY_CLASS = "W" Then
                                    'WARRANTY
                                    If COMPANY_CODE = "HPI" Then

                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "WARRANTY"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "WARRANTY")))
                                    End If
                                Else
                                    'INTERNAL
                                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                                    ElseIf COMPANY_CODE = "HPI" Then

                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "INTERNAL")))
                                    End If
                                End If
                            End If
                        End If

                        If PMIOS_TYPE = "A" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES")))
                                End If
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                'WARRANTY
                                If COMPANY_CODE = "HPI" Then

                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES")))
                                End If
                            Else
                                'INTERNAL
                                If COMPANY_CODE <> "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES")))
                                End If
                            End If
                        End If
                        If PMIOS_TYPE = "M" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                                ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("MATERIALS")))
                                End If
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                'WARRANTY
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("MATERIALS")))
                                End If
                            Else
                                'INTERNAL
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("MATERIALS")))
                                End If
                            End If
                        End If
                        J_GROSS = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                        J_TAX = Round(NumericVal(Round((PMIOS_TTLINVAMT / 1.12), 2) * 0.12), 2)
                        J_NET = Round(NumericVal(PMIOS_TTLINVAMT) - Round(NumericVal(J_TAX), 2), 2)
                        J_DEBIT = 0
                        J_CREDIT = NumericVal(J_NET)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                        If COMPANY_CODE = "HPI" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                            End If
                        Else
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                        End If
                        'COST OF SALES
                        J_JITEMNO = "'0002'"
                        If PMIOS_TYPE = "P" Then
                            If COMPANY_CODE = "HBK" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                            ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "COUNTER"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "COUNTER")))
                            Else
                                If PMIOS_PAY_CLASS = "C" Then
                                    'CUSTOMER PAID
                                    If COMPANY_CODE = "HPI" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "COUNTER"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "COUNTER")))
                                    End If
                                ElseIf PMIOS_PAY_CLASS = "W" Then
                                    'WARRANTY
                                    If COMPANY_CODE = "HPI" Then
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "WARRANTY"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "WARRANTY")))
                                    End If
                                Else
                                    'INTERNAL
                                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "INTERNAL")))
                                    End If
                                End If
                            End If
                        End If

                        If PMIOS_TYPE = "A" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES")))
                                End If
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                'WARRANTY
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES")))
                            Else
                                'INTERNAL
                                If COMPANY_CODE <> "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES")))
                                End If
                            End If
                        End If
                        If PMIOS_TYPE = "M" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                'CUSTOMER PAID
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                                ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("MATERIALS")))
                                End If
                            ElseIf PMIOS_PAY_CLASS = "W" Then
                                'WARRANTY
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("MATERIALS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("MATERIALS")))
                            Else
                                'INTERNAL
                                If COMPANY_CODE = "HPI" Then

                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("MATERIALS")))
                                End If
                            End If
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(NumericVal(PMIOS_NETCOST), 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        If COMPANY_CODE = "HPI" Then
                            If PMIOS_PAY_CLASS = "C" Then
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                            End If
                        Else
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                        End If
                        'INVENTORY
                        J_JITEMNO = "'0003'"
                        If PMIOS_TYPE = "P" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                        End If
                        If PMIOS_TYPE = "A" Then
                            If COMPANY_CODE = "HPI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                            End If
                        End If
                        If PMIOS_TYPE = "M" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                        End If
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(PMIOS_NETCOST), 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If
                    '---ALEND
                    'OUTPUT TAX

                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        If PMIOS_PAY_CLASS = "I" Then

                        Else
                            J_JITEMNO = "'0004'"
                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(Round((PMIOS_NETINVAMT / 1.12), 2) * 0.12), 2)
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            NEW_LogAudit "MM", "JOURNAL IMPORT", SQL_STATEMENT, J_VOUCHERNO, "", CSMIOS_REP_OR, J_JTYPE, J_JNO
                        End If
                    Else
                        J_JITEMNO = "'0004'"
                        J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(Round((PMIOS_NETINVAMT / 1.12), 2) * 0.12), 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If
                    'A/R PARTS
                    J_JITEMNO = "'0005'"
                    'CLEARING ACCOUNT - HGC - FML - 12/19/2007
                    If PMIOS_PAY_CLASS = "C" Then
                        If COMPANY_CODE = "HMH" Then
                            If PMIOS_TRANTYPE = "CHG" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                            End If
                        ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                        ElseIf COMPANY_CODE = "HAS" Then
                            J_ACCT_CODE = N2Str2Null("11-02000-00")
                            J_ACCT_NAME = N2Str2Null(Setacctname("11-02000-00"))
                        ElseIf COMPANY_CODE = "HCC" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                        End If
                    Else
                        If COMPANY_CODE <> "HPI" Then
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INSTALLED"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INSTALLED")))
                            End If
                        End If
                    End If
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAS" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCC" Then
                        If PMIOS_PAY_CLASS = "I" Then
                            'Update bY: BTT - 07092008
                            'J_GROSS = Round(NumericVal(PMIOS_TTLINVAMT), 2)

                            If COMPANY_CODE = "HPI" Then
                                '                            If PMIOS_TYPE = "P" Or PMIOS_TYPE = "M" Then
                                '                                J_ACCT_CODE = "'71-37100-20'"
                                '                                J_ACCT_NAME = N2Str2Null(Setacctname("'71-37100-20'"))
                                '                                J_DEBIT = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                                '                            End If
                            Else
                                J_TAX = Round(NumericVal(Round((PMIOS_TTLINVAMT / 1.12), 2) * 0.12), 2)
                                J_NET = Round(NumericVal(PMIOS_TTLINVAMT) - Round(NumericVal(J_TAX), 2), 2)
                                J_DEBIT = NumericVal(J_NET)
                                J_CREDIT = 0
                            End If
                        Else
                            J_DEBIT = Round(NumericVal(PMIOS_NETINVAMT), 2)
                        End If
                    End If
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    If COMPANY_CODE = "HPI" Then
                        If PMIOS_TYPE = "P" Or PMIOS_TYPE = "M" Then
                            J_ACCT_CODE = "'71-37100-20'"
                            J_ACCT_NAME = N2Str2Null(Setacctname("'71-37100-20'"))
                            J_DEBIT = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                        Else
                            If J_ACCT_CODE = "'11-02009-00'" Then
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                            End If
                        End If
                    Else
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If

                    If PMIOS_DS_AMT1 > 0 Then
                        'DISCOUNT
                        J_JITEMNO = "'0006'"
                        If PMIOS_TYPE = "P" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS")))
                        End If
                        If PMIOS_TYPE = "A" Then
                            If COMPANY_CODE = "HPI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("ACCESSORIES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("ACCESSORIES")))
                            End If
                        End If
                        If PMIOS_TYPE = "M" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS")))
                        End If
                        J_GROSS = Round(NumericVal(PMIOS_DS_AMT1), 2)
                        J_TAX = Round(NumericVal(Round((PMIOS_DS_AMT1 / 1.12), 2) * 0.12), 2)
                        J_NET = Round(NumericVal(PMIOS_DS_AMT1) - Round(NumericVal(J_TAX), 2), 2)
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                        J_CREDIT = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "AA", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    Else
                        If PMIOS_PAY_CLASS = "I" Then
                            If PMIOS_TYPE = "A" Then
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SALES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SALES")))
                                    J_DEBIT = Round(NumericVal(PMIOS_TTLINVAMT), 2)
                                    J_CREDIT = 0
                                    J_STATUS = "'N'"

                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "AA", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                                End If
                            End If
                        End If
                    End If
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                    ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    Grid3.Cell(GridImport, 1).Text = 1
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                End If
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
NextGrid:
DuplicateInvoiceNo:
    Next
    ImportPMISSales = True
    Exit Function
ErrorCode:
    ImportPMISSales = False
End Function

Function ImportSMISSales() As Boolean
    On Error GoTo ErrorCode

    Dim i                                              As Integer
    Dim GridImport                                     As Integer
    Dim rsSMIS_PURCHAGREE                              As ADODB.Recordset
    Dim SMIS_VI_NO                                     As String
    Dim SMIS_DATERELEASED                              As String
    Dim SMIS_CODE                                      As String
    Dim SMIS_NETSALESPRICE                             As Double
    Dim SMIS_OTHERS                                    As Double
    Dim SMIS_FOB                                       As Double
    Dim SMIS_TOTALCOST                                 As Double
    Dim SMIS_TERMS                                     As String
    'VEHICLE SALES TRANSACTION
    '=========================================================================================================
    i = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Set rsSMIS_PURCHAGREE = New ADODB.Recordset
            Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & Grid3.Cell(GridImport, 3).Text & "' AND STATUS = 'P' AND CONVERT(VarChar, DateReleased, 101)  <= '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
            If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
                If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
                    GoTo DuplicateInvoiceNo
                Else
                    SMIS_VI_NO = Null2String(rsSMIS_PURCHAGREE!VI_NO)
                    SMIS_DATERELEASED = Null2String(rsSMIS_PURCHAGREE!DATERELEASED)
                    SMIS_CODE = Null2String(rsSMIS_PURCHAGREE!Code)
                    SMIS_FOB = N2Str2Zero(rsSMIS_PURCHAGREE!FREIGHT)
                    SMIS_OTHERS = N2Str2Zero(rsSMIS_PURCHAGREE!OTHERS)
                    SMIS_TOTALCOST = N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL_COST)
                    SMIS_TERMS = Null2String(rsSMIS_PURCHAGREE!TERM)
                    If Null2String(rsSMIS_PURCHAGREE!TERM) = "F" Then
                        SMIS_NETSALESPRICE = (N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE) + SMIS_FOB)    '- SMIS_OTHERS
                    Else
                        SMIS_NETSALESPRICE = N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)
                    End If

                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'HEADER
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_JNO = "'000001'"
                    End If

                    J_JDATE = N2Date2Null(SMIS_DATERELEASED)
                    J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(SMIS_CODE)

                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0

                    J_AMOUNTTOPAY = 0
                    J_INVOICEAMT = SMIS_NETSALESPRICE
                    J_BALANCE = SMIS_NETSALESPRICE
                    J_AMOUNTPAID = 0

                    J_STATUS = "'N'"

                    J_INVOICEDATE = N2Date2Null(SMIS_DATERELEASED)
                    J_INVOICENO = N2Str2Null(SMIS_VI_NO)
                    J_CHECKNO = "NULL"
                    J_DUEDATE = N2Date2Null(SMIS_DATERELEASED)
                    J_PAYTYPE = "NULL"
                    J_INVOICETYPE = "'VI'"
                    J_CHECKDATE = "NULL"
                    J_BANKCODE = "NULL"
                    J_REFNO = "NULL"
                    J_REFDATE = N2Date2Null(SMIS_DATERELEASED)
                    J_TERMS = "NULL"
                    J_DEALER = "NULL"
                    J_PAIDSTATUS = "'N'"
                    J_RECEIVESTATUS = "'N'"


                    'SALES
                    J_JITEMNO = "'0001'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", "VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", "VEHICLES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!Model)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!Model))))
                    End If
                    If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                        'J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                        J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        J_TAX = Round(NumericVal(Round((J_GROSS / 1.12), 2) * 0.12), 2)
                        'J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                        J_NET = Round(NumericVal(J_GROSS) - Round(NumericVal(J_TAX), 2), 2)
                    Else
                        'J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_OTHERS), 2)
                        J_GROSS = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        J_TAX = 0
                        J_NET = Round(J_GROSS, 2)
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    'COST OF SALES
                    J_JITEMNO = "'0002'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SALES", "VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SALES", "VEHICLES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SALES", Null2String(rsSMIS_PURCHAGREE!Model)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SALES", Null2String(rsSMIS_PURCHAGREE!Model))))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0

                    'J_DEBIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                    J_DEBIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 1.12) * 0.12, 2)
                    J_CREDIT = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    'INVENTORY
                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!Model)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", Null2String(rsSMIS_PURCHAGREE!Model))))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", ""))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 1.12) * 0.12, 2)
                    'J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                    J_JITEMNO = "'0004'"
                    If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                        'OUTPUT TAX
                        J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(Round((SMIS_NETSALESPRICE / 1.12), 2) * 0.12), 2)
                        'J_CREDIT = Round(NumericVal(SMIS_NETSALESPRICE / 9.3333), 2)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                        gconDMIS.Execute SQL_STATEMENT

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)



                        J_JITEMNO = "'0005'"
                    Else
                        J_JITEMNO = "'0004'"
                    End If


                    '            'A/R VEHICLES
                    '            J_JITEMNO = "'0001'"
                    If SMIS_TERMS = "COD" Then
                        If COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                        ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                        End If
                    Else
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("FINANCE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("FINANCE")))
                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HCI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("FINANCING"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("FINANCING")))
                        End If
                    End If

                    If COMPANY_CODE <> "HMH" Then
                        If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        Else
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        End If
                    Else
                        If SMIS_TERMS = "F" Or SMIS_TERMS = "BPO" Then
                            J_DEBIT = Round(NumericVal(SMIS_BALTOFINANCED), 2)
                        Else
                            If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                                J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                            Else
                                J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                            End If
                        End If
                    End If

                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    If COMPANY_CODE = "HMH" Then
                        If SMIS_TERMS = "F" Or SMIS_TERMS = "BPO" Then
                            If J_JITEMNO = "'0004'" Then
                                J_JITEMNO = "'0005'"
                            Else
                                J_JITEMNO = "'0006'"
                            End If

                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                            J_DEBIT = Round(NumericVal(SMIS_DOWNPAYMENT), 2)

                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                        End If
                    End If

                    '            'SALES PARTS
                    '
                    '            'If SMIS_FOB > 0 Then
                    '            '   'FREIGHT AND HANDLING
                    '            '   J_JITEMNO = "'0003'"
                    '            '   J_ACCT_CODE = COA_COSTOFSALES_DISCOUNT_VEHICLES
                    '            '   J_ACCT_NAME = N2Str2Null(SetAcctName(COA_COSTOFSALES_DISCOUNT_VEHICLES))
                    '            '   J_GROSS = NumericVal(SMIS_FOB)
                    '            '   J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                    '            '   J_NET = NumericVal(J_GROSS) - Round(NumericVal(J_TAX),2)
                    '            '   J_DEBIT = NumericVal(J_NET)
                    '            '   J_CREDIT = 0
                    '            '   J_STATUS = "'N'"
                    '            '   TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '
                    '            '   gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 '                '                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                 '                '                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 '                '                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 '                '                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    '            '
                    '            '   J_JITEMNO = "'0004'"
                    '            'Else
                    '            '   J_JITEMNO = "'0003'"
                    '            'End If
                    '            J_JITEMNO = "'0003'"
                    '            If SMIS_OTHERS > 0 Then
                    '                'DISCOUNT
                    '                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL)))
                    '                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!MODEL))))
                    '                J_GROSS = NumericVal(SMIS_OTHERS)
                    '                J_TAX = Round(NumericVal(J_GROSS) / 9.3333, 2)
                    '                J_NET = NumericVal(J_GROSS) - Round(NumericVal(J_TAX),2)
                    '                J_DEBIT = NumericVal(J_NET)
                    '                J_CREDIT = 0
                    '                J_STATUS = "'N'"
                    '                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '
                    '                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     '                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     '                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     '                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     '                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    '                J_JITEMNO = "'0004'"
                    '            End If
                    '
                    '
                    '            If J_JITEMNO = "'0005'" Then
                    '                J_JITEMNO = "'0006'"
                    '            Else
                    '                J_JITEMNO = "'0005'"
                    '            End If


                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                    ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                    Grid3.Cell(GridImport, 1).Text = 1
                End If
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
DuplicateInvoiceNo:
    Next
    ImportSMISSales = True
    Exit Function
ErrorCode:
    ImportSMISSales = False
End Function

Function ImportPurelyInternal() As Boolean
    On Error GoTo ErrorCode

    Dim GridImport                                     As Integer
    Dim i                                              As Integer
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & Grid3.Cell(GridImport, 2).Text & "' and dte_comp > '" & GetCUTOFF_DATE & "' AND dte_comp <= '" & CDate(dtpTranDate) & "' order by invoice ASC")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)
                CSMIOS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)

                'INTERNAL - COMPANY
                '====================================================================================================================================================================================

                COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2) Else COMPANY_DIRECT_EXPENSE_LABOR = 0

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2) Else COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2) Else COMPANY_DIRECT_EXPENSE_GOL = 0
                '====================================================================================================================================================================================

                'INTERNAL - SALES DEPARTMENT
                '====================================================================================================================================================================================

                SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                    SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                Else
                    SALES_DIRECT_EXPENSE_LABOR = 0
                End If

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                Else
                    SALES_DIRECT_EXPENSE_SPAREPARTS = 0
                End If

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                Else
                    SALES_DIRECT_EXPENSE_GOL = 0
                End If

                '====================================================================================================================================================================================

                '=========================================================================================================================================================
                'ENTRY FOR PURELY INTERNAL
                'UPDATED BY: JUN - UPDATE DUE TO ERROR IN SAVING JNO
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    WARRANTY_JNO = "000001"
                End If
                'UPDATED BY: JUN

                If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT = 0 Then

                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"

                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'": J_REMARKS = "NULL": J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0: J_AMOUNTTOPAY = 0
                    CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)

                    J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_AMOUNTPAID = 0
                    J_STATUS = "'N'"

                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)

                    J_CHECKNO = "NULL": J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL): J_PAYTYPE = "NULL": J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL": J_BANKCODE = "NULL": J_REFNO = N2Str2Null(CSMIOS_REP_OR): J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM): J_DEALER = "NULL": J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"

                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()), "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'LABOR

                    'to check if there is more than 1 purely charge to internal : update By BTT
                    If PosibleDoubleInternal(CSMIOS_REP_OR) = True Then Exit Function

                    INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0:
                    INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0:

                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!detcost)
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                        End If
                        J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                        J_TAX = 0
                        J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                        If COMPANY_CODE = "HPI" Then
                            J_CREDIT = 0
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                        Else
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                        End If

                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO

                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        Else
                            'COST OF SALES
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            End If
                            J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        Else
                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If

                    'PARTS
                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' ) AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_PARTS_AMT > 0 Then
                            If COMPANY_CODE = "HPI" Then
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                                End If
                                J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If
                            If COMPANY_CODE = "HPI" Then
                            Else
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    End If
                                End If
                                J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                End If
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If

                    End If

                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_MATERIALS_AMT > 0 Then
                            If COMPANY_CODE = "HPI" Then
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                End If

                                J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If

                            If COMPANY_CODE = "HPI" Then
                            Else
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        If COMPANY_CODE = "HSB" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        End If
                                    Else
                                        If COMPANY_CODE = "HSB" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        End If
                                    End If
                                End If
                                J_GROSS = 0
                                J_TAX = 0
                                J_NET = 0
                                J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                End If
                            End If
                            J_GROSS = 0: J_TAX = 0: J_NET = 0
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If
                    End If


                    'OUT PUT TAX
                    If COMPANY_CODE = "HPI" Then
                        If INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST > 0 Then
                            If CSMIOS_VAT_EXEMPT = False Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                    Else

                                        If COMPANY_CODE = "HBK" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                        End If
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(Round(((INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST)), 2) * 0.12), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                J_TAX = 0: J_GROSS = 0: J_NET = 0
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        End If
                    End If


                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                    CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                    CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)

                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)

                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                    ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    ':-)=================================
                    Grid3.Cell(GridImport, 1).Text = 1
                End If
                '=========================================================================================================================================================
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    ImportPurelyInternal = True
    Exit Function
ErrorCode:
    ImportPurelyInternal = False
End Function

Sub ImportPurelyInternalNew()
    On Error GoTo ErrorCode

            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & Grid2.Cell(GridImport, 2).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)
                CSMIOS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)

                'INTERNAL - COMPANY
                '====================================================================================================================================================================================

                COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2) Else COMPANY_DIRECT_EXPENSE_LABOR = 0

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2) Else COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2) Else COMPANY_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                           "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                           "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                           "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                           "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'C') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    COMPANY_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES))
                Else
                    COMPANY_DIRECT_EXPENSE_ACCESSORIES = 0
                End If

                '====================================================================================================================================================================================

                'INTERNAL - SALES DEPARTMENT
                '====================================================================================================================================================================================

                SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0

                Set rsCSMIOS_LABOR = New ADODB.Recordset
                Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
                    SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
                Else
                    SALES_DIRECT_EXPENSE_LABOR = 0
                End If

                Set rsCSMIOS_PARTS = New ADODB.Recordset
                Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
                    SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
                Else
                    SALES_DIRECT_EXPENSE_SPAREPARTS = 0
                End If

                Set rsCSMIOS_MATERIALS = New ADODB.Recordset
                Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
                If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
                    SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
                Else
                    SALES_DIRECT_EXPENSE_GOL = 0
                End If

                Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
                rsCSMIOS_ACCESSORIES.Open ("SELECT SUM(CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) AS ACCESSORIES, " & _
                                           "SUM((CSMS_Ro_Det.DETVOL * CSMS_Ro_Det.DETPRC) * (CSMS_Ro_Det.DISCRATE / 100)) As DISCOUNT " & _
                                           "FROM CSMS_Repor INNER JOIN CSMS_Ro_Det " & _
                                           "ON CSMS_Repor.REP_OR = CSMS_Ro_Det.REP_OR " & _
                                           "WHERE (CSMS_Ro_Det.LIVIL = '4') AND (CSMS_Ro_Det.WCODE = 'S') AND  CSMS_Repor.REP_OR= " & N2Str2Null(CSMIOS_REP_OR)), gconDMIS, adOpenForwardOnly
                If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
                    SALES_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES))
                Else
                    SALES_DIRECT_EXPENSE_ACCESSORIES = 0
                End If
                '====================================================================================================================================================================================

                '=========================================================================================================================================================
                'ENTRY FOR PURELY INTERNAL
                'UPDATED BY: JUN - UPDATE DUE TO ERROR IN SAVING JNO
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    WARRANTY_JNO = "000001"
                End If
                'UPDATED BY: JUN

                'If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT = 0 Then
                If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL + SALES_DIRECT_EXPENSE_LABOR + SALES_DIRECT_EXPENSE_SPAREPARTS + SALES_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT = 0 Then

                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"

                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'": J_REMARKS = "NULL": J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0: J_AMOUNTTOPAY = 0
                    CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)

                    J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_AMOUNTPAID = 0
                    J_STATUS = "'N'"

                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)

                    J_CHECKNO = "NULL": J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL): J_PAYTYPE = "NULL": J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL": J_BANKCODE = "NULL": J_REFNO = N2Str2Null(CSMIOS_REP_OR): J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM): J_DEALER = "NULL": J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"

                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()), "000000"))
                    WARRANTY_ItemCnt = 0
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'LABOR

                    'to check if there is more than 1 purely charge to internal : update By BTT
                    If PosibleDoubleInternal(CSMIOS_REP_OR) = True Then Exit Sub

                    INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0:
                    INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0:

                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!detcost)
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    If COMPANY_CODE = "HMH" Then
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost)), 2)
                                    Else
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    End If
                                    'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If COMPANY_CODE = "HMH" Then

                        Else
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                            End If
                            J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                            If COMPANY_CODE = "HPI" Then
                                J_CREDIT = 0
                                J_DEBIT = Round(NumericVal(J_NET), 2)
                            Else
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                            End If

                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                        End If

                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HMH" Then
                        Else
                            'COST OF SALES
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                            End If
                            J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        Else
                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If

                    'PARTS
                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' ) AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    If COMPANY_CODE = "HMH" Then
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                    Else
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    End If
                                    'J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_PARTS_AMT > 0 Then
                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                                End If
                                J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If

                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                            Else
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                                    End If
                                End If
                                J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                End If
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If

                    End If

                    'MATERIALS
                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                        rsINTERNAL_RO_DET.MoveFirst
                        Do While Not rsINTERNAL_RO_DET.EOF
                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    If COMPANY_CODE = "HMH" Then
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                    Else
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    End If
                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                                    J_CREDIT = 0
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If
                            End If
                            rsINTERNAL_RO_DET.MoveNext
                        Loop

                        If INTERNAL_MATERIALS_AMT > 0 Then
                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                End If

                                J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                Else
                                    J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                End If
                                J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                            End If

                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                            Else
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        End If
                                    Else
                                        If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                        End If
                                    End If
                                End If
                                J_GROSS = 0
                                J_TAX = 0
                                J_NET = 0
                                J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                                J_CREDIT = 0
                                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If

                            'INVENTORY
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                            Else
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                End If
                            End If
                            J_GROSS = 0: J_TAX = 0: J_NET = 0
                            J_DEBIT = 0
                            J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If
                    End If

                    If COMPANY_CODE = "HMH" Then
                        'ACCESSORIES
                        Set rsINTERNAL_RO_DET = New ADODB.Recordset
                        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '4' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                            rsINTERNAL_RO_DET.MoveFirst
                            Do While Not rsINTERNAL_RO_DET.EOF
                                If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                                    INTERNAL_ACCESSORIES_AMT = INTERNAL_ACCESSORIES_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                                    INTERNAL_ACCESSORIES_COST = INTERNAL_ACCESSORIES_AMT + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                                    If COMPANY_CODE = "HPI" Then
                                    Else
                                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                                        J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                        J_TAX = 0
                                        If COMPANY_CODE = "HMH" Then
                                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL)), 2)
                                        Else
                                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                        End If
                                        J_DEBIT = Round(NumericVal(J_NET), 2)
                                        J_CREDIT = 0
                                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                        ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                        gconDMIS.Execute SQL_STATEMENT
                                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                    End If
                                End If
                                rsINTERNAL_RO_DET.MoveNext
                            Loop

                            If INTERNAL_MATERIALS_AMT > 0 Then
                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                    End If

                                    J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                    Else
                                        J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    End If
                                    J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                                    J_DEBIT = 0
                                    J_CREDIT = Round(NumericVal(J_NET), 2)
                                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & _
                                                    ", " & WARRANTY_JNO & ", " & WARRANTY_J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                                End If

                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                                Else
                                    'COST OF SALES
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                    Else
                                        If CSMIOS_TERM = "CSH" Then
                                            If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                            Else
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                            End If
                                        Else
                                            If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                            Else
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                                            End If
                                        End If
                                    End If
                                    J_GROSS = 0
                                    J_TAX = 0
                                    J_NET = 0
                                    J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                                    J_CREDIT = 0
                                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If

                                'INVENTORY
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES")))
                                    End If
                                End If
                                J_GROSS = 0: J_TAX = 0: J_NET = 0
                                J_DEBIT = 0
                                J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        End If
                    End If
                    'OUT PUT TAX
                    If COMPANY_CODE = "HPI" Then
                        If INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST > 0 Then
                            If CSMIOS_VAT_EXEMPT = False Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                    Else

                                        If COMPANY_CODE = "HBK" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                        End If
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(Round(((INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST)), 2) * 0.12), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                J_TAX = 0: J_GROSS = 0: J_NET = 0
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        End If
                    End If


                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                    CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                    CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)

                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)

                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)

                    WARRANTY_J_AMOUNTTOPAY = 0
                    WARRANTY_J_INVOICEAMT = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_BALANCE = Round(NumericVal(COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL), 2)
                    WARRANTY_J_AMOUNTPAID = 0
                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                    ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT
                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(WARRANTY_VOUCHERNO), J_JTYPE, WARRANTY_JNO
                    ':-)=================================
                End If
                '=========================================================================================================================================================
        End If
ErrorCode:
    ShowVBError
End Sub

Function ImportCSMSSales() As Boolean
    On Error GoTo ErrorCode
    Dim i                                              As Integer
    Dim GridImport                                     As Integer
    Dim rsSublet                                       As ADODB.Recordset
    i = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset

            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            'ORIGINAL CODE:
            SQL_STATEMENT = "Select * from CSMS_REPOR where REP_OR = '" & Grid3.Cell(GridImport, 2).Text & "' and dte_comp > '" & GetCUTOFF_DATE & "' AND dte_comp <= '" & CDate(dtpTranDate) & "' order by invoice ASC"

            'FOR DEBUGGING PURPOSE ONLY
            'SQL_STATEMENT = "Select * from CSMS_REPOR where REP_OR = 'R-00003743' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC"

            Set rsCSMIOS_REPOR = gconDMIS.Execute(SQL_STATEMENT)
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!INVOICE)) = True Then
                    GoTo DuplicateInvoiceNo
                Else
                    ItemCnt = 0
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                    CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                    CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!INVOICE)
                    CSMIOS_VAT_EXEMPT = Null2Bool(rsCSMIOS_REPOR!VAT_EXEMPT)
                    J_JNO = "'" & Format(GetJNo(), "000000") & "'"
                    CSMIOS_RO_AMOUNT = Round(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT), 2)
                    'INITIALIZE ALL VARIABLES
                    TOTAL_DISCOUNT_AMOUNT = 0
                    INITIALIZE_ACCESSORIESCOST CSMIOS_REP_OR
                    INITIALIZE_SALES_AND_DISCOUNT CSMIOS_REP_OR
                    INITIALIZE_COST_VARIABLES CSMIOS_REP_OR
                    INITIALIZE_WARRANTY CSMIOS_REP_OR
                    INITIALIZE_INTERNAL_COMPANY CSMIOS_REP_OR
                    INITIALIZE_INTERNAL_SALES_DEPARTMENT CSMIOS_REP_OR
                    INITIALIZE_INSURANCE CSMIOS_REP_OR
                    '====================================================================================================================================================================================
                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    'HEADER
                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    J_JDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_VOUCHERNO = N2Str2Null(GetSJVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
                    J_AMOUNTTOPAY = 0
                    CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)
                    J_INVOICEAMT = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_BALANCE = Round(NumericVal(CSMIOS_RO_AMOUNT), 2)
                    J_AMOUNTPAID = 0
                    J_STATUS = "'N'"
                    J_INVOICEDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_INVOICENO = N2Str2Null(CSMIOS_INVOICE)
                    J_CHECKNO = "NULL": J_DUEDATE = N2Date2Null(CSMIOS_DTE_REL): J_PAYTYPE = "NULL": J_INVOICETYPE = "'SI'"
                    J_CHECKDATE = "NULL": J_BANKCODE = "NULL": J_REFNO = N2Str2Null(CSMIOS_REP_OR): J_REFDATE = N2Date2Null(CSMIOS_DTE_REL)
                    J_TERMS = N2Str2Null(CSMIOS_TERM): J_DEALER = "NULL": J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
                    'DETAIL



                    If J_INVOICEAMT > 0 Then
                        'SALES SERVICE
                        If COMPANY_CODE = "HCC" Then
                            If CSMIOS_LABOR > 0 Or CSMIOS_SUBLET > 0 Or CSMIOS_PMS > 0 Then
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                J_GROSS = NumericVal(CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_PMS)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_PMS), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_PMS), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_PMS / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_PMS) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                'COST OF SALES
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                                J_DEBIT = Round(CSMIOS_LABOR_COST + CSMIOS_SUBLET_COST + CSMIOS_PMS_COST, 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                'INVENTORY
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                J_DEBIT = 0
                                J_CREDIT = Round(CSMIOS_LABOR_COST + CSMIOS_SUBLET_COST + CSMIOS_PMS_COST, 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        Else
                            If CSMIOS_LABOR > 0 Then
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                J_GROSS = NumericVal(CSMIOS_LABOR)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_LABOR), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_LABOR / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_LABOR) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                Else
                                    If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CSMIOS_PARTICIPAT <> "" And INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                                    Else
                                        'COST OF SALES
                                        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                                        J_DEBIT = Round(CSMIOS_LABOR_COST, 2)
                                        J_CREDIT = 0
                                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                        'INVENTORY
                                        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                        J_DEBIT = 0
                                        J_CREDIT = Round(CSMIOS_LABOR_COST, 2)
                                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                    End If
                                End If
                            End If
                        End If
                        'SUBLET
                        If COMPANY_CODE = "HCC" Then
                        Else
                            If CSMIOS_SUBLET > 0 Then
                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL")))
                                End If
                                J_GROSS = NumericVal(CSMIOS_SUBLET)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_SUBLET), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_SUBLET / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_SUBLET) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HOT" Or ((COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CSMIOS_PARTICIPAT <> "" And INSURANCE_DIRECT_EXPENSE_LABOR > 0) Or (COMPANY_CODE = "HCI" And CSMIOS_PARTICIPAT <> "") Then
                                Else
                                    'COST OF SALES
                                    ItemCnt = ItemCnt + 1
                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL")))
                                    End If
                                    J_DEBIT = Round(CSMIOS_SUBLET_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If

                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or ((COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CSMIOS_PARTICIPAT <> "" And INSURANCE_DIRECT_EXPENSE_LABOR > 0) Or (COMPANY_CODE = "HCI" And CSMIOS_PARTICIPAT <> "") Then
                                Else
                                    'INVENTORY
                                    ItemCnt = ItemCnt + 1
                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HCI" Then
                                        Set rsSublet = New ADODB.Recordset
                                        rsSublet.Open "SELECT DETCDE FROM CSMS_RO_DET WHERE REP_OR='" & CSMIOS_REP_OR & "' AND LEFT(DETCDE,2)='SR'", gconDMIS, adOpenForwardOnly
                                        If Not rsSublet.EOF And Not rsSublet.BOF Then
                                            DetCodeLen = Len(Null2String(rsSublet!DETCDE))
                                            SubletType = UCase(Mid(Null2String(rsSublet!DETCDE), 3, DetCodeLen))
                                        End If
                                        Set rsSublet = Nothing
                                        If SubletType = "LABOR" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "LABOR"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "LABOR")))
                                        ElseIf SubletType = "PARTS" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "PARTS")))
                                        ElseIf SubletType = "MATERIALS" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "MATERIALS")))
                                        End If
                                    Else
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                        End If
                                    End If
                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_SUBLET_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If
                            End If
                        End If
                        'BODY AND PAINT
                        If CSMIOS_TINSPAINT > 0 Then
                            ItemCnt = ItemCnt + 1
                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                ' SALES
                                If COMPANY_CODE = "HBK" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                ElseIf COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY", "RETAIL")))
                                End If
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "BODY", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "BODY", "RETAIL")))
                            End If
                            J_GROSS = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                            If CSMIOS_INVOICE = "PDI RO" Then
                                J_TAX = 0
                                J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                            Else
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_TINSPAINT), 2)
                                Else
                                    J_TAX = Round(NumericVal(Round((CSMIOS_TINSPAINT / 1.12), 2) * 0.12), 2)
                                    J_NET = Round(NumericVal(CSMIOS_TINSPAINT) - Round(NumericVal(J_TAX), 2), 2)
                                End If
                            End If
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            Else
                                If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CSMIOS_PARTICIPAT <> "" And INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                                Else
                                    'COST OF SALES
                                    ItemCnt = ItemCnt + 1
                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "BODY", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "BODY", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "BODY", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "BODY", "RETAIL")))
                                    End If
                                    J_DEBIT = Round(CSMIOS_TINSPAINT_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
    
                                    'INVENTORY
                                    ItemCnt = ItemCnt + 1
                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                    End If
                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_TINSPAINT_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If
                            End If
                        End If


                        'PMS
                        If COMPANY_CODE = "HCC" Then
                        Else
                            If CSMIOS_PMS > 0 Then
                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    'SALES
                                    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DGI" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC")))
                                    End If
                                End If
                                J_GROSS = Round(NumericVal(CSMIOS_PMS), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PMS), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_PMS), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_PMS / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_PMS) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                Else
                                    If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And CSMIOS_PARTICIPAT <> "" And INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                                    Else
                                        'COST OF SALES
                                        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                        If CSMIOS_TERM = "CSH" Then
                                            If COMPANY_CODE = "HSB" Or COMPANY_CODE = "DGI" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                                            Else
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "DIAGNOSTIC"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "DIAGNOSTIC")))
                                            End If
                                        Else
                                            If COMPANY_CODE = "HSB" Then
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                                            Else
                                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "DIAGNOSTIC"))
                                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "DIAGNOSTIC")))
                                            End If
                                        End If
                                        J_DEBIT = Round(CSMIOS_PMS_COST, 2)
                                        J_CREDIT = 0
                                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
    
                                        'INVENTORY
                                        ItemCnt = ItemCnt + 1
                                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                                        End If
                                        J_DEBIT = 0
                                        J_CREDIT = Round(CSMIOS_PMS_COST, 2)
                                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                    End If
                                End If
                            End If
                        End If
                        'PARTS ISSUED
                        If COMPANY_CODE = "HCC" Then
                            If CSMIOS_PARTS > 0 Or CSMIOS_MATERIALS > 0 Or CSMIOS_ACCESSORIES > 0 Then
                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    'SALES
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL")))
                                End If
                                J_GROSS = Round(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'COST OF SALES
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                End If
                                J_DEBIT = Round(CSMIOS_PARTS_COST + CSMIOS_MATERIALS_COST + CSMS_ACCCOST, 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'INVENTORY
                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(CSMIOS_PARTS_COST + CSMIOS_MATERIALS_COST + CSMS_ACCCOST, 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        Else
                            If CSMIOS_PARTS > 0 Then
                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    'SALES
                                    If COMPANY_CODE = "HBK" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "CASH"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "CASH")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HBK" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "CASH"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "CASH")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                                    End If
                                End If
                                J_GROSS = Round(NumericVal(CSMIOS_PARTS), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = Round(NumericVal(CSMIOS_PARTS), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_PARTS / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_PARTS) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                
                                If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And INSURANCE_DIRECT_EXPENSE_SPAREPARTS > 0 And CSMIOS_PARTICIPAT <> "" Then
                                Else
                                    'COST OF SALES
                                    ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        If COMPANY_CODE = "HBK" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "PARTS", "CASH"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "PARTS", "CASH")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                        End If
                                    Else
                                        If COMPANY_CODE = "HBK" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "PARTS", "CASH"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "PARTS", "CASH")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                        End If
                                    End If
                                    J_DEBIT = Round(CSMIOS_PARTS_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
    
                                    'INVENTORY
                                    ItemCnt = ItemCnt + 1
                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                        End If
                                    Else
                                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                        End If
                                    End If
                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_PARTS_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If
                            End If
                        End If
                        'MATERIAL ISSUED
                        If COMPANY_CODE = "HCC" Then
                        Else
                            If CSMIOS_MATERIALS > 0 Then
                                Dim SHOPSUP            As Double
                                Dim SHOPSUP_TAX        As Double
                                Dim SHOPSUP_NET        As Double

                                ItemCnt = ItemCnt + 1
                                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                                End If
                                J_GROSS = Round(NumericVal(CSMIOS_MATERIALS), 2)
                                SHOPSUP = Round(ReturnShopSuplies(CSMIOS_REP_OR), 2)
                                SHOPSUP_TAX = Round(NumericVal(Round((SHOPSUP / 1.12), 2) * 0.12), 2)
                                SHOPSUP_NET = Round(NumericVal(SHOPSUP) - Round(NumericVal(SHOPSUP_TAX), 2), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0
                                    J_NET = Round(NumericVal(CSMIOS_MATERIALS), 2) - SHOPSUP_NET
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0
                                        J_NET = (Round(NumericVal(CSMIOS_MATERIALS), 2)) - SHOPSUP_NET
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS / 1.12), 2) * 0.12), 2) - SHOPSUP_TAX
                                        J_NET = Round(NumericVal(CSMIOS_MATERIALS) - Round(NumericVal(J_TAX), 2), 2) - SHOPSUP_NET
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = (Round(NumericVal(J_NET), 2)) - SHOPSUP_TAX
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'SHOP SUPLIES/MISC
                                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                    If (ReturnShopSuplies(CSMIOS_REP_OR)) > 0 Then
                                        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                        J_ACCT_CODE = "'71-61000-20'"
                                        J_ACCT_NAME = N2Str2Null(Setacctname("'71-61000-20'"))
                                        J_GROSS = SHOPSUP: J_TAX = SHOPSUP_TAX: J_NET = SHOPSUP_NET:
                                        J_DEBIT = 0
                                        J_CREDIT = Round(NumericVal(SHOPSUP_NET), 2)
                                        ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                    End If
                                End If

                                If (COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI") And INSURANCE_DIRECT_EXPENSE_GOL > 0 And CSMIOS_PARTICIPAT <> "" Then
                                Else
                                    'COST OF SALES
                                    ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HBK" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS")))
                                    Else
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                        End If
                                    End If
                                    J_GROSS = 0: J_TAX = 0: J_NET = 0
                                    J_DEBIT = Round(CSMIOS_MATERIALS_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                    'INVENTORY
                                    ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                        End If
                                    Else
                                        If COMPANY_CODE = "HSB" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "VEHICLES"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "VEHICLES")))
                                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                        End If
                                    End If
                                    J_GROSS = 0: J_TAX = 0: J_NET = 0
                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_MATERIALS_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                End If
                            End If
                        End If
                        If COMPANY_CODE = "HCC" Then
                        Else
                            Call IMPORT_ACCESSORIES
                        End If

                        'DISCOUNTS
                        '==============================================================================================================================
                        'LABOR
                        If CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT > 0 Then
                            'DISCOUNT
                            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "LABOR"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "LABOR")))
                            End If
                            J_GROSS = NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)
                            If CSMIOS_INVOICE = "PDI RO" Then
                                J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT), 2)
                            Else
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT), 2)
                                Else
                                    J_TAX = Round(Round(((CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT) / 1.12), 2) * 0.12, 2)
                                    J_NET = Round(NumericVal((CSMIOS_LABOR_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT)) - Round(NumericVal(J_TAX), 2), 2)
                                End If
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If

                        If COMPANY_CODE = "HCC" Then
                            If CSMIOS_PARTS_DISCOUNT > 0 Or CSMIOS_MATERIALS_DISCOUNT > 0 Then
                                'DISCOUNT
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                                End If
                                J_GROSS = NumericVal(CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_PARTS_DISCOUNT / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = NumericVal(J_NET)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        Else
                            If CSMIOS_PARTS_DISCOUNT > 0 Then
                                'DISCOUNT
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                                End If
                                J_GROSS = NumericVal(CSMIOS_PARTS_DISCOUNT)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_PARTS_DISCOUNT / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = NumericVal(J_NET)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If

                            If CSMIOS_MATERIALS_DISCOUNT > 0 Then
                                'DISCOUNT
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HPI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                                Else
                                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                                        End If
                                    ElseIf COMPANY_CODE = "HMH" Then
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                                        End If
                                    ElseIf COMPANY_CODE = "HCI" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS")))
                                    Else
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "PARTS")))
                                        End If
                                    End If
                                End If

                                J_GROSS = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((CSMIOS_MATERIALS_DISCOUNT / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(CSMIOS_MATERIALS_DISCOUNT) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = Round(NumericVal(J_NET), 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT

                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        End If

                        '================================================================================================
                        'ACCESSORIES : Update By BTT - 06252008
                        '=================================================================================================
                        ' THIS IS TEMPORAY COMENTED DUE TO PROCEDURE TO LARGE
                        '                                    If COMPANY_CODE = "HBK" Then
                        '
                        '                                        If CSMIOS_ACCESSORIES > 0 Then
                        '                                        ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        '                                        If CSMIOS_TERM = "CSH" Then
                        '                                            J_ACCT_CODE = N2Str2Null("41-03003-30")
                        '                                            J_ACCT_NAME = N2Str2Null(Setacctname("41-03003-30"))
                        '                                        Else
                        '                                            J_ACCT_CODE = N2Str2Null("41-03003-30")
                        '                                            J_ACCT_NAME = N2Str2Null(Setacctname("41-03003-30"))
                        '                                        End If
                        '                                            J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                        '                                        If CSMIOS_INVOICE = "PDI RO" Then
                        '                                            J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                        '                                        Else
                        '                                            If CSMIOS_VAT_EXEMPT = True Then
                        '                                               J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                        '                                            Else
                        '                                               J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES / 1.12), 2) * 0.12), 2)
                        '                                               J_NET = Round(NumericVal(CSMIOS_ACCESSORIES) - Round(NumericVal(J_TAX), 2), 2)
                        '                                            End If
                        '                                        End If
                        '                                            J_DEBIT = 0
                        '                                            J_CREDIT = Round(NumericVal(J_NET), 2)
                        '                                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        '                                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        '
                        '                                         'COST OF SALES
                        '                                            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        '                                            If CSMIOS_TERM = "CSH" Then
                        '                                                J_ACCT_CODE = N2Str2Null("61-03003-30")
                        '                                                J_ACCT_NAME = N2Str2Null(Setacctname("61-03003-30"))
                        '                                            Else
                        '                                                 J_ACCT_CODE = N2Str2Null("61-03003-30")
                        '                                                J_ACCT_NAME = N2Str2Null(Setacctname("61-03003-30"))
                        '                                            End If
                        '                                            J_DEBIT = Round(CSMS_ACCCOST, 2)
                        '                                            J_CREDIT = 0
                        '                                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                        '                                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        '
                        '                                            'INVENTORY
                        '                                            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        '                                            If CSMIOS_TERM = "CSH" Then
                        '                                                J_ACCT_CODE = N2Str2Null("11-05004-00")
                        '                                                J_ACCT_NAME = N2Str2Null(Setacctname("11-05004-00"))
                        '                                            Else
                        '                                                 J_ACCT_CODE = N2Str2Null("11-05004-00")
                        '                                                 J_ACCT_NAME = N2Str2Null(Setacctname("11-05004-00"))
                        '                                            End If
                        '                                            J_DEBIT = 0
                        '                                            J_CREDIT = Round(CSMS_ACCCOST, 2)
                        '                                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        '                                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        '                                        End If
                        '                                    End If
                        'End Of Accessories
                        '=================================================================================================================================

                        'OUTPUT TAX
                        If J_INVOICEAMT > 0 Then
                            If CSMIOS_VAT_EXEMPT = False Then
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                    Else

                                        If COMPANY_CODE = "HBK" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                        End If
                                    End If
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                End If
                                J_DEBIT = 0
                                If COMPANY_CODE = "HCC" Then
                                    J_CREDIT = Round(NumericVal(J_INVOICEAMT - (AmountLessVAT(J_VOUCHERNO, J_JTYPE) - DiscountLessVAT(J_VOUCHERNO, J_JTYPE))), 2)
                                Else
                                    J_CREDIT = Round(NumericVal(Round((J_INVOICEAMT / 1.12), 2) * 0.12), 2)
                                End If
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                If Round(Round(ALL_CREDIT, 2) - Round(ALL_DEBIT + J_INVOICEAMT - SHOPSUP_NET, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                                If Round(Round(ALL_DEBIT + J_INVOICEAMT - SHOPSUP_NET, 2) - Round(ALL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                                J_TAX = 0: J_GROSS = 0: J_NET = 0
                                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            If CSMIOS_INVOICE = "PDI RO" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAccountCode(CSMIOS_INVOICE))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode(CSMIOS_INVOICE)))
                            Else
                                If COMPANY_CODE = "HMH" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                    End If
                                ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                    If CSMIOS_TERM = "CHG" Then
                                        'CLEARING
                                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                    Else
                                        'CLEARING
                                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                    End If
                                ElseIf COMPANY_CODE = "HAS" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                ElseIf COMPANY_CODE = "HCI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                                End If
                            End If
                            J_DEBIT = Round(J_INVOICEAMT, 2)
                            J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                        " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                        ", " & J_JNO & ", " & ALL_DEBIT & ", " & ALL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_hd", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", J_VOUCHERNO, J_JTYPE, J_JNO
                    End If

                    'END OF CUSTOMER PAID ENTRIES

                    '=========================================================================================================================================================
                    'ENTRY FOR INTERNAL
                    Call ImportPurelyInternalNew
                    Call IMPORT_INTERNAL

                    '                If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL + SALES_DIRECT_EXPENSE_LABOR + SALES_DIRECT_EXPENSE_SPAREPARTS + SALES_DIRECT_EXPENSE_GOL > 0 Then
                    '                    'LABOR
                    '
                    '                    INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0
                    '                    INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0
                    '
                    '                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    '                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DETPRC > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    '                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                    '                        rsINTERNAL_RO_DET.MoveFirst
                    '                        Do While Not rsINTERNAL_RO_DET.EOF
                    '                            If N2Str2Zero(rsINTERNAL_RO_DET!detprc) > 0 Then
                    '                                INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!detprc)
                    '                                INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!detcost)
                    '                                If COMPANY_CODE = "HPI" Then
                    '                                Else
                    '                                    ItemCnt = ItemCnt + 1
                    '                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                    '                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detprc)), 2)
                    '                                    J_TAX = 0
                    '                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detprc)), 2)
                    '                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    '                                    J_CREDIT = 0
                    '
                    '                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                                End If
                    '                            End If
                    '                            rsINTERNAL_RO_DET.MoveNext
                    '                        Loop
                    '
                    '                        ItemCnt = ItemCnt + 1
                    '                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                        If COMPANY_CODE = "HBK" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                    '                        Else
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                    '                        End If
                    '                        J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    '                        J_TAX = 0
                    '                        J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                    '                        J_DEBIT = 0
                    '                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    '
                    '                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '
                    '                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LABOR"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LABOR")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LABOR", "INTERNAL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LABOR", "INTERNAL")))
                    '                            End If
                    '                            J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                    '                            J_CREDIT = 0
                    '                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                            'INVENTORY
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If CSMIOS_TERM = "CSH" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    '                            End If
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                    '                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '                    End If
                    '
                    '                    'PARTS
                    '                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    '                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S' OR WCODE='W') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    '                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                    '                        rsINTERNAL_RO_DET.MoveFirst
                    '                        Do While Not rsINTERNAL_RO_DET.EOF
                    '                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                    '                                INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                    '                                INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))
                    '
                    '                                If COMPANY_CODE = "HPI" Then
                    '                                Else
                    '                                    ItemCnt = ItemCnt + 1
                    '                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                    '                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                    '                                    J_TAX = 0
                    '                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                    '                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    '                                    J_CREDIT = 0
                    '
                    '                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                                End If
                    '
                    '                            End If
                    '                            rsINTERNAL_RO_DET.MoveNext
                    '                        Loop
                    '
                    '                        If INTERNAL_PARTS_AMT > 0 Then
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                    '                            End If
                    '                            J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                    '                            If CSMIOS_VAT_EXEMPT = True Then
                    '                                J_TAX = 0
                    '                            Else
                    '                                J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                    '                            End If
                    '                            J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(NumericVal(J_NET), 2)
                    '                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "PARTS", "INTERNAL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "PARTS", "INTERNAL")))
                    '                            End If
                    '                            J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                    '                            J_CREDIT = 0
                    '                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                            'INVENTORY
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If CSMIOS_TERM = "CSH" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    '                            End If
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                    '                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                        End If
                    '
                    '                    End If
                    '
                    '                    Set rsINTERNAL_RO_DET = New ADODB.Recordset
                    '                    Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
                    '                    If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                    '                        rsINTERNAL_RO_DET.MoveFirst
                    '                        Do While Not rsINTERNAL_RO_DET.EOF
                    '                            If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                    '                                INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                    '                                INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))
                    '                                If COMPANY_CODE = "HPI" Then
                    '                                Else
                    '                                    ItemCnt = ItemCnt + 1
                    '                                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                    '                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                    '                                    J_TAX = 0
                    '                                    J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                    '                                    J_DEBIT = Round(NumericVal(J_NET), 2)
                    '                                    J_CREDIT = 0
                    '                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                                End If
                    '                            End If
                    '                            rsINTERNAL_RO_DET.MoveNext
                    '                        Loop
                    '
                    '                        If INTERNAL_MATERIALS_AMT > 0 Then
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                    '                            End If
                    '                            J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                    '                            If CSMIOS_VAT_EXEMPT = True Then
                    '                                J_TAX = 0
                    '                            Else
                    '                                J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                    '                            End If
                    '                            J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(NumericVal(J_NET), 2)
                    '                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                    '                            End If
                    '                            J_GROSS = 0
                    '                            J_TAX = 0
                    '                            J_NET = 0
                    '                            J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                    '                            J_CREDIT = 0
                    '                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                            'INVENTORY
                    '                            ItemCnt = ItemCnt + 1
                    '                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    '                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    '                            J_GROSS = 0: J_TAX = 0: J_NET = 0
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                    '                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                        End If
                    '                    End If
                    '                    'END OF NOT PURELY INTERNAL RO
                    '                    ':-)=================================
                    '                End If
                    '=========================================================================================================================================================

                    '=========================================================================================================================================================
                    'ENTRY FOR WARRANTY
                    '                If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                    '
                    '                    Set rsJournal_HDDup = New ADODB.Recordset
                    '                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    '                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    '
                    '                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetSJVoucherNo()) + 1, "000000"))
                    '                    WARRANTY_ItemCnt = 0
                    '                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                    TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                    '                    'LABOR
                    '                    If WARRANTY_DIRECT_EXPENSE_LABOR > 0 Then
                    '                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                        If COMPANY_CODE = "HGC" Then
                    '                            If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                    '                            End If
                    '                        Else
                    '                            If COMPANY_CODE = "HAI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                    '                            End If
                    '                        End If
                    '                        J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR), 2)
                    '                        If CSMIOS_VAT_EXEMPT = True Then
                    '                            J_TAX = 0
                    '                        Else
                    '                            J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                    '                        End If
                    '                        J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) - Round(NumericVal(J_TAX), 2), 2)
                    '                        J_DEBIT = 0
                    '                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    '                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                        If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            If CSMIOS_TERM = "CSH" Then
                    '                                If COMPANY_CODE = "HGC" Then
                    '                                    If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                    '                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                    '                                    Else
                    '                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                    '                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    '                                    End If
                    '                                Else
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    '                                End If
                    '                            Else
                    '                                If COMPANY_CODE = "HGC" Then
                    '                                    If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                    '                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                    '                                    Else
                    '                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                    '                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    '                                    End If
                    '                                Else
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    '                                End If
                    '                            End If
                    '                            J_DEBIT = Round(WARRANTY_DIRECT_EXPENSE_LABOR_COST, 2)
                    '                            J_CREDIT = 0
                    '                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                            'INVENTORY
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            If CSMIOS_TERM = "CSH" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                    '                            End If
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(WARRANTY_DIRECT_EXPENSE_LABOR_COST, 2)
                    '                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '                    End If
                    '
                    '                    'PARTS
                    '                    If WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY")))
                    '                            ElseIf COMPANY_CODE = "HAI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "WARRANTY"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "WARRANTY")))
                    '                            End If
                    '                            J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS), 2)
                    '                            If CSMIOS_VAT_EXEMPT = True Then
                    '                                J_TAX = 0
                    '                            Else
                    '                                J_TAX = Round(Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2), 2)
                    '                            End If
                    '                            J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) - Round(NumericVal(J_TAX), 2), 2)
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(NumericVal(J_NET), 2)
                    '                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                    '                            ElseIf COMPANY_CODE = "HAI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                    '                            Else
                    '                                If CSMIOS_TERM = "CSH" Then
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "WARRANTY"))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "WARRANTY")))
                    '                                Else
                    '                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "WARRANTY"))
                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "WARRANTY")))
                    '                                End If
                    '                            End If
                    '                            J_DEBIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                    '                            J_CREDIT = 0
                    '                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '                        'INVENTORY
                    '                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                        If CSMIOS_TERM = "CSH" Then
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    '                            ElseIf COMPANY_CODE = "HPI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                    '                            ElseIf COMPANY_CODE = "HAI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    '                            End If
                    '                        Else
                    '                            'COMMENTED DUE PROCEDURE TO LARGE
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    '                            End If
                    '                        End If
                    '                        J_DEBIT = 0
                    '                        J_CREDIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                    '                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '
                    '                    End If
                    '
                    '
                    '
                    '                    'Materials
                    '                    If WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                    '                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                        If COMPANY_CODE = "HBK" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    '                        ElseIf COMPANY_CODE = "HGC" Then
                    '                            J_ACCT_CODE = "'41-02013-20'"
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname("'41-02013-20'"))
                    '                        ElseIf COMPANY_CODE = "HSB" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                    '                        ElseIf COMPANY_CODE = "HAI" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    '                        Else
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "WARRANTY"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "WARRANTY")))
                    '                        End If
                    '
                    '                        J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL), 2)
                    '                        If CSMIOS_VAT_EXEMPT = True Then
                    '                            J_TAX = 0
                    '                        Else
                    '                            J_TAX = Round(Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333, 2), 2)
                    '                        End If
                    '
                    '                        J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) - Round(NumericVal(J_TAX), 2), 2)
                    '                        J_DEBIT = 0
                    '                        J_CREDIT = Round(NumericVal(J_NET), 2)
                    '                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        If COMPANY_CODE = "HPI" Then
                    '                        Else
                    '                            'COST OF SALES
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            If COMPANY_CODE = "HBK" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS")))
                    '                            ElseIf COMPANY_CODE = "HGC" Then
                    '                                J_ACCT_CODE = "'61-02011-20'"
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname("'61-02011-20'"))
                    '                            ElseIf COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAI" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "WARRANTY"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "WARRANTY")))
                    '                            End If
                    '                            J_GROSS = 0
                    '                            J_TAX = 0
                    '                            J_NET = 0
                    '                            J_DEBIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                    '                            J_CREDIT = 0
                    '                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '                        End If
                    '
                    '                        'INVENTORY
                    '                        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                        If COMPANY_CODE = "HAI" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                    '                        Else
                    '                            If CSMIOS_TERM = "CSH" Then
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    '                            Else
                    '                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                    '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    '                            End If
                    '                        End If
                    '                        J_GROSS = 0: J_TAX = 0: J_NET = 0
                    '                        J_DEBIT = 0
                    '                        J_CREDIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                    '                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                    End If
                    '
                    '                    'OUTPUT TAX
                    '                    If CSMIOS_VAT_EXEMPT = False Then
                    '                        If NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) > 0 Then
                    '                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                            'Commented Due to procedure to Large
                    '                            'If COMPANY_CODE = "HBK" Then
                    '                            '   J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                    '                            '   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                    '                            'Else
                    '                            J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                    '                            'End If
                    '                            J_DEBIT = 0
                    '                            J_CREDIT = Round(NumericVal(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333), 2)
                    '                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '                            If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT + WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                    '                            If Round(Round(TOTAL_DEBIT + WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                    '                            J_TAX = 0: J_GROSS = 0: J_NET = 0
                    '                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                        End If
                    '                    End If
                    '
                    '                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    '                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    '                    If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                        If COMPANY_CODE = "HAS" Then
                    '                            J_ACCT_CODE = N2Str2Null("11-02000-00")
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname("11-02000-00"))
                    '                        ElseIf COMPANY_CODE = "HPI" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                    '                        Else
                    '                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                    '                        End If
                    '                    Else
                    '                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
                    '                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
                    '                    End If
                    '                    J_DEBIT = NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL)
                    '                    J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
                    '                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    '                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    '
                    '                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    '                    'To return the Selling Dealer
                    '                    If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                        CSMIOS_NIYM = Null2String(SetSellingDealerName(ReturnCodeSellingDealer(ReturnPlateNo(CSMIOS_REP_OR))))
                    '                        CSMIOS_ACCT_NO = N2Str2Null(ReturnCode(CSMIOS_NIYM))
                    '                        J_CUSTOMERCODE = N2Str2Null(ReturnCode(CSMIOS_NIYM))
                    '                    Else
                    '                        If COMPANY_CODE = "HPI" Then
                    '                            CSMIOS_ACCT_NO = Null2String("H00416")
                    '                            CSMIOS_NIYM = Null2String(SetCustomerName("H00416"))
                    '                            J_CUSTOMERCODE = N2Str2Null("H00416")
                    '                        Else
                    '                            CSMIOS_ACCT_NO = Null2String("H00001")
                    '                            CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    '                            J_CUSTOMERCODE = N2Str2Null("H00001")
                    '                        End If
                    '                    End If
                    '                    WARRANTY_J_AMOUNTTOPAY = 0
                    '                    WARRANTY_J_INVOICEAMT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL), 2)
                    '                    WARRANTY_J_BALANCE = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL), 2)
                    '                    WARRANTY_J_AMOUNTPAID = 0
                    '                    SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                         '                                    " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                         '                                    " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                         '                                    ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                    '                    gconDMIS.Execute SQL_STATEMENT
                    '                    TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_hd", "X", J_JTYPE, "Jtype"))
                    '                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO
                    '                    ':-)=================================
                    '                End If
                    '=========================================================================================================================================================
                    Call Import_Warranty
                    '=========================================================================================================================================================
                    'ENTRY FOR INSURANCE

                    If INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                        Set rsJournal_HDDup = New ADODB.Recordset
                        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"

                        WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetSJVoucherNo()) + 1, "000000"))
                        WARRANTY_ItemCnt = 0: WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("AR"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("AR")))
                        ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HGH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("INSURANCE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("INSURANCE")))
                        End If
                        J_DEBIT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                        J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        If INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HBK" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("INSURANCE", "AR"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("INSURANCE", "AR")))
                            ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "SUBLET"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "SUBLET")))
                            End If
                            J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR), 2)
                            If CSMIOS_VAT_EXEMPT = True Then
                                J_TAX = 0
                            Else
                                J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
                            End If
                            J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR) - Round(NumericVal(J_TAX), 2), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then

                            Else
                                If CSMIOS_LABOR_COST > 0 Then
                                    'COST OF SALES
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET")))
                                    End If
                                    J_DEBIT = Round(CSMIOS_LABOR_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                    'INVENTORY
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                    End If
                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_LABOR_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                End If

                                If CSMIOS_SUBLET_COST > 0 Then
                                    'COST OF SALES
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET")))
                                    End If
                                    J_DEBIT = Round(CSMIOS_SUBLET_COST, 2)
                                    J_CREDIT = 0
                                    ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                    'INVENTORY
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    If COMPANY_CODE = "HCI" Then
                                        Set rsSublet = New ADODB.Recordset
                                        rsSublet.Open "SELECT DETCDE FROM CSMS_RO_DET WHERE REP_OR='" & CSMIOS_REP_OR & "' AND LEFT(DETCDE,2) ='SR'", gconDMIS, adOpenForwardOnly
                                        If Not rsSublet.EOF And Not rsSublet.BOF Then
                                            DetCodeLen = Len(Null2String(rsSublet!DETCDE))
                                            SubletType = UCase(Mid(Null2String(rsSublet!DETCDE), 3, DetCodeLen))
                                        End If
                                        Set rsSublet = Nothing
                                        If SubletType = "LABOR" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "LABOR"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "LABOR")))
                                        ElseIf SubletType = "PARTS" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "PARTS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "PARTS")))
                                        ElseIf SubletType = "MATERIALS" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnExpense("SERVICE", "MATERIALS"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("SERVICE", "MATERIALS")))
                                        End If
                                    Else
                                        If CSMIOS_TERM = "CSH" Then
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                        Else
                                            J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET"))
                                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET")))
                                        End If
                                    End If

                                    J_DEBIT = 0
                                    J_CREDIT = Round(CSMIOS_SUBLET_COST, 2)
                                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                End If
                            End If
                        End If
                        If INSURANCE_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If COMPANY_CODE = "HBK" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "AR", "INSURANCE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "AR", "INSURANCE")))
                            ElseIf COMPANY_CODE = "HPI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "INSURANCE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "INSURANCE")))
                            ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                            End If


                            J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS), 2)
                            If CSMIOS_VAT_EXEMPT = True Then
                                J_TAX = 0
                            Else
                                J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2)
                            End If
                            J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_SPAREPARTS) - Round(NumericVal(J_TAX), 2), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)


                            If CSMIOS_PARTS_COST > 0 Then
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                                    End If
                                End If
                                J_DEBIT = Round(CSMIOS_PARTS_COST, 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'INVENTORY
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(CSMIOS_PARTS_COST, 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        End If
                        If INSURANCE_DIRECT_EXPENSE_GOL > 0 Then
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                            J_GROSS = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL), 2)
                            If CSMIOS_VAT_EXEMPT = True Then
                                J_TAX = 0
                            Else
                                J_TAX = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) / 9.3333, 2)
                            End If
                            J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_GOL) - Round(NumericVal(J_TAX), 2), 2)
                            J_DEBIT = 0
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            If CSMIOS_MATERIALS_COST > 0 Then
                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                'COMMENTED DUE PROCEDURE TO LARGE : TEMP
                                'If COMPANY_CODE = "HBK" Then
                                '    J_ACCT_CODE = N2Str2Null(ReturnCostOfSales("SERVICE", "LUBRICANTS"))
                                '    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostOfSales("SERVICE", "LUBRICANTS")))
                                'Else
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                                End If
                                'End If
                                J_GROSS = 0: J_TAX = 0: J_NET = 0
                                J_DEBIT = Round(CSMIOS_MATERIALS_COST, 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                                'INVENTORY
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                                Else
                                    If CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                    End If
                                End If
                                J_GROSS = 0: J_TAX = 0: J_NET = 0
                                J_DEBIT = 0
                                J_CREDIT = Round(CSMIOS_MATERIALS_COST, 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        End If

                        If COMPANY_CODE = "HMH" Then
                            If INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED")))
                                End If
                                J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                                If CSMIOS_INVOICE = "PDI RO" Then
                                    J_TAX = 0: J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                                Else
                                    If CSMIOS_VAT_EXEMPT = True Then
                                        J_TAX = 0: J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                                    Else
                                        J_TAX = Round(NumericVal(Round((INSURANCE_DIRECT_EXPENSE_ACCESSORIES / 1.12), 2) * 0.12), 2)
                                        J_NET = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_ACCESSORIES) - Round(NumericVal(J_TAX), 2), 2)
                                    End If
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(J_NET), 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'COST OF SALES
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES", "INSTALLED"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES", "INSTALLED")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES", "INSTALLED"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES", "INSTALLED")))
                                End If
                                J_DEBIT = Round(CSMS_ACCCOST, 2)
                                J_CREDIT = 0
                                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                                'INVENTORY
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                                End If
                                J_DEBIT = 0
                                J_CREDIT = Round(CSMS_ACCCOST, 2)
                                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                            End If
                        End If

                        If CSMIOS_PARTS_DISCOUNT > 0 Then
                            'DISCOUNT
                            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                            End If
                            J_GROSS = NumericVal(CSMIOS_PARTS_DISCOUNT)
                            If CSMIOS_INVOICE = "PDI RO" Then
                                J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                            Else
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT), 2)
                                Else
                                    J_TAX = Round(NumericVal(Round((CSMIOS_PARTS_DISCOUNT / 1.12), 2) * 0.12), 2)
                                    J_NET = Round(NumericVal(CSMIOS_PARTS_DISCOUNT) - Round(NumericVal(J_TAX), 2), 2)
                                End If
                            End If
                            J_DEBIT = NumericVal(J_NET)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT
                            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                        End If

                        'OUTPUT TAX
                        If CSMIOS_VAT_EXEMPT = False Then
                            If NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + CSMIOS_PARTS_DISCOUNT + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) > 0 Then
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HBK" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                                End If

                                J_DEBIT = 0
                                J_CREDIT = Round(NumericVal(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES) / 9.3333), 2)
                                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                                If Round(Round(TOTAL_CREDIT, 2) - Round(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                                If Round(Round(TOTAL_CREDIT, 2) - Round(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2), 2) = -0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                                If Round(Round(TOTAL_DEBIT + INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                                If Round(Round(TOTAL_DEBIT + INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2) - Round(TOTAL_CREDIT, 2), 2) = -0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                                J_TAX = 0
                                J_GROSS = 0
                                J_NET = 0
                                J_STATUS = "'N'"
                                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                            End If
                        End If
                        'END CREDIT ENTRY
                        CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                        CSMIOS_ACCT_NO = Null2String("H00001")
                        CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                        CSMIOS_NIYM = Null2String(SetCustomerName(CSMIOS_PARTICIPAT))
                        WARRANTY_J_AMOUNTTOPAY = 0
                        WARRANTY_J_INVOICEAMT = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                        WARRANTY_J_BALANCE = Round(NumericVal(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES), 2)
                        WARRANTY_J_AMOUNTPAID = 0

                        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                        " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & N2Str2Null(CSMIOS_PARTICIPAT) & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                                        ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        'AUDIT
                        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_hd", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO
                    End If
                    Grid3.Cell(GridImport, 1).Text = 1
                End If
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
DuplicateInvoiceNo:
    Next

    ImportCSMSSales = True
    Exit Function
ErrorCode:
    ImportCSMSSales = False
End Function

Sub IMPORT_INTERNAL()
    If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL + SALES_DIRECT_EXPENSE_LABOR + SALES_DIRECT_EXPENSE_SPAREPARTS + SALES_DIRECT_EXPENSE_GOL > 0 And CSMIOS_RO_AMOUNT > 0 Then
        'LABOR

        INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0
        INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0

        Set rsINTERNAL_RO_DET = New ADODB.Recordset
        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DETPRC > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
            rsINTERNAL_RO_DET.MoveFirst
            Do While Not rsINTERNAL_RO_DET.EOF
                If N2Str2Zero(rsINTERNAL_RO_DET!DETPRC) > 0 Then
                    INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DETPRC)
                    INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!detcost)
                    If COMPANY_CODE = "HPI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                        J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETPRC)), 2)
                        J_TAX = 0
                        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost)), 2)
                        Else
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DETPRC)), 2)
                        End If
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                        J_CREDIT = 0

                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                rsINTERNAL_RO_DET.MoveNext
            Loop

            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
            Else
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "INTERNAL")))
                End If
                J_GROSS = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                J_TAX = 0
                J_NET = Round(NumericVal(INTERNAL_LABOR_AMT), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)

                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
            End If

            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
            Else
                'COST OF SALES
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "INTERNAL")))
                End If
                J_DEBIT = Round(INTERNAL_LABOR_COST, 2)
                J_CREDIT = 0
                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
            End If
            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
            Else
                'INVENTORY
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                If CSMIOS_TERM = "CSH" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                End If
                J_DEBIT = 0
                J_CREDIT = Round(INTERNAL_LABOR_COST, 2)
                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
            End If
        End If

        'PARTS
        If COMPANY_CODE = "HCC" Then
        Else
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                        INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))

                        If COMPANY_CODE = "HPI" Then
                        Else
                            ItemCnt = ItemCnt + 1
                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            If COMPANY_CODE = "HMH" Then
                                J_NET = Round(NumericVal(N2Str2Zero(INTERNAL_PARTS_COST)), 2)
                            Else
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0

                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If

                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop

                If INTERNAL_PARTS_AMT > 0 Then
                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                        End If
                        J_GROSS = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_PARTS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If

                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                    Else
                        'COST OF SALES
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS")))
                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                        End If
                        J_DEBIT = Round(INTERNAL_PARTS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If

                    'INVENTORY
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                    J_DEBIT = 0
                    If COMPANY_CODE = "HMH" Then
                        J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                    Else
                        J_CREDIT = Round(INTERNAL_PARTS_COST, 2)
                    End If
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                End If
            End If
        End If
        If COMPANY_CODE = "HCC" Then
        Else
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                        INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            ItemCnt = ItemCnt + 1
                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            J_TAX = 0
                            If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                                J_NET = Round(NumericVal(N2Str2Zero(INTERNAL_MATERIALS_COST)), 2)
                            Else
                                J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            ALL_CREDIT = ALL_CREDIT + J_CREDIT
                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop

                If INTERNAL_MATERIALS_AMT > 0 Then
                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                        End If
                        J_GROSS = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        If CSMIOS_VAT_EXEMPT = True Then
                            J_TAX = 0
                        Else
                            J_TAX = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        End If
                        J_NET = Round(NumericVal(INTERNAL_MATERIALS_AMT), 2)
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(J_NET), 2)
                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If

                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Then
                    Else
                        'COST OF SALES
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS")))
                        ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "INTERNAL")))
                        End If
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = Round(INTERNAL_MATERIALS_COST, 2)
                        J_CREDIT = 0
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If

                    'INVENTORY
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    End If
                    J_GROSS = 0: J_TAX = 0: J_NET = 0
                    J_DEBIT = 0
                    J_CREDIT = Round(INTERNAL_MATERIALS_COST, 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                End If
            End If
        End If
        '---al for HCC combined entry
        If COMPANY_CODE = "HCC" Then
            Set rsINTERNAL_RO_DET = New ADODB.Recordset
            Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL IN ('2','3','4') AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
            If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
                rsINTERNAL_RO_DET.MoveFirst
                Do While Not rsINTERNAL_RO_DET.EOF
                    If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                        INTERNAL_ACCESSORIES_AMT = INTERNAL_ACCESSORIES_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                        INTERNAL_ACCESSORIES_COST = INTERNAL_ACCESSORIES_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!DETVOL))
                        If COMPANY_CODE = "HPI" Then
                        Else
                            ItemCnt = ItemCnt + 1
                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code)))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!Code))))
                            J_GROSS = Round(NumericVal(INTERNAL_ACCESSORIES_AMT), 2)
                            J_TAX = 0
                            J_NET = Round(NumericVal(N2Str2Zero(INTERNAL_ACCESSORIES_AMT)), 2)
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                        End If
                    End If
                    rsINTERNAL_RO_DET.MoveNext
                Loop
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                J_GROSS = Round(NumericVal(INTERNAL_ACCESSORIES_AMT), 2)
                If CSMIOS_VAT_EXEMPT = True Then
                    J_TAX = 0
                Else
                    J_TAX = Round(NumericVal(INTERNAL_ACCESSORIES_AMT), 2)
                End If
                J_NET = Round(NumericVal(INTERNAL_ACCESSORIES_AMT), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                'COST OF SALES
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                J_GROSS = 0
                J_TAX = 0
                J_NET = 0
                J_DEBIT = Round(INTERNAL_ACCESSORIES_COST, 2)
                J_CREDIT = 0
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                'INVENTORY
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                J_GROSS = 0: J_TAX = 0: J_NET = 0
                J_DEBIT = 0
                J_CREDIT = Round(INTERNAL_ACCESSORIES_COST, 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            End If
        End If


        'OUT PUT TAX
        If COMPANY_CODE = "HPI" Then
            If INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST > 0 Then
                If CSMIOS_VAT_EXEMPT = False Then
                    ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        If CSMIOS_TERM = "CHG" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                        Else

                            If COMPANY_CODE = "HBK" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                            End If
                        End If
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                    End If
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(Round(((INTERNAL_PARTS_COST + INTERNAL_MATERIALS_COST) / 1.12), 2) * 0.12), 2)
                    ALL_CREDIT = ALL_CREDIT + J_CREDIT
                    J_TAX = 0: J_GROSS = 0: J_NET = 0
                    Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
            End If
        End If

        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DGI" Then
            Dim rsCheckHD                              As ADODB.Recordset
            Set rsCheckHD = New ADODB.Recordset
            rsCheckHD.Open "Select VoucherNo, JType from AMIS_JOURNAL_HD WHERE JTYPE=" & J_JTYPE & " AND VOUCHERNO =" & J_VOUCHERNO & " AND STATUS <> 'P'", gconDMIS, adOpenKeyset
            If Not rsCheckHD.EOF And Not rsCheckHD.BOF Then
            Else
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT
                'AUDIT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_hd", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", J_VOUCHERNO, J_JTYPE, J_JNO
            End If
        End If
        'END OF NOT PURELY INTERNAL RO
        ':-)=================================
    End If
End Sub


Sub InsertToJournalDet(vJ_JDATE As Variant, vJ_VOUCHERNO As Variant, vJ_JTYPE As Variant, vJ_JNO As Variant, vJ_JITEMNO As Variant, vJ_ACCT_CODE As Variant, vJ_ACCT_NAME As Variant, vJ_DEBIT As Variant, vJ_CREDIT As Variant, vJ_TAX As Variant, vJ_GROSS As Variant, vJ_NET As Variant, vJ_STATUS As Variant)
    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                     " values (" & vJ_JDATE & ", " & vJ_VOUCHERNO & ", " & vJ_JTYPE & ", " & vJ_JNO & ", " & vJ_JITEMNO & ", " & vJ_ACCT_CODE & ", " & vJ_ACCT_NAME & ", " & vJ_DEBIT & ", " & vJ_CREDIT & ", " & vJ_TAX & "," & vJ_GROSS & "," & vJ_NET & ", " & vJ_STATUS & ")"

    TransactionID = (FindTransactionID(N2Str2Null(vJ_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(vJ_VOUCHERNO), J_JTYPE, N2Str2Zero(vJ_JNO)
End Sub

Sub INITIALIZE_ACCESSORIESCOST(XXX As String)
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
        Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIES Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            CSMIOS_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
            CSMIOS_ACCESSORIES_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!DISCOUNT), 2)
        Else
            CSMIOS_ACCESSORIES = 0: CSMIOS_ACCESSORIES_DISCOUNT = 0
        End If

        Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
        Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIESWarranty Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
        End If

        Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
        Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            WARRANTY_CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACC_COST), 2)
        Else
            WARRANTY_CSMIOS_ACCESSORIES_COST = 0:
        End If
    Else
        'Do nothing
    End If
End Sub

Sub INITIALIZE_SALES_AND_DISCOUNT(XXX As String)
    CSMIOS_LABOR = 0: CSMIOS_LABOR_DISCOUNT = 0: CSMIOS_LABOR_COST = 0
    CSMIOS_PARTS = 0: CSMIOS_PARTS_DISCOUNT = 0
    CSMIOS_MATERIALS = 0: CSMIOS_MATERIALS_DISCOUNT = 0
    CSMIOS_SUBLET = 0: CSMIOS_SUBLET_DISCOUNT = 0: CSMIOS_SUBLET_COST = 0
    CSMIOS_TINSPAINT = 0: CSMIOS_TINSPAINT_DISCOUNT = 0: CSMIOS_TINSPAINT_COST = 0
    CSMIOS_PMS = 0: CSMIOS_PMS_DISCOUNT = 0: CSMIOS_PMS_COST = 0

    CSMIOS_PARTS_COST = 0: CSMIOS_MATERIALS_COST = 0: CSMIOS_ACCESSORIES_COST = 0:

    WARRANTY_DIRECT_EXPENSE_LABOR = 0: WARRANTY_DIRECT_EXPENSE_SPAREPARTS = 0: WARRANTY_DIRECT_EXPENSE_GOL = 0: WARRANTY_CSMIOS_PARTS_COST = 0: WARRANTY_CSMIOS_MATERIALS_COST = 0: WARRANTY_CSMIOS_ACCESSORIES_COST = 0:

    COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0
    SALES_DIRECT_EXPENSE_LABOR = 0: SALES_DIRECT_EXPENSE_SPAREPARTS = 0: SALES_DIRECT_EXPENSE_GOL = 0
    COMPANY_DIRECT_EXPENSE_LABOR = 0: COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0: COMPANY_DIRECT_EXPENSE_GOL = 0


    'SALES AND DISCOUNTS
    '====================================================================================================================================================================================
    Set rsCSMIOS_LABOR = New ADODB.Recordset
    Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DETCOST),2) AS COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABOR Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
        CSMIOS_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
        CSMIOS_LABOR_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_LABOR!DISCOUNT), 2)
        CSMIOS_LABOR_COST = Round(N2Str2Zero(rsCSMIOS_LABOR!Cost), 2)
    End If

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTS Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
        CSMIOS_PARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)
        CSMIOS_PARTS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_PARTS!DISCOUNT), 2)
    End If

    Set rsCSMIOS_MATERIALS = New ADODB.Recordset
    Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALS Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
        CSMIOS_MATERIALS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
        CSMIOS_MATERIALS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_MATERIALS!DISCOUNT), 2)
    End If
    Set rsCSMIOS_SUBLET = New ADODB.Recordset
    Set rsCSMIOS_SUBLET = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS SUBLET,ROUND(sum(DETCOST),2) AS COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_SUBLET Where WCODE IS NULL AND REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_SUBLET.EOF And Not rsCSMIOS_SUBLET.BOF Then
        CSMIOS_SUBLET = Round(N2Str2Zero(rsCSMIOS_SUBLET!SUBLET), 2)
        CSMIOS_SUBLET_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_SUBLET!DISCOUNT), 2)
        CSMIOS_SUBLET_COST = Round(N2Str2Zero(rsCSMIOS_SUBLET!Cost), 2)
    End If

    Set rsCSMIOS_TINSMITH = New ADODB.Recordset
    Set rsCSMIOS_TINSMITH = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS TINSPAINT,ROUND(sum(DETCOST),2) AS COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_TinsPaint Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_TINSMITH.EOF And Not rsCSMIOS_TINSMITH.BOF Then
        CSMIOS_TINSPAINT = Round(N2Str2Zero(rsCSMIOS_TINSMITH!TINSPAINT), 2)
        CSMIOS_TINSPAINT_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_TINSMITH!DISCOUNT), 2)
        CSMIOS_TINSPAINT_COST = Round(N2Str2Zero(rsCSMIOS_TINSMITH!Cost), 2)
    End If

    Set rsCSMIOS_PMS = New ADODB.Recordset
    Set rsCSMIOS_PMS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PMS,ROUND(sum(DETCOST),2) AS COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PMS Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_PMS.EOF And Not rsCSMIOS_PMS.BOF Then
        CSMIOS_PMS = Round(N2Str2Zero(rsCSMIOS_PMS!PMS), 2)
        CSMIOS_PMS_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_PMS!DISCOUNT), 2)
        CSMIOS_PMS_COST = Round(N2Str2Zero(rsCSMIOS_PMS!Cost), 2)
    End If
    TOTAL_DISCOUNT_AMOUNT = Round(CSMIOS_LABOR_DISCOUNT + CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT, 2)

    '====================================================================================================================================================================================
End Sub

Sub INITIALIZE_COST_VARIABLE(XXX As String)
'COST
'====================================================================================================================================================================================
    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST ),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST ),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!MAT_COST), 2)
    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
            CSMS_ACCCOST = Round(N2Str2Zero(rsCSMIOS_PARTS!ACC_COST), 2)
        End If
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!ACC_COST), 2)
    '====================================================================================================================================================================================
End Sub

Sub INITIALIZE_WARRANTY(XXX As String)
'WARRANTY
'====================================================================================================================================================================================

    WARRANTY_DIRECT_EXPENSE_LABOR = 0: WARRANTY_DIRECT_EXPENSE_SPAREPARTS = 0: WARRANTY_DIRECT_EXPENSE_GOL = 0: WARRANTY_DIRECT_EXPENSE_LABOR_COST = 0

    Set rsCSMIOS_LABOR = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORWarranty Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
            WARRANTY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
            WARRANTY_DIRECT_EXPENSE_LABOR_COST = 0
        End If
    Else
        Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DETCOST),2) AS LABOR_COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORWarranty Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then
            WARRANTY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)
            WARRANTY_DIRECT_EXPENSE_LABOR_COST = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR_COST), 2)
        End If
    End If

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSWarranty Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then WARRANTY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)

    Set rsCSMIOS_MATERIALS = New ADODB.Recordset
    Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSWarranty Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then WARRANTY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST),2) AS PARTS_COST from CSMIOS_vw_WARRANTY_PARTSCOST Where REP_OR = " & N2Str2Null(CSMIOS_REP_OR))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_WARRANTY_PARTSCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then WARRANTY_CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST ),2) AS MAT_COST from CSMIOS_vw_WARRANTY_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_WARRANTY_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then WARRANTY_CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!MAT_COST), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST),2) AS ACC_COST from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then WARRANTY_CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!ACC_COST), 2)
    '====================================================================================================================================================================================
End Sub

Sub INITIALIZE_INTERNAL_COMPANY(XXX As String)
'INTERNAL - COMPANY
'====================================================================================================================================================================================

    Set rsCSMIOS_LABOR = New ADODB.Recordset
    Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORCompany Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then COMPANY_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSCompany Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then COMPANY_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)

    Set rsCSMIOS_MATERIALS = New ADODB.Recordset
    Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSCompany Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then COMPANY_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
    '====================================================================================================================================================================================
End Sub

Sub INITIALIZE_INTERNAL_SALES_DEPARTMENT(XXX As String)
'INTERNAL - SALES DEPARTMENT
'====================================================================================================================================================================================
    Set rsCSMIOS_LABOR = New ADODB.Recordset
    Set rsCSMIOS_LABOR = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS LABOR,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_LABORSales Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_LABOR.EOF And Not rsCSMIOS_LABOR.BOF Then SALES_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_LABOR!LABOR), 2)

    Set rsCSMIOS_PARTS = New ADODB.Recordset
    Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS PARTS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_PARTSSales Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then SALES_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS), 2)

    Set rsCSMIOS_MATERIALS = New ADODB.Recordset
    Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS MATERIALS,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_MATERIALSSales Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then SALES_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!MATERIALS), 2)
    '====================================================================================================================================================================================
End Sub

Sub INITIALIZE_INSURANCE(XXX As String)
'INSURANCE
'====================================================================================================================================================================================
    INSURANCE_DIRECT_EXPENSE_LABOR = 0: INSURANCE_DIRECT_EXPENSE_SPAREPARTS = 0: INSURANCE_DIRECT_EXPENSE_GOL = 0: INSURANCE_DIRECT_EXPENSE_ACCESSORIES = 0
    TOTAL_INSURANCE_AMOUNT = 0
    Set rsCSMIOS_MATERIALS = New ADODB.Recordset
    Set rsCSMIOS_MATERIALS = gconDMIS.Execute("Select * from CSMIOS_INSURANCE Where REP_OR = " & N2Str2Null(XXX))
    If Not rsCSMIOS_MATERIALS.EOF And Not rsCSMIOS_MATERIALS.BOF Then
        INSURANCE_DIRECT_EXPENSE_LABOR = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR), 2)
        INSURANCE_DIRECT_EXPENSE_SPAREPARTS = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSPARTS), 2)
        INSURANCE_DIRECT_EXPENSE_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSACCESSORIES), 2)
        INSURANCE_DIRECT_EXPENSE_GOL = Round(N2Str2Zero(rsCSMIOS_MATERIALS!INSMATERIALS), 2)

        If (CSMIOS_LABOR + CSMIOS_SUBLET + CSMIOS_TINSPAINT + CSMIOS_PMS) - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
            If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                CSMIOS_LABOR = CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR
                GoTo PAKSIW
            Else
                If CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                    CSMIOS_LABOR = CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR
                    GoTo PAKSIW
                Else
                    INSURANCE_DIRECT_EXPENSE_LABOR = INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR
                    CSMIOS_LABOR = 0
                End If
            End If
            If CSMIOS_SUBLET > 0 And CSMIOS_LABOR - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                CSMIOS_SUBLET = CSMIOS_SUBLET - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR)
                GoTo PAKSIW
            Else
                If CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                    CSMIOS_SUBLET = CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR
                    GoTo PAKSIW
                Else
                    INSURANCE_DIRECT_EXPENSE_LABOR = INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_SUBLET
                    CSMIOS_SUBLET = 0
                End If
            End If
            If CSMIOS_TINSPAINT > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                CSMIOS_TINSPAINT = CSMIOS_TINSPAINT - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET)
                GoTo PAKSIW
            Else
                If CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                    CSMIOS_TINSPAINT = CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR
                    GoTo PAKSIW
                Else
                    INSURANCE_DIRECT_EXPENSE_LABOR = INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_TINSPAINT
                    CSMIOS_TINSPAINT = 0
                End If
            End If
            If CSMIOS_PMS > 0 And CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT - INSURANCE_DIRECT_EXPENSE_LABOR >= 0 Then
                CSMIOS_PMS = CSMIOS_PMS - (INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_LABOR - CSMIOS_SUBLET - CSMIOS_TINSPAINT)
                GoTo PAKSIW
            Else
                If CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR > 0 Then
                    CSMIOS_PMS = CSMIOS_PMS - INSURANCE_DIRECT_EXPENSE_LABOR
                Else
                    INSURANCE_DIRECT_EXPENSE_LABOR = INSURANCE_DIRECT_EXPENSE_LABOR - CSMIOS_PMS
                    CSMIOS_PMS = 0
                End If
            End If
PAKSIW:             INSURANCE_DIRECT_EXPENSE_LABOR = N2Str2Zero(rsCSMIOS_MATERIALS!INSLABOR)
        Else
            CSMIOS_LABOR = 0
            CSMIOS_SUBLET = 0
            CSMIOS_TINSPAINT = 0
            CSMIOS_PMS = 0
        End If
        If CSMIOS_PARTS > 0 Then
            CSMIOS_PARTS = CSMIOS_PARTS - INSURANCE_DIRECT_EXPENSE_SPAREPARTS
        End If
        If CSMIOS_MATERIALS > 0 Then
            CSMIOS_MATERIALS = CSMIOS_MATERIALS - INSURANCE_DIRECT_EXPENSE_GOL

        End If

        If CSMIOS_ACCESSORIES > 0 Then
            CSMIOS_ACCESSORIES = CSMIOS_ACCESSORIES - INSURANCE_DIRECT_EXPENSE_ACCESSORIES
        End If

        TOTAL_INSURANCE_AMOUNT = Round(INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_ACCESSORIES, 2)
    End If
End Sub

Private Sub cboMonth_Click()
    Dim iCount, xCount                                 As Integer
    Dim Indx                                           As Integer
    iCount = 0
    Select Case cboMonth.Text
    Case "January": Indx = 31
    Case "February": Indx = 28
    Case "March": Indx = 31
    Case "April": Indx = 30
    Case "May": Indx = 31
    Case "June": Indx = 30
    Case "July": Indx = 31
    Case "August": Indx = 31
    Case "September": Indx = 30
    Case "October": Indx = 31
    Case "November": Indx = 30
    Case "December": Indx = 31
    Case Else: Indx = -1
    End Select
    For xCount = 1 To 30
        lab.Item(xCount).Visible = False
    Next
    Do While iCount <= Indx - 1
        lab.Item(iCount).BackColor = &HE0E0E0
        lab.Item(iCount).Visible = True
        iCount = iCount + 1
    Loop
End Sub

Private Sub cboYear_Click()
    Dim iCount, xCount                                 As Integer
    Dim Indx                                           As Integer
    iCount = 0
    Select Case cboMonth.Text
    Case "January": Indx = 31
    Case "February"
        If NumericVal(cboYear.Text) Mod 4 = 0 Then
            Indx = 29
        Else
            Indx = 28
        End If
    Case "March": Indx = 31
    Case "April": Indx = 30
    Case "May": Indx = 31
    Case "June": Indx = 30
    Case "July": Indx = 31
    Case "August": Indx = 31
    Case "September": Indx = 30
    Case "October": Indx = 31
    Case "November": Indx = 30
    Case "December": Indx = 31
    Case Else: Indx = -1
    End Select
    For xCount = 1 To 30
        lab.Item(xCount).Visible = False
    Next
    Do While iCount <= Indx - 1
        lab.Item(iCount).BackColor = &HE0E0E0
        lab.Item(iCount).Visible = True
        iCount = iCount + 1
    Loop
End Sub

Private Sub cmdCheck_Click()
    If Function_Access(LOGID, "Acess_Process", "IMPORT CASH RECEIPTS") = False Then Exit Sub
    Screen.MousePointer = 11
    Dim str_MSG                                        As String


    str_MSG = "Error Appear In During @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Imported Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    If dtpTranDate.Value <= CDate(GetCUTOFF_DATE) Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Date selected is less than Cut-Off Date."
        cmdExit.Enabled = True
        Screen.MousePointer = 0
        Call cmdShowImp_Click
    Else
        If Option1.Value = True Then
            If COMPANY_CODE = "DGI" Then
                ImportSalesNew
                'ImportUndepositedNew
                gconDMIS.BeginTrans
                If ImportCSMSSales = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Import CSMS Sales")
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                
                If ImportUnDeposit = False Then
                    str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Undeposited")
                    MsgBox str_MSG, vbCritical, "Importing Error"
                    cmdExit.Enabled = True
                    gconDMIS.RollbackTrans
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                    
                gconDMIS.CommitTrans
                Call cmdShowImp_Click
                MsgBox "Import Successfully Completed!", vbInformation, "Finish"
                LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
                Exit Sub
            Else
                gconDMIS.BeginTrans
                    If ImportPMISSales = False Then
                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import PMIS Sales")
                        MsgBox str_MSG, vbCritical, "Importing Error"
                        cmdExit.Enabled = True
                        gconDMIS.RollbackTrans
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
    
'                    If ImportPurelyInternal = False Then
'                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Purely Internal")
'                        MsgBox str_MSG, vbCritical, "Importing Error"
'                        cmdExit.Enabled = True
'                        gconDMIS.RollbackTrans
'                        Screen.MousePointer = 0
'                        Exit Sub
'                    End If
    
                    If ImportCSMSSales = False Then
                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import CSMS Sales")
                        MsgBox str_MSG, vbCritical, "Importing Error"
                        cmdExit.Enabled = True
                        gconDMIS.RollbackTrans
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
    
                    If ImportSMISSales = False Then
                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import SMIS Sales")
                        MsgBox str_MSG, vbCritical, "Importing Error"
                        cmdExit.Enabled = True
                        gconDMIS.RollbackTrans
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
    
                If COMPANY_CODE = M_COMPANY_CODE Then
                    If ImportUnDepositM = False Then
                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Undeposited")
                        MsgBox str_MSG, vbCritical, "Importing Error"
                        cmdExit.Enabled = True
                        gconDMIS.RollbackTrans
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                Else
                    If ImportUnDeposit = False Then
                        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Undeposited")
                        MsgBox str_MSG, vbCritical, "Importing Error"
                        cmdExit.Enabled = True
                        gconDMIS.RollbackTrans
                        Screen.MousePointer = 0
                        Exit Sub
                    End If
                End If
                gconDMIS.CommitTrans
                Call cmdShowImp_Click
                MsgBox "Import Successfully Completed!", vbInformation, "Finish"
                LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
                Exit Sub
        End If
        If Option2.Value = True Then
            gconDMIS.BeginTrans
            If ImportDeposited = False Then
                str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Deposited")
                MsgBox str_MSG, vbCritical, "Importing Error"
                cmdExit.Enabled = True
                gconDMIS.RollbackTrans
                Screen.MousePointer = 0
                Exit Sub
            End If
            gconDMIS.CommitTrans
            Call cmdShowImp_Click
            MsgBox "Import Successfully Completed!", vbInformation, "Finish"
            LogAudit "R", "CASH RECEIPTS IMPORT", dtpTranDate
            Exit Sub
        End If
    End If
ErrorCode:
    '    SaveLogFile
    ShowVBError
End Sub

Sub ImportUndepositedNew()
    Dim i                                              As Long

    'PARTS / ACCESSORIES / MATERIALS IMPORTING =============================================================
    i = 0
    Dim GridImport                                     As Integer
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Call ImportUndeposited2("RECEIPTS", "VAT", Grid1.Cell(GridImport, 3).Text, CDate(dtpTranDate))
            Grid1.Cell(GridImport, 1).Text = 1
        End If

        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
End Sub

Private Sub cmdClearJournals_Click()
    If Option1.Value = True Then
        Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
        Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "' AND STATUS <> 'C'")
        If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
            Screen.MousePointer = 0
            If LOGLEVEL = "ADM" Then
                If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                    Screen.MousePointer = 11
                    gconDMIS.Execute ("delete from AMIS_Reference Where Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "'")
                    gconDMIS.Execute ("delete from AMIS_CRJ_Detail Where STATUS <> 'P' AND Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    cmdShowTrans.Value = True
                    Screen.MousePointer = 0
                    MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                End If
            End If
        End If
    End If
    If Option2.Value = True Then
        Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
        Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "'")
        If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
            Screen.MousePointer = 0
            If LOGLEVEL = "ADM" Then
                If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
                    Screen.MousePointer = 11
                    gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
                    cmdShowTrans.Value = True
                    Screen.MousePointer = 0
                    MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
                End If
            End If
            Exit Sub
        End If
    End If
    Call cmdShowImp_Click
    Call cmdShowTrans_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdShowImp_Click()
    Screen.MousePointer = 11
    If cmdCheck.Value = True Then
    Else
        cmdCheck.Value = False
InitGrids:         DoEvents:
        Grid1.Rows = 1
        Grid2.Rows = 1
        Grid3.Rows = 1
    End If
    Dim tmpDate                                        As String
    Dim iCount                                         As Integer
    MonthIndex
    Dim rsCheckCRJData                                 As ADODB.Recordset
    Dim rsCheckCRJData2                                As ADODB.Recordset

    Set rsCheckCRJData = New ADODB.Recordset
    '    rsCheckCRJData.Open "SELECT * FROM (SELECT DISTINCT OR_DATE AS TRANDATE FROM CMIS_OFF_HD WHERE (PAIDNA = 1 OR STATUS = 'P') AND CANCEL =0 AND MONTH(OR_DATE) = '" & Indx & "' AND YEAR(OR_DATE)='" & cboYear.Text & "' " & _
         '                        "UNION SELECT DISTINCT DATDEPOSIT AS TRANDATE FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND CANCEL = 0 AND MONTH(DATDEPOSIT) > '" & Indx & "' AND YEAR(DATDEPOSIT) > '" & cboYear.Text & "') X ORDER BY TRANDATE ASC", gconDMIS, adOpenKeyset
    '    Do While Not rsCheckCRJData.EOF
    '        iCount = 0
    '        Dim rsCheckJournalHD                      As ADODB.Recordset
    '        Set rsCheckJournalHD = New ADODB.Recordset
    '        rsCheckJournalHD.Open "SELECT DISTINCT JDATE,JTYPE FROM AMIS_JOURNAL_HD WHERE JTYPE = 'CRJ' AND JDATE IN (SELECT * FROM (SELECT DISTINCT OR_DATE AS TRANDATE FROM CMIS_OFF_HD WHERE (PAIDNA = 1 OR STATUS = 'P') AND CANCEL =0 " & _
             '                              "UNION SELECT DISTINCT DATDEPOSIT AS TRANDATE FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND CANCEL = 0) X WHERE TRANDATE = '" & Null2String(rsCheckCRJData!trandate) & "')", gconDMIS, adOpenKeyset
    '        If Not rsCheckJournalHD.EOF And Not rsCheckJournalHD.BOF Then
    '            Do While iCount <= lab.Count - 1
    '                If lab.Item(iCount).Caption = Format(Null2String(rsCheckJournalHD!JDate), "d") Then
    '                    lab.Item(iCount).BackColor = &HC0FFC0
    '                    GoTo Skip
    '                End If
    '                tmpDate = Indx & "/" & iCount & "/" & cboYear.Text
    '
    '                iCount = iCount + 1
    '            Loop
    'Skip:
    '        Else
    '            Do While iCount <= lab.Count - 1
    '                If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckCRJData!trandate)), "d") Then
    '                    lab.Item(iCount).BackColor = &HFFFF&
    '                End If
    '                iCount = iCount + 1
    '            Loop
    '        End If
    '        rsCheckCRJData.MoveNext
    '    Loop
    '----------------------------------------
    rsCheckCRJData.Open "SELECT DISTINCT DATE FROM AMIS_VW_IMPORTED_OR WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckCRJData.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckCRJData!Date)), "d") Then
                lab.Item(iCount).BackColor = &HC0FFC0
                DoEvents
            End If
            iCount = iCount + 1
        Loop

        rsCheckCRJData.MoveNext
        DoEvents
    Loop

    Set rsCheckCRJData2 = New ADODB.Recordset
    rsCheckCRJData2.Open "SELECT DISTINCT DATE FROM AMIS_VW_UNIMPORTED_OR WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckCRJData2.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckCRJData2!Date)), "d") Then
                lab.Item(iCount).BackColor = &HFFFF&
                DoEvents
            End If
            iCount = iCount + 1
        Loop

        rsCheckCRJData2.MoveNext
        DoEvents
    Loop

    Screen.MousePointer = 0
End Sub

Private Sub cmdShowTrans_Click()
    Screen.MousePointer = 11
InitGrids:     DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
    Grid3.AutoRedraw = False
    Grid1.Rows = 1: Grid2.Rows = 1: Grid3.Rows = 1: KIM = 0: LIM = 0
    Dim ORType                                         As String
    Dim IS_Exist                                       As Byte
    Dim rsOR_UNDEPOSITED                               As ADODB.Recordset
    Dim rsOR_DEPOSITED                                 As ADODB.Recordset
    Dim rsUNDEPOSITED_INVOICES                         As ADODB.Recordset

    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
        GoTo ShowTrans
    Else
        If CheckImportedOR(dtpTranDate) = True Then
            MsgBox "Previous transactions dated " & TRANSACTIONDATE & " are not yet imported.", vbExclamation, "Message"
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo ShowTransactions
        End If

ShowTransactions:
    End If
ShowTrans:

    Set rsOR_UNDEPOSITED = New ADODB.Recordset
    'Set rsOR_UNDEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD where OR_DATE = '" & CDate(dtpTranDate) & "' AND (DEPOSIT = FALSE OR DEPOSIT = 0) order by OR_NUM ASC")
    Set rsOR_UNDEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD where (PAIDNA = 1 OR STATUS = 'P') AND OR_DATE = '" & CDate(dtpTranDate) & "' and cancel =0  order by OR_NUM ASC")
    If Not rsOR_UNDEPOSITED.EOF And Not rsOR_UNDEPOSITED.BOF Then
        rsOR_UNDEPOSITED.MoveFirst: KIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsOR_UNDEPOSITED.EOF
            KIM = KIM + 1
            If CheckCRJExisting(Null2String(rsOR_UNDEPOSITED!OR_NUM), N2Str2Zero(rsOR_UNDEPOSITED!VAT)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            If N2Str2Zero(rsOR_UNDEPOSITED!VAT) = 1 Then
                ORType = "VAT"
            Else
                ORType = "NON VAT"
            End If
            Grid1.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_UNDEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_UNDEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_UNDEPOSITED!CUSNAME)
            Set rsUNDEPOSITED_INVOICES = New ADODB.Recordset
            Set rsUNDEPOSITED_INVOICES = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE (TRANTYPE = 'VI' OR TRANTYPE = 'SI' OR TRANTYPE = 'PI' OR TRANTYPE = 'AI' OR TRANTYPE = 'MI') AND OR_NUM = " & N2Str2Null(rsOR_UNDEPOSITED!OR_NUM) & " AND VAT = " & N2Str2Zero(rsOR_UNDEPOSITED!VAT))
            If Not rsUNDEPOSITED_INVOICES.EOF And Not rsUNDEPOSITED_INVOICES.BOF Then
                ShowUnImportedPaidInvoices Null2String(rsUNDEPOSITED_INVOICES!TranType), Null2String(rsUNDEPOSITED_INVOICES!INVOICENO)
            End If
            rsOR_UNDEPOSITED.MoveNext
            DoEvents
        Loop
        If KIM = 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If

    Set rsOR_DEPOSITED = New ADODB.Recordset
    'Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' and Cancel = 0 Order by OR_NUM ASC")
    If COMPANY_CODE = M_COMPANY_CODE Then
        Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD_DepositedM Where (DEPOSIT1 = 1 OR DEPOSIT2 =1) AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' and Cancel = 0 Order by OR_NUM ASC")
    Else
        Set rsOR_DEPOSITED = gconDMIS.Execute("Select * from CMIS_OFF_HD_Deposited Where DEPOSIT = 1 AND DATDEPOSIT = '" & CDate(dtpTranDate) & "' and Cancel = 0 Order by OR_NUM ASC")
    End If
    If Not rsOR_DEPOSITED.EOF And Not rsOR_DEPOSITED.BOF Then
        rsOR_DEPOSITED.MoveFirst: KIM = 0
        Grid2.AutoRedraw = False
        Do While Not rsOR_DEPOSITED.EOF
            KIM = KIM + 1
            If COMPANY_CODE = M_COMPANY_CODE Then
                If CheckDRJExistingM(Null2String(rsOR_DEPOSITED!OR_NUM), N2Str2Zero(rsOR_DEPOSITED!VAT), NumericVal(rsOR_DEPOSITED!OR_AMT)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
            Else
                If CheckDRJExisting(Null2String(rsOR_DEPOSITED!OR_NUM), N2Str2Zero(rsOR_DEPOSITED!VAT)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
            End If
            If N2Str2Zero(rsOR_DEPOSITED!VAT) = 1 Then
                ORType = "VAT"
            Else
                ORType = "NON VAT"
            End If
            Grid2.AddItem IS_Exist & Chr(9) & ORType & Chr(9) & Null2String(rsOR_DEPOSITED!OR_NUM) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOR_DEPOSITED!OR_AMT)) & Chr(9) & Null2String(rsOR_DEPOSITED!CUSNAME)
            rsOR_DEPOSITED.MoveNext
            DoEvents
        Loop
        If KIM = 0 Then Grid2.RemoveItem 1
        Grid2.AutoRedraw = True
        Grid2.Refresh
    End If


    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
    End If
    If LIM = 0 Then Grid3.RemoveItem 1
    Grid3.AutoRedraw = True
    Grid3.Refresh
    Screen.MousePointer = 0
End Sub

Private Sub dtpTranDate_Change()
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorCode
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    dtpTranDate = LOGDATE
    InitGrids
    InitCombo
    InitNoDays
    If COMPANY_CODE = "HBK" Then Option2.Enabled = False

    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.DESCRIPTION, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Function GetEntityCode(XXX As String) As String
    Dim rsEntity                                       As ADODB.Recordset
    Set rsEntity = gconDMIS.Execute("select * from All_Entity where Code = '" & XXX & "'")
    If Not rsEntity.EOF And Not rsEntity.BOF Then
        GetEntityCode = rsEntity!ENTITYCODE
    End If
    Set rsEntity = Nothing
End Function

Function CheckImportedOR(xOR_Date As String) As Boolean
'Update: ACL 11162009
    Dim xNoDays                                        As Integer
    Dim rsCheckCRJ                                     As ADODB.Recordset
    Dim rsCheckOR                                      As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='CRJ' AND STATUS <> 'C' AND JDATE < '" & CDate(xOR_Date) & "' ORDER BY JDATE DESC", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        Set rsCheckOR = New ADODB.Recordset
        rsCheckOR.Open "Select * FROM (SELECT TOP 1 OR_DATE AS TRANDATE FROM CMIS_OFF_HD WHERE (PAIDNA = 1 OR STATUS = 'P') AND OR_DATE > '" & Null2Date(rsCheckCRJ!JDATE) & "' AND CANCEL =0  ORDER BY OR_DATE ASC " & _
                       "UNION SELECT TOP 1 DATDEPOSIT AS TRANDATE FROM CMIS_OFF_HD_DEPOSITED WHERE DEPOSIT = 1 AND DATDEPOSIT > '" & Null2Date(rsCheckCRJ!JDATE) & "' AND CANCEL = 0 ORDER BY DATDEPOSIT ASC) X ORDER BY TRANDATE ASC", gconDMIS, adOpenKeyset
        If Not rsCheckOR.EOF And Not rsCheckOR.BOF Then
            TRANSACTIONDATE = Null2Date(rsCheckOR!trandate)
            xNoDays = DateDiff("d", TRANSACTIONDATE, dtpTranDate)
            'If Format(CDate(dtpTranDate), "mm/dd/yyyy") > Format(CDate(TRANSACTIONDATE), "mm/dd/yyyy") Then
            If xNoDays > 0 Then
                CheckImportedOR = True
            End If
        End If
    End If
    Set rsCheckCRJ = Nothing
End Function

Function CheckToClearOR(xOR_Date As String) As Boolean
    Dim rsCheckCRJ                                     As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'CRJ' and Jdate > '" & CDate(xOR_Date) & "' Order by JDate Asc", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        CheckToClearOR = True
    End If
End Function

Function CheckToClearDRJ(xOR_Date As String) As Boolean
    Dim rsCheckCRJ                                     As ADODB.Recordset
    Set rsCheckCRJ = New ADODB.Recordset
    rsCheckCRJ.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'DRJ' and Jdate > '" & CDate(xOR_Date) & "' Order by JDate Asc", gconDMIS, adOpenKeyset
    If Not rsCheckCRJ.EOF And Not rsCheckCRJ.BOF Then
        CheckToClearDRJ = True
    End If
End Function

Function CheckDepositedOR(xOR_Date As String) As Boolean
    Dim rsCheckDRJ                                     As ADODB.Recordset
    Set rsCheckDRJ = New ADODB.Recordset
    rsCheckDRJ.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='DRJ' AND JDATE<'" & CDate(xOR_Date) & "'", gconDMIS, adOpenKeyset
    If Not rsCheckDRJ.EOF And Not rsCheckDRJ.BOF Then
        CheckDepositedOR = True
    End If
    Set rsCheckDRJ = Nothing
End Function

Function CheckIfBank(xCUSCDE As String) As Boolean
    Dim rsCheckCode                                    As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCUSCDE) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                            As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Private Sub lab_Click(Index As Integer)
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
    cmdCheck.Enabled = False
    cmdClearJournals.Enabled = False
    MonthIndex
    xTranDate = Indx & "/" & lab.Item(Index).Caption & "/" & cboYear.Text
    dtpTranDate.Value = xTranDate
    Call cmdShowTrans_Click
End Sub

Sub MonthIndex()
    Select Case cboMonth.Text
    Case "January": Indx = 1
    Case "February": Indx = 2
    Case "March": Indx = 3
    Case "April": Indx = 4
    Case "May": Indx = 5
    Case "June": Indx = 6
    Case "July": Indx = 7
    Case "August": Indx = 8
    Case "September": Indx = 9
    Case "October": Indx = 10
    Case "November": Indx = 11
    Case "December": Indx = 12
    Case Else: Indx = -1
    End Select
End Sub

Sub InitNoDays()
    Dim iCount                                         As Integer
    For iCount = 1 To 31
        lab.Item(iCount - 1).Caption = iCount
    Next
End Sub

Sub InitCombo()
    Dim NoDays                                         As Integer
    With cboYear
        .AddItem ("2005")
        .AddItem ("2006")
        .AddItem ("2007")
        .AddItem ("2008")
        .AddItem ("2009")
        .AddItem ("2010")
        .AddItem ("2011")
        .AddItem ("2012")
        .AddItem ("2013")
        .AddItem ("2014")
        .AddItem ("2015")
    End With
    cboYear.Text = Format(LOGDATE, "yyyy")

    With cboMonth
        .AddItem ("January")
        .AddItem ("February")
        .AddItem ("March")
        .AddItem ("April")
        .AddItem ("May")
        .AddItem ("June")
        .AddItem ("July")
        .AddItem ("August")
        .AddItem ("September")
        .AddItem ("October")
        .AddItem ("November")
        .AddItem ("December")
    End With
    cboMonth.ListIndex = Month(LOGDATE) - 1
End Sub

Function ImportUnDepositM() As Boolean
'HEADER
    On Error GoTo ErrorCode
    Dim J_JDATE                                        As String
    Dim J_VOUCHERNO                                    As String
    Dim J_JTYPE                                        As String
    Dim J_JNO                                          As String
    Dim J_REMARKS                                      As String
    Dim J_VENDORCODE                                   As String
    Dim J_CUSTOMERCODE                                 As String
    Dim J_CUSTOMERCODE2                                As String
    Dim J_CHECKNO                                      As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE            As String
    Dim J_INVOICETYPE, J_INVOICENO                     As String
    Dim J_CHECKDATE, J_BANKCODE                        As String
    Dim J_REFNO, J_REFDATE                             As String
    Dim J_TERMS, J_DEALER                              As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                  As String
    Dim J_REFERENCENO                                  As String
    Dim J_ENTITY                                       As String
    Dim J_OUTBALANCE                                   As Double
    Dim J_AMOUNTTOPAY                                  As Double
    Dim J_INVOICEAMT                                   As Double
    Dim J_BALANCE                                      As Double
    Dim J_AMOUNTPAID                                   As Double

    'DETAIL
    Dim J_ACCT_CODE                                    As String
    Dim J_ACCT_NAME                                    As String
    Dim J_STATUS                                       As String
    Dim J_JITEMNO                                      As String
    Dim J_JITEMNO_2                                    As String
    Dim J_ALLENTITY                                    As String
    Dim J_DEBIT                                        As Double
    Dim J_CREDIT                                       As Double
    Dim J_TAX                                          As Double
    Dim J_GROSS                                        As Double
    Dim J_NET                                          As Double

    Dim rsJournal_HDDup                                As ADODB.Recordset
    Dim CMIS_OR_NUM                                    As String
    Dim CMIS_OR_DATE                                   As String
    Dim CMIS_OR_AMT                                    As String
    Dim CMIS_DISCOUNT                                  As String
    Dim CMIS_TAX                                       As String
    Dim CMIS_CUSCDE                                    As String
    Dim CMIS_CUSNAME                                   As String
    Dim CMIS_DEPOSIT                                   As String
    Dim CMIS_BANKCODE                                  As String
    Dim CMIS_BANK                                      As String
    Dim CMIS_TSEKE                                     As String
    Dim CMIS_CHECKDATE                                 As String
    Dim CMIS_STATUS                                    As String
    Dim CMIS_TYPE_PAYMENT1                             As String
    Dim CMIS_TYPE_PAYMENT2                             As String
    Dim CMIS_TYPE_PAYMENT3                             As String
    Dim CMIS_DT_TRANTYPE                               As String
    Dim CMIS_DT_REFERENCE                              As String
    Dim CMIS_DT_CUSCDE                                 As String
    Dim CMIS_DT_DESCRIPT                               As String
    Dim CMIS_DT_REFERENCENO                            As String
    Dim CMIS_DT_DOCDTE                                 As String
    Dim CMIS_DT_PAIDFOR                                As String
    Dim CMIS_CASHAMOUNT                                As Double
    Dim CMIS_CHKAMOUNT                                 As Double
    Dim CMIS_CARDAMOUNT                                As Double
    Dim TOTAL_DEBIT                                    As Double
    Dim TOTAL_CREDIT                                   As Double
    Dim CMIS_DT_AMOUNT                                 As Double
    Dim CMIS_DT_PAYMENT                                As Double
    Dim CMIS_DT_DISCOUNT                               As Double
    Dim CMIS_DT_TAX                                    As Double
    Dim CMIS_IS_VAT                                    As Boolean
    Dim i                                              As Long

    Dim rsOFF_HD                                       As ADODB.Recordset
    Dim rsOFF_DT                                       As ADODB.Recordset
    Dim rsSJ_DATA                                      As ADODB.Recordset
    Dim rsCheckJournal_HD                              As ADODB.Recordset

    Dim PV_MRRNO                                       As String
    Dim PV_INVNO                                       As String
    Dim PV_PRODNO                                      As String
    Dim J_JVOUCHERNO                                   As String
    Dim PV_STATUS, PV_ITEMNO                           As String
    Dim PV_AMOUNT                                      As Double
    Dim SJ_PV_ITEMNO                                   As Integer
    Dim GridImport                                     As Integer

    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsOFF_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 1 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            Else
                Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_OFF_HD Where OR_NUM = '" & Grid1.Cell(GridImport, 3).Text & "' AND VAT = 0 AND OR_DATE = '" & CDate(dtpTranDate) & "' Order by OR_NUM ASC")
            End If
            If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
                CMIS_OR_NUM = Null2String(rsOFF_HD!OR_NUM)
                CMIS_OR_DATE = Null2Date(rsOFF_HD!OR_DATE)
                CMIS_OR_AMT = Null2String(rsOFF_HD!OR_AMT)
                CMIS_DISCOUNT = Null2String(rsOFF_HD!DISCOUNT)
                CMIS_TAX = Null2String(rsOFF_HD!tax)
                CMIS_CASHAMOUNT = Round(N2Str2Zero(rsOFF_HD!CashAmount), 2)
                CMIS_CHKAMOUNT = Round(N2Str2Zero(rsOFF_HD!ChkAmount), 2)
                CMIS_CARDAMOUNT = Round(N2Str2Zero(rsOFF_HD!cardamount), 2)
                CMIS_CUSCDE = Null2String(rsOFF_HD!CUSCDE)
                CMIS_CUSNAME = Null2String(rsOFF_HD!CUSNAME)
                CMIS_DEPOSIT = Null2String(rsOFF_HD!DEPOSIT1)
                CMIS_BANKCODE = Null2String(rsOFF_HD!bankcode)
                CMIS_BANK = Null2String(rsOFF_HD!Bank)
                CMIS_TSEKE = Null2String(rsOFF_HD!Tseke) & Null2String(rsOFF_HD!cardnumber)
                CMIS_TYPE_PAYMENT1 = Null2String(rsOFF_HD!TOF1)
                CMIS_TYPE_PAYMENT2 = Null2String(rsOFF_HD!TOF2)
                CMIS_TYPE_PAYMENT3 = Null2String(rsOFF_HD!TOF3)
                CMIS_BANKCODE = Null2String(rsOFF_HD!bankcode)
                CMIS_ENTITYCODE = GetEntityCode(Null2String(rsOFF_HD!CUSCDE))
                CMIS_ALLENTITY = CMIS_ENTITYCODE + CMIS_CUSCDE
                If Null2Date(rsOFF_HD!CheckDate) = "" Then
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!carddate)
                Else
                    CMIS_CHECKDATE = Null2Date(rsOFF_HD!CheckDate)
                End If
                CMIS_STATUS = Null2String(rsOFF_HD!Status)
                CMIS_IS_VAT = Null2Bool(rsOFF_HD!VAT)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0

                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If

                J_JDATE = N2Date2Null(CMIS_OR_DATE)
                J_VOUCHERNO = N2Str2Null(GetCRJVoucherNo())
                J_JTYPE = "'CRJ'"

                'INSERTED SEPTEMBER 8, 2007
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 1 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT WHERE VAT = 0 AND OR_NUM = '" & CMIS_OR_NUM & "'")
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst
                    Do While Not rsOFF_DT.EOF
                        J_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_INVOICENUM = Null2String(rsOFF_DT!INVOICENO)

                        If Null2String(rsOFF_DT!TranType) = "OTH" Then
                            J_REMARKS = SetOtherTransaction(Null2String(rsOFF_DT!PAIDFOR)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        Else
                            J_REMARKS = SetTransaction(Null2String(rsOFF_DT!TranType)) & ": " & Null2String(rsOFF_DT!Reference) & " " & ToDoubleNumber(N2Str2Zero(rsOFF_DT!payment))
                        End If
                        rsOFF_DT.MoveNext
                        If Not rsOFF_DT.EOF Then J_REMARKS = "" & Chr(9)
                    Loop
                    J_REMARKS = N2Str2Null(J_REMARKS)
                Else
                    J_REMARKS = "NULL"
                End If
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(CMIS_CUSCDE)
                J_DEPOSIT = 0
                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                J_INVOICEAMT = Round(NumericVal(CMIS_OR_AMT), 2)
                J_BALANCE = 0
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(CMIS_OR_DATE)
                If CMIS_IS_VAT = True Then
                    J_INVOICENO = N2Str2Null(Left(CMIS_OR_NUM, 6))
                Else
                    J_INVOICENO = N2Str2Null("NV" & Left(CMIS_OR_NUM, 6))
                End If
                J_CHECKNO = N2Str2Null(CMIS_TSEKE)
                J_DUEDATE = N2Date2Null(CMIS_CHECKDATE)
                If Null2String(rsOFF_HD!TOF1) = "1" Then
                    J_PAYTYPE = "'CASH'"
                ElseIf Null2String(rsOFF_HD!TOF2) = "2" Then
                    J_PAYTYPE = "'CHECK'"
                ElseIf Null2String(rsOFF_HD!TOF3) = "3" Then
                    J_PAYTYPE = "'CARD'"
                Else
                    J_PAYTYPE = "NULL"
                End If
                J_INVOICETYPE = "'CI'"
                J_CHECKDATE = N2Str2Null(CMIS_CHECKDATE)
                J_BANKCODE = N2Str2Null(CMIS_BANKCODE)
                J_REFNO = N2Str2Null(CMIS_TSEKE)
                J_REFDATE = N2Date2Null(CMIS_CHECKDATE)
                J_ENTITY = N2Str2Null(CMIS_ENTITYCODE)
                J_ALLENTITY = N2Str2Null(CMIS_ALLENTITY)
                J_TERMS = "NULL"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'CASH ON HAND
                'DESCRIPTION: CASH
                If CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = NumericVal(CMIS_CARDAMOUNT)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT = 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT = 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT = 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                ElseIf CMIS_CASHAMOUNT > 0 And CMIS_CHKAMOUNT > 0 And CMIS_CARDAMOUNT > 0 Then
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                End If

                J_JITEMNO = 0
                If CMIS_TYPE_PAYMENT1 = "1" Or CMIS_TYPE_PAYMENT2 = "2" Then
                    J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
                    J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CASH ON HAND"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CASH ON HAND")))
                    J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2) + Round(NumericVal(CMIS_CHKAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                End If
                If CMIS_TYPE_PAYMENT3 = "3" Then
                    J_JITEMNO = Format(NumericVal(J_JITEMNO + 1), "0000")
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAccountCode("CARD ON HAND"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccountCode("CARD ON HAND")))
                    End If
                    J_DEBIT = Round(NumericVal(CMIS_CARDAMOUNT), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                    gconDMIS.Execute SQL_STATEMENT
                End If

                'DESCRIPTION: Checking for Customer Deposits and ADD upon full payment of the invoice amount
                If J_INVOICENUM <> "" Then
                    Dim rsCheckDeposit                 As ADODB.Recordset
                    Set rsCheckDeposit = New ADODB.Recordset
                    rsCheckDeposit.Open "Select * from CMIS_Deposits where InvoiceNo = '" & J_INVOICENUM & "'", gconDMIS, adOpenForwardOnly
                    If Not rsCheckDeposit.EOF And Not rsCheckDeposit.BOF Then
                        J_JITEMNO_2 = 1
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(Null2String(rsCheckDeposit!PAIDFOR))))
                        J_DEBIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                        J_DEPOSIT = Round(NumericVal(rsCheckDeposit!amount), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    Set rsCheckDeposit = Nothing
                End If

                'DESCRIPTION: Bank Charges for credit card payment receive
                If CheckIfBank(CMIS_CUSCDE) = True Then
                    If J_REFERENCENO <> "" Then
                        J_JITEMNO_2 = 1
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null("71-35000-20")
                        J_ACCT_NAME = N2Str2Null(Setacctname("71-35000-20"))
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)
                        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            If CMIS_CASHAMOUNT > 0 Then
                                J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                            Else
                                J_DEBIT = Round((NumericVal(CMIS_CHKAMOUNT) / 0.97) * 0.025, 2)
                            End If
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        'DESCRIPTION: Creditable Withholding Tax for credit card payment receive
                        J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                        J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                        J_ACCT_CODE = N2Str2Null("11-07000-00")
                        J_ACCT_NAME = N2Str2Null(Setacctname("11-07000-00"))
                        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                            If CMIS_CASHAMOUNT > 0 Then
                                J_DEBIT = Round(NumericVal(CMIS_CASHAMOUNT), 2)
                            Else
                                J_DEBIT = Round((NumericVal(CMIS_CHKAMOUNT) / 0.97) * 0.005, 2)
                            End If
                        Else
                            J_DEBIT = Round(NumericVal(CMIS_CHKAMOUNT), 2)
                        End If
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                End If
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                'DESCRIPTION: CREDIT
                Set rsOFF_DT = New ADODB.Recordset
                If Grid1.Cell(GridImport, 2).Text = "VAT" Then
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 1 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                Else
                    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_OFF_DT Where VAT = 0 AND OR_NUM = " & N2Str2Null(CMIS_OR_NUM))
                End If
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    rsOFF_DT.MoveFirst: SJ_PV_ITEMNO = 0
                    Do While Not rsOFF_DT.EOF
                        CMIS_DT_TRANTYPE = Null2String(rsOFF_DT!TranType)
                        CMIS_DT_REFERENCE = Null2String(rsOFF_DT!INVOICENO)
                        CMIS_DT_CUSCDE = Null2String(rsOFF_DT!CUSCDE)
                        CMIS_DT_DESCRIPT = Null2String(rsOFF_DT!DESCRIPT)
                        CMIS_DT_AMOUNT = N2Str2Zero(rsOFF_DT!amount)
                        CMIS_DT_DOCDTE = Null2String(rsOFF_DT!DOCDTE)
                        CMIS_DT_PAYMENT = N2Str2Zero(rsOFF_DT!payment)
                        CMIS_DT_DISCOUNT = N2Str2Zero(rsOFF_DT!DISCOUNT)
                        CMIS_DT_TAX = N2Str2Zero(rsOFF_DT!tax)
                        CMIS_DT_PAIDFOR = Null2String(rsOFF_DT!PAIDFOR)

                        'DESCRIPTION: For credit card detail
                        CMIS_DT_REFERENCENO = Null2String(rsOFF_DT!ReferenceNo)
                        J_JVOUCHERNO = J_VOUCHERNO
                        SJ_PV_ITEMNO = SJ_PV_ITEMNO + 1
                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                        PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                        PV_AMOUNT = CMIS_DT_PAYMENT   ' AMOUNT
                        PV_STATUS = "'N'"
                        PV_INVDATE = N2Date2Null(rsOFF_DT!ORDATE)

                        PV_MRRNO = "'" & CMIS_DT_TRANTYPE & "'"
                        Set rsSJ_DATA = New ADODB.Recordset
                        Set rsSJ_DATA = gconDMIS.Execute("Select * from AMIS_Journal_HD Where jtype = 'SJ' and invoicetype = " & PV_MRRNO & " and invoiceno = " & N2Str2Null(CMIS_DT_REFERENCE))
                        If Not rsSJ_DATA.EOF And Not rsSJ_DATA.BOF Then
                            J_JVOUCHERNO = J_VOUCHERNO
                            PV_ITEMNO = N2Str2Null(Format(SJ_PV_ITEMNO, "0000"))
                            PV_INVNO = N2Str2Null(CMIS_DT_REFERENCE)    ' NO
                            PV_PRODNO = N2Date2Null(rsSJ_DATA!invoicedate)    ' DATE
                            PV_AMOUNT = CMIS_DT_PAYMENT + J_DEPOSIT    ' AMOUNT
                            PV_STATUS = "'N'"

                            'DESCRIPTION: Modify database structure - added fields: CusCde - Insert J_Class and Customer Code upon importing (,J_Class,CusCde) NOT YET
                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                            "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                            " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                            ", " & PV_STATUS & ")"
                            gconDMIS.Execute SQL_STATEMENT

                            'DESCRIPTION: Customer Deposits
                            If J_INVOICENUM <> "" Then
                                Dim rsCheckDetail      As ADODB.Recordset
                                Set rsCheckDetail = New ADODB.Recordset
                                rsCheckDetail.Open "SELECT ReferenceNo FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                                If Not rsCheckDetail.EOF And Not rsCheckDetail.BOF Then
                                    Dim rsCheckDetail2 As ADODB.Recordset
                                    Set rsCheckDetail2 = New ADODB.Recordset
                                    rsCheckDetail2.Open "SELECT * FROM CMIS_OFF_DT WHERE REFERENCENO=" & N2Str2Null(rsCheckDetail!ReferenceNo) & " AND OR_NUM NOT IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE INVOICENO =" & PV_INVNO & ")", gconDMIS, adOpenForwardOnly
                                    If Not rsCheckDetail2.EOF And Not rsCheckDetail2.BOF Then
                                        Dim rsCheckVoucher As ADODB.Recordset
                                        Set rsCheckVoucher = New ADODB.Recordset
                                        rsCheckVoucher.Open "SELECT VOUCHERNO,InvoiceAmt from AMIS_JOURNAL_HD WHERE InvoiceNo =" & N2Str2Null(rsCheckDetail2!OR_NUM) & "", gconDMIS, adOpenForwardOnly
                                        If Not rsCheckVoucher.EOF And Not rsCheckVoucher.BOF Then
                                            PV_AMOUNT = rsCheckVoucher!InvoiceAmt / 0.97
                                            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                            "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                            " values (" & N2Str2Null(rsCheckVoucher!VOUCHERNO) & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                                            ", " & PV_STATUS & ")"
                                            gconDMIS.Execute SQL_STATEMENT
                                        End If
                                        Set rsCheckVoucher = Nothing
                                    End If
                                    Set rsCheckDetail2 = Nothing
                                End If
                                Set rsCheckDetail = Nothing
                            End If

                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(PV_MRRNO)

                            Set rsCheckJournal_HD = New ADODB.Recordset
                            Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
                            If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'Y' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"

                                    gconDMIS.Execute SQL_STATEMENT

                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "MM", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)



                                Else
                                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                    " ReceiveStatus = 'N' " & "," & _
                                                    " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                    " Balance = Balance - " & PV_AMOUNT & _
                                                    " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
                                    gconDMIS.Execute SQL_STATEMENT

                                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                    NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                End If
                            Else
                                Set rsCheckJournal_HD = New ADODB.Recordset
                                Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
                                If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
                                    If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then

                                        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                                                         " ReceiveStatus = 'Y' " & "," & _
                                                         " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                         " Balance = Balance - " & PV_AMOUNT & _
                                                         " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                    Else
                                        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                                        " ReceiveStatus = 'N' " & "," & _
                                                        " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                                                        " Balance = Balance - " & PV_AMOUNT & _
                                                        " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
                                        gconDMIS.Execute SQL_STATEMENT

                                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                                        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                                    End If
                                End If
                            End If
                        Else
                            'DESCRIPTION: Insert detail for A/R CREDIT CARD / Credit Card payment receive
                            Set rsCheckCreditCardSJ = gconDMIS.Execute("select * from CMIS_OFF_DT where ReferenceNo = " & N2Str2Null(CMIS_DT_REFERENCENO) & " and TRANTYPE <>'OTH'")
                            If Not rsCheckCreditCardSJ.EOF And Not rsCheckCreditCardSJ.BOF Then
                                PV_INVNO = N2Str2Null(rsCheckCreditCardSJ!INVOICENO)
                                PV_MRRNO = "'" & rsCheckCreditCardSJ!TranType & "'"
                                PV_AMOUNT = (NumericVal(PV_AMOUNT) / 0.97)
                                SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                                                "(VoucherNo,Jdate,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                                                " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & _
                                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_INVDATE & ", " & PV_AMOUNT & _
                                                ", " & PV_STATUS & ")"
                                gconDMIS.Execute SQL_STATEMENT

                            End If
                        End If

                        'J_JITEMNO = "'0002'"
                        J_JITEMNO = Format(NumericVal(J_JITEMNO) + 1, "0000")
                        'RO  - SERVICE REPAIR ORDER
                        If CMIS_DT_TRANTYPE = "RO" Or CMIS_DT_TRANTYPE = "SI" Then
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                If ReturnSITerm(CMIS_DT_REFERENCE) = "CHG" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                                End If
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
                            End If
                        End If
                        'CSH - PARTS CASH INVOICE
                        'CHG - PARTS CHARGE INVOICE
                        'If CMIS_DT_TRANTYPE = "CSH" Or CMIS_DT_TRANTYPE = "CHG" Then
                        If CMIS_DT_TRANTYPE = "PI" Then
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                            End If
                            'J_ACCT_CODE = COA_AR_TRADE_PARTS
                            'J_ACCT_NAME = N2Str2Null(Setacctname(COA_AR_TRADE_PARTS))
                        End If
                        'VI  - VEHICLE INVOICE
                        If CMIS_DT_TRANTYPE = "VI" Then
                            If COMPANY_CODE = "HMH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                            ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                            End If
                        End If
                        'EST - SERVICE ESTIMATE
                        If CMIS_DT_TRANTYPE = "EST" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDeposit_AccountCode("SERVICE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeposit_AccountCode("SERVICE")))
                        End If
                        '* IOC - INTER OFFICE COLLECTION
                        'BRA - BRANCH PAYMENT
                        'If CMIS_DT_TRANTYPE = "BRA" Then
                        '    J_ACCT_CODE = COA_BRANCH_LEGASPI
                        '    J_ACCT_NAME = N2Str2Null(Setacctname(COA_BRANCH_LEGASPI))
                        'End If
                        'WAR - A/R WARRANTY CLAIMS
                        'INV - INVENTORIES - GAS,OIL,LUBS
                        'CRD - CREDIT CARD PAYMENT
                        'OTH
                        If COMPANY_CODE = "HBK" Then  ' BTT
                            If CMIS_DT_TRANTYPE = "AI" Then
                                J_ACCT_CODE = N2Str2Null("11-02104-00")
                                J_ACCT_NAME = N2Str2Null(Setacctname("11-02104-00"))
                            End If
                        End If
                        If CMIS_DT_TRANTYPE = "OTH" Then
                            CMIS_DT_AMOUNT = CMIS_DT_PAYMENT
                            'OTHER TRANSACTION
                            J_ACCT_CODE = N2Str2Null(SetOTHChartCodes(CMIS_DT_PAIDFOR))
                            J_ACCT_NAME = N2Str2Null(Setacctname(SetOTHChartCodes(CMIS_DT_PAIDFOR)))
                        End If

                        If J_INVOICENUM <> "" Then
                            Dim rsCheckDeposit2        As ADODB.Recordset
                            Set rsCheckDeposit2 = New ADODB.Recordset
                            rsCheckDeposit2.Open "Select * from CMIS_Deposits where InvoiceNo =' " & J_INVOICENUM & "'", gconDMIS, adOpenForwardOnly
                            If Not rsCheckDeposit2.EOF And Not rsCheckDeposit2.BOF Then
                                J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                                J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                                J_DEPOSIT = Round(NumericVal(rsCheckDeposit2!amount), 2)
                            End If
                            Set rsCheckDeposit2 = Nothing
                        End If
                        J_GROSS = Round(NumericVal(CMIS_DT_PAYMENT), 2)
                        J_TAX = 0
                        J_NET = Round(NumericVal(CMIS_DT_PAYMENT), 2) + Round(NumericVal(J_DEPOSIT), 2)
                        J_DEBIT = 0
                        If J_ACCT_CODE = N2Str2Null("11-02002-00") Then
                            J_JITEMNO_2 = NumericVal(J_JITEMNO_2) + 1
                            J_JITEMNO = N2Str2Null(Format(J_JITEMNO_2, "0000"))
                            J_CREDIT = Round(NumericVal(J_NET) / 0.97, 2)
                        Else
                            J_CREDIT = Round(NumericVal(J_NET), 2)
                        End If

                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        J_CUSTOMERCODE2 = N2Str2Null(CMIS_DT_CUSCDE)

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ReferenceNo,Entity)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_CUSTOMERCODE2 & "," & J_ALLENTITY & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        'DESCRIPTION: Insert in AMIS_Reference to be used in Customer Ledger for Customer and Credit Card Company
                        SQL_STATEMENT = "insert into AMIS_REFERENCE (VoucherNo,Jtype,ReferenceNo,JDate) values (" & J_VOUCHERNO & "," & J_JTYPE & "," & J_CUSTOMERCODE2 & "," & J_JDATE & ")"
                        gconDMIS.Execute SQL_STATEMENT

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_Det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                        rsOFF_DT.MoveNext
                    Loop
                End If
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,ReferenceNo,Bank,Entity_Class)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & "," & J_CUSTOMERCODE & "," & N2Str2Null(CMIS_BANK) & "," & J_ENTITY & ")"

                gconDMIS.Execute SQL_STATEMENT
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)

                Grid1.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    Screen.MousePointer = 0

    ImportUnDepositM = True
    Exit Function
ErrorCode:
    ImportUnDepositM = False
End Function

Function ReturnCostOFSALES_PARTS(STR_TRANNO As String, STR_TRANTYPE As String) As Double
    Dim RSCOST                                         As ADODB.Recordset
    Set RSCOST = gconDMIS.Execute("SELECT ISNULL(SUM(TRANUCOST),0) AS TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANNO=" & N2Str2Null(STR_TRANNO) & " AND TRANTYPE=" & N2Str2Null(STR_TRANTYPE))
    If Not RSCOST.EOF Or Not RSCOST.BOF Then
        ReturnCostOFSALES_PARTS = Round(NumericVal(RSCOST!TRANUCOST), 2)
    End If

End Function
Sub INITIALIZE_COST_VARIABLES(XXX As String)
'COST
'====================================================================================================================================================================================
    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST ),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS PARTS_COST from CSMIOS_vw_PARTSCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_PARTS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!PARTS_COST), 2)
    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST ),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS MAT_COST from CSMIOS_vw_MATCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_MATERIALS_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!MAT_COST), 2)
    Set rsCSMIOS_PARTS = New ADODB.Recordset
    If COMPANY_CODE = "HBK" Then
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then
            CSMS_ACCCOST = Round(N2Str2Zero(rsCSMIOS_PARTS!ACC_COST), 2)
        End If
    Else
        Set rsCSMIOS_PARTS = gconDMIS.Execute("Select ROUND(sum(DETCOST * DETVOL),2) AS ACC_COST from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
    End If
    If Not rsCSMIOS_PARTS.EOF And Not rsCSMIOS_PARTS.BOF Then CSMIOS_ACCESSORIES_COST = Round(N2Str2Zero(rsCSMIOS_PARTS!ACC_COST), 2)
    '====================================================================================================================================================================================
End Sub

Private Sub IMPORT_ACCESSORIES()
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
        If CSMIOS_ACCESSORIES > 0 Then
            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
            If CSMIOS_TERM = "CSH" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED")))
            Else
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("ACCESSORIES", "INSTALLED")))
            End If
            J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
            If CSMIOS_INVOICE = "PDI RO" Then
                J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
            Else
                If CSMIOS_VAT_EXEMPT = True Then
                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES), 2)
                Else
                    J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES / 1.12), 2) * 0.12), 2)
                    J_NET = Round(NumericVal(CSMIOS_ACCESSORIES) - Round(NumericVal(J_TAX), 2), 2)
                End If
            End If
            J_DEBIT = 0
            J_CREDIT = Round(NumericVal(J_NET), 2)
            ALL_CREDIT = ALL_CREDIT + J_CREDIT
            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            'COST OF SALES
            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
            If CSMIOS_TERM = "CSH" Then
                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES", "INSTALLED"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES", "INSTALLED")))
            Else
                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("ACCESSORIES", "INSTALLED"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("ACCESSORIES", "INSTALLED")))
            End If
            J_DEBIT = Round(CSMS_ACCCOST, 2)
            J_CREDIT = 0
            ALL_DEBIT = ALL_DEBIT + J_DEBIT
            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            'INVENTORY
            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
            If CSMIOS_TERM = "CSH" Then
                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
            Else
                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
            End If
            J_DEBIT = 0
            J_CREDIT = Round(CSMS_ACCCOST, 2)
            ALL_CREDIT = ALL_CREDIT + J_CREDIT
            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
        End If
    End If
End Sub

Function ReturnExpense(XXX As String, Optional YYY As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If Trim(YYY) = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnExpense = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Private Sub Import_Warranty()
    If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL > 0 Then

        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"

        WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetSJVoucherNo()) + 1, "000000"))
        WARRANTY_ItemCnt = 0
        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
        'LABOR
        If WARRANTY_DIRECT_EXPENSE_LABOR > 0 Then
            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
            If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                End If
            Else
                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                ElseIf COMPANY_CODE = "HCC" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                End If
            End If
            J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR), 2)
            If CSMIOS_VAT_EXEMPT = True Then
                J_TAX = 0
            Else
                J_TAX = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) / 9.3333, 2)
            End If
            J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR) - Round(NumericVal(J_TAX), 2), 2)
            J_DEBIT = 0
            J_CREDIT = Round(NumericVal(J_NET), 2)
            TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
            Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
            Else
                'COST OF SALES
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If CSMIOS_TERM = "CSH" Then
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                        End If
                    ElseIf COMPANY_CODE = "HCC" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    End If
                Else
                    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                        End If
                    ElseIf COMPANY_CODE = "HCC" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    End If
                End If
                J_DEBIT = Round(WARRANTY_DIRECT_EXPENSE_LABOR_COST, 2)
                J_CREDIT = 0
                ALL_DEBIT = ALL_DEBIT + J_DEBIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                'INVENTORY
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If CSMIOS_TERM = "CSH" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("IN-PROCESS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("IN-PROCESS")))
                End If
                J_DEBIT = 0
                J_CREDIT = Round(WARRANTY_DIRECT_EXPENSE_LABOR_COST, 2)
                ALL_CREDIT = ALL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
            End If
        End If

        'PARTS
        If COMPANY_CODE = "HCC" Then
            If WARRANTY_DIRECT_EXPENSE_GOL > 0 Or WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "RETAIL")))
                J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL), 2)
                If CSMIOS_VAT_EXEMPT = True Then
                    J_TAX = 0
                Else
                    J_TAX = Round(Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333, 2), 2)
                End If
                J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) - Round(NumericVal(J_TAX), 2), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                'COST OF SALES
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                J_DEBIT = Round(WARRANTY_CSMIOS_PARTS_COST + WARRANTY_CSMIOS_MATERIALS_COST, 2)
                J_CREDIT = 0
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

                'INVENTORY
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                J_DEBIT = 0
                J_CREDIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)


            End If
        Else
            If WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
                If COMPANY_CODE = "HPI" Then
                Else
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "WARRANTY")))
                    End If
                    J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS), 2)
                    If CSMIOS_VAT_EXEMPT = True Then
                        J_TAX = 0
                    Else
                        J_TAX = Round(Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) / 9.3333, 2), 2)
                    End If
                    J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_SPAREPARTS) - Round(NumericVal(J_TAX), 2), 2)
                    J_DEBIT = 0
                    J_CREDIT = Round(NumericVal(J_NET), 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If

                If COMPANY_CODE = "HPI" Then
                Else
                    'COST OF SALES
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                    Else
                        If CSMIOS_TERM = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "WARRANTY")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "WARRANTY")))
                        End If
                    End If
                    J_DEBIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                    J_CREDIT = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If
                'INVENTORY
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If CSMIOS_TERM = "CSH" Then
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    ElseIf COMPANY_CODE = "HPI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                    ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                Else
                    'COMMENTED DUE PROCEDURE TO LARGE
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                    End If
                End If
                J_DEBIT = 0
                J_CREDIT = Round(WARRANTY_CSMIOS_PARTS_COST, 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)


            End If
        End If


        'Materials
        If COMPANY_CODE = "HCC" Then
        Else
            If WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                    J_ACCT_CODE = "'41-02013-20'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'41-02013-20'"))
                ElseIf COMPANY_CODE = "HSB" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
                ElseIf COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "WARRANTY")))
                End If

                J_GROSS = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL), 2)
                If CSMIOS_VAT_EXEMPT = True Then
                    J_TAX = 0
                Else
                    J_TAX = Round(Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333, 2), 2)
                End If

                J_NET = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_GOL) - Round(NumericVal(J_TAX), 2), 2)
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(J_NET), 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                If COMPANY_CODE = "HPI" Then
                Else
                    'COST OF SALES
                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS")))
                    ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
                        J_ACCT_CODE = "'61-02011-20'"
                        J_ACCT_NAME = N2Str2Null(Setacctname("'61-02011-20'"))
                    ElseIf COMPANY_CODE = "HSB" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "RETAIL")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS", "WARRANTY")))
                    End If
                    J_GROSS = 0
                    J_TAX = 0
                    J_NET = 0
                    J_DEBIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                    J_CREDIT = 0
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                End If

                'INVENTORY
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                Else
                    If CSMIOS_TERM = "CSH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                    End If
                End If
                J_GROSS = 0: J_TAX = 0: J_NET = 0
                J_DEBIT = 0
                J_CREDIT = Round(WARRANTY_CSMIOS_MATERIALS_COST, 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            End If
        End If
        'OUTPUT TAX
        If CSMIOS_VAT_EXEMPT = False Then
            If NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) > 0 Then
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                'Commented Due to procedure to Large
                'If COMPANY_CODE = "HBK" Then
                '   J_ACCT_CODE = N2Str2Null(ReturnDeferredOutPutTax())
                '   J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDeferredOutPutTax()))
                'Else
                J_ACCT_CODE = N2Str2Null(ReturnOutputTax())
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnOutputTax()))
                'End If
                J_DEBIT = 0
                J_CREDIT = Round(NumericVal(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL) / 9.3333), 2)
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                If Round(Round(TOTAL_CREDIT, 2) - Round(TOTAL_DEBIT + WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) - 0.01
                If Round(Round(TOTAL_DEBIT + WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL, 2) - Round(TOTAL_CREDIT, 2), 2) = 0.01 Then J_CREDIT = Round(J_CREDIT, 2) + 0.01
                J_TAX = 0: J_GROSS = 0: J_NET = 0
                Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

            End If
        End If

        WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
            If COMPANY_CODE = "HAS" Then
                J_ACCT_CODE = N2Str2Null("11-02000-00")
                J_ACCT_NAME = N2Str2Null(Setacctname("11-02000-00"))
            ElseIf COMPANY_CODE = "HPI" Then
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
            ElseIf COMPANY_CODE = "HCC" Then
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
            Else
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
            End If
        Else
            If COMPANY_CODE = "HCC" Then
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SERVICE"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SERVICE")))
            Else
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
            End If
        End If
        J_DEBIT = NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL)
        J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

        CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
        'To return the Selling Dealer
        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
            CSMIOS_NIYM = Null2String(SetSellingDealerName(ReturnCodeSellingDealer(ReturnPlateNo(CSMIOS_REP_OR))))
            CSMIOS_ACCT_NO = N2Str2Null(ReturnCode(CSMIOS_NIYM))
            J_CUSTOMERCODE = N2Str2Null(ReturnCode(CSMIOS_NIYM))
        Else
            If COMPANY_CODE = "HPI" Then
                CSMIOS_ACCT_NO = Null2String("H00416")
                CSMIOS_NIYM = Null2String(SetCustomerName("H00416"))
                J_CUSTOMERCODE = N2Str2Null("H00416")
            Else
                CSMIOS_ACCT_NO = Null2String("H00001")
                CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                J_CUSTOMERCODE = N2Str2Null("H00001")
            End If
        End If
        WARRANTY_J_AMOUNTTOPAY = 0
        WARRANTY_J_INVOICEAMT = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL), 2)
        WARRANTY_J_BALANCE = Round(NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL), 2)
        WARRANTY_J_AMOUNTPAID = 0

        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                        " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                        " values (" & J_JDATE & ", " & WARRANTY_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & WARRANTY_J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & WARRANTY_J_AMOUNTTOPAY & "," & WARRANTY_J_BALANCE & "," & WARRANTY_J_AMOUNTPAID & _
                        ", " & WARRANTY_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
        gconDMIS.Execute SQL_STATEMENT
        TransactionID = (FindTransactionID(N2Str2Null(WARRANTY_VOUCHERNO), "voucherno", "AMIS_journal_hd", "X", J_JTYPE, "Jtype"))
        NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, TransactionID, "", WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO

        ':-)=================================
    End If
End Sub

Function GetCustomerORAmount(XXX As String) As Double
    Dim rsGetCustomerORAmount                          As ADODB.Recordset
    Set rsGetCustomerORAmount = New ADODB.Recordset
    Dim ORDATE                                         As String
    Dim ORAMOUNT                                       As Double
    If Len(XXX) > 6 Then
        rsGetCustomerORAmount.Open "SELECT PAYMENT FROM CMIS_OFF_DT WHERE INVOICENO='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORAMOUNT = NumericVal(rsGetCustomerORAmount!payment)
    Else
        rsGetCustomerORAmount.Open "SELECT OR_AMT FROM CMIS_OFF_HD WHERE OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORAMOUNT = NumericVal(rsGetCustomerORAmount!OR_AMT)
    End If
    If Not rsGetCustomerORAmount.EOF And Not rsGetCustomerORAmount.BOF Then
        GetCustomerORAmount = ORAMOUNT
    End If
    Set rsGetCustomerORAmount = Nothing
End Function

Function GetCustomerORDate(XXX As String) As String
    Dim rsGetCustomerORDate                            As ADODB.Recordset
    Set rsGetCustomerORDate = New ADODB.Recordset
    Dim ORDATE                                         As String
    If Len(XXX) > 6 Then
        rsGetCustomerORDate.Open "SELECT ORDATE FROM CMIS_OFF_DT WHERE INVOICENO='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORDATE = N2Date2Null(rsGetCustomerORDate!ORDATE)
    Else
        rsGetCustomerORDate.Open "SELECT OR_DATE FROM CMIS_OFF_HD WHERE OR_NUM='" & XXX & "'", gconDMIS, adOpenForwardOnly
        ORDATE = N2Date2Null(rsGetCustomerORDate!OR_DATE)
    End If
    If Not rsGetCustomerORDate.EOF And Not rsGetCustomerORDate.BOF Then
        GetCustomerORDate = ORDATE
    End If
    Set rsGetCustomerORAmount = Nothing
End Function

Function CreditCardPayment(XXX As String, YYY As String, zzz As String) As Double
    Dim xACCT_CODE                                     As String
    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HGH" Then
        xACCT_CODE = "CARD"
    Else
        xACCT_CODE = "CARD ON HAND"
    End If
    Dim rsAMOUNT                                       As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND ACCT_CODE = " & N2Str2Null(ReturnAccountCode(xACCT_CODE)) & "", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        CreditCardPayment = NumericVal(N2Str2Zero(rsAMOUNT!CREDIT))
    End If
    Set rsAMOUNT = Nothing
End Function

Sub ImportUndeposited2(xTRANTYPE As String, XTYPE As String, XTRANNO As String, xTranDate As Date)
    Dim CMD                                            As New ADODB.Command

    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_IMPORTING_TEMPLATE"
    CMD.CommandTimeout = 500

    With CMD.Parameters
        .Append CMD.CreateParameter("@TRANTYPE", adVarChar, adParamInput, 15, xTRANTYPE)
        .Append CMD.CreateParameter("@TYPE", adVarChar, adParamInput, 15, XTYPE)
        .Append CMD.CreateParameter("@TRANNO", adVarChar, adParamInput, 10, XTRANNO)
        .Append CMD.CreateParameter("@TRANDATE", adDate, adParamInput, 8, xTranDate)
    End With

    CMD.Execute
End Sub

Sub ImportSalesNew()
    Dim i                                              As Long
    'PARTS / ACCESSORIES / MATERIALS IMPORTING =============================================================
    i = 0
    Dim GridImport                                     As Integer
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            If Grid3.Cell(GridImport, 6).Text = "PARTS" Or Grid3.Cell(GridImport, 6).Text = "ACCESSORIES" Or Grid3.Cell(GridImport, 6).Text = "MATERIALS" Then
                Call ImportSales2("SALES", Grid3.Cell(GridImport, 2).Text, Right(Grid3.Cell(GridImport, 3).Text, 6), CDate(dtpTranDate))
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If

        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    'SERVICE TRANSACTION
    '=========================================================================================================
    '    i = 0
    '    For GridImport = 1 To Grid3.Rows - 1
    '        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
    '            If Left(Grid3.Cell(GridImport, 2).Text, 1) = "R" Then
    '                Call ImportSales2("SALES", "SERVICE", Grid3.Cell(GridImport, 2).Text, CDate(dtpTranDate))
    '                Grid3.Cell(GridImport, 1).Text = 1
    '            End If
    '        End If
    '        i = i + 1
    '        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
    '        labCPB.Caption = Int(progCPB.Value) & "% Completed"
    '        DoEvents
    '    Next

    'VEHICLE SALES TRANSACTION
    '=========================================================================================================
    i = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            If Grid3.Cell(GridImport, 6).Text = "VEHICLE" Then
                Call ImportSales2("SALES", "VEHICLES", Right(Grid3.Cell(GridImport, 3).Text, 6), CDate(dtpTranDate))
                Grid3.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
End Sub

Sub ImportSales2(xTRANTYPE As String, XTYPE As String, XTRANNO As String, xTranDate As Date)
    Dim CMD                                            As New ADODB.Command

    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_IMPORTING_TEMPLATE"
    CMD.CommandTimeout = 500

    With CMD.Parameters
        .Append CMD.CreateParameter("@TRANTYPE", adVarChar, adParamInput, 15, xTRANTYPE)
        .Append CMD.CreateParameter("@TYPE", adVarChar, adParamInput, 15, XTYPE)
        .Append CMD.CreateParameter("@TRANNO", adVarChar, adParamInput, 10, XTRANNO)
        .Append CMD.CreateParameter("@TRANDATE", adDate, adParamInput, 8, xTranDate)
        .Append CMD.CreateParameter("@VOUCHERNO", adVarChar, adParamOutput, 6)
        .Append CMD.CreateParameter("@JTYPE", adVarChar, adParamOutput, 3)
        .Append CMD.CreateParameter("@JNO", adVarChar, adParamOutput, 6)
    End With
    CMD.Execute
    J_VOUCHERNO = CMD("@VOUCHERNO")
    J_JTYPE = CMD("@JTYPE")
    J_JNO = CMD("@JNO")
    TransactionID = FindNewID(N2Str2Null(J_VOUCHERNO), "VOUCHERNO", "AMIS_JOURNAL_HD", N2Str2Null(J_JTYPE), "JTYPE")
    SQL_STATEMENT = "INSERT INTO AMIS_JOURNAL_HD " & J_JTYPE & "-" & J_VOUCHERNO
    NEW_LogAudit "R", "SALES JOURNAL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
End Sub


Function AmountLessVAT(XXX As String, YYY As String) As Double
    Dim rsAMOUNT                                       As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT SUM(ISNULL(CREDIT,0)) AS CREDIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND (LEFT(ACCT_CODE,2)='41' OR LEFT(ACCT_CODE,5)='71-61')", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        AmountLessVAT = NumericVal(rsAMOUNT!CREDIT)
    End If
    Set rsAMOUNT = Nothing
End Function

Function DiscountLessVAT(XXX As String, YYY As String) As Double
    Dim rsAMOUNT                                       As ADODB.Recordset
    Set rsAMOUNT = New ADODB.Recordset
    rsAMOUNT.Open "SELECT SUM(ISNULL(DEBIT,0)) AS DEBIT FROM AMIS_JOURNAL_DET WHERE VOUCHERNO =" & XXX & " AND JTYPE= " & YYY & " AND LEFT(ACCT_CODE,1)='5'", gconDMIS, adOpenForwardOnly
    If Not rsAMOUNT.EOF And Not rsAMOUNT.BOF Then
        DiscountLessVAT = N2Str2Zero(rsAMOUNT!DEBIT)
    End If
    Set rsAMOUNT = Nothing
End Function


