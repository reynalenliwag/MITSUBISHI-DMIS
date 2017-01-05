VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmSALESImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES Import Process"
   ClientHeight    =   7800
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   13980
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13980
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
      Left            =   13185
      MouseIcon       =   "SALESImport.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "SALESImport.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Exit Window"
      Top             =   6810
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
      Left            =   12480
      MouseIcon       =   "SALESImport.frx":04B8
      MousePointer    =   99  'Custom
      Picture         =   "SALESImport.frx":060A
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Process Importing of Cash Receipts "
      Top             =   6810
      Width           =   720
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3120
      ScaleHeight     =   465
      ScaleWidth      =   4905
      TabIndex        =   21
      Top             =   7380
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
         TabIndex        =   27
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
         MousePointer    =   99  'Custom
         TabIndex        =   26
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
         TabIndex        =   25
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
         MousePointer    =   99  'Custom
         TabIndex        =   24
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
         TabIndex        =   23
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
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   90
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   -90
      ScaleHeight     =   855
      ScaleWidth      =   3195
      TabIndex        =   16
      Top             =   6690
      Width           =   3195
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
         Picture         =   "SALESImport.frx":08A5
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   30
         Width           =   915
      End
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
         TabIndex        =   18
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
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   30
         Width           =   1455
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   120
         Width           =   585
      End
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   1200
      TabIndex        =   15
      Top             =   7980
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
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
      Left            =   3600
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Import of SALES"
      Top             =   120
      Width           =   2040
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
      Top             =   120
      Width           =   1935
   End
   Begin FlexCell.Grid Grid1 
      Height          =   4905
      Left            =   90
      TabIndex        =   2
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   8040
      TabIndex        =   3
      Top             =   6300
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "SALESImport.frx":1927
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarPicture      =   "SALESImport.frx":1943
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
      TabIndex        =   4
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
      Format          =   63045633
      CurrentDate     =   38216
   End
   Begin wizProgBar.Prg Prg1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   6300
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   556
      Picture         =   "SALESImport.frx":195F
      ForeColor       =   255
      Appearance      =   2
      BorderStyle     =   2
      BarColor        =   4194304
      BarPicture      =   "SALESImport.frx":197B
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
   Begin FlexCell.Grid Grid2 
      Height          =   4905
      Left            =   4740
      TabIndex        =   6
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin FlexCell.Grid Grid3 
      Height          =   4905
      Left            =   9390
      TabIndex        =   7
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8652
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      ReadOnlyFocusRect=   0
      Rows            =   2
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   3090
      ScaleHeight     =   765
      ScaleWidth      =   5805
      TabIndex        =   28
      Top             =   6720
      Width           =   5805
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   60
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
         Index           =   1
         Left            =   420
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
         Index           =   2
         Left            =   780
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
         Index           =   3
         Left            =   1140
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
         Index           =   4
         Left            =   1500
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
         Index           =   5
         Left            =   1860
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
         Index           =   6
         Left            =   2220
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
         Index           =   7
         Left            =   2580
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
         Index           =   8
         Left            =   2940
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
         Index           =   9
         Left            =   3300
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
         Index           =   10
         Left            =   3660
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
         Index           =   11
         Left            =   4020
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
         Index           =   12
         Left            =   4380
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
         Index           =   13
         Left            =   4740
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
         Index           =   14
         Left            =   5100
         MousePointer    =   99  'Custom
         TabIndex        =   45
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
         Index           =   16
         Left            =   420
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
         Index           =   17
         Left            =   780
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
         Index           =   18
         Left            =   1140
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
         Index           =   19
         Left            =   1500
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
         Index           =   20
         Left            =   1860
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
         Index           =   21
         Left            =   2220
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
         Index           =   22
         Left            =   2580
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
         Index           =   23
         Left            =   2940
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
         Index           =   24
         Left            =   3300
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
         Index           =   25
         Left            =   3660
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
         Index           =   26
         Left            =   4020
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
         Index           =   27
         Left            =   4380
         MousePointer    =   99  'Custom
         TabIndex        =   32
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
         Index           =   29
         Left            =   5100
         MousePointer    =   99  'Custom
         TabIndex        =   30
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
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   390
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack2 
      Height          =   3765
      Left            =   120
      TabIndex        =   63
      Top             =   1680
      Visible         =   0   'False
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.Label labCPB 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Completion"
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
      Left            =   8070
      TabIndex        =   14
      Top             =   6030
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
      TabIndex        =   13
      Top             =   210
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Over All Completion"
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
      TabIndex        =   12
      Top             =   6030
      Width           =   5835
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported                          Invoices can be Imported"
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
      Height          =   675
      Left            =   8970
      TabIndex        =   11
      Top             =   7140
      Width           =   3465
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PMIS"
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
      Left            =   120
      TabIndex        =   10
      Top             =   630
      Width           =   4515
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CSMS"
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
      Left            =   4740
      TabIndex        =   9
      Top             =   630
      Width           =   4545
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SMIS"
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
      Left            =   9390
      TabIndex        =   8
      Top             =   630
      Width           =   4545
   End
End
Attribute VB_Name = "frmSALESImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TransactionID                                 As String
Dim rsPMIOS_ORD_HD                                As ADODB.Recordset
Dim rsTdayTran                                    As ADODB.Recordset
Dim rsCSMIOS_REPOR                                As ADODB.Recordset
Dim rsCSMIOS_SUBLET                               As ADODB.Recordset
Dim rsCSMIOS_PMS                                  As ADODB.Recordset
Dim rsCSMIOS_LABOR                                As ADODB.Recordset
Dim rsCSMIOS_PARTS                                As ADODB.Recordset
Dim rsCSMIOS_MATERIALS                            As ADODB.Recordset
Dim rsSMIS_PURCHAGREE                             As ADODB.Recordset
Dim rsVENDOR                                      As New ADODB.Recordset
Dim rsCSMIOS_ACCESSORIES                          As New ADODB.Recordset
Dim rsCSMIOS_TINSMITH                             As New ADODB.Recordset
Dim rsINTERNAL_RO_DET                             As ADODB.Recordset
Dim rsJournal_HDDup                               As ADODB.Recordset
'DISCOUNT
Dim CSMIOS_PMS_DISCOUNT                           As Double
'WARRANTY
Dim WARRANTY_CSMIOS_PARTS_COST                    As Double
Dim WARRANTY_CSMIOS_MATERIALS_COST                As Double
Dim WARRANTY_CSMIOS_ACCESSORIES_COST              As Double
Dim WARRANTY_JNO                                  As String
Dim WARRANTY_VOUCHERNO                            As String
Dim WARRANTY_ItemCnt                              As Integer
Dim WARRANTY_J_JITEMNO                            As String
Dim WARRANTY_DIRECT_EXPENSE_LABOR_COST            As Double
Dim WARRANTY_J_AMOUNTTOPAY                        As Double
Dim WARRANTY_J_INVOICEAMT                         As Double
Dim WARRANTY_J_BALANCE                            As Double
Dim WARRANTY_J_AMOUNTPAID                         As Double
'ACCESSOREIS
Dim CSMIOS_ACCESSORIES_DISCOUNT                   As Double
Dim CSMIOS_REP_OR                                 As String
Dim CSMIOS_ACCT_NO                                As String
Dim CSMIOS_PARTICIPAT                             As String
Dim CSMIOS_PLATE_NO                               As String
Dim CSMIOS_NIYM                                   As String
Dim CSMIOS_TERM                                   As String
Dim CSMIOS_DTE_REL                                As String
Dim CSMIOS_INVOICE                                As String
Dim CSMIOS_VAT_EXEMPT                             As Boolean
Dim CSMIOS_RO_AMOUNT                              As Double
Dim CSMIOS_LABOR                                  As Double
Dim CSMIOS_PARTS                                  As Double
Dim CSMIOS_MATERIALS                              As Double
Dim CSMIOS_ACCESSORIES                            As Double
Dim CSMIOS_LABOR_COST                             As Double
Dim CSMIOS_PARTS_COST                             As Double
Dim CSMIOS_MATERIALS_COST                         As Double
Dim CSMIOS_ACCESSORIES_COST                       As Double
Dim CSMS_ACCCOST                                  As Double
Dim CSMIOS_PMS_COST                               As Double
'DIM CSMIOS_RO_AMOUNT                                  AS DOUBLE
Dim CSMIOS_TINSPAINT                              As Double
Dim CSMIOS_SUBLET                                 As Double
Dim CSMIOS_AIRCON                                 As Double
Dim CSMIOS_TINSPAINT_DISCOUNT                     As Double
Dim CSMIOS_SUBLET_DISCOUNT                        As Double
Dim CSMIOS_LABOR_DISCOUNT                         As Double
Dim CSMIOS_PARTS_DISCOUNT                         As Double
Dim CSMIOS_MATERIALS_DISCOUNT                     As Double
Dim CSMIOS_SUBLET_COST                            As Double
Dim CSMIOS_TINSPAINT_COST                         As Double
Dim WARRANTY_DIRECT_EXPENSE_LABOR                 As Double
Dim WARRANTY_DIRECT_EXPENSE_SPAREPARTS            As Double
Dim WARRANTY_DIRECT_EXPENSE_GOL                   As Double
Dim COMPANY_DIRECT_EXPENSE_LABOR                  As Double
Dim COMPANY_DIRECT_EXPENSE_SPAREPARTS             As Double
Dim COMPANY_DIRECT_EXPENSE_GOL                    As Double
Dim SALES_DIRECT_EXPENSE_LABOR                    As Double
Dim SALES_DIRECT_EXPENSE_SPAREPARTS               As Double
Dim SALES_DIRECT_EXPENSE_GOL                      As Double
Dim INSURANCE_DIRECT_EXPENSE_LABOR                As Double
Dim INSURANCE_DIRECT_EXPENSE_SPAREPARTS           As Double
Dim INSURANCE_DIRECT_EXPENSE_ACCESSORIES          As Double
Dim INSURANCE_DIRECT_EXPENSE_GOL                  As Double
Dim CSMIOS_PMS                                    As Double
Dim TOTAL_INSURANCE_AMOUNT                        As Double
Dim TOTAL_DISCOUNT_AMOUNT                         As Double
Dim ALL_DEBIT                                     As Double
Dim ALL_CREDIT                                    As Double
'INTERNAL
Dim INTERNAL_LABOR_AMT                            As Double
Dim INTERNAL_PARTS_AMT                            As Double
Dim INTERNAL_MATERIALS_AMT                        As Double
Dim INTERNAL_LABOR_COST                           As Double
Dim INTERNAL_PARTS_COST                           As Double
Dim INTERNAL_PARTS_INV                            As Double
Dim INTERNAL_MATERIALS_COST                       As Double
Dim J_ACCT_CODE                                   As String
Dim J_ACCT_NAME                                   As String
Dim J_GROSS                                       As Double
Dim J_TAX                                         As Double
Dim J_NET                                         As Double
Dim J_DEBIT                                       As Double
Dim J_CREDIT                                      As Double
Dim TOTAL_DEBIT                                   As Double
Dim TOTAL_CREDIT                                  As Double
Dim ItemCnt                                       As Integer
Dim J_JDATE                                       As String
Dim J_VOUCHERNO                                   As String
Dim J_JTYPE                                       As String
Dim J_JNO                                         As String
Dim J_REMARKS                                     As String
Dim J_VENDORCODE                                  As String
Dim J_CUSTOMERCODE                                As String
Dim J_OUTBALANCE                                  As Double
Dim J_AMOUNTTOPAY                                 As Double
Dim J_INVOICEAMT                                  As Double
Dim J_BALANCE                                     As Double
Dim J_AMOUNTPAID                                  As Double
Dim J_CHECKNO                                     As String
Dim J_INVOICEDATE                                 As String
Dim J_DUEDATE                                     As String
Dim J_PAYTYPE                                     As String
Dim J_INVOICETYPE                                 As String
Dim J_INVOICENO                                   As String
Dim J_CHECKDATE                                   As String
Dim J_BANKCODE                                    As String
Dim J_REFNO                                       As String
Dim J_REFDATE                                     As String
Dim J_TERMS                                       As String
Dim J_DEALER                                      As String
Dim J_PAIDSTATUS                                  As String
Dim J_RECEIVESTATUS                               As String
Dim J_STATUS                                      As String
Dim J_JITEMNO                                     As String
Dim TRANSACTIONDATE                               As String
Dim Indx                                          As Integer
Dim J_ITEMCNT                                     As Integer
Dim xTranDate                                     As String
Dim SubletType                                    As String
Dim DetCodeLen                                    As Integer


Function SetCustomerName(VVV As Variant) As String
    Dim rsCustomer2                               As ADODB.Recordset
    Set rsCustomer2 = New ADODB.Recordset
    rsCustomer2.Open "Select CustCode,acctname from ALL_CUSTMASTER_AMIS where CustCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer2.EOF And Not rsCustomer2.BOF Then
        SetCustomerName = UCase(Null2String(rsCustomer2!AcctName))
    Else
        SetCustomerName = ""
    End If
End Function

Function Setacctname(VVV As String) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    If Left(VVV, 1) = "'" Then
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & VVV, gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = '" & VVV & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function GetVoucherNo() As String
    Dim rsJournal_HD                              As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'SJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function GetJNo() As String
    Dim rsJournal_HD                              As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(JNo AS int) AS MAX_JNO from AMIS_Journal_HD Order by MAX_JNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetJNo = Format(NumericVal(rsJournal_HD!MAX_JNO) + 1, "000000")
    Else
        GetJNo = "000001"
    End If
End Function

Function SetRONiym(XXX As String)
    Dim rsRO_Niym                                 As ADODB.Recordset
    Set rsRO_Niym = New ADODB.Recordset
    Set rsRO_Niym = gconDMIS.Execute("Select NIYM from CSMS_REPOR WHERE INVOICE = '" & XXX & "'")
    If Not rsRO_Niym.EOF And Not rsRO_Niym.BOF Then
        SetRONiym = Null2String(rsRO_Niym!Niym)
    End If
End Function

'UPDATED: ACL 06042010 DUE TO CASH/CHARGE TRANSACTIONS
Function CheckSJExisting(VarInvoiceType As String, VarInvoiceNo As String, Optional VarTranType As String) As Boolean
    Dim rsCheckSJ_Journal_HD                      As ADODB.Recordset
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
    Dim rsCheckSJ_Journal_HD                      As ADODB.Recordset
    Set rsCheckSJ_Journal_HD = New ADODB.Recordset
    Set rsCheckSJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'SJ' AND Status <> 'C' AND InvoiceType = " & N2Str2Null(VarInvoiceType) & " AND RefNo = " & N2Str2Null(VarInvoiceNo))
    If Not rsCheckSJ_Journal_HD.EOF And Not rsCheckSJ_Journal_HD.BOF Then
        CheckRefNoExisting = True
    Else
        CheckRefNoExisting = False
    End If
    Set rsCheckSJ_Journal_HD = Nothing
End Function

Function ReturnClearing_AccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'CLEARING' AND TRANTYPE1 = '" & Trim(XXX) & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnClearing_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAR_AccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AR' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAR_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnSales_AccountCode(INVTYPE As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
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

Function ReturnDiscount_AccountCode(INVTYPE As String, Optional OTHERTYPE As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
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
Function ReturnCostOFSALES_PARTS(STR_TRANNO As String, STR_TRANTYPE As String, STR_TYPE As String) As Double
    Dim RSCOST                                    As ADODB.Recordset
    'UPDATED BY: ACL 03092010
    'TO GET TOTAL COST NOT THE ITEM COST ONLY
    'Set RSCOST = gconDMIS.Execute("SELECT ISNULL(SUM(TRANUCOST),0) AS TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANNO=" & N2Str2Null(STR_TRANNO) & " AND TRANTYPE=" & N2Str2Null(STR_TRANTYPE))
    Set RSCOST = gconDMIS.Execute("SELECT ISNULL(TRANUCOST,0) * TRANQTY AS TRANUCOST FROM PMIS_ALLDAYTRAN WHERE TRANNO=" & N2Str2Null(STR_TRANNO) & " AND TRANTYPE=" & N2Str2Null(STR_TRANTYPE) & " AND TYPE=" & N2Str2Null(STR_TYPE))
    If Not RSCOST.EOF Or Not RSCOST.BOF Then
        Do While Not RSCOST.EOF
            ReturnCostOFSALES_PARTS = ReturnCostOFSALES_PARTS + Round(NumericVal(RSCOST!TRANUCOST), 2)
            RSCOST.MoveNext
        Loop
    End If
    Set RSCOST = Nothing
End Function

Function ReturnCostofSales(INVTYPE As String, Optional OTHERTYPE As String, Optional NEXTTYPE As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
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
    Dim rsChartAccount                            As ADODB.Recordset
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
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnOutputTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnDeferredOutPutTax()
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'DEFERRED OUTPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnDeferredOutPutTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInternalAccountCode(XXX As String)
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select ChartCodes from CMIS_SBOOK where BOOK = 'S' and CODE = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInternalAccountCode = Null2String(rsChartAccount!CHARTCODES)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAccountCode(XXX As String)
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                          As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
        If Null2String(rsSupplierMaster!NONVAT) = "Y" Then
            CheckIfNonVatSup = True
        Else
            CheckIfNonVatSup = False
        End If
    Else
        CheckIfNonVatSup = False
    End If
End Function

Function ReturnPlateNo(XXX As String) As String
' Update By BTT - 07092008
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

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
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

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
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

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
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    Dim MARK                                      As String
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
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

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

Function CheckIfPMS(XXX As String) As Boolean
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "select Status1 from CSMS_ro_det where jobtype='PMS' and livil='1' and wcode='W' and Rep_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        CheckIfPMS = True
    Else
        CheckIfPMS = False
    End If
    Set RS = Nothing
End Function

Function ReturnShopSuplies(XXX As String) As Integer
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT DETAMT,detprc FROM CSMS_RO_DET WHERE LIVIL = '3' AND DETCDE = 'SVCMAT0068' and REP_or='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        ReturnShopSuplies = NumericVal(RS!detprc)
    Else
        ReturnShopSuplies = 0
    End If
    Set RS = Nothing
End Function

Function PosibleDoubleInternal(XXX As String) As Boolean
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT * FROM AMIS_journal_hd where refno='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        PosibleDoubleInternal = True
    Else
        PosibleDoubleInternal = False
    End If
    Set RS = Nothing
End Function

Sub InsertToJournalDet(vJ_JDATE As Variant, vJ_VOUCHERNO As Variant, vJ_JTYPE As Variant, vJ_JNO As Variant, vJ_JITEMNO As Variant, vJ_ACCT_CODE As Variant, vJ_ACCT_NAME As Variant, vJ_DEBIT As Variant, vJ_CREDIT As Variant, vJ_TAX As Variant, vJ_GROSS As Variant, vJ_NET As Variant, vJ_STATUS As Variant)
'UPDATE DUE TO NEW AUDIT : BTT 08292008
    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                    " values (" & vJ_JDATE & ", " & vJ_VOUCHERNO & ", " & vJ_JTYPE & ", " & vJ_JNO & ", " & vJ_JITEMNO & ", " & vJ_ACCT_CODE & ", " & vJ_ACCT_NAME & ", " & vJ_DEBIT & ", " & vJ_CREDIT & ", " & vJ_TAX & "," & vJ_GROSS & "," & vJ_NET & ", " & vJ_STATUS & ")"
    gconDMIS.Execute SQL_STATEMENT
    TransactionID = (FindTransactionID(N2Str2Null(vJ_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(vJ_VOUCHERNO), J_JTYPE, J_JNO
End Sub

Function ImportPurelyInternal() As Boolean
    On Error GoTo ErrorCode

    Dim GridImport                                As Integer
    Dim i                                         As Integer
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset
            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where REP_OR = '" & Grid2.Cell(GridImport, 2).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                ItemCnt = 0
                'If Null2String(rsCSMIOS_REPOR!REP_OR) = "R-00007017" Then Stop
                CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!invoice)
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

                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
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
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
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
                            If COMPANY_CODE = "HAI" Then
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

                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Then
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
                        If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
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
                                INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                                    J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                                    J_TAX = 0
                                    If COMPANY_CODE = "HMH" Then
                                        J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol)), 2)
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
                                If COMPANY_CODE = "HAI" Then
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
                                If COMPANY_CODE = "HAI" Then
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
                            If COMPANY_CODE = "HAI" Then
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
                                INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))
                                If COMPANY_CODE = "HPI" Then
                                Else
                                    WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                    WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                    J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
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
                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                            Else
                                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                                If COMPANY_CODE = "HAI" Then
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
                                If COMPANY_CODE = "HAI" Then
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
                            If COMPANY_CODE = "HAI" Then
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
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!invoice)

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
                    Grid2.Cell(GridImport, 1).Text = 1
                End If
                '=========================================================================================================================================================
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    ImportPurelyInternal = True
    Exit Function
ErrorCode:
    ImportPurelyInternal = False
End Function

Sub ProcessPartsCost()
    Dim rsRo_det                                  As ADODB.Recordset
    Dim rsOrd_Hd                                  As ADODB.Recordset
    Dim rsRR_HD                                   As ADODB.Recordset
    Dim rsDAYTRAN                                 As ADODB.Recordset
    Dim rsPartMas                                 As ADODB.Recordset
    Dim i                                         As Integer

    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double

    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'P' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        i = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'P', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'P', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!TRANTYPE)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2Zero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'P' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                            vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            i = i + 1
            progCPB.Value = (i / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If

    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '2' and CSMS_RO_DET.DETCOST <=0 AND  CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        i = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'P' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'P' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!DETCDE) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                             " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            i = i + 1
            progCPB.Value = (i / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Sub ProcessMaterialsCost()
    Dim rsRo_det                                  As ADODB.Recordset
    Dim rsOrd_Hd                                  As ADODB.Recordset
    Dim rsRR_HD                                   As ADODB.Recordset
    Dim rsDAYTRAN                                 As ADODB.Recordset
    Dim rsPartMas                                 As ADODB.Recordset
    Dim i                                         As Integer

    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode, vSTOCKDESC As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double

    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'M' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') and status <> 'N' AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        i = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'M', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'M', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!TRANTYPE)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2IntZero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            If rsPartMas.EOF And rsPartMas.BOF Then
                vSTOCKDESC = "'NO DESCRIPTION'"
                'gconDMIS.Execute ("Insert into PMIS_STOCKMAS ([TYPE],STOCKNO,STOCKDESC,date_entered) values ('M'," & N2Str2Null(rsTdayTran!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTdayTran!trandate) & ")")
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = 'M' WHERE STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_STOCKMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'M' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                            vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            i = i + 1
            progCPB.Value = (i / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If


    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '3' AND CSMS_RO_DET.DETCOST <=0 and CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        i = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'M' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'M' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!DETCDE) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                             " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            i = i + 1
            progCPB.Value = (i / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Sub ProcessINITIALIZE_ACCESSORIESCOST()
    Dim rsRo_det                                  As ADODB.Recordset
    Dim rsOrd_Hd                                  As ADODB.Recordset
    Dim rsRR_HD                                   As ADODB.Recordset
    Dim rsDAYTRAN                                 As ADODB.Recordset
    Dim rsPartMas                                 As ADODB.Recordset
    Dim i                                         As Integer

    Dim vTDTranDate, vTDTranType, vTDTranno, vSupCode, vSTOCKDESC As String
    Dim vTDTranQTY, vTotTranCost, vTotTranInvAmt, vMAC, vPMOnhand, vVatAmt As Double

    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,ItemNo,trantype,tranno,STOCK_ORD,tranqty,status,in_out,tranucost,traninvamt,trandate from PMIS_AllDayTran where tranucost <= 0 AND [TYPE] = 'A' AND trantype <> 'ADB' and (status = 'P' OR status = 'B') and status <> 'N' AND trandate = " & N2Date2Null(dtpTranDate) & " order by id asc", gconDMIS
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        rsTdayTran.MoveFirst
        Screen.MousePointer = 11
        i = 0
        Do While Not rsTdayTran.EOF
            gconDMIS.Execute "update PMIS_DayTran set TYPE = 'A', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            gconDMIS.Execute "update PMIS_TDayTran set TYPE = 'A', ItemNo = '" & Format(Null2String(rsTdayTran!ItemNo), "0000") & "' where ID = " & rsTdayTran!ID
            vTDTranDate = N2Date2Null(rsTdayTran!trandate)
            vTDTranType = Null2String(rsTdayTran!TRANTYPE)
            vTDTranno = Null2String(rsTdayTran!TRANNO)
            vTDTranQTY = N2Str2IntZero(rsTdayTran!tranqty)
            If N2Str2Zero(rsTdayTran!TRANUCOST) > 0 Then
                vTotTranCost = N2Str2Zero(rsTdayTran!TRANUCOST) * vTDTranQTY
            Else
                vTotTranCost = 0
            End If
            If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                vTotTranInvAmt = N2Str2Zero(rsTdayTran!TRANINVAMT) * vTDTranQTY
            Else
                vTotTranInvAmt = 0
            End If
            Set rsPartMas = New ADODB.Recordset
            Set rsPartMas = gconDMIS.Execute("select STOCKNO from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            If rsPartMas.EOF And rsPartMas.BOF Then
                vSTOCKDESC = "'NO DESCRIPTION'"
                'gconDMIS.Execute ("Insert into PMIS_STOCKMAS ([TYPE],STOCKNO,STOCKDESC,date_entered) values ('A'," & N2Str2Null(rsTdayTran!STOCK_ORD) & "," & vSTOCKDESC & "," & N2Str2Null(rsTdayTran!trandate) & ")")
            Else
                gconDMIS.Execute ("Update PMIS_STOCKMAS SET ACTIVE = 'Y', TYPE = 'A' WHERE STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord))
            End If
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "select id,STOCKNO,mac,tissqty,trecqty,poqty,tprqty,prqty,issuances,receipts,onhand,ONREQUEST,REQSERVED,SERVED,ONORDER,ORDERED,tpoqty,S_REQSERVED,S_ONREQUEST,purchases from PMIS_STOCKMAS where TYPE = 'A' AND STOCKNO = " & N2Str2Null(rsTdayTran!stock_ord), gconDMIS
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                vMAC = N2Str2Zero(rsPartMas!Mac)
                vPMOnhand = N2Str2IntZero(rsPartMas!ONHAND)
                If Null2String(rsTdayTran!IN_OUT) = "O" And vTDTranQTY <> 0 Then
                    gconDMIS.Execute "update PMIS_DayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TDayTran set tranucost = " & vMAC & " where ID = " & rsTdayTran!ID
                End If

                If Null2String(rsTdayTran!IN_OUT) = "I" And vTDTranQTY <> 0 Then
                    Set rsRR_HD = New ADODB.Recordset
                    rsRR_HD.Open "Select id,recvd_code,ds1,status,classcode,rrno from PMIS_RR_HD where [TYPE] = 'A' AND rrno = '" & Format(vTDTranno, "000000") & "' and status <> 'C'", gconDMIS
                    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                        vSupCode = Null2String(rsRR_HD!RECVD_CODE)
                        vVatAmt = N2Str2Zero(rsRR_HD!DS1)
                        If Null2String(rsRR_HD!CLASSCODE) = "PCG" Or Null2String(rsRR_HD!CLASSCODE) = "PCS" Then
                            gconDMIS.Execute "update PMIS_RR_HD set ds1 = " & vVatAmt & " where id = " & rsRR_HD!ID
                            If CheckIfNonVatSup(Null2String(rsRR_HD!RECVD_CODE)) = False Then
                                If vSupCode <> VPAMCOR And vVatAmt <= 0 Then
                                    vTotTranCost = vTotTranCost / ConvertToBIRDecimalFormat(vVatAmt)
                                End If
                            Else
                                vTotTranCost = vTotTranInvAmt
                                If N2Str2Zero(rsTdayTran!TRANINVAMT) > 0 Then
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = " & rsTdayTran!TRANINVAMT & " Where id = " & rsTdayTran!ID)
                                Else
                                    gconDMIS.Execute ("update PMIS_DayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                    gconDMIS.Execute ("update PMIS_TDayTran Set tranucost = 0 Where id = " & rsTdayTran!ID)
                                End If
                                gconDMIS.Execute ("update PMIS_RR_HD Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                                gconDMIS.Execute ("update PMIS_REC_HIST Set DS1 = 0, DS_AMT1 = 0 WHERE ID = " & rsRR_HD!ID)
                            End If
                        End If
                    Else
                        If vTDTranType = "ADJ" Or vTDTranType = "BEG" Then
                            vTotTranCost = vMAC * vTDTranQTY
                        End If
                    End If
                    If vPMOnhand <= 0 Then
                        vMAC = vTotTranCost / vTDTranQTY
                    Else
                        vMAC = ((vMAC * vPMOnhand) + vTotTranCost) / (vTDTranQTY + vPMOnhand)
                    End If
                    gconDMIS.Execute "update PMIS_dayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                    gconDMIS.Execute "update PMIS_TdayTran set tranucost = " & vMAC & ", mac = " & vMAC & ", status = 'P' where id = " & rsTdayTran!ID
                End If
            End If
            DoEvents
            i = i + 1
            progCPB.Value = (i / rsTdayTran.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed"
            DoEvents
            rsTdayTran.MoveNext
        Loop
        DoEvents
        Screen.MousePointer = 0
    End If


    Set rsRo_det = New ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("Select CSMS_Ro_Det.ID,CSMS_Ro_Det.DetAmt,CSMS_Ro_Det.Rep_Or,detcde,CSMS_repor.Dte_rel from CSMS_Ro_Det inner join CSMS_repor on CSMS_ro_det.rep_or = CSMS_repor.rep_or where CSMS_RO_DET.livil = '4'  AND CSMS_RO_DET.DETCOST <=0 and CSMS_repor.dte_Comp = " & N2Date2Null(dtpTranDate) & " Order by CSMS_ro_det.Rep_Or asc")
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        rsRo_det.MoveFirst
        i = 0
        Do While Not rsRo_det.EOF
            Set rsOrd_Hd = New ADODB.Recordset
            Set rsOrd_Hd = gconDMIS.Execute("Select tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' AND trantype = 'RIV' and RONO = '" & Null2String(rsRo_det!REP_OR) & "'")
            If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                rsOrd_Hd.MoveFirst
                Do While Not rsOrd_Hd.EOF
                    Set rsDAYTRAN = New ADODB.Recordset
                    Set rsDAYTRAN = gconDMIS.Execute("Select tranucost from PMIS_vw_IS_DETHIST where [TYPE] = 'A' AND trantype = 'RIV' and tranno = '" & rsOrd_Hd!TRANNO & "' and STOCK_ORD = " & N2Str2Null(rsRo_det!DETCDE) & " order by trandate desc")
                    If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
                        If N2Str2Zero(rsDAYTRAN!TRANUCOST) > 0 Then
                            gconDMIS.Execute "Update CSMS_Ro_Det Set " & _
                                             "DetCost = " & N2Str2Zero(rsDAYTRAN!TRANUCOST) & _
                                             " Where id = " & rsRo_det!ID
                        End If
                    End If
                    rsOrd_Hd.MoveNext
                Loop
            End If
            i = i + 1
            progCPB.Value = (i / rsRo_det.RecordCount) * 100
            labCPB.Caption = Int(progCPB.Value) & "% Completed": DoEvents
            rsRo_det.MoveNext
        Loop
    End If
End Sub

Sub InitGrids()
    With Grid1
        .Rows = 1
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
        .Cell(0, 2).Text = "RO No."
        .Cell(0, 3).Text = "Inv. No."
        .Cell(0, 4).Text = "Inv. Amt."
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

    With Grid3
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "CS No."
        .Cell(0, 3).Text = "Inv. No."
        .Cell(0, 4).Text = "Inv. Amt."
        .Cell(0, 5).Text = "Customer"

        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 80
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
    DoEvents
End Sub

Function ImportPMISSales() As Boolean
    On Error GoTo ErrorCode
    Dim PMIOS_TRANTYPE                            As String
    Dim PMIOS_TRANNO                              As String
    Dim PMIOS_TRANDATE                            As String
    Dim PMIOS_cuscde                              As String
    Dim PMIOS_AcctName                            As String
    Dim PMIOS_TTLINVAMT                           As Double
    Dim PMIOS_DS_AMT1                             As Double
    Dim PMIOS_NETINVAMT                           As Double
    Dim PMIOS_NETCOST                             As Double
    Dim PMIOS_PAY_CLASS                           As String
    Dim PMIOS_TYPE                                As String
    Dim PMIS_TYPE                                 As String
    Dim TOTAL_DEBIT, TOTAL_CREDIT                 As Double
    Dim i                                         As Long
    Dim J_ITEMCOUNT                               As Integer

    i = 0

    Dim GridImport                                As Integer
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsPMIOS_ORD_HD = New ADODB.Recordset
            If Grid1.Cell(GridImport, 2).Text = "Accessories" Then
                If Len(Grid1.Cell(GridImport, 3).Text) = 10 Then
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                Else
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'A' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 2) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                End If
                PMIS_TYPE = "AI"
            ElseIf Grid1.Cell(GridImport, 2).Text = "Materials" Then
                If Len(Grid1.Cell(GridImport, 3).Text) = 10 Then
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'M' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                Else
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'M' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 2) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                End If
                PMIS_TYPE = "MI"
            ElseIf Grid1.Cell(GridImport, 2).Text = "Parts" Then
                If Len(Grid1.Cell(GridImport, 3).Text) = 10 Then
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'P' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 3) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                Else
                    Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where Type = 'P' AND TranType = '" & Left(Grid1.Cell(GridImport, 3).Text, 2) & "' and Tranno = '" & Right(Grid1.Cell(GridImport, 3).Text, 6) & "' AND STATUS = 'P' AND trandate = '" & CDate(dtpTranDate) & "' order by Tranno ASC")
                End If
                PMIS_TYPE = "PI"
            Else
                GoTo NextGrid
            End If

            If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
                If CheckSJExisting(PMIS_TYPE, Null2String(rsPMIOS_ORD_HD!TRANNO), Null2String(rsPMIOS_ORD_HD!TRANTYPE)) = True Then
                    GoTo DuplicateInvoiceNo
                Else
                PMIOS_TRANTYPE = Null2String(rsPMIOS_ORD_HD!TRANTYPE)
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
                PMIOS_TYPE = Null2String(rsPMIOS_ORD_HD!Type)
                PMIOS_NETCOST = ReturnCostOFSALES_PARTS(PMIOS_TRANNO, PMIOS_TRANTYPE, PMIOS_TYPE)



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
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'SJ'"
                J_REMARKS = "NULL"
                J_VENDORCODE = "'999999'"
                J_CUSTOMERCODE = N2Str2Null(PMIOS_cuscde)

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = 0
                If COMPANY_CODE = "HCI" And PMIOS_PAY_CLASS = "I" Then
                    J_INVOICEAMT = Round(PMIOS_NETCOST, 2)
                Else
                    J_INVOICEAMT = Round(PMIOS_NETINVAMT, 2)
                End If
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

                J_ITEMCOUNT = 0
                'J_JITEMNO = "'0001'"
                If COMPANY_CODE = "HCI" And PMIOS_PAY_CLASS = "I" Then
                Else
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                End If
                If PMIOS_TYPE = "P" Then
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS")))
                    ElseIf COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                    
                    ElseIf COMPANY_CODE = "HAI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "COUNTER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "COUNTER")))
                    Else
                        If PMIOS_PAY_CLASS = "C" Then
                            'CUSTOMER PAID
                            If COMPANY_CODE = "HPI" Then
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
                            If COMPANY_CODE = "HGC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "INTERNAL")))
                            ElseIf COMPANY_CODE = "HPI" Or COMPANY_CODE = "HCI" Then

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
                        ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HOT" Then
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
                ElseIf COMPANY_CODE = "HCI" And PMIOS_PAY_CLASS = "I" Then
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
                'J_JITEMNO = "'0002'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If PMIOS_TYPE = "P" Then
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS")))
                    ElseIf COMPANY_CODE = "HAI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("PARTS", "COUNTER"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("PARTS", "COUNTER")))
                    Else
                        If PMIOS_PAY_CLASS = "C" Then
                            'CUSTOMER PAID
                            If COMPANY_CODE = "HOT" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                            ElseIf COMPANY_CODE = "HPI" Then
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
                            If COMPANY_CODE = "HGC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "INTERNAL")))
                            ElseIf COMPANY_CODE = "HCI" Then
                                If UCase(PMIOS_cuscde) = "H00068" Then
                                    '                                    J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN"))
                                    '                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN")))
                                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP", "TAGBILARAN"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP", "TAGBILARAN")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL")))
                                End If
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
                        If COMPANY_CODE = "HPI" Then
                        ElseIf COMPANY_CODE = "HCI" Then
                            If UCase(PMIOS_cuscde) = "H00068" Then
                                'J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN"))
                                'J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN")))
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA", "TAGBILARAN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA", "TAGBILARAN")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL")))
                            End If
                        Else
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
                        ElseIf COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HOT" Then
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
                        ElseIf COMPANY_CODE = "HCI" Then
                            If UCase(PMIOS_cuscde) = "H00068" Then
                                '                                J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN"))
                                '                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL", "TAGBILARAN")))
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM", "TAGBILARAN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM", "TAGBILARAN")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnExpense("PARTS", "INTERNAL"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnExpense("PARTS", "INTERNAL")))
                            End If
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
                'J_JITEMNO = "'0003'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If PMIOS_TYPE = "P" Then
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                End If
                If PMIOS_TYPE = "A" Then
                    If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                    End If
                End If
                If PMIOS_TYPE = "M" Then
                    If COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                    End If
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
                'OUTPUT TAX

                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
                    If PMIOS_PAY_CLASS = "I" Then

                    Else
                        'J_JITEMNO = "'0004'"
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
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
                    'J_JITEMNO = "'0004'"
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
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
                'J_JITEMNO = "'0005'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                'CLEARING ACCOUNT - HGC - FML - 12/19/2007
                If PMIOS_PAY_CLASS = "C" Then
                    If COMPANY_CODE = "HMH" Then
                        If PMIOS_TRANTYPE = "CSH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                        End If
                    ElseIf COMPANY_CODE = "HGC" Then
                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("CASH"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("CASH")))
                    ElseIf COMPANY_CODE = "HAS" Then
                        J_ACCT_CODE = N2Str2Null("11-02000-00")
                        J_ACCT_NAME = N2Str2Null(Setacctname("11-02000-00"))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                    End If
                Else
                    If COMPANY_CODE <> "HPI" Then
                        If COMPANY_CODE = "HAI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INSTALLED"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INSTALLED")))
                        End If
                    End If
                End If
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HOT" Or COMPANY_CODE = "HAS" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Then
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
                ElseIf COMPANY_CODE = "HCI" And PMIOS_PAY_CLASS = "I" Then
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
                    'J_JITEMNO = "'0006'"
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
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
                        If COMPANY_CODE = "HPI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS")))
                        End If
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
                Grid1.Cell(GridImport, 1).Text = 1
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
            End If
        End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
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

    Dim i                                         As Integer
    Dim GridImport                                As Integer
    Dim rsSMIS_PURCHAGREE                         As ADODB.Recordset
    Dim SMIS_VI_NO                                As String
    Dim SMIS_DATERELEASED                         As String
    Dim SMIS_CODE                                 As String
    Dim SMIS_NETSALESPRICE                        As Double
    Dim SMIS_OTHERS                               As Double
    Dim SMIS_FOB                                  As Double
    Dim SMIS_TOTALCOST                            As Double
    Dim SMIS_DOWNPAYMENT                          As Double
    Dim SMIS_BALTOFINANCED                        As Double
    Dim SMIS_INSURANCE                            As Double
    Dim SMIS_LTOREGFEE                            As Double
    Dim SMIS_CHMOFEE                              As Double
    Dim SMIS_TERMS                                As String
    Dim J_ITEMCOUNT                               As Integer
    'VEHICLE SALES TRANSACTION
    '=========================================================================================================
    i = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Set rsSMIS_PURCHAGREE = New ADODB.Recordset
            Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where VI_NO = '" & Grid3.Cell(GridImport, 3).Text & "' AND STATUS = 'P' AND CONVERT(VarChar, DateReleased, 101)  = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
            If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
                'If rsSMIS_PURCHAGREE!VI_NO = "003034" Then Stop
                If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
                    GoTo DuplicateInvoiceNo
                Else
                SMIS_VI_NO = Null2String(rsSMIS_PURCHAGREE!VI_NO)
                SMIS_DATERELEASED = Null2String(rsSMIS_PURCHAGREE!DATERELEASED)
                SMIS_CODE = Null2String(rsSMIS_PURCHAGREE!code)
                SMIS_OTHERS = N2Str2Zero(rsSMIS_PURCHAGREE!OTHERS)
                SMIS_TOTALCOST = N2Str2Zero(rsSMIS_PURCHAGREE!TOTAL_COST)
                SMIS_DOWNPAYMENT = N2Str2Zero(rsSMIS_PURCHAGREE!DOWNPAYMENT)
                SMIS_BALTOFINANCED = N2Str2Zero(rsSMIS_PURCHAGREE!BALTOFINANCED)
                SMIS_TERMS = Null2String(rsSMIS_PURCHAGREE!TERM)

                SMIS_INSURANCE = N2Str2Zero(rsSMIS_PURCHAGREE!INSURANCE)
                SMIS_LTOREGFEE = N2Str2Zero(rsSMIS_PURCHAGREE!LTOREGFEE)
                SMIS_CHMOFEE = N2Str2Zero(rsSMIS_PURCHAGREE!CHMOFEE)
                SMIS_FOB = N2Str2Zero(rsSMIS_PURCHAGREE!FREIGHT)

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
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
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
                J_ITEMCOUNT = 0
                'J_JITEMNO = "'0001'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", "VEHICLES"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", "VEHICLES")))
                ElseIf COMPANY_CODE = "HAS" Then
                    If rsSMIS_PURCHAGREE!Model = "I10" Then
                        J_ACCT_CODE = N2Str2Null("41-01016-10")
                        J_ACCT_NAME = N2Str2Null(Setacctname("41-01016-10"))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!Model)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SALES", Null2String(rsSMIS_PURCHAGREE!Model))))
                    End If
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
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                'COST OF SALES
                'J_JITEMNO = "'0002'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SALES", "VEHICLES"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SALES", "VEHICLES")))
                ElseIf COMPANY_CODE = "HAS" Then
                    If rsSMIS_PURCHAGREE!Model = "I10" Then
                        J_ACCT_CODE = N2Str2Null("61-01016-10")
                        J_ACCT_NAME = N2Str2Null(Setacctname("61-01016-10"))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SALES", Null2String(rsSMIS_PURCHAGREE!Model)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SALES", Null2String(rsSMIS_PURCHAGREE!Model))))
                    End If
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

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO

                'INVENTORY
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
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
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO

                'J_JITEMNO = "'0004'"
                If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                    'OUTPUT TAX
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
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
                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                End If


                '            'A/R VEHICLES
                '            J_JITEMNO = "'0001'"
                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If SMIS_TERMS = "COD" Then
                    If COMPANY_CODE = "HMH" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                    ElseIf COMPANY_CODE = "HGC" Then
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
                    ElseIf COMPANY_CODE = "HAI" Or COMPANY_CODE = "HCI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("SALES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("SALES")))
                    ElseIf COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("VEHICLES")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("FINANCING"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("FINANCING")))
                    End If
                End If

                If COMPANY_CODE <> "HMH" Then
                    If Null2Bool(rsSMIS_PURCHAGREE!ZERORATED) = False Then
                        If COMPANY_CODE = "HCI" Then
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_INSURANCE) + NumericVal(SMIS_LTOREGFEE) + NumericVal(SMIS_CHMOFEE) + NumericVal(SMIS_FOB), 2)
                        Else
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        End If
                    Else
                        If COMPANY_CODE = "HCI" Then
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE) + NumericVal(SMIS_INSURANCE) + NumericVal(SMIS_LTOREGFEE) + NumericVal(SMIS_CHMOFEE) + NumericVal(SMIS_FOB), 2)
                        Else
                            J_DEBIT = Round(NumericVal(SMIS_NETSALESPRICE), 2)
                        End If
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
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO

                If COMPANY_CODE = "HMH" Then
                    If SMIS_TERMS = "F" Or SMIS_TERMS = "BPO" Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
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

                If COMPANY_CODE = "HCI" Then
                    If SMIS_INSURANCE <> 0 Then

                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnAccommodation("INSURANCE"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccommodation("INSURANCE")))
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(SMIS_INSURANCE, 2)
                        'J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "ELLAINE", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If

                    If SMIS_LTOREGFEE <> 0 Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnAccommodation("LTO"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccommodation("LTO")))
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(SMIS_LTOREGFEE, 2)
                        'J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "ELLAINE", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If

                    If SMIS_CHMOFEE <> 0 Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnAccommodation("CHATTEL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccommodation("CHATTEL")))
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(SMIS_CHMOFEE, 2)
                        'J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "ELLAINE", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL IMPORT DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
                    End If

                    If SMIS_FOB <> 0 Then
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnAccommodation("FOB"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAccommodation("FOB")))
                        J_GROSS = 0
                        J_TAX = 0
                        J_NET = 0
                        J_DEBIT = 0
                        J_CREDIT = Round(SMIS_FOB, 2)
                        'J_CREDIT = Round(SMIS_TOTALCOST - (NumericVal(SMIS_TOTALCOST) / 9.3333), 2)
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "ELLAINE", J_JTYPE, "Jtype"))
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
                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_HD", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, J_JNO
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

Sub INITIALIZE_ACCESSORIESCOST(XXX As String)
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HCI" Then
        Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
        'Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_ACCESSORIES Where REP_OR = " & N2Str2Null(XXX))
        Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMIOS_vw_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
        If Not rsCSMIOS_ACCESSORIES.EOF And Not rsCSMIOS_ACCESSORIES.BOF Then
            CSMIOS_ACCESSORIES = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!ACCESSORIES), 2)
            CSMIOS_ACCESSORIES_DISCOUNT = Round(N2Str2Zero(rsCSMIOS_ACCESSORIES!DISCOUNT), 2)
        Else
            CSMIOS_ACCESSORIES = 0: CSMIOS_ACCESSORIES_DISCOUNT = 0
        End If

        Set rsCSMIOS_ACCESSORIES = New ADODB.Recordset
        'CSMS_vw_ACCESSORIESWarranty
        Set rsCSMIOS_ACCESSORIES = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS ACCESSORIES,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMIOS_vw_WARRANTY_ACCCOST Where REP_OR = " & N2Str2Null(XXX))
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
    CSMIOS_LABOR = 0
    CSMIOS_LABOR_DISCOUNT = 0
    CSMIOS_LABOR_COST = 0
    CSMIOS_PARTS = 0
    CSMIOS_PARTS_DISCOUNT = 0
    CSMIOS_MATERIALS = 0
    CSMIOS_MATERIALS_DISCOUNT = 0
    CSMIOS_SUBLET = 0
    CSMIOS_SUBLET_DISCOUNT = 0
    CSMIOS_SUBLET_COST = 0
    CSMIOS_TINSPAINT = 0
    CSMIOS_TINSPAINT_DISCOUNT = 0
    CSMIOS_TINSPAINT_COST = 0
    CSMIOS_PMS = 0
    CSMIOS_PMS_DISCOUNT = 0
    CSMIOS_PMS_COST = 0
    CSMIOS_PARTS_COST = 0
    CSMIOS_MATERIALS_COST = 0
    CSMIOS_ACCESSORIES_COST = 0:
    WARRANTY_DIRECT_EXPENSE_LABOR = 0
    WARRANTY_DIRECT_EXPENSE_SPAREPARTS = 0
    WARRANTY_DIRECT_EXPENSE_GOL = 0
    WARRANTY_CSMIOS_PARTS_COST = 0
    WARRANTY_CSMIOS_MATERIALS_COST = 0
    WARRANTY_CSMIOS_ACCESSORIES_COST = 0
    COMPANY_DIRECT_EXPENSE_LABOR = 0
    COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0
    COMPANY_DIRECT_EXPENSE_GOL = 0
    SALES_DIRECT_EXPENSE_LABOR = 0
    SALES_DIRECT_EXPENSE_SPAREPARTS = 0
    SALES_DIRECT_EXPENSE_GOL = 0
    COMPANY_DIRECT_EXPENSE_LABOR = 0
    COMPANY_DIRECT_EXPENSE_SPAREPARTS = 0
    COMPANY_DIRECT_EXPENSE_GOL = 0
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
    TOTAL_DISCOUNT_AMOUNT = Round(CSMIOS_LABOR_DISCOUNT + CSMIOS_PARTS_DISCOUNT + CSMIOS_MATERIALS_DISCOUNT + CSMIOS_SUBLET_DISCOUNT + CSMIOS_TINSPAINT_DISCOUNT + CSMIOS_PMS_DISCOUNT + CSMIOS_ACCESSORIES_DISCOUNT, 2)
    '====================================================================================================================================================================================
End Sub

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
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HCI" Then
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
    CSMIOS_TINSPAINT = CSMIOS_TINSPAINT - WARRANTY_DIRECT_EXPENSE_LABOR

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
        'UPDATED: ACL 04162010 COST FOR INSURANCE
        If COMPANY_CODE = "HMH" Then
            '            Dim rsINSURANCECOST                                  As ADODB.Recordset
            '            Set rsINSURANCECOST = gconDMIS.Execute("SELECT ROUND(SUM(DETCOST),2) AS DETCOST From CSMS_Ro_Det " & _
                         '                                                   " WHERE ((LIVIL = '1') AND (POCODE = '99' OR POCODE = '60') OR (LIVIL = '1') AND (JOBTYPE = 'BP')) AND REP_OR = " & N2Str2Null(XXX))
            '            If Not rsINSURANCECOST.EOF And Not rsINSURANCECOST.BOF Then
            '                CSMIOS_LABOR_COST = Round(N2Str2Zero(rsINSURANCECOST!detcost), 2)
            '            End If


            '            Dim rsINSURANCECOST2                                  As ADODB.Recordset
            '            Set rsINSURANCECOST2 = gconDMIS.Execute("Select ROUND(sum(DETPRC),2) AS SUBLET,ROUND(sum(DETCOST),2) AS COST,ROUND(sum(DISCOUNT),2) AS DISCOUNT from CSMS_vw_SUBLET " & _
                         '                                                   " WHERE WCODE IS NULL AND REP_OR = " & N2Str2Null(XXX))
            '            If Not rsINSURANCECOST2.EOF And Not rsINSURANCECOST2.BOF Then
            '                CSMIOS_SUBLET_COST = Round(N2Str2Zero(rsINSURANCECOST2!Cost), 2)
            '            End If

        End If
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
    Dim iCount, xCount                            As Integer
    Dim Indx                                      As Integer
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
    Dim iCount, xCount                            As Integer
    Dim Indx                                      As Integer
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
    If Function_Access(LOGID, "Acess_Process", "IMPORT SALES ENTRIES") = False Then Exit Sub
    Screen.MousePointer = 11
    Dim str_MSG                                   As String
    str_MSG = "Error Appear In During @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Imported Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf
    gconDMIS.BeginTrans

    Prg1.Max = 6:
    Prg1.Value = 1: cmdCheck.Enabled = False: cmdExit.Enabled = False

    If ImportPMISSales = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import PMIS Sales")
        MsgBox str_MSG, vbCritical, "Import Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    Prg1.Value = 4

    If ImportPurelyInternal = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Purely Internal")
        MsgBox str_MSG, vbCritical, "Import Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    If ImportSales = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Sales")
        MsgBox str_MSG, vbCritical, "Import Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If


    Prg1.Value = 5

    If ImportSMISSales = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Import SMIS Sales")
        MsgBox str_MSG, vbCritical, "Import Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    Prg1.Value = 6
    gconDMIS.CommitTrans
    cmdCheck.Enabled = True
    cmdExit.Enabled = True
    Screen.MousePointer = 0
    Call cmdShowImp_Click
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    '=========================================================================================================
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function ImportSales() As Boolean
    On Error GoTo ErrorCode
    Dim i                                         As Integer
    Dim GridImport                                As Integer
    Dim Duo                                       As String
    i = 0
    Dim rsSublet                                  As ADODB.Recordset
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            ALL_DEBIT = 0: ALL_CREDIT = 0
            Set rsCSMIOS_REPOR = New ADODB.Recordset

            'Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice = '" & Grid2.Cell(GridImport, 3).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC")
            'ORIGINAL CODE:
            SQL_STATEMENT = "Select * from CSMS_REPOR where REP_OR = '" & Grid2.Cell(GridImport, 2).Text & "' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC"

            'FOR DEBUGGING PURPOSE ONLY
            'SQL_STATEMENT = "Select * from CSMS_REPOR where REP_OR = 'R-00003743' and dte_comp = '" & CDate(dtpTranDate) & "' order by invoice ASC"

            Set rsCSMIOS_REPOR = gconDMIS.Execute(SQL_STATEMENT)


            If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
                'If Null2String(rsCSMIOS_REPOR!REP_OR) = "R-00008518" Then Stop
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!invoice)) = True Then
                    If Null2String(rsCSMIOS_REPOR!invoice) = "INT RO" Then
                        GoTo ImportINTRO
                    Else
                        GoTo DuplicateInvoiceNo
                    End If
                Else
ImportINTRO:
                    ItemCnt = 0
                    CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
                    CSMIOS_ACCT_NO = Null2String(rsCSMIOS_REPOR!ACCT_NO)
                    CSMIOS_PARTICIPAT = Null2String(rsCSMIOS_REPOR!PARTICIPAT)
                    CSMIOS_PLATE_NO = Null2String(rsCSMIOS_REPOR!PLATE_NO)
                    CSMIOS_NIYM = Null2String(rsCSMIOS_REPOR!Niym)
                    CSMIOS_TERM = Null2String(rsCSMIOS_REPOR!TERM)
                    CSMIOS_DTE_REL = Null2Date(rsCSMIOS_REPOR!dte_comp)
                    CSMIOS_INVOICE = Null2String(rsCSMIOS_REPOR!invoice)
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
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    J_JTYPE = "'SJ'"
                    J_REMARKS = "NULL"
                    J_VENDORCODE = "'999999'"
                    J_CUSTOMERCODE = N2Str2Null(CSMIOS_ACCT_NO)
                    J_DEBIT = 0
                    J_CREDIT = 0
                    J_TAX = 0
                    J_OUTBALANCE = 0
                    J_AMOUNTTOPAY = 0
                    If CSMIOS_TINSPAINT < 0 Then
                        CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)
                    Else
                        CSMIOS_RO_AMOUNT = Round((CSMIOS_LABOR + CSMIOS_AIRCON + CSMIOS_TINSPAINT + CSMIOS_SUBLET + CSMIOS_PMS + CSMIOS_PARTS + CSMIOS_MATERIALS + CSMIOS_ACCESSORIES) - TOTAL_DISCOUNT_AMOUNT, 2)
                    End If
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
                        If CSMIOS_LABOR <> 0 Then
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
                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
                            Else
                                If NumericVal(Duo) = 1 Then
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

                        'SUBLET
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

                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "hot" Then
                            Else
                                'COST OF SALES
                                If NumericVal(Duo) = 1 Then
                                Else
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
                            End If

                            If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
                            Else
                                If NumericVal(Duo) = 1 Then
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
                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
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


                        'PMS
                        If CSMIOS_PMS <> 0 Then
                            ItemCnt = ItemCnt + 1
                            J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                            If CSMIOS_TERM = "CSH" Then
                                'SALES
                                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                                ElseIf COMPANY_CODE = "HOT" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "DIAGNOSTIC")))
                                End If
                            Else
                                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HSB" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
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

                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
                            Else
                                'COST OF SALES
                                ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                                If CSMIOS_TERM = "CSH" Then
                                    If COMPANY_CODE = "HSB" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                                    ElseIf COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
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
                        'PARTS ISSUED
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

                            If COMPANY_CODE = "HMH" And INSURANCE_DIRECT_EXPENSE_SPAREPARTS > 0 Then
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
                                    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                                    Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
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
                        'MATERIAL ISSUED
                        If CSMIOS_MATERIALS > 0 Then
                            Dim SHOPSUP           As Double
                            Dim SHOPSUP_TAX       As Double
                            Dim SHOPSUP_NET       As Double

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
                            If COMPANY_CODE = "HGC" Then
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

                            If COMPANY_CODE = "HMH" And INSURANCE_DIRECT_EXPENSE_GOL > 0 Then
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
                                    If COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAI" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVA", "MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVA", "MATERIALS")))
                                    ElseIf COMPANY_CODE = "HOT" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                                   Else
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS")))
                                    End If
                                Else
                                    If COMPANY_CODE = "HSB" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "VEHICLES"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "VEHICLES")))
                                    ElseIf COMPANY_CODE = "HAI" Then
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

                        Call IMPORT_ACCESSORIES

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
                                If COMPANY_CODE = "HGC" Then
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
                                        J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("MATERIALS", "MATERIALS")))
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
                        If CSMIOS_ACCESSORIES_DISCOUNT > 0 Then
                            'DISCOUNT
                            ItemCnt = ItemCnt + 1: J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"

                            If COMPANY_CODE = "HAI" Then
                                If CSMIOS_TERM = "CSH" Then
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "ACCESSORIES")))
                                Else
                                    J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("SERVICE", "ACCESSORIES"))
                                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("SERVICE", "ACCESSORIES")))
                                End If
                            ElseIf COMPANY_CODE = "HCI" Then
                                J_ACCT_CODE = N2Str2Null(ReturnDiscount_AccountCode("PARTS", "PARTS"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnDiscount_AccountCode("PARTS", "PARTS")))
                            End If


                            J_GROSS = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                            If CSMIOS_INVOICE = "PDI RO" Then
                                J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                            Else
                                If CSMIOS_VAT_EXEMPT = True Then
                                    J_TAX = 0: J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT), 2)
                                Else
                                    J_TAX = Round(NumericVal(Round((CSMIOS_ACCESSORIES_DISCOUNT / 1.12), 2) * 0.12), 2)
                                    J_NET = Round(NumericVal(CSMIOS_ACCESSORIES_DISCOUNT) - Round(NumericVal(J_TAX), 2), 2)
                                End If
                            End If
                            J_DEBIT = Round(NumericVal(J_NET), 2)
                            J_CREDIT = 0
                            ALL_DEBIT = ALL_DEBIT + J_DEBIT

                            Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

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
                                J_CREDIT = Round(NumericVal(Round((J_INVOICEAMT / 1.12), 2) * 0.12), 2)
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
                                ElseIf COMPANY_CODE = "HGC" Then
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
                    Call Import_Warranty
                    '                If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
                    '
                    '                    Set rsJournal_HDDup = New ADODB.Recordset
                    '                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                    '                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    '
                    '                    WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
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
                    '                        ElseIf COMPANY_CODE = "HPI" Then
                    '                            J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    '                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
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
                    '                    If CheckIfPMS(CSMIOS_REP_OR) = True Then
                    '                        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    '                            If COMPANY_CODE = "HPI" Then
                    '                                CSMIOS_ACCT_NO = Null2String("H00416")
                    '                                CSMIOS_NIYM = Null2String(SetCustomerName("H00416"))
                    '                                J_CUSTOMERCODE = N2Str2Null("H00416")
                    '                            Else
                    '                                CSMIOS_ACCT_NO = Null2String("H00001")
                    '                                CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    '                                J_CUSTOMERCODE = N2Str2Null("H00001")
                    '                            End If
                    '                        Else
                    '                            CSMIOS_NIYM = Null2String(SetSellingDealerName(ReturnCodeSellingDealer(ReturnPlateNo(CSMIOS_REP_OR))))
                    '                            CSMIOS_ACCT_NO = N2Str2Null(ReturnCode(CSMIOS_NIYM))
                    '                            J_CUSTOMERCODE = N2Str2Null(ReturnCode(CSMIOS_NIYM))
                    '                        End If
                    '                    Else
                    '                        CSMIOS_ACCT_NO = Null2String("H00001")
                    '                        CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    '                        J_CUSTOMERCODE = N2Str2Null("H00001")
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

                    '=========================================================================================================================================================
                    'ENTRY FOR INSURANCE

                    If INSURANCE_DIRECT_EXPENSE_LABOR + INSURANCE_DIRECT_EXPENSE_SPAREPARTS + INSURANCE_DIRECT_EXPENSE_GOL + INSURANCE_DIRECT_EXPENSE_ACCESSORIES > 0 Then
                        Set rsJournal_HDDup = New ADODB.Recordset
                        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
                        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 2, "000000") & "'"

                        WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
                        WARRANTY_ItemCnt = 0: WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
                        If COMPANY_CODE = "HBK" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("AR"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("AR")))
                        ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
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
                            ElseIf COMPANY_CODE = "HAI" Then
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

                            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then

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
                                    If COMPANY_CODE = "HOT" Or CSMIOS_TERM = "CSH" Then
                                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL"))
                                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "SUBLET", "RETAIL")))
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
                            ElseIf COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
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
                                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
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
                                If COMPANY_CODE = "HAI" Then
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
                    Grid2.Cell(GridImport, 1).Text = 1
                End If
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
DuplicateInvoiceNo:
    Next

    ImportSales = True
    Exit Function
ErrorCode:
    ImportSales = False
End Function

Sub IMPORT_INTERNAL()
    If COMPANY_DIRECT_EXPENSE_LABOR + COMPANY_DIRECT_EXPENSE_SPAREPARTS + COMPANY_DIRECT_EXPENSE_GOL + SALES_DIRECT_EXPENSE_LABOR + SALES_DIRECT_EXPENSE_SPAREPARTS + SALES_DIRECT_EXPENSE_GOL > 0 Then
        'LABOR

        INTERNAL_LABOR_AMT = 0: INTERNAL_PARTS_AMT = 0: INTERNAL_MATERIALS_AMT = 0
        INTERNAL_LABOR_COST = 0: INTERNAL_PARTS_COST = 0: INTERNAL_MATERIALS_COST = 0

        Set rsINTERNAL_RO_DET = New ADODB.Recordset
        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '1' AND DETPRC > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
            rsINTERNAL_RO_DET.MoveFirst
            Do While Not rsINTERNAL_RO_DET.EOF
                If N2Str2Zero(rsINTERNAL_RO_DET!detprc) > 0 Then
                    INTERNAL_LABOR_AMT = INTERNAL_LABOR_AMT + N2Str2Zero(rsINTERNAL_RO_DET!detprc)
                    INTERNAL_LABOR_COST = INTERNAL_LABOR_COST + N2Str2Zero(rsINTERNAL_RO_DET!detcost)
                    If COMPANY_CODE = "HPI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                        J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detprc)), 2)
                        J_TAX = 0
                        If COMPANY_CODE = "HMH" Then
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detcost)), 2)
                        Else
                            J_NET = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!detprc)), 2)
                        End If
                        J_DEBIT = Round(NumericVal(J_NET), 2)
                        J_CREDIT = 0

                        ALL_CREDIT = ALL_CREDIT + J_CREDIT
                        Call InsertToJournalDet(J_JDATE, J_VOUCHERNO, J_JTYPE, J_JNO, J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)
                    End If
                End If
                rsINTERNAL_RO_DET.MoveNext
            Loop

            If COMPANY_CODE = "HMH" Then
            Else
                ItemCnt = ItemCnt + 1
                J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR")))
                ElseIf COMPANY_CODE = "HAI" Then
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

            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Then
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
            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
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
        Set rsINTERNAL_RO_DET = New ADODB.Recordset
        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '2' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
            rsINTERNAL_RO_DET.MoveFirst
            Do While Not rsINTERNAL_RO_DET.EOF
                If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                    INTERNAL_PARTS_AMT = INTERNAL_PARTS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                    INTERNAL_PARTS_COST = INTERNAL_PARTS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))

                    If COMPANY_CODE = "HPI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
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
                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                Else
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS")))
                    ElseIf COMPANY_CODE = "HAI" Then
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

                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                Else
                    'COST OF SALES
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS")))
                    ElseIf COMPANY_CODE = "HAI" Then
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
                ElseIf COMPANY_CODE = "HAI" Then
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

        Set rsINTERNAL_RO_DET = New ADODB.Recordset
        Set rsINTERNAL_RO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET WHERE LIVIL = '3' AND DET_AMT > 0 AND (WCODE = 'C' OR WCODE = 'S') AND REP_OR = '" & CSMIOS_REP_OR & "'")
        If Not rsINTERNAL_RO_DET.EOF And Not rsINTERNAL_RO_DET.BOF Then
            rsINTERNAL_RO_DET.MoveFirst
            Do While Not rsINTERNAL_RO_DET.EOF
                If N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT) > 0 Then
                    INTERNAL_MATERIALS_AMT = INTERNAL_MATERIALS_AMT + N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)
                    INTERNAL_MATERIALS_COST = INTERNAL_MATERIALS_COST + (N2Str2Zero(rsINTERNAL_RO_DET!detcost) * N2Str2Zero(rsINTERNAL_RO_DET!detvol))
                    If COMPANY_CODE = "HPI" Then
                    Else
                        ItemCnt = ItemCnt + 1
                        J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code)))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInternalAccountCode(Null2String(rsINTERNAL_RO_DET!code))))
                        J_GROSS = Round(NumericVal(N2Str2Zero(rsINTERNAL_RO_DET!DET_AMT)), 2)
                        J_TAX = 0
                        If COMPANY_CODE = "HMH" Then
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
                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                Else
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
                    ElseIf COMPANY_CODE = "HAI" Then
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

                If COMPANY_CODE = "HPI" Or COMPANY_CODE = "HMH" Then
                Else
                    'COST OF SALES
                    ItemCnt = ItemCnt + 1
                    J_JITEMNO = "'" & Format(ItemCnt, "0000") & "'"
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LUBRICANTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LUBRICANTS")))
                    ElseIf COMPANY_CODE = "HAI" Then
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
                If COMPANY_CODE = "HAI" Then
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

        If COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAI" Then
            Dim rsCheckHD                         As ADODB.Recordset
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

Private Sub Import_Warranty()
    If WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_journal_hd order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then WARRANTY_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"

        WARRANTY_VOUCHERNO = N2Str2Null(Format(NumericVal(GetVoucherNo()) + 1, "000000"))
        WARRANTY_ItemCnt = 0
        WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
        TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
        'LABOR
        If WARRANTY_DIRECT_EXPENSE_LABOR > 0 Then
            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
            If COMPANY_CODE = "HGC" Then
                If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "RETAIL")))
                Else
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
                End If
            Else
                If COMPANY_CODE = "HAI" Then
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

            If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HAI" Then
            Else
                'COST OF SALES
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1: WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If CSMIOS_TERM = "CSH" Then
                    If COMPANY_CODE = "HGC" Then
                        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                        End If
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                    End If
                Else
                    If COMPANY_CODE = "HGC" Then
                        If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "RETAIL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "RETAIL")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "LABOR", "WARRANTY")))
                        End If
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
        If WARRANTY_DIRECT_EXPENSE_SPAREPARTS > 0 Then
            If COMPANY_CODE = "HPI" Then
            Else
                WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
                WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
                If COMPANY_CODE = "HBK" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("PARTS", "PARTS", "WARRANTY")))
                ElseIf COMPANY_CODE = "HAI" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "RETAIL")))
                ElseIf COMPANY_CODE = "HOT" Then
                    J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "PARTS", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "PARTS", "WARRANTY")))
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
                ElseIf COMPANY_CODE = "HAI" Then
                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "RETAIL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "RETAIL")))
                ElseIf COMPANY_CODE = "HOT" Then
                    J_ACCT_CODE = N2Str2Null(ReturnCostofSales("SERVICE", "PARTS", "WARRANTY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnCostofSales("SERVICE", "PARTS", "WARRANTY")))
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
                ElseIf COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
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



        'Materials
        If WARRANTY_DIRECT_EXPENSE_GOL > 0 Then
            WARRANTY_ItemCnt = WARRANTY_ItemCnt + 1
            WARRANTY_J_JITEMNO = "'" & Format(WARRANTY_ItemCnt, "0000") & "'"
            If COMPANY_CODE = "HBK" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
            ElseIf COMPANY_CODE = "HGC" Then
                J_ACCT_CODE = "'41-02013-20'"
                J_ACCT_NAME = N2Str2Null(Setacctname("'41-02013-20'"))
            ElseIf COMPANY_CODE = "HSB" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "INTERNAL")))
            ElseIf COMPANY_CODE = "HAI" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS")))
            ElseIf COMPANY_CODE = "HPI" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LABOR", "WARRANTY")))
            ElseIf COMPANY_CODE = "HMH" Then
                J_ACCT_CODE = N2Str2Null(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "RETAIL"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnSales_AccountCode("SERVICE", "LUBRICANTS", "RETAIL")))
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
                ElseIf COMPANY_CODE = "HGC" Then
                    J_ACCT_CODE = "'61-02011-20'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'61-02011-20'"))
                ElseIf COMPANY_CODE = "HSB" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HMH" Then
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
            If COMPANY_CODE = "HAI" Then
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
            Else
                J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
            End If
        Else
            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("WARRANTY"))
            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("WARRANTY")))
        End If
        J_DEBIT = NumericVal(WARRANTY_DIRECT_EXPENSE_LABOR + WARRANTY_DIRECT_EXPENSE_SPAREPARTS + WARRANTY_DIRECT_EXPENSE_GOL)
        J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
        Call InsertToJournalDet(J_JDATE, WARRANTY_VOUCHERNO, J_JTYPE, WARRANTY_JNO, WARRANTY_J_JITEMNO, J_ACCT_CODE, J_ACCT_NAME, J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET, J_STATUS)

        CSMIOS_REP_OR = Null2String(rsCSMIOS_REPOR!REP_OR)
        'To return the Selling Dealer
        If CheckIfPMS(CSMIOS_REP_OR) = True Then
            If CheckIfPMS_Ik_to_5k(CSMIOS_REP_OR) = True Then
                If COMPANY_CODE = "HPI" Then
                    CSMIOS_ACCT_NO = Null2String("H00416")
                    CSMIOS_NIYM = Null2String(SetCustomerName("H00416"))
                    J_CUSTOMERCODE = N2Str2Null("H00416")
                Else
                    CSMIOS_ACCT_NO = Null2String("H00001")
                    CSMIOS_NIYM = Null2String(SetCustomerName("H00001"))
                    J_CUSTOMERCODE = N2Str2Null("H00001")
                End If
            Else
                CSMIOS_NIYM = Null2String(SetSellingDealerName(ReturnCodeSellingDealer(ReturnPlateNo(CSMIOS_REP_OR))))
                CSMIOS_ACCT_NO = N2Str2Null(ReturnCode(CSMIOS_NIYM))
                J_CUSTOMERCODE = N2Str2Null(ReturnCode(CSMIOS_NIYM))
            End If
        Else
            If COMPANY_CODE = "HCI" Then
                CSMIOS_ACCT_NO = Null2String("H00042")
                CSMIOS_NIYM = Null2String(SetCustomerName("H00042"))
                J_CUSTOMERCODE = N2Str2Null("H00042")
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

Private Sub cmdClearJournals_Click()
    Dim rsCHATCheckControlIfExistRecordInJournalHD As ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "' and status <> 'C'")
    If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
        Screen.MousePointer = 0
        If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Purge Data") = vbYes Then
            Screen.MousePointer = 11
            gconDMIS.Execute ("delete from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("delete from AMIS_Journal_DET Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("delete from AMIS_PV_DETAIL Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate = '" & CDate(dtpTranDate) & "'")
            cmdShowTrans.Value = True
            Screen.MousePointer = 0
            MsgBox "Existing Data Successfully deleted.", vbInformation, "Deleted"
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
        Grid1.Rows = 1: Grid2.Rows = 1: Grid3.Rows = 1
    End If
    Dim tmpDate                                   As String
    Dim iCount                                    As Integer
    MonthIndex
    Dim rsCheckSALESData                          As ADODB.Recordset
    Dim rsCheckSALESData2                         As ADODB.Recordset
'--------------------------------
'    Set rsCheckSALESData = New ADODB.Recordset
'    'rsCheckSALESData.Open "SELECT * FROM (SELECT DISTINCT TRANDATE AS TRANDATE FROM PMIS_VW_ISS_HISTORY WHERE (TRANTYPE = 'CSH' OR  TRANTYPE = 'CHG') AND STATUS = 'P' AND MONTH(TRANDATE) = '" & Indx & "' AND YEAR(TRANDATE)='" & cboYear.Text & "' " & _
'                          "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR  WHERE INVOICE = 'INT RO' AND MONTH(DTE_COMP) = '" & Indx & "' AND YEAR(DTE_COMP)='" & cboYear.Text & "'  " & _
'                          "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR WHERE RO_AMOUNT > 0 AND INVOICE <> 'INT RO' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND MONTH(DTE_COMP) = '" & Indx & "' AND YEAR(DTE_COMP)='" & cboYear.Text & "' " & _
'                          "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND INVOICE <> 'INT RO' AND MONTH(DTE_COMP) = '" & Indx & "' AND YEAR(DTE_COMP)='" & cboYear.Text & "' " & _
'                          "UNION SELECT DISTINCT CONVERT(VARCHAR, DATERELEASED, 101) AS TRANDATE FROM SMIS_PURCHAGREE WHERE STATUS = 'P' AND (MONTH(CONVERT(VARCHAR, DATERELEASED, 101))> '" & Indx & "' AND YEAR(CONVERT(VARCHAR, DATERELEASED, 101)) > '" & cboYear.Text & "')) X ORDER BY TRANDATE ASC", gconDMIS, adOpenKeyset
'    rsCheckSALESData.Open "SELECT * FROM AMIS_VW_UNIMPORTED_ISSUANCES WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "' ", gconDMIS, adOpenForwardOnly
'    Do While Not rsCheckSALESData.EOF
'        iCount = 0
'        Dim rsCheckJournalHD                      As ADODB.Recordset
'        Set rsCheckJournalHD = New ADODB.Recordset
'        'rsCheckJournalHD.Open "SELECT DISTINCT JDATE,JTYPE FROM AMIS_JOURNAL_HD WHERE JTYPE = 'SJ' AND JDATE IN (SELECT * FROM (SELECT DISTINCT TRANDATE AS TRANDATE FROM PMIS_VW_ISS_HISTORY WHERE (TRANTYPE = 'CSH' OR  TRANTYPE = 'CHG') AND STATUS = 'P' " & _
'                              "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR WHERE INVOICE = 'INT RO' " & _
'                              "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR WHERE RO_AMOUNT > 0 AND INVOICE <> 'INT RO' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' " & _
'                              "UNION SELECT DISTINCT DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND INVOICE <> 'INT RO' " & _
'                              "UNION SELECT DISTINCT CONVERT(VARCHAR, DATERELEASED, 101) AS TRANDATE FROM SMIS_PURCHAGREE WHERE STATUS = 'P') X WHERE TRANDATE = '" & Null2String(rsCheckSALESData!trandate) & "')", gconDMIS, adOpenKeyset
'
''        rsCheckJournalHD.Open "SELECT * FROM AMIS_VW_UNIMPORTED_ISSUANCES WHERE [DATE] = '" & Null2String(Format(rsCheckSALESData!Date, "mm/dd/yyyy")) & "'", gconDMIS, adOpenForwardOnly
'        If Not rsCheckJournalHD.EOF And Not rsCheckJournalHD.BOF Then
'            Do While iCount <= lab.Count - 1
'                If lab.Item(iCount).Caption = Format(Null2String(rsCheckJournalHD!Date), "d") Then
'                    lab.Item(iCount).BackColor = &HC0FFC0 'GREEN
'                    GoTo Skip
'                End If
'                tmpDate = Indx & "/" & iCount & "/" & cboYear.Text
'
'                iCount = iCount + 1
'            Loop
'Skip:
'        Else
'            Do While iCount <= lab.Count - 1
'                If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckSALESData!Date)), "d") Then
'                    lab.Item(iCount).BackColor = &HFFFF& 'YELLOW
'                End If
'                iCount = iCount + 1
'            Loop
'        End If
'        rsCheckSALESData.MoveNext
'    Loop
'    Screen.MousePointer = 0
'----------------------------------

    Set rsCheckSALESData2 = New ADODB.Recordset
    rsCheckSALESData2.Open "SELECT DISTINCT CONVERT(VARCHAR,DATE,101) AS DATE FROM AMIS_VW_IMPORTED_ISSUANCES WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "' ", gconDMIS, adOpenForwardOnly
    Do While Not rsCheckSALESData2.EOF
        iCount = 0
            Do While iCount <= lab.Count - 1
                If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckSALESData2!Date)), "d") Then
                    lab.Item(iCount).BackColor = &HC0FFC0 'YELLOW
                End If
                iCount = iCount + 1
                DoEvents
            Loop
        rsCheckSALESData2.MoveNext
        DoEvents
    Loop
    
    Set rsCheckSALESData = New ADODB.Recordset
    rsCheckSALESData.Open "SELECT DISTINCT CONVERT(VARCHAR,DATE,101) AS DATE FROM AMIS_VW_UNIMPORTED_ISSUANCES WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "' ", gconDMIS, adOpenForwardOnly
    Do While Not rsCheckSALESData.EOF
        iCount = 0
            Do While iCount <= lab.Count - 1
                If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckSALESData!Date)), "d") Then
                    lab.Item(iCount).BackColor = &HFFFF& 'YELLOW
                End If
                iCount = iCount + 1
                DoEvents
            Loop
        rsCheckSALESData.MoveNext
        DoEvents
    Loop
    Screen.MousePointer = 0
End Sub

Private Sub cmdShowTrans_Click()
    Dim KIM                                       As Integer
    Screen.MousePointer = 11
InitGrids:     DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
    Grid1.Rows = 2: Grid2.Rows = 2: Grid3.Rows = 2
    Dim InvoiceType, InvoiceTypeCode              As String
    Dim IS_Exist                                  As Byte

    If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
        GoTo ShowTrans
    Else
        If CheckImportedSales(dtpTranDate) = True Then
            MsgBox "Previous transactions dated " & TRANSACTIONDATE & " are not yet imported.", vbExclamation, "Message"
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo ShowTransactions
        End If

ShowTransactions:
    End If

ShowTrans:
    Set rsPMIOS_ORD_HD = New ADODB.Recordset
    If COMPANY_CODE = "HCI" Then
        Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where (TranType = 'CSH' OR  TranType = 'CHG' OR  TranType = 'DR') and STATUS = 'P' AND trandate ='" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by Tranno ASC")
    Else
        Set rsPMIOS_ORD_HD = gconDMIS.Execute("Select * from PMIS_vw_ISS_HISTORY Where (TranType = 'CSH' OR  TranType = 'CHG') and STATUS = 'P' AND trandate ='" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by Tranno ASC")
    End If
    If Not rsPMIOS_ORD_HD.EOF And Not rsPMIOS_ORD_HD.BOF Then
        rsPMIOS_ORD_HD.MoveFirst: KIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsPMIOS_ORD_HD.EOF
            KIM = KIM + 1
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

            If CheckSJExisting(InvoiceTypeCode, Null2String(rsPMIOS_ORD_HD!TRANNO), Null2String(rsPMIOS_ORD_HD!TRANTYPE)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid1.AddItem IS_Exist & Chr(9) & InvoiceType & Chr(9) & Null2String(rsPMIOS_ORD_HD!TRANTYPE) & "-" & Null2String(rsPMIOS_ORD_HD!TRANNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPMIOS_ORD_HD!NETINVAMT)) & Chr(9) & Null2String(rsPMIOS_ORD_HD!CUSTNAME)

            rsPMIOS_ORD_HD.MoveNext
        Loop

        If KIM > 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If


    'kim = 0:
    Grid2.AutoRedraw = False
    'PURELY INTERNAL

    Set rsCSMIOS_REPOR = New ADODB.Recordset
    Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,CSMS_REPOR.RO_AMOUNT from CSMS_REPOR  WHERE INVOICE = 'INT RO' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' ORDER BY CSMS_REPOR.REP_OR ASC")
    If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
        rsCSMIOS_REPOR.MoveFirst:
        Do While Not rsCSMIOS_REPOR.EOF
            KIM = KIM + 1
            If CheckRefNoExisting("SI", Null2String(rsCSMIOS_REPOR!REP_OR)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!RO_AMOUNT)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!invoice))
            rsCSMIOS_REPOR.MoveNext
        Loop
    End If
    'CHARGE TO COMPANY / SALES
    Set rsCSMIOS_REPOR = New ADODB.Recordset
    Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETPRC) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE INVOICE <> 'INT RO' AND RO_AMOUNT = 0 AND DETAMT > 0 AND (WCODE = 'S' OR WCODE = 'C') AND invoice <> 'NO CHG' AND invoice <> 'PDI RO' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
    If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
        rsCSMIOS_REPOR.MoveFirst:
        Do While Not rsCSMIOS_REPOR.EOF
            KIM = KIM + 1
            If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!invoice)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!invoice))
            rsCSMIOS_REPOR.MoveNext
        Loop
    End If

    Set rsCSMIOS_REPOR = New ADODB.Recordset
    Set rsCSMIOS_REPOR = gconDMIS.Execute("Select * from CSMS_REPOR where RO_AMOUNT > 0 AND invoice <> 'INT RO' AND invoice <> 'NO CHG' AND invoice <> 'PDI RO' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by invoice ASC")
    If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
        rsCSMIOS_REPOR.MoveFirst:
        Grid2.AutoRedraw = False
        Do While Not rsCSMIOS_REPOR.EOF
            KIM = KIM + 1
            If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!invoice)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & Null2String(rsCSMIOS_REPOR!Niym)
            rsCSMIOS_REPOR.MoveNext
        Loop
    End If

    'PURELY WARRANTY
        Set rsCSMIOS_REPOR = New ADODB.Recordset
        Set rsCSMIOS_REPOR = gconDMIS.Execute("Select CSMS_REPOR.REP_OR,CSMS_REPOR.INVOICE,SUM(CSMS_RO_DET.DETAMT) AS AMOUNT from CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND invoice <> 'NO CHG' AND invoice <> 'PDI RO' AND invoice <> 'INT RO' and dte_comp = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' GROUP BY CSMS_REPOR.REP_OR,INVOICE")
        If Not rsCSMIOS_REPOR.EOF And Not rsCSMIOS_REPOR.BOF Then
            rsCSMIOS_REPOR.MoveFirst:
            Do While Not rsCSMIOS_REPOR.EOF
                KIM = KIM + 1
                If CheckSJExisting("SI", Null2String(rsCSMIOS_REPOR!invoice)) = True Then
                    IS_Exist = 1
                Else
                    IS_Exist = 0
                End If
                Grid2.AddItem IS_Exist & Chr(9) & Null2String(rsCSMIOS_REPOR!REP_OR) & Chr(9) & Null2String(rsCSMIOS_REPOR!invoice) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsCSMIOS_REPOR!amount)) & Chr(9) & SetRONiym(Null2String(rsCSMIOS_REPOR!invoice))
                rsCSMIOS_REPOR.MoveNext
            Loop
        End If
    If KIM > 0 Then Grid2.RemoveItem 1
    Grid2.AutoRedraw = True
    Grid2.Refresh

    Set rsSMIS_PURCHAGREE = New ADODB.Recordset
    Set rsSMIS_PURCHAGREE = gconDMIS.Execute("Select * from SMIS_PurchAgree Where STATUS = 'P' AND CONVERT(VarChar, DateReleased, 101)  = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' order by VI_NO ASC")
    If Not rsSMIS_PURCHAGREE.EOF And Not rsSMIS_PURCHAGREE.BOF Then
        rsSMIS_PURCHAGREE.MoveFirst: KIM = 0
        Grid3.AutoRedraw = False
        Do While Not rsSMIS_PURCHAGREE.EOF
            KIM = KIM + 1
            If CheckSJExisting("VI", Null2String(rsSMIS_PURCHAGREE!VI_NO)) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid3.AddItem IS_Exist & Chr(9) & Null2String(rsSMIS_PURCHAGREE!IGNKEY_NO) & Chr(9) & Null2String(rsSMIS_PURCHAGREE!VI_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSMIS_PURCHAGREE!NETSALESPRICE)) & Chr(9) & SetCustomerName(Null2String(rsSMIS_PURCHAGREE!code))
            rsSMIS_PURCHAGREE.MoveNext
        Loop
        If KIM > 0 Then Grid3.RemoveItem 1
        Grid3.AutoRedraw = True
        Grid3.Refresh
    End If
    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    Screen.MousePointer = 11
    Prg1.Max = 6:
    Prg1.Value = 1
    cmdCheck.Enabled = False
    cmdExit.Enabled = False
    Dim i
    For i = 1 To Grid1.Rows - 1

        Select Case UCase(Grid1.Cell(i, 2).Text)
        Case "PARTS"
            gconDMIS.Execute "SP_IMPORT_PMIS_SALES  'P', '" & Left(Grid1.Cell(i, 3).Text, 3) & "','" & Right(Grid1.Cell(i, 3).Text, 6) & "','" & LOGCODE & "'"
        Case "MATERIALS"
            gconDMIS.Execute "SP_IMPORT_PMIS_SALES  'M', '" & Left(Grid1.Cell(i, 3).Text, 3) & "','" & Right(Grid1.Cell(i, 3).Text, 6) & "','" & LOGCODE & "'"
        Case "ACCESSORIES"
        End Select
        Grid1.Cell(i, 1).Text = 1
        Grid1.TopRow = i
        DoEvents
    Next
    Screen.MousePointer = 0
    MsgBox "finished"
End Sub

Private Sub dtpTranDate_Change()
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
    Grid3.Rows = 1
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
    cmdCheck.Enabled = False
    Screen.MousePointer = 0

    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Function CheckImportedSales(xTranDate As String) As Boolean
'Update: ACL 11162009
    Dim xNoDays                                   As Integer
    Dim rsCheckSales                              As ADODB.Recordset
    Dim rsCheckTrans                              As ADODB.Recordset
    Set rsCheckSales = New ADODB.Recordset
    rsCheckSales.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='SJ' AND STATUS <> 'C' AND JDATE < '" & CDate(xTranDate) & "' ORDER BY JDATE DESC", gconDMIS, adOpenKeyset
    If Not rsCheckSales.EOF And Not rsCheckSales.BOF Then
        Set rsCheckTrans = New ADODB.Recordset
        rsCheckTrans.Open "SELECT * FROM (SELECT TOP 1 TRANDATE AS TRANDATE FROM PMIS_VW_ISS_HISTORY WHERE (TRANTYPE = 'CSH' OR  TRANTYPE = 'CHG') AND STATUS = 'P' AND TRANDATE > '" & Null2Date(rsCheckSales!JDate) & "' ORDER BY TRANDATE ASC " & _
                          "UNION SELECT TOP 1 DTE_COMP AS TRANDATE FROM CSMS_REPOR  WHERE INVOICE = 'INT RO' AND DTE_COMP > '" & Null2Date(rsCheckSales!JDate) & "' ORDER BY DTE_COMP ASC " & _
                          "UNION SELECT TOP 1 DTE_COMP AS TRANDATE FROM CSMS_REPOR WHERE RO_AMOUNT > 0 AND INVOICE <> 'INT RO' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND DTE_COMP > '" & Null2Date(rsCheckSales!JDate) & "' ORDER BY DTE_COMP ASC " & _
                          "UNION SELECT TOP 1 DTE_COMP AS TRANDATE FROM CSMS_REPOR INNER JOIN CSMS_RO_DET ON CSMS_REPOR.REP_OR = CSMS_RO_DET.REP_OR WHERE RO_AMOUNT = 0 AND DETAMT > 0 AND WCODE = 'W' AND INVOICE <> 'NO CHG' AND INVOICE <> 'PDI RO' AND INVOICE <> 'INT RO' AND DTE_COMP > '" & Null2Date(rsCheckSales!JDate) & "' ORDER BY DTE_COMP ASC " & _
                          "UNION SELECT TOP 1 CONVERT(VARCHAR, DATERELEASED, 101) AS TRANDATE FROM SMIS_PURCHAGREE WHERE STATUS = 'P' AND (MONTH(CONVERT(VARCHAR, DATERELEASED, 101))> '" & Month(CDate(rsCheckSales!JDate)) & "' AND DAY(CONVERT(VARCHAR, DATERELEASED, 101)) > '" & Day(CDate(rsCheckSales!JDate)) & "' AND YEAR(CONVERT(VARCHAR, DATERELEASED, 101)) > '" & Year(CDate(rsCheckSales!JDate)) & "') ORDER BY DATERELEASED ASC) X ORDER BY TRANDATE ASC", gconDMIS, adOpenKeyset
        If Not rsCheckTrans.EOF And Not rsCheckTrans.BOF Then
            TRANSACTIONDATE = Null2Date(rsCheckTrans!trandate)
            xNoDays = DateDiff("d", TRANSACTIONDATE, dtpTranDate)
            'If Format(CDate(dtpTranDate), "mm/dd/yyyy") > Format(CDate(TRANSACTIONDATE), "mm/dd/yyyy") Then
            If xNoDays > 0 Then
                CheckImportedSales = True
            End If
        End If
    End If
    Set rsCheckSales = Nothing
End Function

Function CheckToClearSales(xTranDate As String) As Boolean
    Dim rsCheckSJ                                 As ADODB.Recordset
    Set rsCheckSJ = New ADODB.Recordset
    rsCheckSJ.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'SJ' and Jdate > '" & CDate(xTranDate) & "' Order by Jdate ASC", gconDMIS, adOpenKeyset
    If Not rsCheckSJ.EOF And Not rsCheckSJ.BOF Then
        CheckToClearSales = True
    End If
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
    Dim iCount                                    As Integer
    For iCount = 1 To 31
        lab.Item(iCount - 1).Caption = iCount
    Next
End Sub

Sub InitCombo()
    Dim NoDays                                    As Integer
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

Private Sub IMPORT_ACCESSORIES()
    If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HAI" Or COMPANY_CODE = "HCI" Then
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

Function ReturnExpense(XXX As String, Optional YYY As String, Optional zzz As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If Trim(YYY) = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "'")
    Else
        If Trim(zzz) = "" Then
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "'")
        Else
            Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'EXPENSE' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "' AND TRANTYPE4 = '" & zzz & "'")
        End If
    End If

    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnExpense = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAccommodation(XXX As String)
    Dim rsAccommodation                           As ADODB.Recordset
    Set rsAccommodation = New ADODB.Recordset
    Set rsAccommodation = gconDMIS.Execute("SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE TRANTYPE3='ACCOMMODATION' AND TRANTYPE1='" & XXX & "'")
    If Not rsAccommodation.EOF And Not rsAccommodation.BOF Then
        ReturnAccommodation = Null2String(rsAccommodation!ACCTCODE)
    End If
    Set rsAccommodation = Nothing
End Function

Function CheckInternalRO(XXX As String) As Boolean
    Dim rsCheckInternalRO As ADODB.Recordset
    Set rsCheckInternalRO = New ADODB.Recordset
    rsCheckInternalRO.Open "SELECT INVOICENO FROM CMIS_REPOR WHERE INVOICE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsCheckInternalRO.EOF And Not rsCheckInternalRO.BOF Then
        CheckInternalRO = True
    End If
    Set rsCheckInternalRO = Nothing
End Function
