VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmAPJImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Payable Import Process"
   ClientHeight    =   7680
   ClientLeft      =   345
   ClientTop       =   1110
   ClientWidth     =   14055
   ControlBox      =   0   'False
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "APJImport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   14055
   Begin FlexCell.Grid Grid1 
      Height          =   4965
      Left            =   60
      TabIndex        =   60
      Top             =   1140
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   8758
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   2
   End
   Begin VB.CheckBox chkService 
      Caption         =   "SERVICE SUBLET"
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
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   630
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CheckBox chkVehicles 
      Caption         =   "VEHICLES"
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
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   630
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CheckBox chkParts 
      Caption         =   "PARTS, ACCS && MATERIALS"
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
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   630
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   3195
      TabIndex        =   51
      Top             =   6120
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
         Picture         =   "APJImport.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   30
         Width           =   915
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
         ItemData        =   "APJImport.frx":138C
         Left            =   780
         List            =   "APJImport.frx":138E
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   30
         Width           =   1455
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
         TabIndex        =   52
         Top             =   420
         Width           =   1005
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
         TabIndex        =   56
         Top             =   120
         Width           =   585
      End
      Begin VB.Label Label6 
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
         TabIndex        =   55
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   3210
      ScaleHeight     =   465
      ScaleWidth      =   4905
      TabIndex        =   12
      Top             =   6810
      Width           =   4905
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   30
         MouseIcon       =   "APJImport.frx":1390
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label8 
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
         TabIndex        =   17
         Top             =   90
         Width           =   1200
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
         MouseIcon       =   "APJImport.frx":169A
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label9 
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
         TabIndex        =   15
         Top             =   90
         Width           =   720
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
         MouseIcon       =   "APJImport.frx":19A4
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   90
         Width           =   315
      End
      Begin VB.Label Label10 
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
         TabIndex        =   13
         Top             =   90
         Width           =   1305
      End
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
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   150
      Width           =   1935
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
      Left            =   3630
      MouseIcon       =   "APJImport.frx":1CAE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Process Import of SALES"
      Top             =   90
      Width           =   2010
   End
   Begin MSComCtl2.DTPicker dtpTranDate 
      Height          =   405
      Left            =   1770
      TabIndex        =   9
      Top             =   90
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
      Format          =   47710209
      CurrentDate     =   38216
   End
   Begin wizProgBar.Prg progCPB 
      Height          =   315
      Left            =   9480
      TabIndex        =   10
      Top             =   6480
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   556
      Picture         =   "APJImport.frx":1E00
      BarPicture      =   "APJImport.frx":1E1C
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
      MouseIcon       =   "APJImport.frx":1E38
      MousePointer    =   99  'Custom
      Picture         =   "APJImport.frx":1F8A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   6840
      Width           =   735
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
      Left            =   12510
      MouseIcon       =   "APJImport.frx":22F0
      MousePointer    =   99  'Custom
      Picture         =   "APJImport.frx":2442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process Importing of Purchases"
      Top             =   6840
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   3180
      ScaleHeight     =   765
      ScaleWidth      =   5925
      TabIndex        =   19
      Top             =   6150
      Width           =   5925
      Begin VB.Label lab 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   5460
         MouseIcon       =   "APJImport.frx":26DD
         MousePointer    =   99  'Custom
         TabIndex        =   50
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
         MouseIcon       =   "APJImport.frx":29E7
         MousePointer    =   99  'Custom
         TabIndex        =   49
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
         Index           =   28
         Left            =   4740
         MouseIcon       =   "APJImport.frx":2CF1
         MousePointer    =   99  'Custom
         TabIndex        =   48
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
         Index           =   27
         Left            =   4380
         MouseIcon       =   "APJImport.frx":2FFB
         MousePointer    =   99  'Custom
         TabIndex        =   47
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
         MouseIcon       =   "APJImport.frx":3305
         MousePointer    =   99  'Custom
         TabIndex        =   46
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
         MouseIcon       =   "APJImport.frx":360F
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
         Index           =   24
         Left            =   3300
         MouseIcon       =   "APJImport.frx":3919
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
         Index           =   23
         Left            =   2940
         MouseIcon       =   "APJImport.frx":3C23
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
         Index           =   22
         Left            =   2580
         MouseIcon       =   "APJImport.frx":3F2D
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
         Index           =   21
         Left            =   2220
         MouseIcon       =   "APJImport.frx":4237
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
         MouseIcon       =   "APJImport.frx":4541
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
         Index           =   19
         Left            =   1500
         MouseIcon       =   "APJImport.frx":484B
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
         Index           =   18
         Left            =   1140
         MouseIcon       =   "APJImport.frx":4B55
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
         Index           =   17
         Left            =   780
         MouseIcon       =   "APJImport.frx":4E5F
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
         Index           =   16
         Left            =   420
         MouseIcon       =   "APJImport.frx":5169
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
         Index           =   15
         Left            =   60
         MouseIcon       =   "APJImport.frx":5473
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
         Index           =   14
         Left            =   5100
         MouseIcon       =   "APJImport.frx":577D
         MousePointer    =   99  'Custom
         TabIndex        =   34
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
         MouseIcon       =   "APJImport.frx":5A87
         MousePointer    =   99  'Custom
         TabIndex        =   33
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
         MouseIcon       =   "APJImport.frx":5D91
         MousePointer    =   99  'Custom
         TabIndex        =   32
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
         MouseIcon       =   "APJImport.frx":609B
         MousePointer    =   99  'Custom
         TabIndex        =   31
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
         MouseIcon       =   "APJImport.frx":63A5
         MousePointer    =   99  'Custom
         TabIndex        =   30
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
         MouseIcon       =   "APJImport.frx":66AF
         MousePointer    =   99  'Custom
         TabIndex        =   29
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
         MouseIcon       =   "APJImport.frx":69B9
         MousePointer    =   99  'Custom
         TabIndex        =   28
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
         MouseIcon       =   "APJImport.frx":6CC3
         MousePointer    =   99  'Custom
         TabIndex        =   27
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
         MouseIcon       =   "APJImport.frx":6FCD
         MousePointer    =   99  'Custom
         TabIndex        =   26
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
         MouseIcon       =   "APJImport.frx":72D7
         MousePointer    =   99  'Custom
         TabIndex        =   25
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
         MouseIcon       =   "APJImport.frx":75E1
         MousePointer    =   99  'Custom
         TabIndex        =   24
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
         MouseIcon       =   "APJImport.frx":78EB
         MousePointer    =   99  'Custom
         TabIndex        =   23
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
         MouseIcon       =   "APJImport.frx":7BF5
         MousePointer    =   99  'Custom
         TabIndex        =   22
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
         MouseIcon       =   "APJImport.frx":7EFF
         MousePointer    =   99  'Custom
         TabIndex        =   21
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
         Index           =   0
         Left            =   60
         MouseIcon       =   "APJImport.frx":8209
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   60
         Width           =   315
      End
   End
   Begin FlexCell.Grid Grid2 
      Height          =   4965
      Left            =   4710
      TabIndex        =   61
      Top             =   1140
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
      Left            =   9360
      TabIndex        =   62
      Top             =   1140
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
      Caption         =   "SERVICE SUBLET"
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
      Left            =   9360
      TabIndex        =   11
      Top             =   615
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only Un-Imported Invoices can be Imported"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   150
      TabIndex        =   8
      Top             =   7320
      Width           =   7995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PARTS, ACCS && MATERIALS"
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
      Left            =   60
      TabIndex        =   7
      Top             =   630
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VEHICLES"
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
      TabIndex        =   6
      Top             =   615
      Width           =   4575
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
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   1875
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
      Left            =   9480
      TabIndex        =   4
      Top             =   6180
      Width           =   5835
   End
End
Attribute VB_Name = "frmAPJImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TRANSACTIONDATE                               As String
Dim Indx                                          As Integer
Dim xTranDate                                     As String
Dim ATCRate2                                      As Double

Function SetTransaction(XXX As Variant) As String
    Dim rsSBOOKTransaction                        As ADODB.Recordset
    Set rsSBOOKTransaction = New ADODB.Recordset
    Set rsSBOOKTransaction = gconDMIS.Execute("Select * from SBOOK Where BOOK = 'A' and CODE = '" & XXX & "'")
    If Not rsSBOOKTransaction.EOF And Not rsSBOOKTransaction.BOF Then
        SetTransaction = Null2String(rsSBOOKTransaction!DESCNAME)
    End If
    Set rsSBOOKTransaction = Nothing
End Function

Function SetOtherTransaction(XXX As Variant) As String
    Dim rsSBOOKOtherTransaction                   As ADODB.Recordset
    Set rsSBOOKOtherTransaction = New ADODB.Recordset
    Set rsSBOOKOtherTransaction = gconDMIS.Execute("Select * from SBOOK Where BOOK = 'D' and CODE = '" & XXX & "'")
    If Not rsSBOOKOtherTransaction.EOF And Not rsSBOOKOtherTransaction.BOF Then
        SetOtherTransaction = Null2String(rsSBOOKOtherTransaction!DESCNAME)
    End If
    Set rsSBOOKOtherTransaction = Nothing
End Function

Function Setacctname(VVV As Variant) As String
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
    Set rsJournal_HD = gconDMIS.Execute("Select CAST(VoucherNo AS int) AS MAX_VOUCHERNO from AMIS_Journal_HD Where Jtype = 'APJ' Order by MAX_VOUCHERNO desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!MAX_VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function CheckAPJExisting(VarInvoiceNo As String, VarTYPE As Variant) As Boolean
    Dim rsCheckAPJ_Journal_HD                     As ADODB.Recordset
    Set rsCheckAPJ_Journal_HD = New ADODB.Recordset
    If VarTYPE = "PARTS" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'PARTS' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "ACCESSORIES" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'ACCESSORIES' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "MATERIALS" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'MATERIALS' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "VEHICLES" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'VEHICLES' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
     ElseIf VarTYPE = "SUBLET" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = 'SUBLET' AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    ElseIf VarTYPE = "" Then
        Set rsCheckAPJ_Journal_HD = gconDMIS.Execute("Select VoucherNo,Jtype from AMIS_Journal_HD where JTYPE = 'APJ' AND InvoiceType = NULL AND InvoiceNo = " & N2Str2Null(VarInvoiceNo))
    End If
    If Not rsCheckAPJ_Journal_HD.EOF And Not rsCheckAPJ_Journal_HD.BOF Then
        CheckAPJExisting = True
    Else
        CheckAPJExisting = False
    End If
    Set rsCheckAPJ_Journal_HD = Nothing
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

Function ReturnAP_AccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AP' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAP_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAR_AccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'ARPARTS' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAR_AccountCode = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnAR_AccountCode2(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'ARUNITS' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAR_AccountCode2 = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInPutTax()
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnWithholdingTax()
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'EXPANDED'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnWithholdingTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInComeTax(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INCOME TAX' AND TRANTYPE2 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInComeTax = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInventory(XXX As String, Optional YYY As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If Trim(YYY) = "" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & XXX & "'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE3 = 'INVENTORY' AND TRANTYPE2 = '" & XXX & "' AND TRANTYPE1 = '" & YYY & "'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInventory = Null2String(rsChartAccount!ACCTCODE)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnExpense(XXX As String, Optional YYY As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
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

Function SetSellingDealerName(XXX As String) As String
    Dim rsSellingDealer                           As ADODB.Recordset
    Set rsSellingDealer = New ADODB.Recordset
    Set rsSellingDealer = gconDMIS.Execute("Select * from CSMS_SellingDealer Where DealerCode = '" & XXX & "'")
    If Not rsSellingDealer.EOF And Not rsSellingDealer.BOF Then
        SetSellingDealerName = Null2String(rsSellingDealer!DealerName)
    End If
End Function

Function ReturnPartNo(nard As String) As String
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT stock_ord from PMIS_Tdaytran where tranno='" & nard & "' and  status ='P'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        ReturnPartNo = Null2String(RS!stock_ord)
    End If
    Set RS = Nothing
End Function

Function CheckIfORIG(ARNIE As String) As Boolean
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT GENUINE From PMIS_stockmas where Stockno='" & ARNIE & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        If Null2String(RS!genuine) = "Y" Then
            CheckIfORIG = True
        Else
            CheckIfORIG = False
        End If
    End If
    Set RS = Nothing
End Function

Function ReturnCode(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    Dim MARK                                      As String

    MARK = (Replace(XXX, " ", ""))

    SQL = "SELECT Code, replace(Nameofvendor,' ','') from ALL_Vendor_table where REPLACE(Nameofvendor,' ','') like '" & MARK & "%'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.BOF And Not RS.EOF Then
        ReturnCode = Null2String(RS!Code)
    Else
        ReturnCode = ""
    End If
    Set RS = Nothing
End Function

Function CheckSubletifExist(MARK As String, EVAN As String) As Boolean
'Update By BTT - 07092008
    Dim SQL, nard                                 As String
    Dim RS                                        As New ADODB.Recordset

    nard = "SELECT * from AMIS_pv_detail where MRR_no='" & MARK & "' and po_no='" & EVAN & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(nard)

    If Not RS.EOF And Not RS.BOF Then
        CheckSubletifExist = True
    Else
        CheckSubletifExist = False
    End If
    Set RS = Nothing
End Function

Function GetVendorTerms(XXX As String) As String
'Update By BTT - 07092008
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    SQL = "SELECT terms from all_vendor_table where code='" & XXX & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        GetVendorTerms = Null2String(RS!TERMS)
    Else
        GetVendorTerms = ""
    End If
    Set RS = Nothing
End Function

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Supplier"

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
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Supplier"

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
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Contractor"

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

End Sub

Function ImportSublet() As Boolean
    On Error GoTo ErrorCode
    Dim rsJournal_HDDup                           As New ADODB.Recordset
    Dim rsSublet                                  As New ADODB.Recordset
    Dim RsSublet_Det                              As New ADODB.Recordset
    Dim rsATC                                     As ADODB.Recordset
    Dim SQL                                       As String
    Dim GridImports                               As Integer
    Dim J_JDATE                                   As String
    Dim J_VOUCHERNO                               As String
    Dim J_JVOUCHERNO                              As String
    Dim J_JTYPE                                   As String
    Dim J_JNO                                     As String
    Dim J_REMARKS                                 As String
    Dim J_VENDORCODE                              As String
    Dim J_CUSTOMERCODE                            As String
    Dim J_CUSTOMERNAME                            As String
    Dim J_DEBIT                                   As Double
    Dim J_CREDIT                                  As Double
    Dim J_TAX                                     As Double
    Dim J_OUTBALANCE                              As Double
    Dim J_AMOUNTTOPAY                             As Double
    Dim J_INVOICEAMT                              As Double
    Dim J_BALANCE                                 As Double
    Dim J_AMOUNTPAID                              As Double
    Dim J_STATUS                                  As String
    Dim J_JITEMNO                                 As String
    Dim J_CHECKNO                                 As String
    Dim J_INVOICEDATE                             As String
    Dim J_DUEDATE                                 As String
    Dim J_PAYTYPE                                 As String
    Dim J_INVOICETYPE                             As String
    Dim J_INVOICENO                               As String
    Dim J_CHECKDATE                               As String
    Dim J_BANKCODE                                As String
    Dim J_REFNO                                   As String
    Dim J_REFDATE                                 As String
    Dim J_TERMS                                   As String
    Dim J_DEALER                                  As String
    Dim J_ACCT_CODE                               As String
    Dim J_ACCT_NAME                               As String
    Dim PV_ITEMNO                                 As String
    Dim PV_MRRNO                                  As String
    Dim PV_PONO                                   As String
    Dim PV_INVNO                                  As String
    Dim PV_PRODNO                                 As String
    Dim PV_STATUS                                 As String
    Dim J_ATC                                     As String
    Dim J_RATE                                    As Double
    Dim J_TAXBASE                                 As Double
    Dim J_GROSS                                   As Double
    Dim J_NET                                     As Double
    Dim PV_AMOUNT                                 As Double
    Dim i                                         As Integer
    Dim J_PAIDSTATUS                              As String
    Dim J_RECEIVESTATUS                           As String
    J_JTYPE = "'APJ'"
    Dim TOTAL_CREDIT                              As Double
    Dim TOTAL_DEBIT                               As Double
    Dim TheRO                                     As String
    Dim ThePO                                     As String
    Dim TheSublet_Cost                            As Double
    Dim TheSublet_Vat                             As Double
    Dim TheSublet_Net                             As Double
    Dim TheRRDate                                 As String
    Dim TheRRNO                                   As String
    J_CUSTOMERCODE = "'999999'"
    Dim TheINVOICE_no                             As String
    Dim WCode                                     As String
    Dim TERMS                                     As String
    Dim VENDOR                                    As String
    Dim SubletType                                As String
    Dim DetCodeLen                                As Integer
    Dim J_ITEMCOUNT                               As Integer
    Dim ATCRate                                   As Double
    TOTAL_CREDIT = 0: TOTAL_DEBIT = 0
    i = 0
    For GridImports = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImports, 1).Text) = 0 Then
            SQL = "SELECT * from CSMS_PO_RC_HD where RC_NO='" & Grid3.Cell(GridImports, 3).Text & "' and status='P' and RC_DATE='" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "'"
            Set rsSublet = New ADODB.Recordset
            Set rsSublet = gconDMIS.Execute(SQL)
            If Not rsSublet.EOF And Not rsSublet.BOF Then
                TheRO = Null2String(rsSublet!ro_no)
                ThePO = Null2String(rsSublet!po_no)
                TheRRNO = Null2String(rsSublet!RC_NO)
                TheINVOICE_no = Null2String(rsSublet!Invoice_no)
                J_CUSTOMERNAME = "NULL"
                J_INVOICETYPE = "'SUBLET'"
                J_VENDORCODE = N2Str2Null(rsSublet!CONTRACTOR_CODE)
                TERMS = GetVendorTerms(rsSublet!CONTRACTOR_CODE)
                'J_VENDORCODE = N2Str2Null(ReturnCode(RsSublet_Det!TECHNICIAN))
                'SUBLET DETAIL LOOKUP
                Set RsSublet_Det = New ADODB.Recordset
                Set RsSublet_Det = gconDMIS.Execute("SELECT * FROM CSMS_PO_RC_DT where PO_no='" & ThePO & "'")
                If Not RsSublet_Det.EOF And Not RsSublet_Det.BOF Then
                    TheSublet_Cost = NumericVal(RsSublet_Det!contractamount)    'COST
                    TheSublet_Vat = NumericVal(TheSublet_Cost) / 1.12 * 0.12
                    TheSublet_Net = NumericVal(TheSublet_Cost) - TheSublet_Vat
                    TheRRDate = Null2String(rsSublet!rc_date)
                    WCode = Null2String(RsSublet_Det!WCode)
                    DetCodeLen = Len(Null2String(RsSublet_Det!DETCDE))
                    SubletType = UCase(Mid(Null2String(RsSublet_Det!DETCDE), 3, DetCodeLen))
                End If
                'TERMS = GetVendorTerms(ReturnCode(RsSublet_Det!TECHNICIAN))
               
                'HEADER

                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If


                'VENDOR = Null2String(RsSublet_Det!TECHCODE)
                VENDOR = Null2String(rsSublet!CONTRACTOR_CODE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_OUTBALANCE = 0
                J_INVOICEDATE = N2Date2Null(rsSublet!rc_date): J_BALANCE = 0: J_AMOUNTPAID = 0
                J_DUEDATE = N2Date2Null(TheRRDate)
                J_PAYTYPE = "'" & TERMS & "D'": J_STATUS = "'N'"
                J_TERMS = "'" & TERMS & "D'": J_DEALER = "NULL"
                J_CHECKDATE = "NULL": J_BANKCODE = "NULL"
                J_INVOICEAMT = NumericVal(TheSublet_Cost)
                J_PAIDSTATUS = "'N'": J_RECEIVESTATUS = "'N'"
                J_CHECKNO = "NULL": J_REFDATE = "NULL"
                J_AMOUNTTOPAY = Round(NumericVal(TheSublet_Cost), 2)
                J_JDATE = N2Date2Null(TheRRDate)
'
'                J_REFNO = N2Str2Null(TheRRNO)
'                J_INVOICENO = N2Str2Null(TheINVOICE_no)
                
                J_REFNO = N2Str2Null(TheINVOICE_no)
                J_INVOICENO = N2Str2Null(TheRRNO)
                'AP
                J_REMARKS = N2Str2Null("To Record Sublet Recieving with RR No:" + TheRRNO + " (And Ro No " + TheRO + ")")
                J_ITEMCOUNT = 0

                'INVENTORY
                '                If COMPANY_CODE = "HPI" Then
                '                Else
                'J_JITEMNO = "'0001'"

                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HMH" Then
                    '                    J_ACCT_CODE = "'11-05006-00'"
                    '                    J_ACCT_NAME = N2Str2Null(Setacctname("'11-05006-00'"))
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET", "SUBLET"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET", "SUBLET")))
                ElseIf COMPANY_CODE = "HPI" Then
                    J_ACCT_CODE = "'41-02001-20'"
                    J_ACCT_NAME = N2Str2Null(Setacctname("'41-02001-20'"))
                ElseIf COMPANY_CODE = "HCI" Then
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
                Else                                       ' HSB
                    '                    J_ACCT_CODE = "'11-05006-21'"
                    '                    J_ACCT_NAME = N2Str2Null(Setacctname("'11-05006-21'"))
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("SUBLET", "SUBLET"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SUBLET", "SUBLET")))
                End If
                'J_DEBIT = Round(NumericVal(TheSublet_Cost) / 1.12, 2) - TheSublet_Vat
                If COMPANY_CODE = "HCI" Then
                    If ReturnNONVATVendor(J_VENDORCODE) = False Then
                        J_DEBIT = Round(TheSublet_Cost, 2) - Round(TheSublet_Vat, 2)
                    Else
                        J_DEBIT = Round(TheSublet_Cost, 2)
                    End If
                Else
                    J_DEBIT = Round(TheSublet_Cost, 2) - Round(TheSublet_Vat, 2)
                End If
                J_CREDIT = 0: J_TAX = 0: J_ATC = N2Str2Null("")
                J_RATE = 0: J_TAXBASE = 0
                TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                'SUBLET DETAIL
                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                '                End If

                'TAX
                '                If (WCode = "C" And COMPANY_CODE <> "HPI") Or (WCode = "C" And COMPANY_CODE <> "HCI") Then
                ' Do Nothing
                '                Else
                'J_JITEMNO = "'0002'"

                If (WCode = "C" And COMPANY_CODE = "HPI") Or (WCode = "C" And COMPANY_CODE = "HCI") Then
                    If COMPANY_CODE = "HCI" Then
                        If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = Round(NumericVal(Round((TheSublet_Cost / 1.12), 2) * 0.12), 2)
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            J_GROSS = 0
                            J_NET = 0
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                        End If
                    End If
                ElseIf WCode <> "C" Then
                    If COMPANY_CODE = "HCI" Then
                        If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = Round(NumericVal(Round((TheSublet_Cost / 1.12), 2) * 0.12), 2)
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            J_GROSS = 0
                            J_NET = 0
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                            SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                            "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                            " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                            ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                            ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                            gconDMIS.Execute SQL_STATEMENT
                            TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                            NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                        End If
                    Else
                        If WCode <> "C" And COMPANY_CODE = "HCI" And ReturnNONVATVendor(J_VENDORCODE) = True Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = Round(NumericVal(Round((TheSublet_Cost / 1.12), 2) * 0.12), 2)
                            J_CREDIT = 0
                            J_TAX = 0
                            J_GROSS = 0
                            J_NET = 0
                            J_STATUS = "'N'"
                            J_GROSS = 0
                            J_NET = 0
                            TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                        End If
                        If COMPANY_CODE <> "HOT" Then
                            If VENDOR = "M00002" Or COMPANY_CODE = "M00003" Then
                            Else
                                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                                gconDMIS.Execute SQL_STATEMENT
                                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                            End If
                        End If
                    End If


                End If


                '0003

                J_ITEMCOUNT = J_ITEMCOUNT + 1
                J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HMH" Or COMPANY_CODE = "HCI" Then
                    J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                ElseIf COMPANY_CODE = "HPI" Then
                    J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                Else                                       ' HSB
                    J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("SERVICE"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("SERVICE")))
                End If
                J_DEBIT = 0
                'J_CREDIT = Round(NumericVal(TheSublet_Cost) / 1.12, 2)
                If COMPANY_CODE = "HCI" Then
                    Dim alhci
                    Dim alhci2
                    Set rsATC = New ADODB.Recordset
                    ATCRate = 0
                    rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                    If Not rsATC.EOF And Not rsATC.BOF Then
                        If SubletType = "LABOR" Then
                            ATCRate = ReturnATCRate("WI 160")
                        Else
                            ATCRate = ReturnATCRate("WC 158")
                        End If
                    End If
                    Set rsATC = Nothing
                    If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        J_CREDIT = Round(NumericVal(TheSublet_Cost) - Round(NumericVal(TheSublet_Cost) * ATCRate, 2), 2)
                    Else
                        J_CREDIT = Round(NumericVal(TheSublet_Cost - Round(NumericVal(TheSublet_Cost - (TheSublet_Vat)), 2) * ATCRate), 2)

                    End If
                    J_TAX = 0
                    J_ATC = N2Str2Null("")
                    J_RATE = 0
                    J_TAXBASE = 0
                    J_GROSS = 0
                    J_NET = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                ElseIf COMPANY_CODE = "HOT" Then
                    Set rsATC = New ADODB.Recordset
                    ATCRate = 0
                    rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                    If Not rsATC.EOF And Not rsATC.BOF Then
                        If SubletType = "LABOR" Then
                            ATCRate = ReturnATCRate("WI 160")
                        Else
                            ATCRate = ReturnATCRate("WC 158")
                        End If
                    End If
                    Set rsATC = Nothing
                    If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        J_CREDIT = Round(NumericVal(TheSublet_Cost) - Round(NumericVal(TheSublet_Cost) * ATCRate, 2), 2)
                    Else
                        J_CREDIT = Round(NumericVal(TheSublet_Cost - Round(NumericVal(TheSublet_Cost - (TheSublet_Vat)), 2) * ATCRate), 2)
                    End If
                Else
                    J_CREDIT = Round(NumericVal(TheSublet_Cost))
                End If
                If VENDOR = "M00002" Or VENDOR = "M00003" Then
                    J_TAX = TheSublet_Cost / 1.12 * 0.12
                    J_CREDIT = Round((NumericVal(TheSublet_Cost) - J_TAX), 2)
                ElseIf VENDOR = "S00002" Then
                    J_CREDIT = Round(NumericVal(TheSublet_Cost), 2)
                End If
                J_TAX = 0
                J_ATC = N2Str2Null("")
                J_RATE = 0
                J_TAXBASE = 0
                J_GROSS = 0
                J_NET = 0
                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                'SUBLET DETAIL

                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                If COMPANY_CODE = "HCI" Then
                    Set rsATC = New ADODB.Recordset
                    ATCRate = 0
                    rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                    If Not rsATC.EOF And Not rsATC.BOF Then
                        If SubletType = "LABOR" Then
                            ATCRate = ReturnATCRate("WI 160")
                        Else
                            ATCRate = ReturnATCRate("WC 158")
                        End If
                    End If
                    Set rsATC = Nothing
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnWithholdingTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnWithholdingTax()))
                    J_DEBIT = 0
                    If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        J_CREDIT = Round(NumericVal(TheSublet_Cost), 2) * ATCRate
                        J_TAXBASE = TheSublet_Cost
                    Else
                        J_CREDIT = Round(NumericVal(TheSublet_Cost - TheSublet_Vat), 2) * ATCRate
                        J_TAXBASE = Round(NumericVal(TheSublet_Cost - TheSublet_Vat), 2)
                    End If
                    J_TAX = 0
                    'J_ATC = 0
                    If SubletType = "LABOR" Then
                        J_ATC = "WI 160"
                    Else
                        J_ATC = "WC 158"
                    End If
                    J_RATE = ATCRate * 100

                    J_GROSS = 0
                    J_NET = 0
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                ElseIf COMPANY_CODE = "HOT" Then
                    If VENDOR = "M00002" Or VENDOR = "M00003" Or VENDOR = "S00002" Then
                    Else
                        Set rsATC = New ADODB.Recordset
                        ATCRate = 0
                        rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                        If Not rsATC.EOF And Not rsATC.BOF Then
                            If SubletType = "LABOR" Then
                                ATCRate = ReturnATCRate("WI 160")
                            Else
                                ATCRate = ReturnATCRate("WC 158")
                            End If
                        End If
                        Set rsATC = Nothing
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnWithholdingTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnWithholdingTax()))
                        J_DEBIT = 0
                        If ReturnNONVATVendor(J_VENDORCODE) = True Then
                            J_CREDIT = Round(NumericVal(TheSublet_Cost), 2) * ATCRate
                            J_TAXBASE = TheSublet_Cost
                        Else
                            J_CREDIT = Round(NumericVal(TheSublet_Cost - TheSublet_Vat), 2) * ATCRate
                            J_TAXBASE = Round(NumericVal(TheSublet_Cost - TheSublet_Vat), 2)
                        End If
                        J_AMOUNTTOPAY = J_AMOUNTTOPAY - J_CREDIT
                        If VENDOR = "M00002" Or VENDOR = "M00003" Then
                            J_TAX = (J_AMOUNTTOPAY) / 1.12 * 0.12
                            J_AMOUNTTOPAY = J_AMOUNTTOPAY - J_TAX
                        ElseIf VENDOR = "S00002" Then
                            J_AMOUNTTOPAY = J_AMOUNTTOPAY
                        End If
                        J_TAX = 0
                        'J_ATC = 0
                        If SubletType = "LABOR" Then
                            J_ATC = "WI 160"
                        Else
                            J_ATC = "WC 158"
                        End If
                        J_RATE = ATCRate * 100

                        J_GROSS = 0
                        J_NET = 0
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                        'SUBLET DETAIL
                    End If
                

                SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ATC,RATE,TAXBASE)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & N2Str2Null(J_ATC) & "," & J_RATE & "," & J_TAXBASE & ")"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                End If


                'SUBLET DETAIL
                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(TheRRNO)
                PV_PONO = N2Str2Null(ThePO)
                PV_INVNO = N2Str2Null(TheINVOICE_no)
                PV_PRODNO = "NULL"
                PV_AMOUNT = Round(NumericVal(TheSublet_Cost), 2)
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                "(VoucherNo,JDATE,JTYPE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                                " values (" & J_JVOUCHERNO & "," & J_JDATE & "," & J_JTYPE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                ", " & PV_STATUS & ")"

                gconDMIS.Execute SQL_STATEMENT

                'TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_PV_Detail", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


                'SUBLET HEADER
                gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                 " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                 " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "X", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                
                'Grid3.Cell(i, 1).Text = 1
                Grid3.Cell(GridImports, 1).Text = 1
            End If
        End If
        i = i + 1
    Next

    ImportSublet = True
    Exit Function
ErrorCode:
    ImportSublet = False
End Function

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

Private Sub chkParts_Click()
    If chkParts.Value = False Then
        If MsgBox("Are you sure Parts, Accessories and Materials will not be included during importing?", vbQuestion + vbYesNo, "Import Parts, Accessories & Materials") = vbYes Then
            chkParts.Value = 0
        Else
            chkParts.Value = 1
        End If
    End If
End Sub

Private Sub chkService_Click()
    If chkService.Value = False Then
        If MsgBox("Are you sure Parts, Accessories and Materials will not be included during importing?", vbQuestion + vbYesNo, "Import Parts, Accessories & Materials") = vbYes Then
            chkService.Value = 0
        Else
            chkService.Value = 1
        End If
    End If
End Sub

Private Sub chkVehicles_Click()
    If chkVehicles.Value = False Then
        If MsgBox("Are you sure Parts, Accessories and Materials will not be included during importing?", vbQuestion + vbYesNo, "Import Parts, Accessories & Materials") = vbYes Then
            chkVehicles.Value = 0
        Else
            chkVehicles.Value = 1
        End If
    End If
End Sub

Private Sub cmdCheck_Click()

    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Process", "IMPORT PURCHASES") = False Then Exit Sub

    Screen.MousePointer = 11
    Dim str_MSG                                   As String


    str_MSG = "Error Appear In During @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Imported Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    If COMPANY_CODE = "HCC" Then
        ImportPurchasesNew
    Else
        gconDMIS.BeginTrans
        'HEADER
        If ImportPurchases = False Then
            str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Purchases")
            MsgBox str_MSG, vbCritical, "Import Error "
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
        'Sublet
        If ImportSublet = False Then
            str_MSG = Replace(str_MSG, "@ACL09182716350", "Import Sublet")
            MsgBox str_MSG, vbCritical, "Import Error "
            cmdExit.Enabled = True
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
    
        gconDMIS.CommitTrans
    End If
    
    Screen.MousePointer = 0
    Call cmdShowImp_Click
    MsgBox "Import Successfully Completed!", vbInformation, "Finish"
    LogAudit "R", "ACCOUNTS PAYABLE IMPORT", dtpTranDate

ErrorCode:
'    SaveLogFile
    ShowVBError
End Sub

Sub ImportPurchasesNew()
Dim i                                         As Integer
Dim GridImport                                As Integer
    'PARTS / ACCESSORIES / MATERIALS IMPORTING =============================================================
    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Call ImportPurchase2("PURCHASES", Grid1.Cell(GridImport, 2).Text, Grid1.Cell(GridImport, 3).Text, CDate(dtpTranDate))
            Grid1.Cell(GridImport, 1).Text = 1
        End If
        
        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next
    
    'VEHICLES IMPORTING==================================================================================================================================================================================================================================================================================================================================
    i = 0
    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then
            Call ImportPurchase2("PURCHASES", Grid2.Cell(GridImport, 2).Text, Grid2.Cell(GridImport, 3).Text, CDate(dtpTranDate))
            Grid2.Cell(GridImport, 1).Text = 1
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
    Next
'
'    'SUBLET IMPORTING==================================================================================================================================================================================================================================================================================================================================
    i = 0
    For GridImport = 1 To Grid3.Rows - 1
        If N2Str2Zero(Grid3.Cell(GridImport, 1).Text) = 0 Then
            Call ImportPurchase2("PURCHASES", Grid3.Cell(GridImport, 2).Text, Grid3.Cell(GridImport, 3).Text, CDate(dtpTranDate))
            Grid3.Cell(GridImport, 1).Text = 1
        End If
        i = i + 1
        progCPB.Value = (i / (Grid3.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
    Next
End Sub

Sub ImportPurchase2(xTRANTYPE As String, XTYPE As String, xTRANNO As String, xTranDate As Date)
    Dim CMD                                       As New ADODB.Command

    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "USP_IMPORTING_TEMPLATE"

    With CMD.Parameters
        .Append CMD.CreateParameter("@TRANTYPE", adVarChar, adParamInput, 15, xTRANTYPE)
        .Append CMD.CreateParameter("@TYPE", adVarChar, adParamInput, 15, XTYPE)
        .Append CMD.CreateParameter("@TRANNO", adVarChar, adParamInput, 10, xTRANNO)
        .Append CMD.CreateParameter("@TRANDATE", adDate, adParamInput, 8, xTranDate)
    End With

    CMD.Execute
End Sub

Function ImportPurchases() As Boolean
    On Error GoTo ErrorCode

    Dim J_JDATE                                   As String
    Dim J_VOUCHERNO                               As String
    Dim J_JTYPE                                   As String
    Dim J_JNO
    Dim J_REMARKS                                 As String
    Dim J_VENDORCODE                              As String
    Dim J_CUSTOMERCODE                            As String
    Dim J_OUTBALANCE                              As Double
    Dim J_AMOUNTTOPAY                             As Double
    Dim J_INVOICEAMT                              As Double
    Dim J_BALANCE                                 As Double
    Dim J_AMOUNTPAID                              As Double
    Dim J_CHECKNO                                 As String
    Dim J_INVOICEDATE                             As String
    Dim J_DUEDATE                                 As String
    Dim J_PAYTYPE                                 As String
    Dim J_INVOICETYPE                             As String
    Dim J_INVOICENO                               As String
    Dim J_CHECKDATE                               As String
    Dim J_BANKCODE                                As String
    Dim J_REFNO                                   As String
    Dim J_REFDATE                                 As String
    Dim J_TERMS                                   As String
    Dim J_DEALER                                  As String
    Dim J_PAIDSTATUS                              As String
    Dim J_RECEIVESTATUS                           As String

    'DETAIL
    Dim J_ACCT_CODE                               As String
    Dim J_ACCT_NAME                               As String
    Dim J_DEBIT                                   As Double
    Dim J_CREDIT                                  As Double
    Dim J_TAX                                     As Double
    Dim J_GROSS                                   As Double
    Dim J_NET                                     As Double
    Dim J_STATUS                                  As String
    Dim J_JITEMNO                                 As String

    Dim rsJournal_HDDup                           As ADODB.Recordset
    Dim PMIOS_RRNO                                As String
    Dim PMIOS_ISTATUS                             As String
    Dim PMIOS_RRNO1                               As String
    Dim PMIOS_Notes                               As String
    Dim PMIOS_RRDATE                              As String
    Dim PMIOS_RRDATE1                             As String
    Dim PMIOS_PONO                                As String
    Dim PMIOS_PODATE                              As String
    Dim PMIOS_RECVD_CODE                          As String
    Dim PMIOS_RECVD_FROM                          As String
    Dim PMIOS_DRNO                                As String
    Dim PMIOS_INVNO                               As String
    Dim PMIOS_CLASSCODE                           As String
    Dim PMIOS_CLASSCODE1                          As String
    Dim PMIOS_TERMS                               As String
    Dim PMIOS_TOTALQTY                            As Double
    Dim PMIOS_TTLRRAMT                            As Double
    Dim PMIOS_DS1                                 As Double
    Dim PMIOS_DS_AMT1                             As Double
    Dim PMIOS_NETRRAMT                            As Double
    Dim PMIOS_STATUS                              As String
    Dim PMIOS_TYPE                                As String
    Dim CONDUCTION                                As String
    Dim AMIS_JTYPE                                As String
    Dim J_ITEMCOUNT                               As Integer
    Dim ATCRate                                   As Double
    Dim TOTAL_DEBIT, TOTAL_CREDIT                 As Double
    Dim J_ATC                                     As String
    Dim J_RATE                                    As Double
    Dim J_TAXBASE                                 As Double

    Dim i                                         As Long


    Dim PV_PONO                                   As String
    Dim PV_MRRNO                                  As String
    Dim PV_INVNO                                  As String
    Dim PV_PRODNO                                 As String
    Dim J_JVOUCHERNO                              As String
    Dim PV_AMOUNT                                 As Double
    Dim PV_STATUS, PV_ITEMNO                      As String
    Dim j_transfer1                               As Boolean


    Dim rsrecords                                 As New ADODB.Recordset

    Dim rsRR_HD                                   As ADODB.Recordset
    Dim rsRR_HD1                                  As ADODB.Recordset
    Dim rsATC                                     As ADODB.Recordset

    Dim GridImport                                As Integer
    i = 0
    For GridImport = 1 To Grid1.Rows - 1
        If N2Str2Zero(Grid1.Cell(GridImport, 1).Text) = 0 Then
            Set rsRR_HD = New ADODB.Recordset
            ' Update By BTT : 08132008
            If COMPANY_CODE = "HGC" Then
                Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where RRNO = '" & Grid1.Cell(GridImport, 3).Text & "' AND (CLASSCODE = 'PCG' or CLASSCODE = 'PCS' or CLASSCODE = 'IBT') AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
            Else
                Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where RRNO = '" & Grid1.Cell(GridImport, 3).Text & "' AND (CLASSCODE = 'PCG' or CLASSCODE = 'PCS' or CLASSCODE = 'IBT') AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
            End If
            If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                PMIOS_RRNO = Null2String(rsRR_HD!RRNO)
                PMIOS_RRDATE = Null2String(rsRR_HD!RRDATE)
                PMIOS_PONO = Null2String(rsRR_HD!PONO)
                PMIOS_PODATE = Null2String(rsRR_HD!PODATE)
                PMIOS_RECVD_CODE = Null2String(rsRR_HD!RECVD_CODE)
                PMIOS_RECVD_FROM = Null2String(rsRR_HD!RECVD_FROM)
                PMIOS_DRNO = Null2String(rsRR_HD!DRNO)
                PMIOS_INVNO = Null2String(rsRR_HD!INVNO)
                PMIOS_CLASSCODE = Null2String(rsRR_HD!CLASSCODE)
                'PMIOS_TERMS = Null2String(rsRR_HD!TERMS)
                PMIOS_TERMS = Return_Terms(Null2String(rsRR_HD!TERMS))
                PMIOS_TOTALQTY = Round(N2Str2Zero(rsRR_HD!TOTALQTY), 2)
                PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!TTLRRAMT), 2)
                PMIOS_DS1 = Round(N2Str2Zero(rsRR_HD!DS1), 2)
                PMIOS_DS_AMT1 = Round(N2Str2Zero(rsRR_HD!DS_AMT1), 2)
                PMIOS_NETRRAMT = Round(N2Str2Zero(rsRR_HD!NETRRAMT), 2)
                PMIOS_STATUS = Null2String(rsRR_HD!Status)
                PMIOS_TYPE = Null2String(rsRR_HD!Type)
                TOTAL_DEBIT = 0: TOTAL_CREDIT = 0



                'HEADER
                Set rsJournal_HDDup = New ADODB.Recordset
                Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
                If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                    J_JNO = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                Else
                    J_JNO = "'000001'"
                End If
                J_JDATE = N2Date2Null(PMIOS_RRDATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'APJ'"
                If PMIOS_TYPE = "P" Then
                    AMIS_JTYPE = "PI"
                    J_REMARKS = "'To record spareparts purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                If PMIOS_TYPE = "A" Then
                    AMIS_JTYPE = "AI"
                    J_REMARKS = "'To record accessories purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                If PMIOS_TYPE = "M" Then
                    AMIS_JTYPE = "MI"
                    J_REMARKS = "'To record materials purchases. with Ref# " & PMIOS_RRNO & "'"
                End If
                J_VENDORCODE = N2Str2Null(PMIOS_RECVD_CODE)
                J_CUSTOMERCODE = "'999999'"

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                J_AMOUNTTOPAY = Round(NumericVal(PMIOS_NETRRAMT), 2)
                J_INVOICEAMT = 0
                J_BALANCE = Round(NumericVal(PMIOS_NETRRAMT), 2)
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(PMIOS_RRDATE)
                J_INVOICENO = N2Str2Null(PMIOS_RRNO)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(Format(DateAdd("d", NumericVal(PMIOS_TERMS), Format(PMIOS_RRDATE, "DD-MMM-YY"))))
                If PMIOS_TERMS = "CSH" Then
                    J_PAYTYPE = N2Str2Null(PMIOS_TERMS)
                Else
                    J_PAYTYPE = N2Str2Null(PMIOS_TERMS)
                End If
                If PMIOS_TYPE = "P" Then
                    J_INVOICETYPE = "'PARTS'"
                ElseIf PMIOS_TYPE = "A" Then
                    J_INVOICETYPE = "'ACCESSORIES'"
                ElseIf PMIOS_TYPE = "M" Then
                    J_INVOICETYPE = "'MATERIALS'"
                Else
                    J_INVOICETYPE = "NULL"
                End If
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = "NULL"
                J_TERMS = N2Str2Null(PMIOS_TERMS)
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"



                'CASH ON HAND
                If PMIOS_NETRRAMT > 0 Then
                    'J_JITEMNO = "'0001'"
                    J_ITEMCOUNT = 0
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    If PMIOS_TYPE = "P" Then
                        'Update By BTT: 07042008 to separate the Orig to not Orig
                        If COMPANY_CODE = "HGC" Then
                            If PMIOS_CLASSCODE = "IBT" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            ElseIf CheckIfORIG(ReturnPartNo(PMIOS_RRNO)) = True Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = "'11-05001-00'"
                                J_ACCT_NAME = N2Str2Null(Setacctname("'11-05001-00'"))
                            End If
                        ElseIf COMPANY_CODE = "HOT" Then
                            If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            ElseIf PMIOS_RECVD_CODE = "H00001" Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVN")))
                            End If
                        ElseIf COMPANY_CODE = "HCI" Then
                            If CheckIfORIG(ReturnPartNo(PMIOS_RRNO)) = True Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVPN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVPN")))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                        End If
                    End If
                    If PMIOS_TYPE = "A" Then
                        If COMPANY_CODE = "HPI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                        ElseIf COMPANY_CODE = "HOT" Then
                            If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            ElseIf PMIOS_RECVD_CODE = "H00001" Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVN")))
                            End If
                        ElseIf COMPANY_CODE = "HGC" Then
                            If PMIOS_CLASSCODE = "IBT" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("ACCESSORIES", "INVA"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("ACCESSORIES", "INVA")))
                        End If
                    End If
                    If PMIOS_TYPE = "M" Then
                        If COMPANY_CODE = "HOT" Then
                            If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            ElseIf PMIOS_RECVD_CODE = "H00001" Then
                                'Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                'Non - Original Parts
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVN"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVN")))
                            End If
                        ElseIf COMPANY_CODE = "HSB" Then
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIAL", "INVA"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIAL", "INVA")))
                        ElseIf COMPANY_CODE = "HGC" Then
                            If PMIOS_CLASSCODE = "IBT" Then
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("PARTS", "INVP"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("PARTS", "INVP")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                            End If
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnInventory("MATERIALS", "INVM"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("MATERIALS", "INVM")))
                        End If

                    End If

                    'WITHOUT INPUT TAX
                    'J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                    'WITH INPUT TAX

                    J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    If COMPANY_CODE = "HOT" Or COMPANY_CODE = "HGC" Then
                        If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                            J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        End If
                    End If
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Zero(J_JNO)


                    'NO INPUT TAX - HGC
                    'J_JITEMNO = "'0002'"
                    If COMPANY_CODE = "HCI" Then
                        If ReturnNONVATVendor(J_VENDORCODE) = True Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = NumericVal(PMIOS_DS_AMT1)
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
                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                        End If
                    ElseIf COMPANY_CODE = "HOT" Then
                        If ReturnNONVATVendor(J_VENDORCODE) = True Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = NumericVal(PMIOS_DS_AMT1)
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
                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                        End If
                    ElseIf COMPANY_CODE = "HGC" Then
                        If ReturnNONVATVendor(J_VENDORCODE) = True Or PMIOS_CLASSCODE = "IBT" Then
                        Else
                            J_ITEMCOUNT = J_ITEMCOUNT + 1
                            J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                            J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                            J_DEBIT = NumericVal(PMIOS_DS_AMT1)
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
                            NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                        End If
                    Else
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                        J_DEBIT = NumericVal(PMIOS_DS_AMT1)
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
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    End If

                    '                    J_JITEMNO = "'0003'"
                    '                    J_ACCT_CODE = N2Str2Null(ReturnInComeTax("EXPANDED"))
                    '                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInComeTax("EXPANDED")))
                    '                    J_DEBIT = 0
                    '                    J_CREDIT = NumericVal(PMIOS_NETRRAMT) * 0.01
                    '                    J_TAX = 0
                    '                    J_GROSS = 0
                    '                    J_NET = 0
                    '                    J_STATUS = "'N'"
                    '                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                         '                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                         '                                   " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                         '                                     ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                         '                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    'J_JITEMNO = "'0003'"
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    'AP IS CLEARING ACCOUNT
                    If COMPANY_CODE = "HGC" Then
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        ElseIf PMIOS_CLASSCODE = "IBT" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("BRANCH"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("BRANCH")))
                        Else
                            'J_ACCT_CODE = "'21-01002-00'"
                            'J_ACCT_NAME = N2Str2Null(Setacctname("'21-01002-00'"))
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    ElseIf COMPANY_CODE = "HMH" Then
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    ElseIf COMPANY_CODE = "HAI" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("AP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("AP")))
                    ElseIf COMPANY_CODE = "HPI" Then
                        If PMIOS_RECVD_CODE = "H00001" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode("HYUNDAI")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    ElseIf COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("PARTS"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("PARTS")))
                    ElseIf COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("AP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("AP")))
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVP"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVP")))

                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                    End If
                End If

                If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HOT" Then
                    Set rsATC = New ADODB.Recordset
                    ATCRate = 0
                    rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                    If Not rsATC.EOF And Not rsATC.BOF Then
                        ATCRate = ReturnATCRate("WC 158")
                    End If
                    Set rsATC = Nothing
                    '''
                    If COMPANY_CODE = "HOT" Or COMPANY_CODE = "HGC" Then
                        ' J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2) - Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                        J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2) - Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1) * ATCRate, 2)
                    Else
                        J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2) - Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                    End If
                Else
                    J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                End If

                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HOT" Then
                    If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                        J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                    End If
                End If
                J_DEBIT = 0
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
                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)



                '0004
                'WITHHOLDING TAX PAYABLE - EXPANDED

                If COMPANY_CODE = "HCI" Then
                    Set rsATC = New ADODB.Recordset
                    ATCRate = 0
                    rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                    If Not rsATC.EOF And Not rsATC.BOF Then
                        ATCRate = ReturnATCRate("WC 158")
                    End If
                    Set rsATC = Nothing
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    J_ACCT_CODE = N2Str2Null(ReturnWithholdingTax())
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnWithholdingTax()))
                    J_DEBIT = 0

                    'J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                    J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1) * ATCRate, 2)

                    J_AMOUNTTOPAY = J_AMOUNTTOPAY - J_CREDIT
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    J_ATC = "'WC 158'"
                    J_RATE = ATCRate * 100
                    J_TAXBASE = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2)
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                ElseIf COMPANY_CODE = "HOT" Then
                    If PMIOS_CLASSCODE = "IBT" Or PMIOS_RECVD_CODE = "M00002" Or PMIOS_RECVD_CODE = "M00003" Or PMIOS_RECVD_CODE = "S00002" Then
                    Else
                        Set rsATC = New ADODB.Recordset
                        ATCRate = 0
                        rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                        If Not rsATC.EOF And Not rsATC.BOF Then
                            ATCRate = ReturnATCRate("WC 158")
                        End If
                        Set rsATC = Nothing
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnWithholdingTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnWithholdingTax()))
                        J_DEBIT = 0

                        'J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                        J_CREDIT = Round(NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1) * ATCRate, 2)

                        J_AMOUNTTOPAY = J_AMOUNTTOPAY - J_CREDIT
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        J_ATC = "'WC 158'"
                        J_RATE = ATCRate * 100
                        J_TAXBASE = NumericVal(PMIOS_NETRRAMT - PMIOS_DS_AMT1)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    End If

                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ATC,RATE,TAXBASE)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                End If

                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                PV_PONO = N2Str2Null(PMIOS_PONO)
                PV_INVNO = N2Str2Null(PMIOS_INVNO)
                PV_PRODNO = "NULL"
                PV_AMOUNT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                "(VoucherNo,JTYPE,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                                " values (" & J_JVOUCHERNO & "," & J_JTYPE & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                ", " & PV_STATUS & ")"

                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_PV_Detail", "BERNARD", N2Str2Null(AMIS_JTYPE), "Jtype"))
                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), AMIS_JTYPE, N2Str2Null(PV_MRRNO)
                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                'SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                 '               " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                 '               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & ",'" & AMIS_JTYPE & "'," & PV_INVNO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & PV_MRRNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                 '                 ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT



                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                Grid1.Cell(GridImport, 1).Text = 1

            End If
        End If

        i = i + 1
        progCPB.Value = (i / (Grid1.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
        DoEvents
    Next

    'VEHICLES IMPORTING==================================================================================================================================================================================================================================================================================================================================
    i = 0

    For GridImport = 1 To Grid2.Rows - 1
        If N2Str2Zero(Grid2.Cell(GridImport, 1).Text) = 0 Then

            Set rsRR_HD = New ADODB.Recordset
            Set rsRR_HD = gconDMIS.Execute("Select * from SMIS_MRRINV Where CODE = '" & Grid2.Cell(GridImport, 3).Text & "' AND DateReceived = '" & CDate(dtpTranDate) & "' Order by DateReceived ASC")

            If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
                J_ITEMCOUNT = 0
                PMIOS_RRNO = Null2String(rsRR_HD!Code)
                PMIOS_RRDATE = Null2String(rsRR_HD!DateReceived)
                PMIOS_PONO = Null2String(rsRR_HD!PONO)
                PMIOS_PODATE = Null2String(rsRR_HD!PullOutDate)
                PMIOS_RECVD_CODE = Null2String("H00001")
                PMIOS_RECVD_FROM = Null2String("HYUNDAI ASIA RESOURCES INC.")
                'PMIOS_RECVD_CODE = Null2String(rsRR_HD!Source)
                'PMIOS_RECVD_FROM = SetSellingDealerName(Null2String(rsRR_HD!Source))
                PMIOS_DRNO = Null2String(rsRR_HD!DRNO)
                'PMIOS_INVNO = Null2String(rsRR_HD!VI_NO)
                PMIOS_INVNO = Null2String(rsRR_HD!refpono)
                PMIOS_CLASSCODE = Null2String(rsRR_HD!Model)
                CONDUCTION = Null2String(rsRR_HD!ignkey)
                PMIOS_TERMS = "CSH"
                PMIOS_TOTALQTY = 1
                PMIOS_ISTATUS = Null2String(rsRR_HD!IStatus)


                If COMPANY_CODE = "HBK" Or COMPANY_CODE = "HGC" Or COMPANY_CODE = "HAS" Then
                    'Update By BTT - 06272008 : Net of subsidy
                    PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!PURCHPRICE), 2) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)
                    PMIOS_DS_AMT1 = Round(((N2Str2Zero(rsRR_HD!PURCHPRICE) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)) / 1.12) * 0.12, 2)
                    PMIOS_NETRRAMT = Round((N2Str2Zero(rsRR_HD!PURCHPRICE) - Round(N2Str2Zero(rsRR_HD!mmpcsubs), 2)) - PMIOS_DS_AMT1, 2)
                    PMIOS_DS1 = 12
                Else
                    PMIOS_DS1 = 12
                    PMIOS_TTLRRAMT = Round(N2Str2Zero(rsRR_HD!PURCHPRICE), 2)
                    PMIOS_DS_AMT1 = Round((N2Str2Zero(rsRR_HD!PURCHPRICE) / 1.12) * 0.12, 2)
                    PMIOS_NETRRAMT = Round(N2Str2Zero(rsRR_HD!PURCHPRICE) - PMIOS_DS_AMT1, 2)
                    PMIOS_STATUS = Null2String(rsRR_HD!IStatus)
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

                J_JDATE = N2Date2Null(PMIOS_RRDATE)
                J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                J_JTYPE = "'APJ'"
                If COMPANY_CODE = "HAS" Or COMPANY_CODE = "HPI" Then
                    J_REMARKS = "'To record New Vehicle Purchases. with Ref# " & PMIOS_RRNO + ":" + CONDUCTION & "'"
                ElseIf COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
                    J_REMARKS = "'To record New Vehicle Purchases. with Ref# " & PMIOS_RRNO & ": INV.#" & PMIOS_INVNO & ": P.O.#" & PMIOS_PONO & ": C.S.#" & CONDUCTION & ": " & PMIOS_CLASSCODE & "'"
                Else
                    J_REMARKS = "'To record New Vehicle Purchases. with Ref# " & PMIOS_RRNO + "'"
                End If
                J_VENDORCODE = N2Str2Null(PMIOS_RECVD_CODE)
                J_CUSTOMERCODE = "'999999'"

                J_DEBIT = 0
                J_CREDIT = 0
                J_TAX = 0
                J_OUTBALANCE = 0

                If COMPANY_CODE = "HGC" And PMIOS_ISTATUS = "T" Then
                    J_AMOUNTTOPAY = Round(NumericVal(PMIOS_NETRRAMT), 2)
                Else
                    J_AMOUNTTOPAY = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                End If
                J_INVOICEAMT = 0
                J_BALANCE = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                J_AMOUNTPAID = 0

                J_STATUS = "'N'"

                J_INVOICEDATE = N2Date2Null(PMIOS_RRDATE)
                J_INVOICENO = N2Str2Null(rsRR_HD!Code)
                J_CHECKNO = "NULL"
                J_DUEDATE = N2Date2Null(Format(DateAdd("d", NumericVal(PMIOS_TERMS), Format(PMIOS_RRDATE, "DD-MMM-YY"))))
                J_PAYTYPE = "'" & PMIOS_TERMS & "'"
                J_INVOICETYPE = "'VEHICLES'"
                J_CHECKDATE = "NULL"
                J_BANKCODE = "NULL"
                J_REFNO = "NULL"
                J_REFDATE = "NULL"
                J_TERMS = "'" & PMIOS_TERMS & "'"
                J_DEALER = "NULL"
                J_PAIDSTATUS = "'N'"
                J_RECEIVESTATUS = "'N'"

                'J_JITEMNO = "'0001'"
                If PMIOS_NETRRAMT > 0 Then
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    'Update By : BTT - 06252008
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVENTORY"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVENTORY")))
                        J_DEBIT = 0
                        J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    ElseIf COMPANY_CODE = "HGC" And PMIOS_ISTATUS = "T" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "VEHICLES")))
                        J_CREDIT = 0
                        J_TAX = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) / 1.12 * 0.12)
                        J_DEBIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) - J_TAX)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    ElseIf COMPANY_CODE = "HGC" And PMIOS_RECVD_CODE = "HGC" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "VEHICLES")))
                        J_CREDIT = 0
                        J_DEBIT = NumericVal(PMIOS_TTLRRAMT)
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    ElseIf COMPANY_CODE = "HAI" Or COMPANY_CODE = "HOT" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", PMIOS_CLASSCODE))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", PMIOS_CLASSCODE)))
                        J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    ElseIf COMPANY_CODE = "HSB" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "INVA"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "INVA")))
                        J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    Else

                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "VEHICLES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "VEHICLES")))
                        J_DEBIT = Round(NumericVal(PMIOS_NETRRAMT), 2)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    End If
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                    'J_JITEMNO = "'0002'"
                    If COMPANY_CODE = "HGC" And PMIOS_ISTATUS = "T" Or COMPANY_CODE = "HGC" And PMIOS_RECVD_CODE = "HGC" Then
                    Else
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnInPutTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInPutTax()))
                        J_DEBIT = Round(NumericVal(PMIOS_DS_AMT1), 2)
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

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    End If

                    '                J_JITEMNO = "'0003'"
                    '                J_ACCT_CODE = N2Str2Null(ReturnInComeTax("EXPANDED"))
                    '                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInComeTax("EXPANDED")))
                    '                J_DEBIT = 0
                    '                J_CREDIT = NumericVal(PMIOS_NETRRAMT) * 0.01
                    '                J_TAX = 0
                    '                J_GROSS = 0
                    '                J_NET = 0
                    '                J_STATUS = "'N'"
                    '                TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    '
                    '                gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     '                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     '                                 " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     '                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     '                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & "," & J_STATUS & ")"
                    '                J_JITEMNO = "'0004'"


                    'CASH ON HAND
                    'J_JITEMNO = "'0003'"
                    'UPDATED BY: ACL - 06062010
                    J_ITEMCOUNT = J_ITEMCOUNT + 1
                    J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                    ' Update By : BTT - 06252008
                    If COMPANY_CODE = "HBK" Then
                        J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "SALES"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "SALES")))
                        J_DEBIT = NumericVal(PMIOS_NETRRAMT)
                        J_CREDIT = 0
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    Else

                        If COMPANY_CODE = "HGC" Then
                            If PMIOS_ISTATUS = "T" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("BRANCH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("BRANCH")))
                            ElseIf PMIOS_RECVD_CODE = "HGC" Then
                                J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("BRANCH"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("BRANCH")))
                            Else
                                J_ACCT_CODE = N2Str2Null(ReturnInventory("SALES", "IN TRANSIT"))
                                J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("SALES", "IN TRANSIT")))
                            End If
                        ElseIf COMPANY_CODE = "HPI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAR_AccountCode2("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAR_AccountCode2("HYUNDAI")))
                        ElseIf COMPANY_CODE = "HAI" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("AP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("AP")))
                        ElseIf COMPANY_CODE = "HOT" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("VEHICLE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("VEHICLE")))
                        ElseIf COMPANY_CODE = "HMH" Then
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("FINANCE"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("FINANCE")))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("HYUNDAI"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("HYUNDAI")))
                        End If
                        J_DEBIT = 0
                        
                        'HCI: WITHHOLDING TAX
                        If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HOT" Then
                            Set rsATC = New ADODB.Recordset
                            ATCRate = 0
                            rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                            If Not rsATC.EOF And Not rsATC.BOF Then
                                ATCRate = ReturnATCRate("WC 158")
                            End If
                            Set rsATC = Nothing
                            J_CREDIT = Round(NumericVal(PMIOS_TTLRRAMT), 2) - Round(NumericVal(PMIOS_TTLRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                        Else
                            If COMPANY_CODE = "HGC" And PMIOS_ISTATUS = "T" Then
                                J_TAX = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) / 1.12 * 0.12)
                                J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) - J_TAX)
                            ElseIf COMPANY_CODE = "HGC" And PMIOS_RECVD_CODE = "HGC" Then
                                J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                            Else
                                J_CREDIT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                            End If
                        End If

                        'J_CREDIT = Round(NumericVal(PMIOS_TTLRRAMT), 2)

                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    End If
                    SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                    "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                    " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                    ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                    ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    gconDMIS.Execute SQL_STATEMENT

                    TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "BERNARD", J_JTYPE, "Jtype"))
                    NEW_LogAudit "MM", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)


                    'WITHHOLDING TAX PAYABLE - EXPANDED
                    If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HOT" Then
                        Set rsATC = New ADODB.Recordset
                        ATCRate = 0
                        rsATC.Open "SELECT ATC FROM AMIS_ATC", gconDMIS, adOpenKeyset
                        If Not rsATC.EOF And Not rsATC.BOF Then
                            ATCRate = ReturnATCRate("WC 158")
                        End If
                        Set rsATC = Nothing
                        J_ITEMCOUNT = J_ITEMCOUNT + 1
                        J_JITEMNO = "'" & Format(J_ITEMCOUNT, "0000") & "'"
                        J_ACCT_CODE = N2Str2Null(ReturnWithholdingTax())
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnWithholdingTax()))
                        J_DEBIT = 0
                        J_CREDIT = Round(NumericVal(PMIOS_TTLRRAMT - PMIOS_DS_AMT1), 2) * ATCRate
                        J_AMOUNTTOPAY = J_AMOUNTTOPAY - J_CREDIT
                        J_TAX = 0
                        J_GROSS = 0
                        J_NET = 0
                        J_STATUS = "'N'"
                        J_ATC = "'WC 158'"
                        J_RATE = ATCRate * 100
                        J_TAXBASE = Round(NumericVal(PMIOS_TTLRRAMT - PMIOS_DS_AMT1), 2)
                        TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,ATC,RATE,TAXBASE)" & _
                                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

                        gconDMIS.Execute SQL_STATEMENT

                        TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_journal_det", "X", J_JTYPE, "Jtype"))
                        NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)
                    End If
                End If

                J_JVOUCHERNO = J_VOUCHERNO
                PV_ITEMNO = N2Str2Null("0001")
                PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                PV_PONO = N2Str2Null(PMIOS_PONO)
                PV_INVNO = N2Str2Null(PMIOS_INVNO)
                PV_PRODNO = "NULL"
                If COMPANY_CODE = "HGC" And PMIOS_ISTATUS = "T" Then
                    J_TAX = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) / 1.12 * 0.12)
                    PV_AMOUNT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT) - J_TAX)
                Else
                    PV_AMOUNT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                End If
                PV_STATUS = "'N'"

                SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                                "(VoucherNo,JTYPE,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                                " values (" & J_JVOUCHERNO & "," & J_JTYPE & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                ", " & PV_STATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                NEW_LogAudit "MM", "JOURNAL ENTRY DETAIL ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(PV_MRRNO)

                SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                                " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                ", " & J_JNO & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"
                gconDMIS.Execute SQL_STATEMENT

                TransactionID = (FindTransactionID(N2Str2Null(J_VOUCHERNO), "voucherno", "AMIS_Journal_HD", "BERNARD", J_JTYPE, "Jtype"))
                NEW_LogAudit "M", "JOURNAL ENTRY ", SQL_STATEMENT, N2Str2Zero(TransactionID), "", N2Str2Null(J_VOUCHERNO), J_JTYPE, N2Str2Null(J_JNO)

                '=========================================================================================================================
                'Entry For Clering Accoung PSB : Update By BTT - 06232008
                If COMPANY_CODE = "HBK" Then

                    Set rsJournal_HDDup = New ADODB.Recordset
                    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")

                    Dim J_ClearingNo              As String
                    Dim VendorCode                As String
                    Dim VendorName                As String

                    VendorCode = "P00007"
                    VendorName = "PS Bank"
                    J_AMOUNTTOPAY = NumericVal(0)
                    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then
                        J_ClearingNo = "'" & Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") & "'"
                    Else
                        J_ClearingNo = "'000001'"
                    End If
                    J_VOUCHERNO = N2Str2Null(GetVoucherNo())
                    'Detail
                    J_JITEMNO = "'0001'"
                    J_ACCT_CODE = N2Str2Null(ReturnClearing_AccountCode("INVENTORY"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnClearing_AccountCode("INVENTORY")))
                    J_DEBIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_ClearingNo & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = N2Str2Null(ReturnInventory("AP", "FLOORSTOCK"))
                    J_ACCT_NAME = N2Str2Null(Setacctname(ReturnInventory("AP", "FLOORSTOCK")))
                    J_DEBIT = 0
                    J_CREDIT = NumericVal(N2Str2Zero(PMIOS_TTLRRAMT))
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT

                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                     "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                     ", " & J_ClearingNo & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                     ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"

                    J_JVOUCHERNO = J_VOUCHERNO
                    PV_ITEMNO = N2Str2Null("0001")
                    PV_MRRNO = N2Str2Null(PMIOS_RRNO)
                    PV_PONO = N2Str2Null(PMIOS_PONO)
                    PV_INVNO = N2Str2Null(PMIOS_INVNO)
                    PV_PRODNO = "NULL"
                    PV_AMOUNT = Round(NumericVal(PMIOS_TTLRRAMT), 2)
                    PV_STATUS = "'N'"

                    gconDMIS.Execute "insert into AMIS_PV_Detail " & _
                                     "(VoucherNo,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                                     " values (" & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                                     ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                                     ", " & PV_STATUS & ")"
                    'Header
                    gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                                     " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus)" & _
                                     " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", '" & VendorCode & "','" & VendorName & "', " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                                     ", " & J_ClearingNo & ", " & TOTAL_DEBIT & ", " & TOTAL_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ")"

                End If
                'End of Clearing Acoount
                '*****************************************************************************************************

                Grid2.Cell(GridImport, 1).Text = 1
            End If
        End If
        i = i + 1
        progCPB.Value = (i / (Grid2.Rows - 1)) * 100
        labCPB.Caption = Int(progCPB.Value) & "% Completed"
    Next

    ImportPurchases = True
    Exit Function
ErrorCode:
    ImportPurchases = False
End Function
Private Sub cmdClearJournals_Click()
    Dim rsCHATCheckControlIfExistRecordInJournalHD As ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = New ADODB.Recordset
    Set rsCHATCheckControlIfExistRecordInJournalHD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'APJ' and Jdate = '" & CDate(dtpTranDate) & "' AND STATUS <> 'C'")
    If Not rsCHATCheckControlIfExistRecordInJournalHD.EOF And Not rsCHATCheckControlIfExistRecordInJournalHD.BOF Then
        Screen.MousePointer = 0
        If MsgBox("Clear Unposted Data for this Particular Date?", vbQuestion + vbYesNo, "Confirm...") = vbNo Then
            Exit Sub
        Else
            gconDMIS.Execute ("Delete from AMIS_PV_Detail Where (STATUS <> 'P' AND STATUS <> 'C') AND JTYPE ='APJ' AND JDate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("Delete from AMIS_Journal_Det Where (STATUS <> 'P' AND STATUS <> 'C') AND Jtype = 'APJ' and JDate = '" & CDate(dtpTranDate) & "'")
            gconDMIS.Execute ("Delete from AMIS_Journal_HD Where (STATUS <> 'P' AND STATUS <> 'C') AND Jtype = 'APJ' and JDate = '" & CDate(dtpTranDate) & "'")
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
        Grid1.Rows = 1
        Grid2.Rows = 1
        Grid3.Rows = 1
    End If
    Dim xRRDate                                   As String
    Dim tmpDate                                   As String
    Dim iCount                                    As Integer
    MonthIndex
    Dim rsCheckPMISData                           As ADODB.Recordset
    Dim rsCheckPMISData2                          As ADODB.Recordset
    '---------------------------
    '    Set rsCheckPMISData = New ADODB.Recordset
    '    rsCheckPMISData.Open "SELECT * FROM (SELECT DISTINCT RRDATE AS TRANDATE FROM PMIS_VW_RR_TRANS WHERE STATUS = 'P' AND (CLASSCODE = 'PCG' OR CLASSCODE = 'PCS' ) AND MONTH(RRDATE) = '" & Indx & "'  AND YEAR(RRDATE) = '" & cboYear.Text & "' " & _
         '                         "UNION SELECT DISTINCT DATERECEIVED AS TRANDATE FROM SMIS_MRRINV LEFT OUTER JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE WHERE STATUS = 'P' AND MONTH(DATERECEIVED) = '" & Indx & "' AND YEAR(DATERECEIVED) = '" & cboYear.Text & "' " & _
         '                         "UNION SELECT DISTINCT RC_DATE AS TRANDATE FROM CSMS_PO_RC_HD WHERE STATUS='P' AND MONTH(RC_DATE)='" & Indx & "' AND YEAR(RC_DATE)='" & cboYear.Text & "') X ORDER BY 1 ASC", gconDMIS, adOpenKeyset
    '    Do While Not rsCheckPMISData.EOF
    '        iCount = 0
    '        Dim rsCheckJournalHD                      As ADODB.Recordset
    '        Set rsCheckJournalHD = New ADODB.Recordset
    '        rsCheckJournalHD.Open "SELECT DISTINCT JDATE,JTYPE FROM AMIS_JOURNAL_HD WHERE JTYPE = 'APJ' AND JDATE IN (SELECT * FROM (SELECT RRDATE AS TRANDATE FROM PMIS_VW_RR_TRANS WHERE STATUS = 'P' AND (CLASSCODE = 'PCG' OR CLASSCODE = 'PCS' ) " & _
             '                              "UNION SELECT DISTINCT DATERECEIVED AS TRANDATE FROM SMIS_MRRINV LEFT OUTER JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE WHERE STATUS = 'P' " & _
             '                              "UNION SELECT DISTINCT RC_DATE AS TRANDATE FROM CSMS_PO_RC_HD WHERE STATUS='P' ) X WHERE TRANDATE = '" & Null2String(rsCheckPMISData!trandate) & "')", gconDMIS, adOpenKeyset
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
    '                If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckPMISData!trandate)), "d") Then
    '                    lab.Item(iCount).BackColor = &HFFFF&
    '                End If
    '                iCount = iCount + 1
    '            Loop
    '        End If
    '        rsCheckPMISData.MoveNext
    '    Loop
    '----------------------------------
    Screen.MousePointer = 11
    Set rsCheckPMISData = New ADODB.Recordset
    rsCheckPMISData.Open "SELECT DISTINCT CONVERT(VARCHAR,DATE,101) AS DATE FROM AMIS_VW_IMPORTED_PURCHASE WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckPMISData.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckPMISData!Date)), "d") Then
                lab.Item(iCount).BackColor = &HC0FFC0      'GREEN
            End If
            iCount = iCount + 1
            DoEvents
        Loop
        rsCheckPMISData.MoveNext
        DoEvents
    Loop

    Set rsCheckPMISData2 = New ADODB.Recordset
    rsCheckPMISData2.Open "SELECT DISTINCT CONVERT(VARCHAR,DATE,101) AS DATE FROM AMIS_VW_UNIMPORTED_PURCHASE WHERE MONTH([DATE]) = '" & Indx & "' AND YEAR([DATE])='" & cboYear.Text & "'", gconDMIS, adOpenKeyset
    Do While Not rsCheckPMISData2.EOF
        iCount = 0
        Do While iCount <= lab.Count - 1
            If lab.Item(iCount).Caption = Format(Null2String(Null2String(rsCheckPMISData2!Date)), "d") Then
                lab.Item(iCount).BackColor = &HFFFF&       'YELLOW
            End If
            iCount = iCount + 1
            DoEvents
        Loop
        rsCheckPMISData2.MoveNext
        DoEvents
    Loop
    Screen.MousePointer = 0
End Sub


Private Sub cmdShowTrans_Click()
    Dim KIM                                       As Integer
    Screen.MousePointer = 11
InitGrids:     DoEvents: cmdCheck.Enabled = False: cmdClearJournals.Enabled = False
    Grid1.Rows = 1: Grid2.Rows = 1: KIM = 0
    Dim RRTYPE                                    As String
    Dim IS_Exist                                  As Byte
    Dim rsRR_HD                                   As ADODB.Recordset
    Dim rsPURCH_AGREE                             As ADODB.Recordset
    Dim rsTransferIn                              As ADODB.Recordset
    Dim CostPrice                                 As String
    Dim Code                                      As String

    If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HPI" Then
        GoTo ShowTransactions
    Else
        If CheckImportedAP(dtpTranDate) = True Then
            MsgBox "Previous transaction(s) dated " & TRANSACTIONDATE & " are not yet imported.", vbExclamation, "Message"
            Screen.MousePointer = 0
            Exit Sub
        Else
            GoTo ShowTransactions
        End If
    End If

ShowTransactions:

    Set rsRR_HD = New ADODB.Recordset
    Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_vw_RR_TRANS Where STATUS = 'P' AND (CLASSCODE = 'PCG' or CLASSCODE = 'PCS' or CLASSCODE = 'IBT' ) AND RRDATE = '" & CDate(dtpTranDate) & "' Order by RRNO ASC")
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        rsRR_HD.MoveFirst: KIM = 0
        Grid1.AutoRedraw = False
        Do While Not rsRR_HD.EOF
            KIM = KIM + 1
            If Null2String(rsRR_HD!Type) = "P" Then
                RRTYPE = "PARTS"
            ElseIf Null2String(rsRR_HD!Type) = "A" Then
                RRTYPE = "ACCESSORIES"
            ElseIf Null2String(rsRR_HD!Type) = "M" Then
                RRTYPE = "MATERIALS"
            Else
                RRTYPE = ""
            End If
            If CheckAPJExisting(Null2String(rsRR_HD!RRNO), RRTYPE) = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid1.AddItem IS_Exist & Chr(9) & RRTYPE & Chr(9) & Null2String(rsRR_HD!RRNO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsRR_HD!TTLRRAMT)) & Chr(9) & Null2String(rsRR_HD!RECVD_FROM)
            rsRR_HD.MoveNext
        Loop
        If KIM = 0 Then Grid1.RemoveItem 1
        Grid1.AutoRedraw = True
        Grid1.Refresh
    End If
    Set rsPURCH_AGREE = New ADODB.Recordset
    Set rsPURCH_AGREE = gconDMIS.Execute("Select SMIS_MRRINV.*,CSMS_SELLINGDEALER.DEALERNAME from SMIS_MRRINV left outer JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE Where STATUS = 'P' AND DateReceived = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' Order by DateReceived ASC")
    If Not rsPURCH_AGREE.EOF And Not rsPURCH_AGREE.BOF Then
        rsPURCH_AGREE.MoveFirst: KIM = 0
        Grid2.AutoRedraw = False
        Do While Not rsPURCH_AGREE.EOF
            KIM = KIM + 1
            If CheckAPJExisting(Null2String(rsPURCH_AGREE!Code), "VEHICLES") = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid2.AddItem IS_Exist & Chr(9) & "VEHICLES" & Chr(9) & Null2String(rsPURCH_AGREE!Code) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPURCH_AGREE!PURCHPRICE)) & Chr(9) & Null2String(rsPURCH_AGREE!DealerName)
            rsPURCH_AGREE.MoveNext
        Loop
        If KIM = 0 Then Grid2.RemoveItem 1
        Grid2.AutoRedraw = True
        Grid2.Refresh
    End If
    If KIM > 0 Then
        cmdCheck.Enabled = True
        cmdClearJournals.Enabled = True
    End If

    Screen.MousePointer = 0
    'Update By : BTT 07142008 : to Process the Sublet in CSMS
    Dim rsSublet                                  As ADODB.Recordset
    Set rsSublet = New ADODB.Recordset
    Set rsSublet = gconDMIS.Execute("Select * from CSMS_PO_RC_HD Where STATUS = 'P' AND Rc_DATE = '" & Format(CDate(dtpTranDate), "MM/DD/YYYY") & "' Order by RC_date ASC")
    Grid3.Rows = 1
    If Not rsSublet.EOF And Not rsSublet.BOF Then
        rsSublet.MoveFirst: KIM = 0
        Grid3.AutoRedraw = False
        Do While Not rsSublet.EOF
            KIM = KIM + 1
            If CheckSubletifExist(Null2String(rsSublet!RC_NO), Null2String(rsSublet!po_no)) = True Then
'            If CheckAPJExisting(Null2String(rsSublet!RC_NO), "SUBLET") = True Then
                IS_Exist = 1
            Else
                IS_Exist = 0
            End If
            Grid3.AddItem IS_Exist & Chr(9) & "SUBLET" & Chr(9) & Null2String(rsSublet!RC_NO) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsSublet!sublet_Total_net_AMT)) & Chr(9) & Null2String(rsSublet!contractor_name)
            rsSublet.MoveNext
        Loop
    End If
    If KIM = 0 Then Grid3.RemoveItem 1
    Grid3.AutoRedraw = True
    Grid3.Refresh
    cmdCheck.Enabled = True
    cmdClearJournals.Enabled = True
    'End of Update
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
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error!"
    Unload frmSplash
    cmdCheck.Enabled = False
End Sub

Function CheckImportedAP(xRRDate As String) As Boolean
'Update: ACL 11162009
    Dim xNoDays                                   As Integer
    Dim rsCheckAP                                 As ADODB.Recordset
    Dim rsCheckTrans                              As ADODB.Recordset
    Set rsCheckAP = New ADODB.Recordset
    rsCheckAP.Open "Select TOP 1 * from AMIS_Journal_HD where JTYPE='APJ' AND STATUS <> 'C' AND JDATE < '" & CDate(xRRDate) & "' ORDER BY JDATE DESC", gconDMIS, adOpenKeyset
    If Not rsCheckAP.EOF And Not rsCheckAP.BOF Then
        Set rsCheckTrans = New ADODB.Recordset
        rsCheckTrans.Open "Select * from (SELECT TOP 1 RRDATE AS TRANDATE FROM PMIS_VW_RR_TRANS WHERE STATUS = 'P' AND (CLASSCODE = 'PCG' OR CLASSCODE = 'PCS' ) AND RRDATE > '" & Null2Date(rsCheckAP!JDate) & "' ORDER BY RRDATE ASC " & _
                          "Union SELECT TOP 1 DATERECEIVED AS TRANDATE FROM SMIS_MRRINV LEFT OUTER JOIN CSMS_SELLINGDEALER ON SMIS_MRRINV.SOURCE = CSMS_SELLINGDEALER.DEALERCODE WHERE STATUS = 'P' AND DATERECEIVED > '" & Null2Date(rsCheckAP!JDate) & "' ORDER BY DATERECEIVED ASC " & _
                          "UNION SELECT DISTINCT RC_DATE AS TRANDATE FROM CSMS_PO_RC_HD WHERE STATUS='P' AND RC_DATE > '" & Null2Date(rsCheckAP!JDate) & "') X Order By TranDate Asc", gconDMIS, adOpenKeyset
        If Not rsCheckTrans.EOF And Not rsCheckTrans.BOF Then
            TRANSACTIONDATE = Null2Date(rsCheckTrans!trandate)
            'If Format(dtpTranDate, "mm/dd/yyyy") > Format(TRANSACTIONDATE, "mm/dd/yyyy") Then
            xNoDays = DateDiff("d", TRANSACTIONDATE, dtpTranDate)
            If xNoDays > 0 Then
                CheckImportedAP = True
            End If
        End If
    End If
    Set rsCheckAP = Nothing
End Function

Function CheckToClearAP(xRRDate As String) As Boolean
    Dim rsCheckAP                                 As ADODB.Recordset
    Set rsCheckAP = New ADODB.Recordset
    rsCheckAP.Open "Select TOP 1 * from AMIS_Journal_HD Where STATUS <> 'P' AND Jtype = 'APJ' and Jdate > '" & CDate(xRRDate) & "' Order by JDate Asc", gconDMIS, adOpenKeyset
    If Not rsCheckAP.EOF And Not rsCheckAP.BOF Then
        CheckToClearAP = True
    End If
End Function

Function Return_Terms(xTerm As String) As String
    Dim XXX                                       As String
    If xTerm = "" Then
        xTerm = "CSH"
    ElseIf NumericVal(xTerm) = 0 Then
        xTerm = "CSH"
    Else
        xTerm = xTerm & "D"
    End If
    Dim rsTerm                                    As ADODB.Recordset
    Set rsTerm = New ADODB.Recordset
    rsTerm.Open "Select PAY_CODE from ALL_PAYTERM where PAY_CODE ='" & xTerm & "'", gconDMIS, adOpenKeyset
    If Not rsTerm.EOF And Not rsTerm.BOF Then
        Return_Terms = rsTerm!pay_Code
    End If
End Function

Private Sub lab_Click(Index As Integer)
InitGrids:     DoEvents:
    Grid1.Rows = 1
    Grid2.Rows = 1
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

Function ReturnATCRate(XXX As String) As Double
    Dim rsRerturnATCRate                          As ADODB.Recordset
    Set rsRerturnATCRate = New ADODB.Recordset
    rsRerturnATCRate.Open "SELECT * FROM AMIS_ATC WHERE ATC='" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsRerturnATCRate.EOF And Not rsRerturnATCRate.BOF Then
        ReturnATCRate = Round(ToDoubleNumber(rsRerturnATCRate!Rate)) / 100
    End If
    Set rsRerturnATCRate = Nothing
End Function

Function ReturnNONVATVendor(XXX As String) As Boolean
    Dim rsVENDOR                                  As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "SELECT NONVAT FROM ALL_VENDOR WHERE CODE = " & XXX & "", gconDMIS, adOpenForwardOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        If rsVENDOR!NONVAT = "Y" Then
            ReturnNONVATVendor = True
        Else
            ReturnNONVATVendor = False
        End If
    End If
    Set rsVENDOR = Nothing
End Function

Function ReturnAmount(XXX As String) As Double
    Dim rsReturnAmount                            As ADODB.Recordset
    Set rsReturnAmount = New ADODB.Recordset
    rsReturnAmount.Open "SELECT DESCRIPT,CostPrice FROM ALL_Model WHERE DESCRIPT = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsReturnAmount.EOF And Not rsReturnAmount.BOF Then
        ReturnAmount = rsReturnAmount!CostPrice
    End If
    Set rsReturnAmount = Nothing
End Function

Function ReturnCodeTR(XXX As Variant) As Double
    Dim rsReturnCode                              As ADODB.Recordset
    Set rsReturnCode = New ADODB.Recordset
    rsReturnCode.Open "SELECT CODE, DESCRIPT FROM ALL_Model WHERE DESCRIPT = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsReturnCode.EOF And Not rsReturnCode.BOF Then
        ReturnCodeTR = N2Str2IntZero(rsReturnCode!Code)
    End If
    Set rsReturnCode = Nothing
End Function




