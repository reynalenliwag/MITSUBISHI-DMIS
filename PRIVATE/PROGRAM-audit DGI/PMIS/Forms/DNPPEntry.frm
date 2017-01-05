VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMaster_DNPPEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DISTRIBUTOR Parts Maintenance"
   ClientHeight    =   8025
   ClientLeft      =   1800
   ClientTop       =   435
   ClientWidth     =   6225
   ForeColor       =   &H00DEDFDE&
   Icon            =   "DNPPEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdRegister_parts 
      Caption         =   "Register This To Parts Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   51
      Top             =   7710
      Width           =   2745
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register All MMPC To Parts Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   30
      TabIndex        =   50
      Top             =   7710
      Width           =   2745
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   390
      ScaleHeight     =   795
      ScaleWidth      =   5790
      TabIndex        =   25
      Top             =   6570
      Width           =   5790
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
         Left            =   5040
         MouseIcon       =   "DNPPEntry.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
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
         Left            =   4320
         MouseIcon       =   "DNPPEntry.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   3600
         MouseIcon       =   "DNPPEntry.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   2880
         MouseIcon       =   "DNPPEntry.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   2160
         MouseIcon       =   "DNPPEntry.frx":1B65
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":1CB7
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
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
         Left            =   1440
         MouseIcon       =   "DNPPEntry.frx":1FCA
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":211C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
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
         Left            =   720
         MouseIcon       =   "DNPPEntry.frx":2416
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":2568
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
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
         Left            =   0
         MouseIcon       =   "DNPPEntry.frx":28C0
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":2A12
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   30
      TabIndex        =   6
      Top             =   -30
      Width           =   6135
      Begin VB.CommandButton cmdUpdateDistributorParts 
         Caption         =   "Update Distributor Parts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3420
         MouseIcon       =   "DNPPEntry.frx":2D71
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Move to Previous Record"
         Top             =   150
         Width           =   2655
      End
      Begin VB.TextBox txtDNP3 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   2640
         Width           =   1275
      End
      Begin VB.TextBox txtDNP2 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   2220
         Width           =   1275
      End
      Begin VB.TextBox txtDNP 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1830
         Width           =   1275
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   5
         Text            =   "99,000.00"
         Top             =   2430
         Width           =   1695
      End
      Begin Crystal.CrystalReport rptPrintParts 
         Left            =   6150
         Top             =   2730
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "List of New Part Numbers"
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
      Begin VB.TextBox txtNewPartNo 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the parts new number."
         Top             =   1350
         Width           =   1635
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Type the part number (e.g. MR241052)"
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtModel 
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
         Height          =   735
         Left            =   3030
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type the part's model code (e.g. WRANGLER)"
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox txtICC 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Type the part's vehicle type."
         Top             =   960
         Width           =   315
      End
      Begin VB.TextBox txtPartDesc 
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
         Height          =   345
         Left            =   1320
         MaxLength       =   16
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type the part's description (e.g. BUMPER KIT,FR CO)"
         Top             =   570
         Width           =   4725
      End
      Begin VB.Label LABALLOWREPRINT 
         Caption         =   "Label14"
         Height          =   225
         Left            =   5340
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "20 % Markup"
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
         Height          =   255
         Left            =   2730
         TabIndex        =   41
         Top             =   2670
         Width           =   1545
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Parts DNP"
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
         Height          =   255
         Left            =   90
         TabIndex        =   40
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "32 % Markup"
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
         Height          =   255
         Left            =   2730
         TabIndex        =   39
         Top             =   2280
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Parts DNP"
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
         Height          =   255
         Left            =   90
         TabIndex        =   38
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "40 % Markup"
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
         Height          =   255
         Left            =   2730
         TabIndex        =   37
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Parts DNP"
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
         Height          =   255
         Left            =   90
         TabIndex        =   24
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   3120
         TabIndex        =   15
         Top             =   210
         Width           =   225
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "New No."
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
         Height          =   255
         Left            =   90
         TabIndex        =   14
         Top             =   1410
         Width           =   1245
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Suggested Retail Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   4200
         TabIndex        =   13
         Top             =   1890
         Width           =   1845
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
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
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Height          =   255
         Index           =   1
         Left            =   1830
         TabIndex        =   9
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ICC"
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
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1245
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3495
      Left            =   30
      TabIndex        =   16
      Top             =   3030
      Width           =   6135
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5580
         Top             =   150
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "D&escription [Ctrl + E]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2790
         TabIndex        =   20
         Top             =   150
         Width           =   2385
      End
      Begin VB.OptionButton optPartNo 
         Caption         =   "Pa&rt Number [Ctrl + R]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   19
         Top             =   180
         Value           =   -1  'True
         Width           =   2385
      End
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   540
         Width           =   5955
      End
      Begin MSComctlLib.ListView lstPartsEntry 
         Height          =   2445
         Left            =   90
         TabIndex        =   18
         Top             =   960
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   4313
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "DNPPEntry.frx":2EC3
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PART NO."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6262
         EndProperty
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   960
      ScaleHeight     =   1695
      ScaleWidth      =   4335
      TabIndex        =   43
      Top             =   2880
      Visible         =   0   'False
      Width           =   4365
      Begin wizProgBar.Prg Progressbar1 
         Height          =   555
         Left            =   120
         TabIndex        =   46
         ToolTipText     =   "Update progress"
         Top             =   420
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   979
         Picture         =   "DNPPEntry.frx":3025
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "DNPPEntry.frx":3041
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
      Begin VB.Label labstockno 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1350
         Width           =   4095
      End
      Begin VB.Label labProg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1035
         Width           =   4095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*** Updating Distributor Master File ***"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   44
         Top             =   30
         Width           =   4275
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   4725
      ScaleHeight     =   825
      ScaleWidth      =   1620
      TabIndex        =   34
      Top             =   6540
      Width           =   1620
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
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
         Left            =   720
         MouseIcon       =   "DNPPEntry.frx":305D
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":31AF
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   0
         MouseIcon       =   "DNPPEntry.frx":34ED
         MousePointer    =   99  'Custom
         Picture         =   "DNPPEntry.frx":363F
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   735
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   570
      TabIndex        =   10
      Top             =   210
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label LAB_PStatusDealer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parts Status:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   30
      TabIndex        =   49
      Top             =   7410
      Width           =   6135
   End
End
Attribute VB_Name = "frmPMISMaster_DNPPEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDNPP                                             As ADODB.Recordset
Dim AddorEdit                                          As String

Dim gconHARI_DNP                                       As ADODB.Connection
Dim FILNAME                                            As String

Sub updatednpp()
    Dim rsDNPP                                         As ADODB.Recordset
    Dim rsSTOCKS                                       As ADODB.Recordset

    Set rsDNPP = New ADODB.Recordset
    Set rsDNPP = gconDMIS.Execute("Select * from PMIS_DNPP order by partnumber asc")
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        rsDNPP.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsDNPP.EOF
            Set rsSTOCKS = New ADODB.Recordset
            Set rsSTOCKS = gconDMIS.Execute("Select * from pmis_stockmas where stockno = '" & rsDNPP!PARTNUMBER & "'")
            If Not rsSTOCKS.EOF And Not rsSTOCKS.BOF Then
                gconDMIS.Execute ("Update pmis_stockmas set srp = " & rsDNPP!SRP & " where stockno = '" & rsDNPP!PARTNUMBER & "'")
            End If
            Me.Caption = rsDNPP!PARTNUMBER
            DoEvents
            rsDNPP.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
End Sub

Sub DecideToUpload()
    Dim rsDNPP                                         As ADODB.Recordset
    Dim RSCURR_DNPP                                    As ADODB.Recordset
    Dim PARTNO                                         As String
    Dim PARTDESC                                       As String
    Dim DNP40                                          As Double
    Dim DNP32                                          As Double
    Dim DNP20                                          As Double
    Dim SRP                                            As Double
    Dim MODELAPP                                       As String
    Progressbar1.Value = 0
    Dim i                                              As Long
    Dim lng                                            As Long

    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select * from [ANNEX 1] Order by Field2 asc", gconHARI_DNP, adOpenForwardOnly
    lng = rsDNPP.RecordCount
    Progressbar1.Max = 100
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        rsDNPP.MoveFirst: i = 0: Screen.MousePointer = 11:
        Do While Not rsDNPP.EOF
            If N2Str2Zero(rsDNPP![FIELD7]) > 0 Then
                PARTNO = Repleys(Null2String(rsDNPP!FIELD2))
                PARTDESC = Repleys(Null2String(rsDNPP!FIELD3))
                DNP40 = NumericVal(rsDNPP!FIELD4)
                DNP32 = NumericVal(rsDNPP!FIELD5)
                DNP20 = NumericVal(rsDNPP!FIELD6)
                SRP = NumericVal(rsDNPP!FIELD7)
                MODELAPP = N2Str2Null(ReplaceQuote(Null2String(rsDNPP!FIELD8)))

                Set RSCURR_DNPP = New ADODB.Recordset
                Set RSCURR_DNPP = gconDMIS.Execute("Select DNPP from PMIS_DNPP where PartNumber = '" & PARTNO & "'")

                If Not RSCURR_DNPP.EOF And Not RSCURR_DNPP.BOF Then
                    gconDMIS.Execute ("Update PMIS_DNPP set " & _
                                      "DNPP = " & DNP40 & "," & _
                                      "DNPP2 = " & DNP32 & "," & _
                                      "DNPP3 = " & DNP20 & "," & _
                                      "SRP = " & SRP & "," & _
                                      "MODEL = " & MODELAPP & _
                                    " where PARTNUMBER = '" & PARTNO & "'")
                    labstockno = "UPDATING STOCK NO " & PARTNO
                Else
                    gconDMIS.Execute ("Insert into PMIS_DNPP (DNPP,DNPP2,DNPP3,SRP,MODEL,PARTNUMBER,DESCRIPTIO) values (" & _
                                      DNP40 & "," & DNP32 & "," & DNP20 & "," & SRP & "," & MODELAPP & ",'" & PARTNO & "','" & PARTDESC & "')")

                    labstockno = "REGISTERING NEW STOCK NO " & PARTNO
                End If

            End If
            DoEvents
            i = i + 1
            Progressbar1.Value = (i / lng) * 100
            labProg.Caption = Int(Progressbar1.Value) & "% Completed"
            DoEvents
            Screen.MousePointer = 0
            rsDNPP.MoveNext
        Loop
    End If


    Progressbar1.Value = 0
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select * from [ANNEX 2] Order by Field2 asc", gconHARI_DNP, adOpenForwardOnly
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        rsDNPP.MoveFirst: i = 0: Screen.MousePointer = 11
        Do While Not rsDNPP.EOF
            If N2Str2Zero(rsDNPP![FIELD7]) > 0 Then
                PARTNO = Repleys(Null2String(rsDNPP!FIELD2))
                PARTDESC = Repleys(Null2String(rsDNPP!FIELD3))
                DNP40 = NumericVal(rsDNPP!FIELD4)
                DNP32 = NumericVal(rsDNPP!FIELD5)
                DNP20 = NumericVal(rsDNPP!FIELD6)
                SRP = NumericVal(rsDNPP!FIELD7)
                MODELAPP = N2Str2Null(Repleys(Null2String(rsDNPP!FIELD8)))

                Set RSCURR_DNPP = New ADODB.Recordset
                Set RSCURR_DNPP = gconDMIS.Execute("Select * from PMIS_DNPP where PartNumber = '" & PARTNO & "'")
                If Not RSCURR_DNPP.EOF And Not RSCURR_DNPP.BOF Then
                    gconDMIS.Execute ("Update PMIS_DNPP set " & _
                                      "DNPP = " & DNP40 & "," & _
                                      "DNPP2 = " & DNP32 & "," & _
                                      "DNPP3 = " & DNP20 & "," & _
                                      "SRP = " & SRP & "," & _
                                      "MODEL = " & MODELAPP & _
                                    " where PARTNUMBER = '" & PARTNO & "'")
                Else
                    gconDMIS.Execute ("Insert into PMIS_DNPP (DNPP,DNPP2,DNPP3,SRP,MODEL,PARTNUMBER,DESCRIPTIO) values (" & _
                                      DNP40 & "," & DNP32 & "," & DNP20 & "," & SRP & "," & MODELAPP & ",'" & PARTNO & "','" & PARTDESC & "')")
                End If
            End If
            Screen.MousePointer = 0
            DoEvents
            i = i + 1
            Progressbar1.Value = (i / rsDNPP.RecordCount) * 100
            labProg.Caption = Int(Progressbar1.Value) & "% Completed"
            DoEvents

            rsDNPP.MoveNext
        Loop
    End If


    Progressbar1.Value = 0
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "Select * from [ANNEX 3] Order by Field2 asc", gconHARI_DNP, adOpenForwardOnly
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        rsDNPP.MoveFirst: i = 0: Screen.MousePointer = 11
        Do While Not rsDNPP.EOF
            If N2Str2Zero(rsDNPP![FIELD7]) > 0 Then
                PARTNO = Repleys(Null2String(rsDNPP!FIELD2))
                PARTDESC = Repleys(Null2String(rsDNPP!FIELD3))
                DNP40 = NumericVal(rsDNPP!FIELD4)
                DNP32 = NumericVal(rsDNPP!FIELD5)
                DNP20 = NumericVal(rsDNPP!FIELD6)
                SRP = NumericVal(rsDNPP!FIELD7)
                MODELAPP = N2Str2Null(Repleys(Null2String(rsDNPP!FIELD8)))

                Set RSCURR_DNPP = New ADODB.Recordset
                Set RSCURR_DNPP = gconDMIS.Execute("Select * from PMIS_DNPP where PartNumber = '" & PARTNO & "'")
                If Not RSCURR_DNPP.EOF And Not RSCURR_DNPP.BOF Then
                    gconDMIS.Execute ("Update PMIS_DNPP set " & _
                                      "DNPP = " & DNP40 & "," & _
                                      "DNPP2 = " & DNP32 & "," & _
                                      "DNPP3 = " & DNP20 & "," & _
                                      "SRP = " & SRP & "," & _
                                      "MODEL = " & MODELAPP & _
                                    " where PARTNUMBER = '" & PARTNO & "'")
                Else
                    gconDMIS.Execute ("Insert into PMIS_DNPP (DNPP,DNPP2,DNPP3,SRP,MODEL,PARTNUMBER,DESCRIPTIO) values (" & _
                                      DNP40 & "," & DNP32 & "," & DNP20 & "," & SRP & "," & MODELAPP & ",'" & PARTNO & "','" & PARTDESC & "')")
                End If
            End If
            Screen.MousePointer = 0
            DoEvents
            i = i + 1
            Progressbar1.Value = (i / rsDNPP.RecordCount) * 100
            labProg.Caption = Int(Progressbar1.Value) & "% Completed"
            DoEvents

            rsDNPP.MoveNext
        Loop
    End If

    MsgBox "Upload Completed!", vbInformation, "Info"

End Sub

Sub initMemvars()
    txtPartNo.Text = ""
    txtPartDesc.Text = ""
    txtICC.Text = ""
    txtModel.Text = ""
    txtNewPartNo.Text = ""
    txtDNP.Text = 0
    txtDNP2.Text = 0
    txtDNP3.Text = 0
    txtSRP.Text = 0
End Sub

Sub StoreMemVars()
    If Not rsDNPP.EOF And Not rsDNPP.BOF Then
        labid.Caption = rsDNPP!ID
        txtPartNo.Text = Null2String(rsDNPP!PARTNUMBER)
        txtPartDesc.Text = Null2String(rsDNPP!DESCRIPTIO)
        txtICC.Text = Null2String(rsDNPP!icc)
        txtModel.Text = Null2String(rsDNPP!Model)
        txtNewPartNo.Text = Null2String(rsDNPP!NewPARTNO)

        txtDNP.Text = ToDoubleNumber(N2Str2Zero(rsDNPP!DNPP))
        txtDNP2.Text = ToDoubleNumber(N2Str2Zero(rsDNPP!DNPP2))
        txtDNP3.Text = ToDoubleNumber(N2Str2Zero(rsDNPP!DNPP3))
        txtSRP.Text = ToDoubleNumber(N2Str2Zero(rsDNPP!SRP))
        Dim rsParts                                    As ADODB.Recordset
        Set rsParts = gconDMIS.Execute("SELECT TYPE FROM PMIS_STOCKMAS WHERE stockno=" & N2Str2Null(rsDNPP!PARTNUMBER))
        If Not (rsParts.EOF Or rsParts.BOF) Then
            LAB_PStatusDealer.ForeColor = &H8000&
            
            cmdRegister_parts.Enabled = False
            If Null2String(rsParts!Type) <> "P" Then
                If Null2String(rsParts!Type) = "A" Then
                    LAB_PStatusDealer = "Exists In Dealer Parts Master. Registered as Accessories"
                Else
                    LAB_PStatusDealer = "Exists In Dealer Parts Master. Registered as Materials"
                End If
            Else
                LAB_PStatusDealer = "Exists In Dealer Parts Master"
            End If

        Else
            LAB_PStatusDealer = "Doesn't Exists In Dealer Parts Master"
            LAB_PStatusDealer.ForeColor = vbRed
            cmdRegister_parts.Enabled = True
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsDNPP = New ADODB.Recordset
    rsDNPP.Open "select * from PMIS_DNPP order by PARTNUMBER asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsParts                                        As ADODB.Recordset
    lstPartsEntry.Enabled = False
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select PARTNUMBER, DESCRIPTIO from PMIS_DNPP")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsParts                                        As ADODB.Recordset
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    lstPartsEntry.Enabled = False
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select PARTNUMBER, DESCRIPTIO from PMIS_DNPP WHERE PARTNUMBER like'" & XXX & "%' order by PARTNUMBER Asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsParts                                        As ADODB.Recordset
    lstPartsEntry.Enabled = False
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    lstPartsEntry.Refresh
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select DESCRIPTIO,PARTNUMBER from PMIS_DNPP order by DESCRIPTIO asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
        lstPartsEntry.Enabled = True
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsParts                                        As ADODB.Recordset
    lstPartsEntry.Enabled = False
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsParts = gconDMIS.Execute("select DESCRIPTIO, PARTNUMBER from PMIS_DNPP where DESCRIPTIO like'" & XXX & "%' order by DESCRIPTIO Asc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MASTER HARIPARTS") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Screen.MousePointer = 11
    rptPrintParts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintParts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "HARIParts.rpt", "", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0

    Call NEW_LogAudit("V", "MASTER HARIPARTS", "", labid, "", "", "", "")

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "MASTER HARIPARTS") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    txtPartNo.Enabled = True
    lstPartsEntry.Enabled = False
    textSearch.Enabled = False
    optPartNo.Enabled = False
    optDescription.Enabled = False
    On Error Resume Next
    txtPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtPartNo.Enabled = False
    lstPartsEntry.Enabled = True
    textSearch.Enabled = True
    optPartNo.Enabled = True
    optDescription.Enabled = True
    optPartNo.Enabled = True
    optDescription.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo mmmm
    SQL_STATEMENT = "delete from PMIS_DNPPx where id = " & labid.Caption
    Exit Sub
mmmm:
    Dim err                                            As ADODB.error
    ShowADOErrors gconDMIS

    Exit Sub
    If Function_Access(LOGID, "Acess_Delete", "MASTER HARIPARTS") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsDNPP.BOF Or Not rsDNPP.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from PMIS_DNPP where id = " & labid.Caption

            gconDMIS.Execute SQL_STATEMENT
            Call NEW_LogAudit("X", "MASTER HARIPARTS", SQL_STATEMENT, labid, "", "HARI PARTS CODE: " & txtPartNo, "", "")

            'LogAudit "X", "DNPP ENTRY", txtDNP
            ShowDeletedMsg
            FillGrid
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MASTER HARIPARTS") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstPartsEntry.Enabled = False
    textSearch.Enabled = False
    optPartNo.Enabled = False
    optDescription.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
    'Picture3.Visible = False
End Sub

Private Sub cmdNext_Click()
    rsDNPP.MoveNext
    If rsDNPP.EOF Then
        rsDNPP.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsDNPP.MovePrevious
    If rsDNPP.BOF Then
        rsDNPP.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdRegister_Click()

    If Module_Access(LOGID, "PARTS REGISTRATION", "SYSTEM") = False Then Exit Sub

    If MsgBox("Confirm.", vbInformation + vbYesNo) = vbNo Then Exit Sub

    gconDMIS.Execute ("INSERT INTO PMIS_STOCKMAS (TYPE,STOCKNO, STOCKDESC, DNP,SRP,MODELCODE,ACTIVE,GENUINE,NON_HARI,DATE_ENTERED) SELECT 'P', PARTNUMBER,DESCRIPTIO,  DNPP2  ,SRP, MODEL,'Y','Y','N',GETDATE()   FROM PMIS_DNPP WHERE PARTNUMBER NOT IN(SELECT STOCKNO FROM PMIS_STOCKMAS) AND RIGHT(PARTNUMBER,2)<>'LP'")

    gconDMIS.Execute ("INSERT INTO PMIS_STOCKMAS (TYPE,STOCKNO, STOCKDESC, DNP,SRP,MODELCODE,ACTIVE,GENUINE,NON_HARI,DATE_ENTERED) SELECT 'P', PARTNUMBER,DESCRIPTIO,  DNPP2  ,SRP, MODEL,'Y','N','Y',GETDATE()   FROM PMIS_DNPP WHERE PARTNUMBER NOT IN(SELECT STOCKNO FROM PMIS_STOCKMAS) AND RIGHT(PARTNUMBER,2)='LP'")

    rsRefresh
    rsDNPP.Find ("ID=" & labid)
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim RSFINDDUP                                      As ADODB.Recordset

    Dim vtxtPARTNO, vtxtPARTDESC, VTXTICC              As String
    Dim VTXTmodel                                      As String
    Dim VtxtNewPARTNO                                  As String
    Dim VTXTSRP, VTXTDNP, VTXTDNP2, VTXTDNP3           As Double

    If IsNull(txtPartNo.Text) = True Then
        MsgSpeechBox "Part Number must not be empty"
        On Error Resume Next
        txtPartNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select PARTNUMBER from PMIS_DNPP where PARTNUMBER = '" & txtPartNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Part Number already exist!"
                On Error Resume Next
                txtPartNo.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtPartDesc.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        txtPartDesc.SetFocus
        Exit Sub
    End If

    vtxtPARTNO = N2Str2Null(txtPartNo.Text)
    vtxtPARTDESC = N2Str2Null(txtPartDesc.Text)
    VTXTICC = N2Str2Null(txtICC.Text)
    VTXTmodel = N2Str2Null(txtModel.Text)
    VtxtNewPARTNO = N2Str2Null(txtNewPartNo.Text)
    VTXTDNP = NumericVal(txtDNP.Text)
    VTXTDNP2 = NumericVal(txtDNP2.Text)
    VTXTDNP3 = NumericVal(txtDNP3.Text)
    VTXTSRP = NumericVal(txtSRP.Text)

    If AddorEdit = "ADD" Then

        SQL_STATEMENT = "Insert into PMIS_DNPP" & _
                      " (PARTNUMBER,DESCRIPTIO,ICC,model,newSTOCKNO,srp,dnpp,dnpp2,dnpp3,lastupdate,usercode)" & _
                      " values (" & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & VTXTICC & ", " & _
                      " " & VTXTmodel & ", " & VtxtNewPARTNO & ", " & VTXTSRP & _
                        ", " & VTXTDNP & "," & VTXTDNP2 & "," & VTXTDNP3 & ", '" & LOGDATE & "', '" & LOGCODE & "')"

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", "MASTER HARIPARTS", SQL_STATEMENT, labid, "", "HARI PARTS CODE: " & vtxtPARTNO, "", "")
        'LogAudit "A", "DNPP ENTRY"

    Else
        SQL_STATEMENT = "update PMIS_DNPP set" & _
                      " DESCRIPTIO = " & vtxtPARTDESC & "," & _
                      " ICC = " & VTXTICC & "," & _
                      " model = " & VTXTmodel & "," & _
                      " newPARTNO = " & VtxtNewPARTNO & "," & _
                      " srp = " & VTXTSRP & "," & _
                      " dnpp = " & VTXTDNP & "," & _
                      " dnpp2 = " & VTXTDNP2 & "," & _
                      " dnpp3 = " & VTXTDNP3 & "," & _
                      " lastupdate = '" & LOGDATE & "', " & _
                      " usercode = " & N2Str2Null(LOGCODE) & _
                      " where PARTNUMBER = " & vtxtPARTNO

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", "MASTER HARIPARTS", SQL_STATEMENT, labid, "", "HARI PARTS CODE: " & vtxtPARTNO, "", "")


        'LogAudit "U", "DNPP ENTRY", txtDNP

    End If
    rsRefresh
    rsDNPP.Find "PARTNUMBER =" & vtxtPARTNO
    cmdCancel.Value = True
    FillGrid
    Exit Sub

ErrorCode:
    ShowVBError
    cmdCancel.Value = True
    Exit Sub
End Sub

Private Sub cmdUpdateDistributorParts_Click()
    Dim MYPATH, PAYLNAME                               As String
    MYPATH = App.Path
    labstockno = ""
    CommonDialog1.Filter = "Access Files (*.MDB)|*.MDB"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DefaultExt = "MDB"
    CommonDialog1.DialogTitle = "Open HARI Database"
    CommonDialog1.FileName = ""
    PAYLNAME = CommonDialog1.FileName
    If MYPATH <> "\" Then
        CommonDialog1.FileName = MYPATH & "\" & CommonDialog1.FileName
    End If
    If PAYLNAME = "" Then
        CommonDialog1.FileName = "*.MDB"
    End If
    CommonDialog1.ShowOpen

    If err = 32755 Then Exit Sub
    FILNAME = CommonDialog1.FileName
    Dim CS                                             As String
    CS = wizVar.DecryptAccess("50726F@§É¥¥èï_Å∂†d}oNvmbÜÑlëmîßâN∂•èµÆí•®m±Ö¶®ñ±±p")
    On Error Resume Next
    Set gconHARI_DNP = New ADODB.Connection
    gconHARI_DNP.ConnectionString = CS & FILNAME
    gconHARI_DNP.Open
    gconHARI_DNP.CursorLocation = adUseClient
    If err = 32755 Then Exit Sub
    'SETODBC
    If MsgBox("Upload HARI Parts Now?", vbQuestion + vbYesNo, "New Price List Upload.") = vbYes Then
        Picture3.Visible = True
        Call DecideToUpload
        Picture3.Visible = False
        NEW_LogAudit "R", "MASTER HARIPARTS", "", labid, "", "", "", ""
    End If

    Set gconHARI_DNP = Nothing
    rsRefresh
    cmdCancel.Value = True
    FillGrid
    Exit Sub

ErrorCode:
    ShowADOErrors gconHARI_DNP
    On Error Resume Next
    MsgSpeechBox "Warning: HARI Database is Invalid or Corrupted!" & vbCrLf & _
                 "Inventory Menu will be Unloaded... Contact NETSPEED Inc. Immediately"
    gconHARI_DNP.Close
    Set gconHARI_DNP = Nothing
    Unload Me

    Exit Sub
End Sub

Private Sub cmdRegister_parts_Click()
    If Module_Access(LOGID, "PARTS REGISTRATION", "SYSTEM") = False Then Exit Sub

    If MsgBox("Confirm.", vbInformation + vbYesNo) = vbNo Then Exit Sub

    If Right(txtPartNo, 2) = "LP" Then
        gconDMIS.Execute ("INSERT INTO PMIS_STOCKMAS (TYPE,STOCKNO, STOCKDESC, DNP,SRP,MODELCODE,ACTIVE,GENUINE,NON_HARI,DATE_ENTERED) SELECT 'P', PARTNUMBER,DESCRIPTIO,  DNPP2  ,SRP, MODEL,'Y','N','Y',GETDATE()   FROM PMIS_DNPP WHERE PARTNUMBER=" & N2Str2Null(txtPartNo))
    Else
        gconDMIS.Execute ("INSERT INTO PMIS_STOCKMAS (TYPE,STOCKNO, STOCKDESC, DNP,SRP,MODELCODE,ACTIVE,GENUINE,NON_HARI,DATE_ENTERED) SELECT 'P', PARTNUMBER,DESCRIPTIO,  DNPP2  ,SRP, MODEL,'Y','Y','N',GETDATE()   FROM PMIS_DNPP WHERE PARTNUMBER=" & N2Str2Null(txtPartNo))
    End If




    rsRefresh
    rsDNPP.Find ("ID=" & labid)
    StoreMemVars


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyR
                optPartNo.Value = True: optPARTNO_Click
            Case vbKeyE
                optDescription.Value = True: optDescription_Click
        End Select
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = ""
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISMaster_DNPPEntry = Nothing
    UnloadForm Me
End Sub



Private Sub txtDNP2_Change()
    If AddorEdit = "" Then Exit Sub
    txtSRP.Text = ToDoubleNumber(NumericVal(txtDNP2.Text) * PARTS_MARK_UP_FROM_DNP)

End Sub

Private Sub txtDNP2_Validate(Cancel As Boolean)
    If NumericVal(txtDNP2) <= 0 Then
        MsgBox "Please Encode the DNP 32 For the Parts", vbInformation
        Cancel = True
    End If
End Sub

Private Sub txtSRP_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub lstPartsEntry_GotFocus()
    If optPartNo.Value = True Then
        rsDNPP.Bookmark = rsFind(rsDNPP.Clone, "PARTNUMBER", lstPartsEntry.SelectedItem).Bookmark
    Else
        rsDNPP.Bookmark = rsFind(rsDNPP.Clone, "PARTNUMBER", lstPartsEntry.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstPartsEntry_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optPartNo.Value = True Then
        rsDNPP.Bookmark = rsFind(rsDNPP.Clone, "PARTNUMBER", lstPartsEntry.SelectedItem).Bookmark
    Else
        rsDNPP.Bookmark = rsFind(rsDNPP.Clone, "PARTNUMBER", lstPartsEntry.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstPartsEntry_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPartsEntry
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

Private Sub lstPartsEntry_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstPartsEntry_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optPartNo.Value = True Then
        If Trim(textSearch.Text) = "" Then
            FillGrid
        Else
            FillSearchGrid (textSearch.Text)
        End If
    Else
        If Trim(textSearch.Text) = "" Then
            FillGrid2
        Else
            FillSearchGrid2 (textSearch.Text)
        End If
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstPartsEntry.ListItems.Count > 0 And lstPartsEntry.Enabled = True Then: lstPartsEntry.SetFocus
    End If
End Sub

Private Sub optDescription_Click()
    lstPartsEntry.ColumnHeaders(1).Text = "DESCRIPTION"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    lstPartsEntry.ColumnHeaders(1).Text = "PART NO."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picAdds.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MASTER HARIPARTS)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "MASTER HARIPARTS", "")

    End Select
End Sub

