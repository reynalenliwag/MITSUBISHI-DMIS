VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISPartsEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Parts Data Entry"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PartsEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   8970
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   2700
      TabIndex        =   12
      Top             =   0
      Width           =   6165
      Begin VB.TextBox txtResService 
         Alignment       =   1  'Right Justify
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
         Left            =   4740
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   2910
         Width           =   1245
      End
      Begin VB.TextBox txtSStock 
         Alignment       =   1  'Right Justify
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
         Left            =   4740
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   2160
         Width           =   1275
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   1  'Right Justify
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
         Left            =   4740
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1740
         Width           =   1275
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
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
         Left            =   4740
         MaxLength       =   12
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1350
         Width           =   1275
      End
      Begin VB.TextBox txtDate_Entered 
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
         Left            =   4710
         MaxLength       =   12
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   180
         Width           =   1335
      End
      Begin Crystal.CrystalReport rptPrintParts 
         Left            =   5640
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "List of New Part Numbers"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtGenNo 
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
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Type the parts gen. number."
         Top             =   2910
         Width           =   1965
      End
      Begin VB.TextBox txtNewNo 
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
         TabIndex        =   6
         Text            =   "Text1"
         ToolTipText     =   "Type the parts new number."
         Top             =   2520
         Width           =   1965
      End
      Begin VB.TextBox txtOldNo 
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
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Type the part's old number, if there's any."
         Top             =   2130
         Width           =   1965
      End
      Begin VB.TextBox txtPartNo 
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
         TabIndex        =   0
         Text            =   "Text1"
         ToolTipText     =   "Type the part number (e.g. MR241052)"
         Top             =   180
         Width           =   1755
      End
      Begin VB.TextBox txtLocation 
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
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the location where the part can be found (e.g. Q-1)"
         Top             =   1740
         Width           =   2085
      End
      Begin VB.TextBox txtModelCode 
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
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type the part's model code (e.g. WRANGLER)"
         Top             =   1350
         Width           =   2085
      End
      Begin VB.TextBox txtVehType 
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
         TabIndex        =   29
         Top             =   210
         Width           =   225
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Entered"
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
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reserve for Service"
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
         Height          =   465
         Left            =   3330
         TabIndex        =   26
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Safety Stock Qty"
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
         Height          =   435
         Left            =   3360
         TabIndex        =   25
         Top             =   2160
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Gen No."
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
         Top             =   2940
         Width           =   1245
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
         TabIndex        =   23
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "On Hand"
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
         Left            =   3690
         TabIndex        =   22
         Top             =   1770
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Parts SRP"
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
         Left            =   3600
         TabIndex        =   21
         Top             =   1410
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Old No."
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
         TabIndex        =   20
         Top             =   2190
         Width           =   1245
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
         TabIndex        =   19
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         TabIndex        =   16
         Top             =   1800
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
         Left            =   90
         TabIndex        =   15
         Top             =   1410
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Veh Type"
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   630
         Width           =   1245
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   4245
      Left            =   75
      TabIndex        =   30
      Top             =   0
      Width           =   2595
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
         Height          =   285
         Left            =   210
         TabIndex        =   34
         Top             =   720
         Width           =   2325
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
         Left            =   210
         TabIndex        =   33
         Top             =   450
         Value           =   -1  'True
         Width           =   2325
      End
      Begin VB.TextBox textSearch 
         Appearance      =   0  'Flat
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
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1020
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstPartsEntry 
         Height          =   2775
         Left            =   60
         TabIndex        =   32
         Top             =   1410
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   4895
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
         MouseIcon       =   "PartsEntry.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PART NUMBER"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   3150
      ScaleHeight     =   885
      ScaleWidth      =   5715
      TabIndex        =   36
      Top             =   3390
      Width           =   5715
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
         MouseIcon       =   "PartsEntry.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   15
         Width           =   705
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
         Left            =   705
         MouseIcon       =   "PartsEntry.frx":0EDD
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":102F
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   15
         Width           =   705
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
         Left            =   1425
         MouseIcon       =   "PartsEntry.frx":1387
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":14D9
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   15
         Width           =   705
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
         Left            =   2130
         MouseIcon       =   "PartsEntry.frx":17D3
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":1925
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   15
         Width           =   705
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
         Left            =   2835
         MouseIcon       =   "PartsEntry.frx":1C38
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":1D8A
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   15
         Width           =   705
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
         Left            =   3540
         MouseIcon       =   "PartsEntry.frx":20E6
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":2238
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   15
         Width           =   705
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
         Left            =   4260
         MouseIcon       =   "PartsEntry.frx":2563
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":26B5
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   15
         Width           =   705
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
         Left            =   4965
         MouseIcon       =   "PartsEntry.frx":2A1B
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":2B6D
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   15
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7395
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   45
      Top             =   3405
      Width           =   1470
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
         MouseIcon       =   "PartsEntry.frx":2ED3
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":3025
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   0
         Width           =   705
      End
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
         Left            =   705
         MouseIcon       =   "PartsEntry.frx":3375
         MousePointer    =   99  'Custom
         Picture         =   "PartsEntry.frx":34C7
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   270
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   600
      TabIndex        =   17
      Top             =   210
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmPMISPartsEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPartMas                          As ADODB.Recordset
Dim AddorEdit                          As String

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub
    Screen.MousePointer = 11
    rptPrintParts.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptPrintParts.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptPrintParts, PMIS_REPORT_PATH & "printparts.rpt", "month({partmas.date_entered}) = " & Month(LOGDATE) & " AND year({partmas.date_entered}) = " & Year(LOGDATE), DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemvars
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete") = False Then Exit Sub
    On Error GoTo ErrorCode
    If Not rsPartMas.BOF Or Not rsPartMas.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from PMIS_STOCKMAS where id = " & labid.Caption
            ShowDeletedMsg
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
    If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    textSearch.SetFocus
    'Picture3.Visible = False

    'Dim findStr As String
    'findStr = InputSpeechBox("Please Input Part Number or Description ...", txtPARTNO.Text)
    'If findStr <> "" Then
    '   On Error Resume Next
    '   rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKNO", findStr).Bookmark
    '   If Err.Number = 3021 Then
    '      On Error GoTo ErrorCode
    '      rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "STOCKDESC", findStr).Bookmark
    '   End If
    'End If
    'StoreMemvars
    'Exit Sub

    'ErrorCode:
    'If Err.Number = 3021 Then
    '   ShowCantFind findStr
    '   Resume Next
    'End If
End Sub

Private Sub cmdNext_Click()
    rsPartMas.MoveNext
    If rsPartMas.EOF Then
        rsPartMas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsPartMas.MovePrevious
    If rsPartMas.BOF Then
        rsPartMas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup                      As ADODB.Recordset

    Dim vtxtPARTNO, vtxtPARTDESC, VTXTVehType As String
    Dim VTXTModelCode, VTXTLocation, VTXTOldNo As String
    Dim VTXTNewNo, VTXTGenNo           As String
    Dim VTXTSRP, VTXTDNP               As Double
    Dim VTXTOnHand, VTXTMAC            As Double
    Dim VTXTSStock, VTXTResService     As Double

    If IsNull(txtPartNo.Text) = True Then
        MsgSpeechBox "Part Number must not be empty"
        On Error Resume Next
        txtPartNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select STOCKNO from PMIS_STOCKMAS where STOCKNO = '" & txtPartNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
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
    VTXTVehType = N2Str2Null(txtVehType.Text)
    VTXTModelCode = N2Str2Null(txtModelCode.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    VTXTOldNo = N2Str2Null(txtOldNo.Text)
    VTXTNewNo = N2Str2Null(txtNewNo.Text)
    VTXTGenNo = N2Str2Null(txtGenNo.Text)
    VTXTSRP = NumericVal(txtSRP.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    VTXTSStock = NumericVal(txtSStock.Text)
    VTXTResService = NumericVal(txtResService.Text)

    If AddorEdit = "ADD" Then
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            rsPartMas.MoveLast
            labid.Caption = NumericVal(rsPartMas!ID) + 1
        End If
        gconDMIS.Execute "Insert into PMIS_STOCKMAS" & _
                       " (STOCKNO,STOCKDESC,vehtype,modelcode,location,oldno,newno,genno,srp,dnp,onhand,mac,sstock,resservice,lastupdate,usercode,date_entered)" & _
                       " values (" & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & VTXTVehType & ", " & _
                       " " & VTXTModelCode & ", " & VTXTLocation & _
                         ", " & VTXTOldNo & ", " & VTXTNewNo & ", " & VTXTGenNo & ", " & VTXTSRP & _
                         ", " & VTXTDNP & ", " & VTXTOnHand & ", " & VTXTMAC & ", " & VTXTSStock & ", " & VTXTResService & ", '" & LOGDATE & "', '" & LOGCODE & "', '" & LOGDATE & "')"
        ShowSuccessFullyAdded
        LogAudit "A", "Parts StockMas-N-HARI"
    Else
        gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                       " STOCKNO = " & vtxtPARTNO & "," & _
                       " STOCKDESC = " & vtxtPARTDESC & "," & _
                       " vehtype = " & VTXTVehType & "," & _
                       " modelcode = " & VTXTModelCode & "," & _
                       " location = " & VTXTLocation & "," & _
                       " oldno = " & VTXTOldNo & "," & _
                       " newno = " & VTXTNewNo & "," & _
                       " genno = " & VTXTGenNo & "," & _
                       " srp = " & VTXTSRP & "," & _
                       " dnp = " & VTXTDNP & "," & _
                       " sstock = " & VTXTSStock & "," & _
                       " resservice = " & VTXTResService & ", " & _
                       " lastupdate = '" & LOGDATE & "', " & _
                       " usercode = " & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labid.Caption
        ShowSuccessFullyUpdated
        LogAudit "U", "Parts StockMas-N-HARI", txtPartNo & "-" & txtPartDesc
    End If
    rsRefresh
    On Error Resume Next
    rsPartMas.Find "id =" & labid.Caption
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    ShowVBError
    cmdCancel.Value = True
    Exit Sub
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
    rsRefresh
    Frame1.Enabled = False
    SetFormSettings Me
    textSearch.Text = "":                             'Picture3.ZOrder 0
    InitMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub InitMemvars()
    txtPartNo.Text = ""
    txtPartDesc.Text = ""
    txtVehType.Text = ""
    txtModelCode.Text = ""
    txtLocation.Text = ""
    txtOldNo.Text = ""
    txtNewNo.Text = ""
    txtGenNo.Text = ""
    txtSRP.Text = 0
    txtOnHand.Text = 0
    txtSStock.Text = 0
    txtResService.Text = 0
End Sub

Sub StoreMemVars()
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        labid.Caption = rsPartMas!ID
        txtDate_Entered.Text = Format(Null2Date(rsPartMas!DATE_ENTERED), "DD-MMM-YY")
        txtPartNo.Text = Null2String(rsPartMas!STOCKNO)
        txtPartDesc.Text = Null2String(rsPartMas!STOCKDESC)
        txtVehType.Text = Null2String(rsPartMas!vehtype)
        txtModelCode.Text = Null2String(rsPartMas!ModelCode)
        txtLocation.Text = Null2String(rsPartMas!Location)
        txtOldNo.Text = Null2String(rsPartMas!oldno)
        txtNewNo.Text = Null2String(rsPartMas!newno)
        txtGenNo.Text = Null2String(rsPartMas!genno)
        txtSRP.Text = N2Str2Zero(rsPartMas!SRP)
        txtOnHand.Text = N2Str2IntZero(rsPartMas!ONHAND)
        txtSStock.Text = N2Str2IntZero(rsPartMas!SSTOCK)
        txtResService.Text = N2Str2IntZero(rsPartMas!RESSERVICE)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select * from PMIS_STOCKMAS where MONTH(DATE_ENTERED) = " & Month(LOGDATE) & " order by DATE_ENTERED desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISPartsEntry = Nothing
    UnloadForm Me
End Sub

Private Sub optDescription_Click()
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    textSearch.SetFocus
End Sub

Private Sub optPARTNO_Click()
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    textSearch.SetFocus
End Sub

Private Sub txtOnHand_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtResService_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtSRP_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub txtSStock_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub lstPartsEntry_GotFocus()
    rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "ID", lstPartsEntry.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstPartsEntry_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsPartMas.Bookmark = rsFind(rsPartMas.Clone, "ID", lstPartsEntry.SelectedItem.SubItems(1)).Bookmark
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
    If KeyCode = vbKeyEscape Then textSearch.SetFocus
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

Sub FillGrid()
    Dim rsParts                        As ADODB.Recordset
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKNO,ID from PMIS_STOCKMAS where MONTH(DATE_ENTERED) = " & Month(LOGDATE) & " order by DATE_ENTERED desc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsParts                        As ADODB.Recordset
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKNO, ID from PMIS_STOCKMAS where MONTH(DATE_ENTERED) = " & Month(LOGDATE) & " and STOCKNO like'" & XXX & "%' order by DATE_ENTERED desc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsParts                        As ADODB.Recordset
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKDESC,ID from PMIS_STOCKMAS where MONTH(DATE_ENTERED) = " & Month(LOGDATE) & " order by DATE_ENTERED desc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsParts                        As ADODB.Recordset
    lstPartsEntry.Sorted = False: lstPartsEntry.ListItems.Clear
    Set rsParts = New ADODB.Recordset
    Set rsParts = gconDMIS.Execute("select STOCKDESC, ID from PMIS_STOCKMAS where MONTH(DATE_ENTERED) = " & Month(LOGDATE) & " and STOCKDESC like'" & XXX & "%' order by DATE_ENTERED desc")
    If Not (rsParts.EOF And rsParts.BOF) Then
        lstPartsEntry.Enabled = True
        Listview_Loadval Me.lstPartsEntry.ListItems, rsParts
        lstPartsEntry.Refresh
    Else
        lstPartsEntry.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstPartsEntry.SetFocus
End Sub
