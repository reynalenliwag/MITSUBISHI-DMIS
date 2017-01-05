VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCSMSTechnician_NEO 
   Caption         =   "Technician Master File"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   4155
      Left            =   990
      ScaleHeight     =   4095
      ScaleWidth      =   5655
      TabIndex        =   22
      Top             =   2130
      Width           =   5715
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   1350
         TabIndex        =   33
         Top             =   1980
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53149697
         CurrentDate     =   39310
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1350
         TabIndex        =   32
         Top             =   2910
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1350
         TabIndex        =   31
         Top             =   2430
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
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
         Left            =   3360
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel"
         Top             =   3210
         Width           =   705
      End
      Begin VB.CommandButton Command2 
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
         Left            =   4110
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":0490
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":05E2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel"
         Top             =   3150
         Width           =   705
      End
      Begin VB.CommandButton Command1 
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
         Left            =   4830
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":0920
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":0A72
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel"
         Top             =   3150
         Width           =   705
      End
      Begin VB.TextBox txtSkills 
         Height          =   1185
         Left            =   1350
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   690
         Width           =   4155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sponsor"
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
         Index           =   3
         Left            =   480
         TabIndex        =   28
         Top             =   3030
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Place"
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
         Index           =   2
         Left            =   750
         TabIndex        =   27
         Top             =   2550
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month / Year"
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
         TabIndex        =   26
         Top             =   2070
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Skills"
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
         Left            =   810
         TabIndex        =   24
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   7350
      ScaleHeight     =   1125
      ScaleWidth      =   1410
      TabIndex        =   13
      Top             =   4950
      Width           =   1410
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
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":0DB0
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":0F02
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Save Technician"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   660
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":1252
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":13A4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4845
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   2625
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   11
         Top             =   150
         Width           =   2505
      End
      Begin MSComctlLib.ListView lsvTechnician 
         Height          =   4245
         Left            =   30
         TabIndex        =   12
         Top             =   540
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7488
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
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
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":16E2
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "EMPNO"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   2670
      TabIndex        =   6
      Top             =   0
      Width           =   6105
      Begin MSComctlLib.ListView lsvSKILLS 
         Height          =   1935
         Left            =   60
         TabIndex        =   19
         Top             =   1380
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Skills"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Month / Year"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Place"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Sponsor"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblTECHNAME 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1770
         TabIndex        =   21
         Top             =   600
         Width           =   4185
      End
      Begin VB.Label lblTECHCODE 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1770
         TabIndex        =   20
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technican Skills"
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
         Left            =   90
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technician Code"
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
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Technician Name"
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
         Left            =   90
         TabIndex        =   7
         Top             =   690
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1425
      Left            =   4305
      ScaleHeight     =   1425
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   4020
      Width           =   4515
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
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":1844
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":1996
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to Previous Record"
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
         Left            =   750
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":1CF5
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":1E47
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Move to Next Record"
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
         Left            =   1470
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":219F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":22F1
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Find a Record"
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
         Left            =   2220
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":25EB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":273D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit Selected Technician"
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
         Left            =   2970
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":2A99
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":2BEB
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
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
         Left            =   3720
         MouseIcon       =   "frmCSMSTechnician_NEO.frx":2F51
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSTechnician_NEO.frx":30A3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   540
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   210
      TabIndex        =   16
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "frmCSMSTechnician_NEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
