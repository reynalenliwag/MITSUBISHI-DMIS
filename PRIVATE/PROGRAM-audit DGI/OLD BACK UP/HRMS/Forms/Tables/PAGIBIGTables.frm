VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmHRMSTables_PAGIBIG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAG-IBIG Table"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7905
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PAGIBIGTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   7905
   Begin VB.PictureBox Picture3 
      Height          =   4695
      Left            =   60
      ScaleHeight     =   4635
      ScaleWidth      =   7725
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   7785
      Begin VB.PictureBox Picture7 
         Enabled         =   0   'False
         Height          =   615
         Left            =   180
         ScaleHeight     =   555
         ScaleWidth      =   6975
         TabIndex        =   46
         Top             =   540
         Width           =   7035
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   630
            TabIndex        =   51
            Text            =   "Text3"
            Top             =   120
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Height          =   345
            Left            =   2700
            TabIndex        =   48
            Text            =   "Text1"
            Top             =   90
            Width           =   1065
         End
         Begin VB.TextBox Text2 
            Height          =   345
            Left            =   5160
            TabIndex        =   47
            Text            =   "Text1"
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Code"
            Height          =   345
            Left            =   90
            TabIndex        =   52
            Top             =   150
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Employee Share"
            Height          =   345
            Left            =   1380
            TabIndex        =   50
            Top             =   150
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Employer Share"
            Height          =   345
            Left            =   3930
            TabIndex        =   49
            Top             =   150
            Width           =   1155
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   150
         TabIndex        =   31
         Top             =   1350
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
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
         MouseIcon       =   "PAGIBIGTables.frx":0442
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pagibig Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Employee Share"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Employer Share"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   240
         ScaleHeight     =   855
         ScaleWidth      =   5235
         TabIndex        =   33
         Top             =   3150
         Width           =   5235
         Begin VB.CommandButton cmdDeleteDET 
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
            Left            =   3690
            MouseIcon       =   "PAGIBIGTables.frx":05A4
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":06F6
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Delete Selected Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command7 
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
            Left            =   210
            MouseIcon       =   "PAGIBIGTables.frx":0A21
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":0B73
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command6 
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
            Left            =   900
            MouseIcon       =   "PAGIBIGTables.frx":0ED2
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":1024
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command5 
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
            Left            =   1590
            MouseIcon       =   "PAGIBIGTables.frx":137C
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":14CE
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Find a Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command4 
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
            Left            =   2280
            MouseIcon       =   "PAGIBIGTables.frx":17C8
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":191A
            Style           =   1  'Graphical
            TabIndex        =   36
            ToolTipText     =   "Add Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command3 
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
            Left            =   2970
            MouseIcon       =   "PAGIBIGTables.frx":1C2D
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":1D7F
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command2 
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
            Left            =   4380
            MouseIcon       =   "PAGIBIGTables.frx":20DB
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":222D
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Exit Window"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   3870
         ScaleHeight     =   885
         ScaleWidth      =   1545
         TabIndex        =   40
         Top             =   3150
         Width           =   1545
         Begin VB.CommandButton Command9 
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
            Left            =   75
            MouseIcon       =   "PAGIBIGTables.frx":2593
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":26E5
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Save Entry"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton Command8 
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
            Left            =   765
            MouseIcon       =   "PAGIBIGTables.frx":2A35
            MousePointer    =   99  'Custom
            Picture         =   "PAGIBIGTables.frx":2B87
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   435
         Left            =   0
         TabIndex        =   45
         Top             =   4200
         Width           =   7725
         _Version        =   655364
         _ExtentX        =   13626
         _ExtentY        =   767
         _StockProps     =   14
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   435
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   7725
         _Version        =   655364
         _ExtentX        =   13626
         _ExtentY        =   767
         _StockProps     =   14
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
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PIF Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5250
         TabIndex        =   43
         Top             =   1740
         Width           =   2655
      End
      Begin VB.Image Image2 
         Height          =   1125
         Left            =   6000
         Picture         =   "PAGIBIGTables.frx":2EC5
         Top             =   2100
         Width           =   1200
      End
      Begin VB.Shape Shape4 
         Height          =   2085
         Left            =   5580
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label LABDETID 
         Caption         =   "LABDETID"
         Height          =   315
         Left            =   6090
         TabIndex        =   32
         Top             =   3540
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OVERRIDE TABLE"
      Height          =   435
      Left            =   1590
      TabIndex        =   29
      Top             =   4080
      Width           =   1845
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2730
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   19
      Top             =   3870
      Width           =   5625
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
         Left            =   4380
         MouseIcon       =   "PAGIBIGTables.frx":78E3
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":7A35
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   3690
         MouseIcon       =   "PAGIBIGTables.frx":7D9B
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":7EED
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   3000
         MouseIcon       =   "PAGIBIGTables.frx":8249
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":839B
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   2310
         MouseIcon       =   "PAGIBIGTables.frx":86AE
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":8800
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   1620
         MouseIcon       =   "PAGIBIGTables.frx":8AFA
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":8C4C
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
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
         Left            =   930
         MouseIcon       =   "PAGIBIGTables.frx":8FA4
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":90F6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picPAGIBIGTable 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   45
      ScaleHeight     =   2535
      ScaleWidth      =   7905
      TabIndex        =   6
      Top             =   60
      Width           =   7905
      Begin VB.TextBox txtER 
         Alignment       =   1  'Right Justify
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
         Left            =   2385
         MaxLength       =   100
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1620
         Width           =   1695
      End
      Begin VB.TextBox txtPAG_CODE 
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
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   60
         Width           =   1425
      End
      Begin VB.TextBox txtMAX_MC 
         Alignment       =   1  'Right Justify
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
         Left            =   2370
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2025
         Width           =   1710
      End
      Begin VB.TextBox txtFROM 
         Alignment       =   1  'Right Justify
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
         Left            =   1020
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   450
         Width           =   1440
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
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
         Left            =   3420
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox txtPercent 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1620
         Width           =   1695
      End
      Begin Crystal.CrystalReport rptSalaryGrade 
         Left            =   7380
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Shape Shape3 
         Height          =   2085
         Left            =   5640
         Top             =   90
         Width           =   1905
      End
      Begin VB.Shape Shape2 
         Height          =   945
         Left            =   0
         Top             =   0
         Width           =   5325
      End
      Begin VB.Shape Shape1 
         Height          =   1545
         Left            =   0
         Top             =   930
         Width           =   5325
      End
      Begin VB.Image Image1 
         Height          =   1125
         Left            =   6030
         Picture         =   "PAGIBIGTables.frx":9455
         Top             =   960
         Width           =   1200
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PIF Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5280
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer's Share"
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
         Height          =   255
         Left            =   2385
         TabIndex        =   17
         Top             =   1335
         Width           =   1845
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contributions"
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
         Left            =   150
         TabIndex        =   16
         Top             =   990
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Max MC"
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
         Height          =   255
         Left            =   1530
         TabIndex        =   12
         Top             =   2085
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range1"
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
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range2"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   10
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's Share"
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
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   225
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   90
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   60
      ScaleHeight     =   1305
      ScaleWidth      =   7905
      TabIndex        =   14
      Top             =   2610
      Width           =   7905
      Begin MSComctlLib.ListView lstPAGIBIGTable 
         Height          =   1215
         Left            =   -30
         TabIndex        =   15
         Top             =   0
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
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
         MouseIcon       =   "PAGIBIGTables.frx":DE73
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Pagibig Code"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Range From"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Range To"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Employee Share"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Employer Share"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Maximum MC"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6300
      ScaleHeight     =   885
      ScaleWidth      =   1755
      TabIndex        =   26
      Top             =   3885
      Width           =   1755
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
         Left            =   765
         MouseIcon       =   "PAGIBIGTables.frx":DFD5
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":E127
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
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
         Left            =   75
         MouseIcon       =   "PAGIBIGTables.frx":E465
         MousePointer    =   99  'Custom
         Picture         =   "PAGIBIGTables.frx":E5B7
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSTables_PAGIBIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPAGIBIGTable                                                    As ADODB.Recordset
Dim rsTemp                                                            As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim AddorEditDET                                                      As String

Function CheckIfExist(PAGCODE As Integer, CHOICE As String, DETID As Integer) As Boolean
    CheckIfExist = False
    Dim RSCHECK                                                       As ADODB.Recordset
    Set RSCHECK = New ADODB.Recordset

    If CHOICE = "ADD" Then
        Set RSCHECK = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIGTABLE WHERE PAG_CODE IS NULL AND SETBYUSER IS NOT NULL AND SETBYUSER = '" & PAGCODE & "'")
        If RSCHECK.RecordCount > 0 Then
            CheckIfExist = True
        End If
    ElseIf CHOICE = "EDIT" Then
        Set RSCHECK = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIGTABLE WHERE PAG_CODE IS NULL AND SETBYUSER IS NOT NULL AND SETBYUSER = '" & PAGCODE & "' AND ID <> '" & DETID & "'")
        If RSCHECK.RecordCount > 0 Then
            CheckIfExist = True
        End If
    End If
    Set RSCHECK = Nothing
End Function

Sub rsrefresh()
    Set rsPAGIBIGTable = New ADODB.Recordset
    'rsPAGIBIGTable.Open "select * from HRMS_PAGIBIGTable order by PAG_CODE", gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsPAGIBIGTable.Open "SELECT * FROM HRMS_PAGIBIGTABLE WHERE SETBYUSER = '0' OR SETBYUSER IS NULL ORDER BY PAG_CODE", gconDMIS, adOpenForwardOnly, adLockReadOnly
    FillGrid
End Sub

Sub InitMemvars()
    picPAGIBIGTable.Enabled = True
    txtPAG_CODE.Text = ""
    txtFROM.Text = 0
    txtTo.Text = 0
    txtMAX_MC.Text = ""
    txtPercent.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsPAGIBIGTable.EOF And Not rsPAGIBIGTable.BOF Then
        picPAGIBIGTable.Enabled = False
        labID.Caption = rsPAGIBIGTable!ID
        txtPAG_CODE.Text = Null2String(rsPAGIBIGTable!PAG_CODE)
        txtFROM.Text = Null2String(rsPAGIBIGTable!From)
        txtTo.Text = Null2String(rsPAGIBIGTable!To)
        txtMAX_MC.Text = Null2String(rsPAGIBIGTable!Max_MC)
        txtPercent.Text = Null2String(rsPAGIBIGTable!Percent)
        txtER.Text = Null2String(rsPAGIBIGTable!ER)
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsPAGIBIGTable2                                               As ADODB.Recordset
    lstPAGIBIGTable.Sorted = False
    lstPAGIBIGTable.ListItems.Clear
    lstPAGIBIGTable.Enabled = False
    Set rsPAGIBIGTable2 = New ADODB.Recordset
    'Set rsPAGIBIGTable2 = gconDMIS.Execute("select PAG_CODE,[FROM],[TO],[Percent],ER,Max_MC from HRMS_PAGIBIGTable")
    Set rsPAGIBIGTable2 = gconDMIS.Execute("select PAG_CODE,[FROM],[TO],[Percent],ER,Max_MC from HRMS_PAGIBIGTable WHERE SETBYUSER = '0' OR SETBYUSER IS NULL")
    If Not (rsPAGIBIGTable2.EOF And rsPAGIBIGTable2.BOF) Then
        Listview_Loadval Me.lstPAGIBIGTable.ListItems, rsPAGIBIGTable2
        lstPAGIBIGTable.Refresh
        lstPAGIBIGTable.Enabled = True
    End If
End Sub

Sub rsRefreshDet()
    Set rsTemp = New ADODB.Recordset
    Set rsTemp = gconDMIS.Execute("SELECT SETBYUSER, [FROM], [TO], ID FROM HRMS_PAGIBIGTABLE WHERE SETBYUSER IS NOT NULL")
End Sub

Sub storeMemvarsDet()
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Text1.Text = N2Str2Zero(rsTemp!From)
        Text2.Text = N2Str2Zero(rsTemp!To)
        Text3.Text = N2Str2Zero(rsTemp!SETBYUSER)
        LABDETID.Caption = Null2String(rsTemp!ID)
    Else
        InitMemvarsDet
    End If
End Sub

Sub InitMemvarsDet()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    LABDETID.Caption = ""
End Sub

Sub FillGridDet()
    ListView1.Sorted = False
    ListView1.ListItems.Clear
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        Listview_Loadval Me.ListView1.ListItems, rsTemp
        ListView1.Refresh
        ListView1.Enabled = True
    End If
    Set rsTemp = Nothing
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "PAG-IBIG TABLE") = False Then Exit Sub
    AddorEdit = "ADD"
    InitMemvars
    lstPAGIBIGTable.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtPAG_CODE.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picPAGIBIGTable.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstPAGIBIGTable.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "PAB-IBIG TABLE") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_PAGIBIGTable where id = " & labID.Caption
        Call LogAudit("X", "DELETE PAGIBIG MASTERFILE RECORD", txtPAG_CODE.Text)
        Call ShowDeletedMsg
    End If
    Call rsrefresh
    Call StoreMemVars
    Exit Sub
Errorcode:
    Call ShowVBError
End Sub

Private Sub cmdDeleteDET_Click()
    gconDMIS.Execute ("DELETE FROM HRMS_PAGIBIGTABLE WHERE ID = '" & LABDETID.Caption & "'")
    Command8.Value = True
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "PAG-IBIG TABLE") = False Then Exit Sub
    AddorEdit = "EDIT"
    picPAGIBIGTable.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstPAGIBIGTable.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    lstPAGIBIGTable.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsPAGIBIGTable.MoveNext
    If rsPAGIBIGTable.EOF Then
        rsPAGIBIGTable.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsPAGIBIGTable.MovePrevious
    If rsPAGIBIGTable.BOF Then
        rsPAGIBIGTable.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Dim vtxtPAG_CODE                                                  As String
    Dim vtxtFROM, vtxtTO                                              As Double
    Dim vtxtMax_MC, vtxtPercent, vtxtER
    On Error GoTo Errorcode:
    vtxtPAG_CODE = N2Str2Null(txtPAG_CODE.Text)
    vtxtFROM = NumericVal(txtFROM.Text)
    vtxtTO = NumericVal(txtTo.Text)
    vtxtMax_MC = N2Str2Null(txtMAX_MC.Text)
    vtxtPercent = N2Str2Null(txtPercent.Text)
    vtxtER = N2Str2Null(txtER.Text)
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_PAGIBIGTable " & _
                         "(PAG_CODE,FROM,TO,Max_MC,Percent,ER,LastUpdate,USERCODE) " & _
                       " values (" & vtxtPAG_CODE & ", " & _
                         "" & vtxtFROM & ", " & vtxtTO & ", " & vtxtMax_MC & ", " & vtxtPercent & ", " & vtxtER & ", '" & LOGDATE & "', '" & LOGCODE & "')"

        Call LogAudit("A", "ADD PAGIBIG RECORD", vtxtPAG_CODE)
    Else
        gconDMIS.Execute "update HRMS_PAGIBIGTable set" & _
                       " PAG_CODE = " & vtxtPAG_CODE & "," & _
                       " [FROM] = " & vtxtFROM & "," & _
                       " [TO] = " & vtxtTO & "," & _
                       " Max_MC = " & vtxtMax_MC & "," & _
                       " [Percent] = " & vtxtPercent & "," & _
                       " ER = " & vtxtER & "," & _
                       " LastUpdate = '" & LOGDATE & "'," & _
                       " USERCODE = '" & LOGCODE & "'" & _
                       " where id = " & labID.Caption

        Call LogAudit("E", "UPDATE PAGIBIG TABLE", vtxtPAG_CODE)
    End If
    Call rsrefresh
    On Error Resume Next
    rsPAGIBIGTable.Find "PAG_CODE = " & vtxtPAG_CODE
    cmdCancel.Value = True
    Exit Sub
Errorcode:
    Call ShowVBError
End Sub

Private Sub Command1_Click()

    Picture3.Visible = True
    InitMemvarsDet
    rsRefreshDet
    storeMemvarsDet
    FillGridDet

End Sub

Private Sub Command2_Click()
    Picture3.Visible = False
End Sub

Private Sub Command3_Click()
    AddorEditDET = "EDIT"
    Picture4.Visible = False
    Picture6.Visible = True
    ListView1.Enabled = False
    Picture7.Enabled = True
End Sub

Private Sub Command4_Click()
    InitMemvarsDet
    AddorEditDET = "ADD"
    Picture4.Visible = False
    Picture6.Visible = True
    ListView1.Enabled = False
    Picture7.Enabled = True
End Sub

Private Sub Command8_Click()
    Picture4.Visible = True
    Picture6.Visible = False
    ListView1.Enabled = True
    Picture7.Enabled = False
    rsRefreshDet
    storeMemvarsDet
    FillGridDet
End Sub

Private Sub Command9_Click()
    Dim vCODE                                                         As Integer
    Dim vEMPRSHARE                                                    As Double
    Dim vEMPESHARE                                                    As Double
    Dim vID                                                           As Integer

    vCODE = N2Str2Zero(Text3.Text)
    vEMPRSHARE = N2Str2Zero(Text1.Text)
    vEMPESHARE = N2Str2Zero(Text2.Text)
    vID = N2Str2Zero(LABDETID.Caption)

    If Text3.Text <> "" Then
        If AddorEditDET = "ADD" Then
            If CheckIfExist(vCODE, "ADD", vID) = False Then
                gconDMIS.Execute ("INSERT INTO HRMS_PAGIBIGTABLE ([FROM], [TO], SETBYUSER)" & _
                                " VALUES(" & vEMPESHARE & "," & vEMPESHARE & "," & vCODE & ")")
            Else
                ShowAlreadyExistMsg "CODE"
            End If
        ElseIf AddorEditDET = "EDIT" Then
            If CheckIfExist(vCODE, "EDIT", vID) = False Then
                gconDMIS.Execute "UPDATE HRMS_PAGIBIGTABLE SET" & _
                               " [FROM] = " & vEMPESHARE & "," & _
                               " [TO] = " & vEMPESHARE & "," & _
                               " SETBYUSER = " & vCODE & _
                               " where id = " & LABDETID.Caption
            Else
                ShowAlreadyExistMsg "CODE"
            End If
        Else
            Exit Sub
        End If
        Command8.Value = True
    End If
End Sub

Private Sub Form_KeyDown(KeyPAG_CODE As Integer, Shift As Integer)
    MoveKeyPress KeyPAG_CODE
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsrefresh
    StoreMemVars
    FillGrid
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstPAGIBIGTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPAGIBIGTable
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lstPAGIBIGTable_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstPAGIBIGTable_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    rsPAGIBIGTable.Bookmark = rsFind(rsPAGIBIGTable.Clone, "PAG_CODE", Me.lstPAGIBIGTable.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub ListView1_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    On Error Resume Next
    Dim rstemp2                                                       As ADODB.Recordset
    Set rstemp2 = New ADODB.Recordset
    Set rstemp2 = gconDMIS.Execute("SELECT * FROM HRMS_PAGIBIGTABLE WHERE PAG_CODE IS NULL AND SETBYUSER = '" & ListView1.SelectedItem & "'")
    If Not rstemp2.EOF And Not rstemp2.BOF Then
        Text1.Text = N2Str2Zero(rstemp2!From)
        Text2.Text = N2Str2Zero(rstemp2!To)
        Text3.Text = N2Str2Zero(rstemp2!SETBYUSER)
        LABDETID.Caption = Null2String(rstemp2!ID)
    End If
    Set rstemp2 = Nothing
End Sub

