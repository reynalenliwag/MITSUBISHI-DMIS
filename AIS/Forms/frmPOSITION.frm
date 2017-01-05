VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAISPOSITION 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jobs/ Position"
   ClientHeight    =   6900
   ClientLeft      =   1740
   ClientTop       =   2010
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPOSITION.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9225
   Begin VB.PictureBox picAdd 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   -450
      ScaleHeight     =   960
      ScaleWidth      =   9495
      TabIndex        =   32
      Top             =   5970
      Width           =   9495
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
         Left            =   8760
         MouseIcon       =   "frmPOSITION.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Left            =   8040
         MouseIcon       =   "frmPOSITION.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   36
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
         Left            =   7320
         MouseIcon       =   "frmPOSITION.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   6600
         MouseIcon       =   "frmPOSITION.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Left            =   5880
         MouseIcon       =   "frmPOSITION.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":1977
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         Left            =   5160
         MouseIcon       =   "frmPOSITION.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   4440
         MouseIcon       =   "frmPOSITION.frx":212C
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":227E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Move to First Record"
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
         Left            =   3720
         MouseIcon       =   "frmPOSITION.frx":25DC
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":272E
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   3000
         MouseIcon       =   "frmPOSITION.frx":2A28
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":2B7A
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Left            =   2280
         MouseIcon       =   "frmPOSITION.frx":2ED2
         MousePointer    =   99  'Custom
         Picture         =   "frmPOSITION.frx":3024
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.TabControl tcPOSITION 
      Height          =   5865
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9195
      _Version        =   655364
      _ExtentX        =   16219
      _ExtentY        =   10345
      _StockProps     =   64
      AllowReorder    =   -1  'True
      Appearance      =   9
      Color           =   4
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Main"
      Item(0).Tooltip =   "Main"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "picTAB_MAIN"
      Item(0).Control(1)=   "lsvPOSITION"
      Item(1).Caption =   "Required Education"
      Item(1).Tooltip =   "Required Education"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "picTAB_EDU"
      Item(2).Caption =   "Documents Required"
      Item(2).Tooltip =   "Documents Required"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "picTAB_DOC"
      Begin VB.PictureBox picTAB_DOC 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   -70000
         ScaleHeight     =   5505
         ScaleWidth      =   9165
         TabIndex        =   27
         Top             =   330
         Visible         =   0   'False
         Width           =   9195
         Begin VB.PictureBox picPOSITION_DOC 
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   7320
            ScaleHeight     =   825
            ScaleWidth      =   1875
            TabIndex        =   29
            Top             =   3000
            Width           =   1875
            Begin VB.CommandButton cmdPOSITION_ADD_DOC 
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   840
               Picture         =   "frmPOSITION.frx":3383
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Add Documents Required "
               Top             =   0
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView lsvDOC 
            Height          =   2800
            Left            =   105
            TabIndex        =   13
            Top             =   150
            Width           =   8900
            _ExtentX        =   15690
            _ExtentY        =   4948
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Document Type"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Note"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   18
            EndProperty
         End
      End
      Begin VB.PictureBox picTAB_EDU 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   -70000
         ScaleHeight     =   5505
         ScaleWidth      =   9165
         TabIndex        =   26
         Top             =   330
         Visible         =   0   'False
         Width           =   9195
         Begin VB.PictureBox picPOSITION_EDU 
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   7320
            ScaleHeight     =   825
            ScaleWidth      =   1695
            TabIndex        =   28
            Top             =   3000
            Width           =   1695
            Begin VB.CommandButton cmdPOSITION_ADD_EDU 
               Caption         =   "ADD"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   765
               Left            =   840
               Picture         =   "frmPOSITION.frx":390E
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "Add Required Education"
               Top             =   0
               Width           =   795
            End
         End
         Begin MSComctlLib.ListView lsvEDUC 
            Height          =   2800
            Left            =   105
            TabIndex        =   11
            Top             =   150
            Width           =   8900
            _ExtentX        =   15690
            _ExtentY        =   4948
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Degree"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fields"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Notes"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   9
            EndProperty
         End
      End
      Begin VB.PictureBox picTAB_MAIN 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2805
         Left            =   0
         ScaleHeight     =   2775
         ScaleWidth      =   9135
         TabIndex        =   17
         Top             =   360
         Width           =   9165
         Begin VB.TextBox txtPOSITION_TAKEN 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   6450
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1800
            Width           =   1245
         End
         Begin VB.TextBox txtPOS_ID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   345
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   0
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox txtPOSITION_POS 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2175
            MaxLength       =   50
            TabIndex        =   1
            Top             =   570
            Width           =   6075
         End
         Begin VB.TextBox txtPOSITION_AVAILABLE 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   6450
            MaxLength       =   3
            TabIndex        =   6
            Top             =   1410
            Width           =   1245
         End
         Begin VB.TextBox txtPOSITION_TOAGE 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   4
            Top             =   1800
            Width           =   1305
         End
         Begin VB.TextBox txtPOSITION_FROMAGE 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   3
            Top             =   1380
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker DTPPOSITION_ASOF 
            Height          =   345
            Left            =   2160
            TabIndex        =   2
            Top             =   990
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50724865
            CurrentDate     =   39132
         End
         Begin MSComCtl2.DTPicker DTPPOSITION_DEADLINE 
            Height          =   345
            Left            =   6000
            TabIndex        =   5
            Top             =   1020
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50724865
            CurrentDate     =   39132
         End
         Begin MSComCtl2.DTPicker dtpDATEINACTIVE 
            Height          =   345
            Left            =   2160
            TabIndex        =   31
            Top             =   2220
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   609
            _Version        =   393216
            Format          =   50724865
            CurrentDate     =   39132
         End
         Begin Crystal.CrystalReport rptPosition 
            Left            =   180
            Top             =   2190
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Inactive"
            Height          =   240
            Index           =   0
            Left            =   705
            TabIndex        =   30
            Top             =   2250
            Width           =   1350
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Age"
            Height          =   240
            Index           =   3
            Left            =   1380
            TabIndex        =   25
            Top             =   1800
            Width           =   705
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
            Height          =   240
            Index           =   20
            Left            =   1785
            TabIndex        =   24
            Top             =   300
            Width           =   210
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "As Of"
            Height          =   240
            Index           =   7
            Left            =   1485
            TabIndex        =   23
            Top             =   1050
            Width           =   540
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position Description"
            Height          =   240
            Index           =   6
            Left            =   90
            TabIndex        =   22
            Top             =   630
            Width           =   1935
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deadline"
            Height          =   240
            Index           =   1
            Left            =   5025
            TabIndex        =   21
            Top             =   1020
            Width           =   825
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From Age"
            Height          =   240
            Index           =   2
            Left            =   1110
            TabIndex        =   20
            Top             =   1380
            Width           =   930
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Position Taken"
            Height          =   240
            Index           =   8
            Left            =   4170
            TabIndex        =   19
            Top             =   1830
            Width           =   2115
         End
         Begin VB.Label lblCAP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Position Available"
            Height          =   240
            Index           =   10
            Left            =   3930
            TabIndex        =   18
            Top             =   1470
            Width           =   2385
         End
      End
      Begin MSComctlLib.ListView lsvPOSITION 
         Height          =   2355
         Left            =   210
         TabIndex        =   8
         Top             =   3300
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   4154
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Position Description"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Available Position"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Position Taken"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Date Inactive"
            Object.Width           =   2469
         EndProperty
      End
   End
   Begin VB.PictureBox picSAVE 
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
      Height          =   825
      Left            =   7500
      ScaleHeight     =   825
      ScaleWidth      =   1605
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   1605
      Begin VB.CommandButton cmdCANCEL 
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
         Height          =   765
         Left            =   840
         Picture         =   "frmPOSITION.frx":3E99
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   180
         Picture         =   "frmPOSITION.frx":4415
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Save Jobs and Position"
         Top             =   0
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmAISPOSITION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FUNCTION /FEATURE:Add new Control Button For Prev/Next/First/Last
'DATE STARTED:06/08/2007
'LAST UPDATE:
'DATABASE UPDATE:Create new Table  name CSMS_ConcernResolution
'WHO UPDATE:HardNard
'UPDATING CODE:BTT - 06/07/2007
'**********************************************************************************
Option Explicit
Dim RS                                                                As New ADODB.Recordset    'BTT- 07032007
Dim SAVE_OR_EDIT_POSITION                                             As String

Function GenerateNewPositionID() As Integer
    Dim RSTMP                                                         As ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select Pos_ID From HRMS_Position Order By POS_ID ASC")

    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            GenerateNewPositionID = CInt(RSTMP!POS_ID)
            RSTMP.MoveNext
        Loop
    End If
End Function

Function CheckIfCompletePositionEntry() As Boolean
    If txtPOSITION_POS.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_POS.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If txtPOSITION_FROMAGE.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_FROMAGE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If IsNumeric(txtPOSITION_FROMAGE.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_FROMAGE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If txtPOSITION_TOAGE.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_TOAGE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If IsNumeric(txtPOSITION_TOAGE.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_TOAGE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If txtPOSITION_AVAILABLE.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_AVAILABLE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If IsNumeric(txtPOSITION_AVAILABLE.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_AVAILABLE.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If txtPOSITION_TAKEN.Text = "" Then
        MsgBox "Incomplete Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_TAKEN.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    If IsNumeric(txtPOSITION_TAKEN.Text) = False Then
        MsgBox "Invalid Entry", vbExclamation, "Open Position"
        On Error Resume Next
        txtPOSITION_TAKEN.SetFocus
        CheckIfCompletePositionEntry = False
        Exit Function
    End If
    CheckIfCompletePositionEntry = True
End Function

Function DisplayPOSITION_EDUCATION()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_EDUCATION Where POS_ID = " & CInt(txtPOS_ID) & "")
    lsvEDUC.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvEDUC.ListItems.Add(, , Null2String(RSTMP!DEGREE))
            ITEM.SubItems(1) = Null2String(RSTMP!FIELDS)
            ITEM.SubItems(2) = Null2String(RSTMP!NOTES)
            ITEM.SubItems(3) = RSTMP!Entryid

            RSTMP.MoveNext
        Loop
    End If
End Function

Function DisplayPOSITION_DOCUMENT()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_POSITION_DOCUMENTS Where POS_ID = " & CInt(txtPOS_ID) & "")
    lsvDOC.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvDOC.ListItems.Add(, , Null2String(RSTMP!DocumentType))
            ITEM.SubItems(1) = Null2String(RSTMP!NOTES)
            ITEM.SubItems(2) = RSTMP!Entryid

            RSTMP.MoveNext
        Loop
    End If
End Function

Sub StoreMemVars()
    'BTT - 07032007
    If Not (RS.BOF And RS.EOF) Then
        POSITION_ID = RS!POS_ID
        txtPOS_ID.Text = Null2String(RS!POS_ID)            'Position ID
        txtPOSITION_POS.Text = Null2String(RS!PositionDesc)            'Position Name
        txtPOSITION_AVAILABLE.Text = Null2String(RS!PositionAvailable)            'No of Available
        txtPOSITION_TAKEN.Text = Null2String(RS!PositionTaken)            'No Of Taken
        DTPPOSITION_ASOF.Day = Day(RS!FromDate)            'From Date
        DTPPOSITION_ASOF.MONTH = MONTH(RS!FromDate)
        DTPPOSITION_ASOF.YEAR = YEAR(RS!FromDate)
        DTPPOSITION_DEADLINE.Day = Day(RS!ToDate)          'To Date
        DTPPOSITION_DEADLINE.MONTH = MONTH(RS!ToDate)
        DTPPOSITION_DEADLINE.YEAR = YEAR(RS!ToDate)
        txtPOSITION_FROMAGE.Text = Null2String(RS!fromAge)
        txtPOSITION_TOAGE.Text = Null2String(RS!toAge)
    End If
    Call DisplayPOSITION_EDUCATION
    Call DisplayPOSITION_DOCUMENT
End Sub

Sub rsrefresh()
    Set RS = New ADODB.Recordset
    Call RS.Open("SELECT * FROM HRMS_Position", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Private Sub EnableEntry(COND As Boolean)
    picTAB_DOC.Enabled = COND
    picTAB_EDU.Enabled = COND
    picTAB_MAIN.Enabled = COND
End Sub

Private Sub CleanPositionForm()
    txtPOS_ID.Text = ""
    txtPOSITION_POS.Text = ""
    txtPOSITION_FROMAGE.Text = ""
    txtPOSITION_TOAGE.Text = ""
    txtPOSITION_TAKEN.Text = ""
    txtPOSITION_AVAILABLE.Text = ""

    lsvDOC.ListItems.Clear
    lsvEDUC.ListItems.Clear
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_ADD", "APPLICANT OPEN POSITION") = False Then Exit Sub

    SAVE_OR_EDIT_POSITION = "SAVE"

    Call CleanPositionForm
    Call AlternateButton(False)
    Call EnableEntry(True)

    txtPOS_ID.Text = 0
    txtPOS_ID.Text = GenerateNewPositionID()
    txtPOS_ID.Text = CInt(txtPOS_ID.Text) + 1

    POSITION_ID = CInt(txtPOS_ID)
    Call DeleteTmpFile

    cmdSave.Caption = "&Save"
    txtPOSITION_POS.SetFocus
    lsvPOSITION.Enabled = False


    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub DeleteTmpFile()
    gconDMIS.Execute ("Delete From HRMS_POSITION_EDUCATION Where POS_ID = " & POSITION_ID & "")
    gconDMIS.Execute ("Delete From HRMS_POSITION_DOCUMENTS Where POS_ID = " & POSITION_ID & "")

End Sub

Private Sub AlternateButton(COND As Boolean)
    picAdd.Visible = COND
    picSave.Visible = Not COND
End Sub

Private Sub cmdCancel_Click()
    Call AlternateButton(True)
    Call EnableEntry(False)
    lsvPOSITION.Enabled = True
    lsvPOSITION.Enabled = True
    Call lsvPOSITION_Click

End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "ACESS_DELETE", "APPLICANT OPEN POSITION") = False Then Exit Sub
    On Error GoTo Errorcode:





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_EDIT", "APPLICANT OPEN POSITION") = False Then Exit Sub
    lsvPOSITION.Enabled = False
    Call lsvPOSITION_DblClick



    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MoveFirst
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MoveLast
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdPOSITION_ADD_DOC_Click()
    SAVE_OR_EDIT_PAPERS = "SAVE"

    picSave.Visible = False
    frmAISPOSITION.Enabled = False
    frmAISPOSITION_DOC.Show
    frmAISPOSITION_DOC.cmdDelete.Enabled = False
End Sub

Private Sub cmdPOSITION_ADD_EDU_Click()
    POSITION_SAVE_OR_EDIT_EDU = "SAVE"

    picSave.Visible = False
    frmAISPOSITION.Enabled = False
    frmAISPOSITION_EDUC.Show
    frmAISPOSITION_EDUC.cmdDelete.Enabled = False
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    frmMain.MousePointer = 11

    RS.MovePrevious
    If RS.BOF Then
        RS.MoveNext
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_PRINT", "APPLICANT OPEN POSITION") = False Then Exit Sub
    frmMain.MousePointer = 11

    rptPosition.ReportFileName = AIS_REPORT_PATH & "JobReport.rpt"
    rptPosition.RetrieveDataFiles

    Call PrintSQLReport(rptPosition, AIS_REPORT_PATH & "JobReport.rpt", "{HRMS_POSITION.POS_ID} = " & CLng(txtPOS_ID), AIS_REPORT_Connection, 1)

    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim vtxtID As Integer, vtxtPOSITION_FROMAGE As Integer, vtxtPOSITION_TOAGE As Integer
    Dim vtxtPOSITION_AVAILABLE As Integer, vtxtPOSITION_TAKEN         As Integer
    Dim vdtpPOSITION_FROMDATE As String, vdtpPOSITION_TODATE          As String
    Dim vtxtPosition                                                  As String
    Dim vDateInactive                                                 As String

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    If MsgBox("Save Position", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        If CheckIfCompletePositionEntry = True Then
            vtxtID = CInt(txtPOS_ID)
            vtxtPosition = N2Str2Null(txtPOSITION_POS)
            vtxtPOSITION_FROMAGE = CInt(txtPOSITION_FROMAGE)
            vtxtPOSITION_TOAGE = CInt(txtPOSITION_TOAGE)
            vtxtPOSITION_AVAILABLE = CInt(txtPOSITION_AVAILABLE)
            vtxtPOSITION_TAKEN = CInt(txtPOSITION_TAKEN)
            vdtpPOSITION_FROMDATE = N2Str2Null(DTPPOSITION_ASOF)
            vdtpPOSITION_TODATE = N2Str2Null(DTPPOSITION_DEADLINE)
            vDateInactive = dtpDATEINACTIVE

            Select Case SAVE_OR_EDIT_POSITION
                Case "SAVE":
                    gconDMIS.Execute ("Insert Into HRMS_POSITION Values(" & vtxtID & _
                                      "," & vtxtPosition & _
                                      "," & vdtpPOSITION_FROMDATE & _
                                      "," & vdtpPOSITION_TODATE & _
                                      "," & vtxtPOSITION_AVAILABLE & _
                                      "," & vtxtPOSITION_TAKEN & _
                                      "," & vtxtPOSITION_FROMAGE & _
                                      "," & vtxtPOSITION_TOAGE & _
                                      "," & vDateInactive & ")")

                    Call AlternateButton(True)
                    Call EnableEntry(False)
                    Call DisplayAllPosition
                    lsvPOSITION.Enabled = True
                    tcPOSITION.SelectedItem = 0
                    Call lsvPOSITION_Click
                Case "EDIT":
                    gconDMIS.Execute ("Update HRMS_POSITION Set PositionDesc = " & vtxtPosition & _
                                      ",FromDate = " & vdtpPOSITION_FROMDATE & _
                                      ",ToDate = " & vdtpPOSITION_TODATE & _
                                      ",PositionAvailable = " & vtxtPOSITION_AVAILABLE & _
                                      ",PositionTaken = " & vtxtPOSITION_TAKEN & _
                                      ",FromAge = " & vtxtPOSITION_FROMAGE & _
                                      ",ToAge = " & vtxtPOSITION_TOAGE & _
                                      ",DateInactive = " & vDateInactive & _
                                    " Where POS_ID = " & vtxtID & "")

                    Call AlternateButton(True)
                    Call EnableEntry(False)
                    Call DisplayAllPosition

                    lsvPOSITION.Enabled = True
                    tcPOSITION.SelectedItem = 0
                    Call lsvPOSITION_Click
            End Select
        End If
    End If
    frmMain.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call rsrefresh    'BTT - 07032007
    Call CenterMe(frmMain, Me, 1)
    tcPOSITION.SelectedItem = 0

    frmMain.MousePointer = 11

    Call DisplayAllPosition
    Call EnableEntry(False)

    Call lsvPOSITION_Click
    Call StoreMemVars

    frmMain.MousePointer = 0
End Sub

Private Sub DisplayAllPosition()
    Dim RSTMP                                                         As ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_Position Order By POS_ID")

    lsvPOSITION.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvPOSITION.ListItems.Add(, , Null2String(RSTMP!POS_ID))
            ITEM.SubItems(1) = Null2String(RSTMP!PositionDesc)
            ITEM.SubItems(2) = Null2String(RSTMP!PositionAvailable)
            ITEM.SubItems(3) = Null2String(RSTMP!PositionTaken)

            RSTMP.MoveNext
        Loop
    End If
End Sub

Private Sub lsvDOC_DblClick()
    Dim Index                                                         As Integer

    If Not lsvDOC.ListItems.count = 0 Then
        Index = CInt(lsvDOC.SelectedItem.Index)
        With lsvDOC
            picSave.Visible = False
            SAVE_OR_EDIT_PAPERS = "EDIT"
            POSITION_DOC_ENTRY_ID = CInt(.ListItems(Index).SubItems(2))

            'frmMain.tbMENU.Enabled = False
            frmAISPOSITION.Enabled = False
            frmAISPOSITION_DOC.Show

            frmAISPOSITION_DOC.cboDOC.Text = .ListItems(Index).Text
            frmAISPOSITION_DOC.txtNOTE.Text = .ListItems(Index).SubItems(1)
        End With
    End If
End Sub

Private Sub lsvEDUC_DblClick()
    Dim Index                                                         As Integer

    If Not lsvEDUC.ListItems.count = 0 Then
        Index = CInt(lsvEDUC.SelectedItem.Index)
        With lsvEDUC
            picSave.Visible = False
            POSITION_SAVE_OR_EDIT_EDU = "EDIT"
            POSITION_EDU_ENTRY_ID = CInt(.ListItems(Index).SubItems(3))

            frmAISPOSITION.Enabled = False
            frmAISPOSITION_EDUC.Show

            frmAISPOSITION_EDUC.cboDEGREE.Text = .ListItems(Index).Text
            frmAISPOSITION_EDUC.cboFIELDS.Text = .ListItems(Index).SubItems(1)
            frmAISPOSITION_EDUC.txtNOTE.Text = .ListItems(Index).SubItems(2)
        End With
    End If
End Sub

'Upating Code       : AXP-0707200721:36
Private Sub lsvPOSITION_Click()
    Dim Index                                                         As Integer
    Dim RSTMP                                                         As ADODB.Recordset

    On Error GoTo Errorcode:
    frmMain.MousePointer = 11

    lsvPOSITION.Enabled = False

    If Not lsvPOSITION.ListItems.count = 0 Then
        Index = CInt(lsvPOSITION.SelectedItem.Index)

        With lsvPOSITION
            Set RSTMP = gconDMIS.Execute("Select * From HRMS_Position Where Pos_ID = " & CInt(.ListItems(Index).Text) & "")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                POSITION_ID = RSTMP!POS_ID
                txtPOS_ID.Text = Null2String(RSTMP!POS_ID)    'Position ID
                txtPOSITION_POS.Text = Null2String(RSTMP!PositionDesc)    'Position Name
                txtPOSITION_AVAILABLE.Text = Null2String(RSTMP!PositionAvailable)    'No of Available
                txtPOSITION_TAKEN.Text = Null2String(RSTMP!PositionTaken)    'No Of Taken
                DTPPOSITION_ASOF.Day = Day(RSTMP!FromDate)    'From Date
                DTPPOSITION_ASOF.MONTH = MONTH(RSTMP!FromDate)
                DTPPOSITION_ASOF.YEAR = YEAR(RSTMP!FromDate)
                'DTPPOSITION_DEADLINE.Day = Day(rsTmp!ToDate)  'To Date
                'DTPPOSITION_DEADLINE.Month = Month(rsTmp!ToDate)
                DTPPOSITION_DEADLINE.Value = DateValue(RSTMP!ToDate)
                'DTPPOSITION_DEADLINE.Year = Year(rsTmp!ToDate)
                txtPOSITION_FROMAGE.Text = Null2String(RSTMP!fromAge)
                txtPOSITION_TOAGE.Text = Null2String(RSTMP!toAge)
            End If
        End With
    End If

    Call DisplayPOSITION_EDUCATION
    Call DisplayPOSITION_DOCUMENT

    lsvPOSITION.Enabled = True
    frmMain.MousePointer = 0

    Exit Sub

Errorcode:
    ShowVBError
    frmMain.MousePointer = 0
End Sub

Private Sub lsvPOSITION_DblClick()
    Dim Index                                                         As Integer

    If Not lsvPOSITION.ListItems.count = 0 Then
        SAVE_OR_EDIT_POSITION = "EDIT"
        Index = CInt(lsvPOSITION.SelectedItem.Index)

        With lsvPOSITION
            Call AlternateButton(False)
            Call EnableEntry(True)

            tcPOSITION.SelectedItem = 0
            txtPOSITION_POS.SetFocus
        End With
    End If
End Sub

Private Sub lsvPOSITION_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Call lsvPOSITION_Click
End Sub

