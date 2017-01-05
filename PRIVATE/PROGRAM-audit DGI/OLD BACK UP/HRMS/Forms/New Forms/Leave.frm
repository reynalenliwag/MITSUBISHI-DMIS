VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Leave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Leave Module"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   Icon            =   "Leave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   8490
   Begin VB.PictureBox picLeave 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   3180
      ScaleHeight     =   3345
      ScaleWidth      =   4455
      TabIndex        =   1
      Top             =   1230
      Width           =   4485
      Begin VB.TextBox txtNoOfDays 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3180
         TabIndex        =   21
         Top             =   1110
         Width           =   1185
      End
      Begin VB.TextBox txtNotes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   1980
         Width           =   4185
      End
      Begin VB.ComboBox cboLeaveType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Leave.frx":058A
         Left            =   840
         List            =   "Leave.frx":059D
         TabIndex        =   20
         Text            =   "cboLeaveType"
         Top             =   1140
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   840
         TabIndex        =   17
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53084161
         CurrentDate     =   39456
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   3180
         TabIndex        =   18
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53084161
         CurrentDate     =   39456
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   315
         Left            =   840
         TabIndex        =   19
         Top             =   750
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53084161
         CurrentDate     =   39456
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   43
         Top             =   -30
         Width           =   4455
         _Version        =   655364
         _ExtentX        =   7858
         _ExtentY        =   556
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
         VisualTheme     =   0
         GradientColorLight=   16777215
         GradientColorDark=   14737632
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   42
         Top             =   1770
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Filed"
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
         Left            =   30
         TabIndex        =   41
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type of   Leave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   150
         TabIndex        =   25
         Top             =   1140
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   2850
         TabIndex        =   24
         Top             =   420
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   360
         TabIndex        =   23
         Top             =   450
         Width           =   360
      End
      Begin VB.Label labID 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000006&
         Height          =   495
         Left            =   2850
         TabIndex        =   3
         Top             =   90
         Width           =   285
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "No. of days"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   3210
         TabIndex        =   2
         Top             =   870
         Width           =   915
      End
   End
   Begin VB.TextBox txtDateAsOf 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   6330
      TabIndex        =   39
      Top             =   1770
      Width           =   1335
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5235
      Left            =   30
      Picture         =   "Leave.frx":05B0
      ScaleHeight     =   5205
      ScaleWidth      =   2445
      TabIndex        =   14
      Top             =   120
      Width           =   2475
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   30
         MaxLength       =   35
         TabIndex        =   15
         Top             =   60
         Width           =   2385
      End
      Begin MSComctlLib.ListView lsLeave 
         Height          =   4665
         Left            =   0
         TabIndex        =   16
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   8229
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
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Leave.frx":32EC
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "firstname"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "lastname"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "position"
            Object.Width           =   2
         EndProperty
         Picture         =   "Leave.frx":344E
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1725
      Left            =   2550
      ScaleHeight     =   1725
      ScaleWidth      =   5865
      TabIndex        =   4
      Top             =   30
      Width           =   5865
      Begin VB.TextBox txtMaxSL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3780
         TabIndex        =   8
         Top             =   1290
         Width           =   1335
      End
      Begin VB.TextBox txtMaxVL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3780
         TabIndex        =   7
         Top             =   900
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   90
         Width           =   5685
      End
      Begin VB.TextBox txtPosition 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1050
         TabIndex        =   5
         Top             =   480
         Width           =   4695
      End
      Begin Crystal.CrystalReport rptDeductions 
         Left            =   5280
         Top             =   1140
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
      Begin VB.Label lblEmpNo 
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   1260
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum No. of Sick Leave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   1290
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum No. of Vacation Leave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   570
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   5235
      Left            =   30
      Picture         =   "Leave.frx":171BB
      ScaleHeight     =   5175
      ScaleWidth      =   2415
      TabIndex        =   0
      Top             =   120
      Width           =   2475
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   2640
      ScaleHeight     =   3105
      ScaleWidth      =   5865
      TabIndex        =   12
      Top             =   2280
      Width           =   5865
      Begin MSFlexGridLib.MSFlexGrid grdLeave 
         Height          =   3075
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5424
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   -2147483633
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         MousePointer    =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2820
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   26
      Top             =   5430
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "Leave.frx":2AF18
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2B06A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   4170
         MouseIcon       =   "Leave.frx":2B3D0
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2B522
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3480
         MouseIcon       =   "Leave.frx":2B888
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2B9DA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Delete Selected Record"
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
         Left            =   2790
         MouseIcon       =   "Leave.frx":2BD05
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2BE57
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   2100
         MouseIcon       =   "Leave.frx":2C1B3
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2C305
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Left            =   1410
         MouseIcon       =   "Leave.frx":2C618
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2C76A
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   720
         MouseIcon       =   "Leave.frx":2CA64
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2CBB6
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   30
         MouseIcon       =   "Leave.frx":2CF0E
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2D060
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6960
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   35
      Top             =   5460
      Width           =   1440
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
         MouseIcon       =   "Leave.frx":2D3BF
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2D511
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   30
         MouseIcon       =   "Leave.frx":2D84F
         MousePointer    =   99  'Custom
         Picture         =   "Leave.frx":2D9A1
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "As Of Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2700
      TabIndex        =   40
      Top             =   1770
      Width           =   2775
   End
End
Attribute VB_Name = "frmHRMS_Leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLeave                                                           As ADODB.Recordset
Dim rsLeaveDet                                                        As ADODB.Recordset
Dim rsEmpInfoTable                                                    As ADODB.Recordset

Dim ADDOREDIT                                                         As String

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Sub rsrefresh1(XXX As String)
    Set rsEmpInfoTable = New ADODB.Recordset
    If XXX <> "" Then
        rsEmpInfoTable.Open "select EmpNo, LastName, FirstName, Position from HRMS_EmpInfo where EmpNo like '" & XXX & "' order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsEmpInfoTable.Open "select EmpNo, LastName, FirstName, Position from HRMS_EmpInfo order by lastname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        '        Set rsEmpInfoTable = gconDMIS.Execute("select EmpNo, LastName, FirstName, Position from HRMS_EmpInfo")
    End If

End Sub

Sub rsrefresh2()
    Set rsLeave = New ADODB.Recordset
    rsLeave.Open "select * from HRMS_Leave where EmplNo = '" & lblEmpNo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub rsRefresh3()
    Set rsLeaveDet = New ADODB.Recordset
    rsLeaveDet.Open "select * from HRMS_LeaveDet where EmplNo = '" & lblEmpNo & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitMemvars()
    txtMaxSL = ""
    txtMaxVL = ""
    txtDateAsOf = ""
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    cboLeaveType = ""
    txtNoOfDays = ""
End Sub

Sub initMemvars2()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
    DTPicker3.Value = Date
    cboLeaveType = ""
    txtNoOfDays = ""
End Sub

Sub FillGrid()
    grdLeave.Rows = 1
    If Not rsLeaveDet.BOF And Not rsLeaveDet.EOF Then
        rsLeaveDet.MoveFirst
        Do While Not rsLeaveDet.EOF
            grdLeave.AddItem Null2String(rsLeaveDet!DateFrom) & Chr(9) & _
                             Null2String(rsLeaveDet!DateTo) & Chr(9) & _
                             Null2String(rsLeaveDet!LEAVETYPE) & Chr(9) & _
                             Null2String(rsLeaveDet!DAYS_NO) & Chr(9) & _
                             Null2String(rsLeaveDet!DateFiled) & Chr(9) & _
                             Null2String(rsLeaveDet!ID)
            rsLeaveDet.MoveNext
        Loop
    End If
End Sub

Sub FillLstLeave()
    Listview_Loadval lsLeave.ListItems, rsEmpInfoTable
    rsrefresh1 ""
End Sub

Sub InitGrid()
    With grdLeave
        .Rows = 2
        .ColWidth(0) = 1200: .ColWidth(1) = 1200: .ColWidth(2) = 1300: .ColWidth(3) = 1000: .ColWidth(4) = 1000
        .Row = 0
        .Col = 0: .Text = "From Date"
        .Col = 1: .Text = "To Date"
        .Col = 2: .Text = "Type"
        .Col = 3: .Text = "Total Days"
        .Col = 4: .Text = "Date Filed"
        .Col = 5: .Text = "ID"
    End With
End Sub

Sub storeMemvars1()
    If Not rsEmpInfoTable.BOF And Not rsEmpInfoTable.EOF Then
        lblEmpNo = Null2String(rsEmpInfoTable!EMPNO)
        txtName = Null2String(rsEmpInfoTable!lastname) + ", " + Null2String(rsEmpInfoTable![FIRSTNAME])
        txtPosition = Null2String(rsEmpInfoTable!Position)
    End If
End Sub

Sub storeMemvars2()
    If Not rsLeave.BOF And Not rsLeave.EOF Then
        txtMaxVL = Null2String(rsLeave!MAXVL)
        txtMaxSL = Null2String(rsLeave!MAXSL)
        txtDateAsOf = Null2String(rsLeave!DateAsOf)
    Else
        txtMaxVL = ""
        txtMaxSL = ""
        txtDateAsOf = ""
    End If
End Sub

Sub storeMemvars3()
    If Not rsLeaveDet.BOF And Not rsLeaveDet.EOF Then
        DTPicker1.Value = Null2String(rsLeaveDet!DateFrom)
        DTPicker2.Value = Null2String(rsLeaveDet!DateTo)

        txtNoOfDays = Null2String(rsLeaveDet!DAYS_NO)
        LABID = Null2String(rsLeaveDet!ID)
        DTPicker3.Value = Null2String(rsLeaveDet!DateFiled)
        '        txtNotes = Null2String(rsLeaveDet!NOTES)
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "FILES LEAVE CODES") = False Then Exit Sub
    ADDOREDIT = "ADD"

    grdLeave.Enabled = False
    initMemvars2
    picLeave.Visible = True
    Picture1.Visible = False
End Sub

Private Sub cmdCancel_Click()
    grdLeave.Enabled = True
    picLeave.Visible = False
    Picture1.Visible = True
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "FILES LEAVE CODES") = False Then Exit Sub
    If LABID <> "" Then
        If MsgBox("Delete This Leave Type", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            gconDMIS.Execute "delete from HRMS_LeaveDet where ID = " & LABID
        End If
    Else
        'do nothing
    End If
    rsRefresh3
    storeMemvars3
    FillGrid
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "FILES LEAVE CODES") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    picLeave.Visible = True
    Picture1.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch = ""
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfoTable.MoveNext
    If rsEmpInfoTable.EOF Then
        rsEmpInfoTable.MoveLast
        ShowLastRecordMsg
    End If
    storeMemvars1
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfoTable.MovePrevious
    If rsEmpInfoTable.BOF Then
        rsEmpInfoTable.MoveFirst
        ShowFirstRecordMsg
    End If
    storeMemvars1
End Sub

Private Sub cmdSave_Click()
    Dim vtxtEmplNo                                                    As String
    Dim vcboLeaveType                                                 As String
    Dim vtxtMaxSL                                                     As String
    Dim vtxtMaxVL                                                     As String
    Dim vtxtDateAsof                                                  As String
    Dim vtxtNoOfDays                                                  As Double
    Dim vDateFrom                                                     As String
    Dim vDateTo                                                       As String
    Dim vDateFiled                                                    As String
    Dim vNotes                                                        As String

    vtxtEmplNo = N2Str2Null(lblEmpNo)
    vcboLeaveType = N2Str2Null(cboLeaveType)
    vtxtMaxSL = N2Str2Null(txtMaxSL)
    vtxtMaxVL = N2Str2Null(txtMaxVL)
    vtxtDateAsof = N2Str2Null(txtDateAsOf)

    vtxtNoOfDays = N2Str2Zero(txtNoOfDays.Text)

    vDateFrom = N2Str2Null(DTPicker1.Value)
    vDateTo = N2Str2Null(DTPicker2.Value)
    vDateFiled = N2Str2Null(DTPicker3.Value)
    vNotes = N2Str2Null(txtNotes)

    'check if there is leave type chosen
    If cboLeaveType = "" Then
        MsgBox "Please choose type of Leave.........", vbOKOnly, "Type Of Leave"
        Exit Sub
    End If
    'replace blank value with 0
    If txtNoOfDays = "" Then
        txtNoOfDays = 0
    End If
    'check if DateFrom is greater than DateTo
    If txtNoOfDays <= 0 Then
        MsgBox "Please Check the Entered dates.........", vbOKOnly, "Wrong Dates"
        Exit Sub
    End If

    'if ADD
    If ADDOREDIT = "ADD" Then
        If rsLeave.RecordCount = 0 Then
            'if not existing yet in HRMS_Leave insert the record
            gconDMIS.Execute "Insert into HRMS_Leave (EmplNo, MaxVL, MaxSL, DateAsOf)" & _
                             "values (" & vtxtEmplNo & "," & vtxtMaxVL & "," & vtxtMaxSL & "," & vtxtDateAsof & ")"
        Else
            'if existing already in HRMS_Leave update the record
            gconDMIS.Execute "Update HRMS_Leave set" & _
                           " MaxVL = " & vtxtMaxVL & "," & _
                           " MaxSL = " & vtxtMaxSL & "," & _
                           " DateAsOf = " & vtxtDateAsof & _
                           " where EmplNo = " & vtxtEmplNo
        End If
        'insert the record in HRMS_LeaveDet
        gconDMIS.Execute "Insert into HRMS_LeaveDet (EmplNo, LeaveType, DateFrom, DateTo, Days_No, DateFiled , notes)" & _
                         "values (" & vtxtEmplNo & "," & vcboLeaveType & "," & vDateFrom & "," & vDateTo & "," & vtxtNoOfDays & "," & vDateFiled & "," & vNotes & ")"
        'if EDIT
    Else
        If LABID = "" Then
            MsgBox "Can't Save...No item selected"
            Exit Sub
        End If

        'update HRMS_Leave
        gconDMIS.Execute "Update HRMS_Leave set" & _
                       " MaxVL = " & vtxtMaxVL & "," & _
                       " MaxSL = " & vtxtMaxSL & "," & _
                       " DateAsOf = " & vtxtDateAsof & _
                       " where EmplNo = " & vtxtEmplNo
        'update HRMS_LeaveDet
        'NOT VALID FIELD TYPE NOTES

        '            gconDMIS.Execute "Update HRMS_LeaveDet set" & _
                     '                " LeaveType = " & vcboLeaveType & "," & _
                     '                " DateFrom = " & vDateFrom & "," & _
                     '                " DateTo = " & vDateTo & "," & _
                     '                " Days_No = " & vtxtNoOfDays & "," & _
                     '                " DateFiled = " & vDateFiled & "," & _
                     '                " NOtes = " & vNotes & _
                     '                " where ID = " & labID
        '
        gconDMIS.Execute "Update HRMS_LeaveDet set" & _
                       " LeaveType = " & vcboLeaveType & "," & _
                       " DateFrom = " & vDateFrom & "," & _
                       " DateTo = " & vDateTo & "," & _
                       " Days_No = " & vtxtNoOfDays & "," & _
                       " DateFiled = " & vDateFiled & _
                       " where ID = " & LABID
    End If
    rsrefresh2
    storeMemvars2
    rsRefresh3
    FillGrid
    cmdCancel.Value = True
End Sub

Private Sub DTPicker1_Change()
    txtNoOfDays = NumericVal(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
End Sub

Private Sub DTPicker2_Change()
    txtNoOfDays = NumericVal(DateDiff("d", DTPicker1.Value, DTPicker2.Value)) + 1
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    picLeave.Visible = False

    InitMemvars
    InitGrid
    rsrefresh1 ""

    storeMemvars1
    FillLstLeave
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsLeave = Nothing
    Set rsLeaveDet = Nothing
    Set rsEmpInfoTable = Nothing
End Sub

Private Sub grdLeave_Click()
    If grdLeave.Rows > 1 Then
        On Error Resume Next
        DTPicker1.Value = grdLeave.TextMatrix(grdLeave.RowSel, 0)

        On Error Resume Next
        DTPicker2.Value = grdLeave.TextMatrix(grdLeave.RowSel, 1)

        On Error Resume Next
        cboLeaveType = grdLeave.TextMatrix(grdLeave.RowSel, 2)

        On Error Resume Next
        txtNoOfDays = grdLeave.TextMatrix(grdLeave.RowSel, 3)

        On Error Resume Next
        DTPicker3.Value = grdLeave.TextMatrix(grdLeave.RowSel, 4)

        On Error Resume Next
        LABID = grdLeave.TextMatrix(grdLeave.RowSel, 5)
    End If
End Sub

Private Sub lsLeave_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    lblEmpNo = ITEM.Text
    rsrefresh1 lblEmpNo

    storeMemvars1
    rsrefresh2
    storeMemvars2
    rsRefresh3
    storeMemvars3
    FillGrid
End Sub

Private Sub txtNoOfDays_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "tab"
    Else
        KeyAscii = LimitChar("015.", KeyAscii)
    End If
End Sub

Private Sub txtsearch_Change()
    Listview_Loadval lsLeave.ListItems, gconDMIS.Execute("select EmpNo from hrms_empinfo where EmpNo like '%" & Repleys(txtSearch) & "%'")
End Sub

