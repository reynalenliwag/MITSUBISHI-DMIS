VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPASformA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Performance Appraisal"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPASformA.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10605
   Begin VB.PictureBox picCHOOSE 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   5850
      ScaleHeight     =   3285
      ScaleWidth      =   3570
      TabIndex        =   12
      Top             =   390
      Visible         =   0   'False
      Width           =   3630
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3225
         Left            =   30
         ScaleHeight     =   3195
         ScaleWidth      =   3480
         TabIndex        =   13
         Top             =   30
         Width           =   3510
         Begin VB.CommandButton cmdChooseExit 
            BackColor       =   &H000000FF&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3180
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Search"
            Top             =   0
            Width           =   315
         End
         Begin VB.OptionButton optFNAME 
            BackColor       =   &H00FFFFFF&
            Caption         =   "FIRSTNAME"
            Height          =   240
            Left            =   1800
            TabIndex        =   23
            Top             =   810
            Width           =   1425
         End
         Begin VB.OptionButton optLNAME 
            BackColor       =   &H00FFFFFF&
            Caption         =   "LASTNAME"
            Height          =   240
            Left            =   120
            TabIndex        =   22
            Top             =   810
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.TextBox txtCHOOSE 
            Height          =   375
            Left            =   90
            TabIndex        =   15
            Top             =   360
            Width           =   3315
         End
         Begin MSComctlLib.ListView lsvCHOOSE 
            Height          =   1725
            Left            =   90
            TabIndex        =   14
            Top             =   1140
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   3043
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Employee Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double click to Choose"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   24
            Top             =   2940
            Width           =   1890
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CHOOSE EMPLOYEE"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   21
            Top             =   60
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox picADD_EDIT 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   3240
      ScaleHeight     =   1365
      ScaleWidth      =   6675
      TabIndex        =   37
      Top             =   5790
      Width           =   6675
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
         Height          =   705
         Left            =   5250
         MouseIcon       =   "frmPASformA.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmPASformA.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Print this Record"
         Top             =   0
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
         Height          =   705
         Left            =   5940
         MouseIcon       =   "frmPASformA.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmPASformA.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Exit Window"
         Top             =   0
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
         Height          =   705
         Left            =   4560
         MouseIcon       =   "frmPASformA.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmPASformA.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdadd 
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
         Height          =   705
         Left            =   3870
         MouseIcon       =   "frmPASformA.frx":13A8
         MousePointer    =   99  'Custom
         Picture         =   "frmPASformA.frx":14FA
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add"
         Height          =   795
         Left            =   -2400
         MouseIcon       =   "frmPASformA.frx":180D
         MousePointer    =   99  'Custom
         Picture         =   "frmPASformA.frx":195F
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2070
      Top             =   5880
   End
   Begin VB.PictureBox picMENU 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   330
      ScaleHeight     =   1245
      ScaleWidth      =   4245
      TabIndex        =   16
      Top             =   7260
      Width           =   4305
      Begin VB.PictureBox picSAVE 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2610
         ScaleHeight     =   735
         ScaleWidth      =   2505
         TabIndex        =   17
         Top             =   270
         Width           =   2565
         Begin VB.CommandButton cmdDelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   840
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "SAVE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1680
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   825
         End
      End
   End
   Begin VB.Frame fmeFIELDS 
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2685
      Left            =   90
      TabIndex        =   10
      Top             =   2940
      Width           =   9945
      Begin VB.CommandButton cmdAddDetails 
         Caption         =   "ADD DETAILS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2910
         TabIndex        =   29
         Top             =   1950
         Width           =   2055
      End
      Begin VB.CommandButton cmdAddCate 
         Caption         =   "ADD CATEGORY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   810
         TabIndex        =   28
         Top             =   1950
         Width           =   2055
      End
      Begin MSComctlLib.ListView lsvDetails 
         Height          =   1605
         Left            =   2880
         TabIndex        =   11
         Top             =   300
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   2831
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
            Text            =   "Process/Objectives"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Standard"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Monitor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Score"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lsvCATEGORY 
         Height          =   1605
         Left            =   90
         TabIndex        =   27
         Top             =   300
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2831
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CATEGORY"
            Object.Width           =   4674
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCATEID 
         BackColor       =   &H000000FF&
         Height          =   195
         Left            =   2130
         TabIndex        =   72
         Top             =   2340
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8010
         TabIndex        =   31
         Top             =   1980
         Width           =   1545
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SUBTOTAL:"
         Height          =   195
         Index           =   8
         Left            =   6810
         TabIndex        =   30
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Double Click to Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   26
         Top             =   2340
         Width           =   1590
      End
   End
   Begin VB.Frame fmePERSON 
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   9945
      Begin VB.CommandButton cmdAddRes 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7230
         TabIndex        =   35
         ToolTipText     =   "Add / Edit / Delete"
         Top             =   690
         Width           =   375
      End
      Begin VB.ComboBox cboMajorRes 
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
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   720
         Width           =   4785
      End
      Begin VB.CommandButton cmdChoose 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9390
         TabIndex        =   9
         ToolTipText     =   "Choose Employee"
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox txtJSummary 
         Height          =   1155
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1560
         Width           =   7455
      End
      Begin VB.TextBox txtPerCover 
         Height          =   345
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1140
         Width           =   4785
      End
      Begin VB.Label lblEmpName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   330
         Width           =   5925
      End
      Begin VB.Label lblEMPNO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   2340
         TabIndex        =   7
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "JOB SUMMARY:"
         Height          =   195
         Index           =   4
         Left            =   930
         TabIndex        =   4
         Top             =   1590
         Width           =   1320
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIOD COVERED:"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   3
         Top             =   1230
         Width           =   1665
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAJOR RESPONSIBILITY:"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   2
         Top             =   840
         Width           =   2190
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RATEE:"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   1
         Top             =   450
         Width           =   630
      End
   End
   Begin VB.PictureBox picCAT 
      BackColor       =   &H00C0C0FF&
      Height          =   3675
      Left            =   2940
      ScaleHeight     =   3615
      ScaleWidth      =   5655
      TabIndex        =   57
      Top             =   1530
      Visible         =   0   'False
      Width           =   5715
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3555
         Left            =   30
         ScaleHeight     =   3525
         ScaleWidth      =   5535
         TabIndex        =   58
         Top             =   30
         Width           =   5565
         Begin VB.CommandButton cmdCATCancel 
            Caption         =   "CANCEL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4560
            TabIndex        =   71
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton cmdCATDelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2880
            TabIndex        =   70
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton cmdCATSave 
            Caption         =   "SAVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3720
            TabIndex        =   69
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton cmdCATEDIT 
            Caption         =   "EDIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2040
            TabIndex        =   68
            Top             =   1230
            Width           =   825
         End
         Begin VB.CommandButton cmdCATADD 
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
            Height          =   435
            Left            =   1200
            TabIndex        =   67
            Top             =   1230
            Width           =   825
         End
         Begin VB.Frame fmeCAT 
            BackColor       =   &H00FFFFFF&
            Height          =   795
            Left            =   60
            TabIndex        =   62
            Top             =   330
            Width           =   5355
            Begin VB.TextBox txtCAT 
               Height          =   345
               Left            =   1470
               MaxLength       =   50
               TabIndex        =   64
               Top             =   300
               Width           =   3705
            End
            Begin VB.Label lblCATID 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   720
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblADD_EDIT_CAT 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1200
               TabIndex        =   65
               Top             =   720
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblCAP 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "CATEGORY/ RESULT AREAS"
               Height          =   525
               Index           =   13
               Left            =   60
               TabIndex        =   63
               Top             =   270
               Width           =   1350
            End
         End
         Begin MSComctlLib.ListView lsvCAT 
            Height          =   1365
            Left            =   90
            TabIndex        =   60
            Top             =   1860
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   2408
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
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Category/Result Areas"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   0
            EndProperty
         End
         Begin VB.CommandButton cmdCatExit 
            BackColor       =   &H000000FF&
            Caption         =   "x"
            Height          =   285
            Left            =   5250
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Search"
            Top             =   0
            Width           =   285
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click to Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   14
            Left            =   90
            TabIndex        =   61
            Top             =   3240
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox picAddDetails 
      BackColor       =   &H00C0FFFF&
      Height          =   4575
      Left            =   2970
      ScaleHeight     =   4515
      ScaleWidth      =   6045
      TabIndex        =   32
      Top             =   300
      Visible         =   0   'False
      Width           =   6105
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4455
         Left            =   30
         ScaleHeight     =   4425
         ScaleWidth      =   5955
         TabIndex        =   33
         Top             =   30
         Width           =   5985
         Begin VB.CommandButton cmdDetCancel 
            Caption         =   "CANCEL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4980
            TabIndex        =   89
            Top             =   2310
            Width           =   825
         End
         Begin VB.CommandButton cmdDetDelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3300
            TabIndex        =   88
            Top             =   2310
            Width           =   825
         End
         Begin VB.CommandButton cmdDetEdit 
            Caption         =   "EDIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2460
            TabIndex        =   86
            Top             =   2310
            Width           =   825
         End
         Begin VB.CommandButton cmdDetAdd 
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
            Height          =   435
            Left            =   1620
            TabIndex        =   85
            Top             =   2310
            Width           =   825
         End
         Begin VB.CommandButton cmdExitDetail 
            BackColor       =   &H000000FF&
            Caption         =   "x"
            Height          =   285
            Left            =   5670
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Search"
            Top             =   0
            Width           =   285
         End
         Begin VB.Frame fmeDet 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   1965
            Left            =   90
            TabIndex        =   73
            Top             =   270
            Width           =   5715
            Begin VB.TextBox txtDetScore 
               Height          =   315
               Left            =   4020
               MaxLength       =   10
               TabIndex        =   84
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox txtDetProc 
               Height          =   645
               Left            =   1500
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   76
               Top             =   300
               Width           =   4095
            End
            Begin VB.TextBox txtdetMon 
               Height          =   315
               Left            =   1500
               MaxLength       =   50
               TabIndex        =   75
               Top             =   1440
               Width           =   1545
            End
            Begin VB.TextBox txtDetStd 
               Height          =   345
               Left            =   1500
               MaxLength       =   100
               TabIndex        =   74
               Top             =   1020
               Width           =   4065
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SCORE"
               Height          =   195
               Index           =   18
               Left            =   3240
               TabIndex        =   83
               Top             =   1500
               Width           =   615
            End
            Begin VB.Label lblCAP 
               BackStyle       =   0  'Transparent
               Caption         =   "PROCESS/ OBJECTIVES"
               Height          =   405
               Index           =   17
               Left            =   270
               TabIndex        =   81
               Top             =   270
               Width           =   1170
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "STANDARD"
               Height          =   195
               Index           =   16
               Left            =   360
               TabIndex        =   80
               Top             =   1110
               Width           =   975
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MONITOR"
               Height          =   195
               Index           =   15
               Left            =   570
               TabIndex        =   79
               Top             =   1560
               Width           =   825
            End
            Begin VB.Label lblAdd_EDIT_DET 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   150
               TabIndex        =   78
               Top             =   1860
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblDETID 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1440
               TabIndex        =   77
               Top             =   1860
               Visible         =   0   'False
               Width           =   945
            End
         End
         Begin MSComctlLib.ListView lsvDET 
            Height          =   1395
            Left            =   90
            TabIndex        =   90
            Top             =   2820
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   2461
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
               Text            =   "Process/Objectives"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Standard"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Monitor"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Score"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Object.Width           =   0
            EndProperty
         End
         Begin VB.CommandButton cmdDetSave 
            Caption         =   "SAVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4140
            TabIndex        =   87
            Top             =   2310
            Width           =   825
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click to Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   19
            Left            =   90
            TabIndex        =   91
            Top             =   4230
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox picMAS 
      BackColor       =   &H00C0FFC0&
      Height          =   4905
      Left            =   750
      ScaleHeight     =   4845
      ScaleWidth      =   6525
      TabIndex        =   38
      Top             =   720
      Visible         =   0   'False
      Width           =   6585
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4755
         Left            =   30
         ScaleHeight     =   4725
         ScaleWidth      =   6435
         TabIndex        =   39
         Top             =   30
         Width           =   6465
         Begin VB.Frame fmeMAS 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   2115
            Left            =   90
            TabIndex        =   48
            Top             =   300
            Width           =   6285
            Begin VB.TextBox txtMASperiod 
               Height          =   345
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   50
               Top             =   600
               Width           =   3735
            End
            Begin VB.TextBox txtMASJOB 
               Height          =   1005
               Left            =   2400
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   49
               Top             =   990
               Width           =   3735
            End
            Begin VB.TextBox txtMASMajor 
               Height          =   345
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   51
               Top             =   180
               Width           =   3735
            End
            Begin VB.Label lblADDEDIT 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   1260
               TabIndex        =   56
               Top             =   1590
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblMASID 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BorderStyle     =   1  'Fixed Single
               Height          =   315
               Left            =   180
               TabIndex        =   55
               Top             =   1590
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "JOB SUMMARY:"
               Height          =   195
               Index           =   9
               Left            =   960
               TabIndex        =   54
               Top             =   1020
               Width           =   1320
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PERIOD COVERED:"
               Height          =   195
               Index           =   10
               Left            =   630
               TabIndex        =   53
               Top             =   660
               Width           =   1665
            End
            Begin VB.Label lblCAP 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MAJOR RESPONSIBILITY:"
               Height          =   195
               Index           =   11
               Left            =   90
               TabIndex        =   52
               Top             =   270
               Width           =   2190
            End
         End
         Begin VB.CommandButton cmdMasAdd 
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
            Height          =   435
            Left            =   1980
            TabIndex        =   47
            Top             =   2520
            Width           =   825
         End
         Begin VB.CommandButton cmdMasEdit 
            Caption         =   "EDIT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2820
            TabIndex        =   46
            Top             =   2520
            Width           =   825
         End
         Begin VB.CommandButton cmdMasSave 
            Caption         =   "SAVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4500
            TabIndex        =   45
            Top             =   2520
            Width           =   825
         End
         Begin VB.CommandButton cmdMasDelete 
            Caption         =   "DELETE"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   3660
            TabIndex        =   44
            Top             =   2520
            Width           =   825
         End
         Begin VB.CommandButton cmdMasCancel 
            Caption         =   "CANCEL"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5340
            TabIndex        =   43
            Top             =   2520
            Width           =   825
         End
         Begin VB.CommandButton cmdMasExit 
            BackColor       =   &H000000FF&
            Caption         =   "x"
            Height          =   285
            Left            =   6150
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Search"
            Top             =   0
            Width           =   285
         End
         Begin MSComctlLib.ListView lsvMAS 
            Height          =   1305
            Left            =   60
            TabIndex        =   41
            Top             =   3090
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   2302
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
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Major Responsibility"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Period Covered"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Summary"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label lblCAP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "* Double Click to Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   12
            Left            =   60
            TabIndex        =   42
            Top             =   4470
            Width           =   1590
         End
      End
   End
   Begin VB.Label lblCAP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORM A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   510
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   5940
      Width           =   1500
   End
End
Attribute VB_Name = "frmPASformA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DisabledFrames(COND As Boolean)
    fmeFIELDS.Enabled = COND
    fmePERSON.Enabled = COND
    picADD_EDIT.Enabled = COND
End Sub

Sub DisplayDetails()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ID                                                            As String
    lblCATEID.Caption = ""

    ID = Right(Trim(cboMajorRes), 2)
    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_MASTER Where Empno = '" & lblEMPNO.Caption & _
                                 "' And ID = " & ID & "")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        txtPerCover.Text = Null2String(rsTmp!PerCov)
        txtJSummary.Text = Null2String(rsTmp!JOBSUMM)

        Call DisplayCategoryList
    End If

    Set rsTmp = Nothing
End Sub

Sub DisplayCategoryList()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim rsDET                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim MASID                                                         As String

    lblTotal.Caption = "0"
    MASID = Right(Trim(cboMajorRes.Text), 2)
    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_CATEGORY Where EmpNO = '" & lblEMPNO.Caption & _
                                 "' And ID = '" & MASID & "' Order By CatID ASC")
    lsvDetails.ListItems.Clear
    lsvCATEGORY.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvCATEGORY.ListItems.Add(, , Null2String(rsTmp!CATEGORY))
            ITEM.SubItems(1) = Null2String(rsTmp!CatID)

            Set rsDET = gconDMIS.Execute("Select Score From PAS_FORMA_DETAILS Where Empno = '" & _
                                         lblEMPNO.Caption & "' And Catid = '" & rsTmp!CatID & "'")
            If Not (rsDET.BOF And rsDET.EOF) Then
                Do While Not rsDET.EOF
                    lblTotal.Caption = Val(lblTotal.Caption) + Val(rsDET!Score)

                    rsDET.MoveNext
                Loop
            End If

            rsTmp.MoveNext
        Loop
    End If
    Set rsTmp = Nothing
End Sub

Sub ListViewMas()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_MASTER Where EMPNO = '" & lblEMPNO.Caption & "' Order By ID")
    lsvMAS.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvMAS.ListItems.Add(, , Null2String(rsTmp!MajorRespon))
            ITEM.SubItems(1) = Null2String(rsTmp!PerCov)
            ITEM.SubItems(2) = Null2String(rsTmp!JOBSUMM)
            ITEM.SubItems(3) = Null2String(rsTmp!ID)

            rsTmp.MoveNext
        Loop
    End If

    Set rsTmp = Nothing
End Sub

Sub ListViewCategory()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim MASID                                                         As String

    MASID = Right(Trim(cboMajorRes.Text), 2)
    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_CATEGORY Where EMPNO = '" & lblEMPNO.Caption & _
                                 "' And ID = '" & MASID & "' Order By CatID")
    lsvCAT.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvCAT.ListItems.Add(, , Null2String(rsTmp!CATEGORY))
            ITEM.SubItems(1) = Null2String(rsTmp!CatID)

            rsTmp.MoveNext
        Loop
    End If

    Set rsTmp = Nothing
End Sub

Sub ListViewDet()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem
    Dim MASID                                                         As String

    'MASID = Right(Trim(cboMajorRes.Text), 2)
    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_DETAILS Where EMPNO = '" & lblEMPNO.Caption & _
                                 "' And CatID = '" & lblCATEID.Caption & "' Order By Details_ID")
    lsvDET.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvDET.ListItems.Add(, , Null2String(rsTmp!PROCESS))
            ITEM.SubItems(1) = Null2String(rsTmp!STANDARD)
            ITEM.SubItems(2) = Null2String(rsTmp!MONITOR)
            ITEM.SubItems(3) = Null2String(rsTmp!Score)
            ITEM.SubItems(4) = Null2String(rsTmp!details_ID)

            rsTmp.MoveNext
        Loop
    End If

    Set rsTmp = Nothing
End Sub

Sub CleanDetText()
    txtDetProc.Text = ""
    txtdetMon.Text = ""
    txtdetMon.Text = ""
    txtDetStd.Text = ""
    txtDetScore.Text = ""
End Sub

Sub CleanMasText()
    lblMASID.Caption = ""
    txtMASMajor.Text = ""
    txtMASperiod.Text = ""
    txtMASJOB.Text = ""
End Sub

Sub DisplayListDetails()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set rsTmp = gconDMIS.Execute("Select * From PAS_FORMA_DETAILS Where EmpNO = '" & _
                                 lblEMPNO.Caption & "' And CatID = '" & lblCATEID.Caption & "'")
    lsvDetails.ListItems.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvDetails.ListItems.Add(, , Null2String(rsTmp!PROCESS))
            ITEM.SubItems(1) = Null2String(rsTmp!STANDARD)
            ITEM.SubItems(2) = Null2String(rsTmp!MONITOR)
            ITEM.SubItems(3) = Null2String(rsTmp!Score)
            ITEM.SubItems(4) = Null2String(rsTmp!details_ID)

            rsTmp.MoveNext
        Loop
    End If

    Set rsTmp = Nothing
End Sub

Sub FillcboMajor()
    Dim rsTmp                                                         As New ADODB.Recordset

    Set rsTmp = gconDMIS.Execute("Select MajorRespon,ID From PAS_FORMA_MASTER Where EmpNo = '" & lblEMPNO.Caption & "' Order BY ID ASC")
    cboMajorRes.Clear
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            If Len(Trim(rsTmp!ID)) = 2 Then ZEROS = ""
            If Len(Trim(rsTmp!ID)) = 1 Then ZEROS = "0"
            cboMajorRes.AddItem Null2String(rsTmp!MajorRespon) & " - " & ZEROS & rsTmp!ID

            rsTmp.MoveNext
        Loop
        cboMajorRes.ListIndex = o
    End If
    Set rsTmp = Nothing
End Sub

Private Sub cboMajorRes_Change()
    Call DisplayDetails
End Sub

Private Sub cboMajorRes_Click()
    Call DisplayDetails
End Sub

Private Sub cmdADD_Click()

End Sub

Private Sub cmdAddCate_Click()
    If Not cboMajorRes.Text = "" Then
        Call DisabledFrames(False)
        picCAT.Visible = True
        picCAT.ZOrder 0

        Call ListViewCategory
    End If
End Sub

Private Sub cmdAddDetails_Click()
    If Not cboMajorRes.Text = "" Then
        Call DisabledFrames(False)
        picAddDetails.Visible = True
        picAddDetails.ZOrder 0

        Call CleanDetText
        Call ListViewDet
    End If
End Sub

Private Sub cmdAddRes_Click()
    If Not lblEMPNO.Caption = "" Then
        Call DisabledFrames(False)
        picMAS.Visible = True
        picMAS.ZOrder 0

        txtMASMajor.Text = ""
        txtMASperiod.Text = ""
        txtMASJOB.Text = ""

        Call ListViewMas
    Else
        MsgBox "Choose a Employee First", vbInformation, "Information"
        cmdChoose.SetFocus
    End If
End Sub

Private Sub cmdCATADD_Click()
    lblADD_EDIT_CAT.Caption = "ADD"

    cmdCATADD.Enabled = False
    cmdCATEDIT.Enabled = False
    cmdCATDelete.Enabled = False
    lsvCAT.Enabled = False

    cmdCATSave.Enabled = True
    cmdCATCancel.Enabled = True

    fmeCAT.Enabled = True
    txtCAT.Text = ""
    txtCAT.SetFocus
End Sub

Private Sub cmdCATCancel_Click()
    cmdCATCancel.Enabled = False
    cmdCATSave.Enabled = False

    cmdCATDelete.Enabled = True
    cmdCATADD.Enabled = True
    cmdCATEDIT.Enabled = True
    lsvCAT.Enabled = True

    fmeCAT.Enabled = False
End Sub

Private Sub cmdCATDelete_Click()
    If Not lblCATID.Caption = "" Then
        If MsgBox("Delete This Entry", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("Delete  From PAS_FORMA_CATEGORY Where Empno = '" & lblEMPNO.Caption & _
                              "' And ID = '" & Trim(lblMASID.Caption) & "'")
            gconDMIS.Execute ("Delete From PAS_FORMA_DETAILS Where Empno = '" & lblEMPNO.Caption & _
                              "' And CatID = '" & Trim(lblCATID.Caption) & "'")

            Call CleanMasText
            Call ListViewMas

            fmeMAS.Enabled = False
        End If
    Else
        MsgBox "Choose What To Delete", vbInformation, "Information"
        lsvCAT.SetFocus
    End If
End Sub

Private Sub cmdCATEDIT_Click()
    If Not lblCATID.Caption = "" Then
        lblADD_EDIT_CAT.Caption = "EDIT"

        cmdCATADD.Enabled = False
        cmdCATDelete.Enabled = False
        cmdCATEDIT.Enabled = False
        lsvCAT.Enabled = False

        cmdCATSave.Enabled = True
        cmdCATCancel.Enabled = True

        fmeCAT.Enabled = True
        txtCAT.SetFocus
    Else
        MsgBox "Choose What Entry to Edit", vbInformation, "Information"
        lsvCAT.SetFocus
    End If
End Sub

Private Sub cmdCatExit_Click()
    Call DisabledFrames(True)
    picCAT.Visible = False
    picCAT.ZOrder 1

    Call DisplayCategoryList
End Sub

Private Sub cmdCATSave_Click()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ID                                                            As Integer
    Dim CATEGORY                                                      As String
    Dim MASID                                                         As String

    If MsgBox("Save Category Entry", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        CATEGORY = N2Str2Null(txtCAT.Text)
        MASID = Right(Trim(cboMajorRes.Text), 2)

        If lblADD_EDIT_CAT.Caption = "ADD" Then
            Set rsTmp = gconDMIS.Execute("Select CatID From PAS_FORMA_CATEGORY Where EmpNo = '" & Trim(lblEMPNO.Caption) & "' Order By CatID DESC")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                ID = rsTmp!CatID
            End If
            ID = ID + 1

            gconDMIS.Execute ("Insert Into PAS_FORMA_CATEGORY Values('" & Trim(lblEMPNO.Caption) & _
                              "','" & MASID & "'," & CATEGORY & ",'" & ID & "')")
        Else
            gconDMIS.Execute ("Update PAS_FORMA_CATEGORY Set Category = " & CATEGORY & ",PerCov = " & PERIODCON & _
                            " Where EmpNO = '" & Trim(lblEMPNO.Caption) & "' And ID = '" & ID & "' And CatID = '" & lblCATID.Caption & "'")
        End If

        Call ListViewCategory
        Call cmdCATCancel_Click
    End If
End Sub

Private Sub cmdChoose_Click()
    Call DisabledFrames(False)
    picCHOOSE.Visible = True
    picCHOOSE.ZOrder 0

    txtCHOOSE.Text = ""
    txtCHOOSE.SetFocus
End Sub

Private Sub cmdChooseExit_Click()
    picCHOOSE.Visible = False
    picCHOOSE.ZOrder 1
    Call DisabledFrames(True)
End Sub

Private Sub cmdDetAdd_Click()
    lblAdd_EDIT_DET.Caption = "ADD"

    cmdDetAdd.Enabled = False
    cmdDetEdit.Enabled = False
    cmdDetDelete.Enabled = False
    lsvDET.Enabled = False

    cmdDetSave.Enabled = True
    cmdDetCancel.Enabled = True

    Call CleanDetText
    fmeDet.Enabled = True
    txtDetProc.SetFocus
End Sub

Private Sub cmdDetCancel_Click()
    cmdDetCancel.Enabled = False
    cmdDetSave.Enabled = False

    cmdDetDelete.Enabled = True
    cmdDetAdd.Enabled = True
    cmdDetEdit.Enabled = True
    lsvDET.Enabled = True

    fmeDet.Enabled = False
End Sub

Private Sub cmdDetDelete_Click()
    If Not lblDETID.Caption = "" Then
        If MsgBox("Delete This Entry", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("Delete  From PAS_FORMA_DETAILS Where Empno = '" & lblEMPNO.Caption & _
                              "' And CatID = '" & lblCATEID & "' And Details_ID = '" & Trim(lblDETID.Caption) & "'")

            Call CleanDetText
            Call ListViewDet

            fmeDet.Enabled = False
        End If
    Else
        MsgBox "Choose What To Delete", vbInformation, "Information"
    End If
End Sub

Private Sub cmdDetEdit_Click()
    If Not lblDETID.Caption = "" Then
        lblAdd_EDIT_DET.Caption = "EDIT"

        cmdDetAdd.Enabled = False
        cmdDetDelete.Enabled = False
        cmdDetEdit.Enabled = False
        lsvDET.Enabled = False

        cmdDetSave.Enabled = True
        cmdDetCancel.Enabled = True

        fmeDet.Enabled = True
        txtDetProc.SetFocus
    Else
        MsgBox "Choose What Entry to Edit", vbInformation, "Information"
    End If
End Sub

Private Sub cmdDetSave_Click()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ID                                                            As Integer
    Dim PROCESS                                                       As String
    Dim STANDARD                                                      As String
    Dim MONITOR                                                       As String
    Dim Score                                                         As String

    If MsgBox("Save Entry", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        PROCESS = N2Str2Null(txtDetProc.Text)
        STANDARD = N2Str2Null(txtDetStd.Text)
        MONITOR = N2Str2Null(txtdetMon.Text)
        Score = Null2String(txtDetScore.Text)

        If lblAdd_EDIT_DET.Caption = "ADD" Then
            Set rsTmp = gconDMIS.Execute("Select Details_ID From PAS_FORMA_DETAILS Where EmpNo = '" & Trim(lblEMPNO.Caption) & _
                                         "' And CatID = '" & lblCATEID.Caption & "' Order By Details_ID DESC")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                ID = rsTmp!details_ID
            End If
            ID = ID + 1

            gconDMIS.Execute ("Insert Into PAS_FORMA_Details Values('" & Trim(lblEMPNO.Caption) & _
                              "','" & lblCATEID.Caption & "'," & PROCESS & "," & STANDARD & "," & MONITOR & _
                              "," & Score & ",'" & ID & "')")
        Else
            ID = lblDETID.Caption
            gconDMIS.Execute ("Update PAS_FORMA_DETAILS Set Process = " & PROCESS & ",Standard = " & STANDARD & ",Monitor = " & MONITOR & ",Score = " & Score & " Where EmpNO = '" & Trim(lblEMPNO.Caption) & _
                              "' And Details_ID = '" & lblDETID.Caption & "'")

        End If

        Call ListViewDet
        Call cmdDetCancel_Click
    End If
End Sub



Private Sub cmdEDIT_Click()

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExitDetail_Click()
    Call DisabledFrames(True)
    picAddDetails.Visible = False
    picAddDetails.ZOrder 1

    Call DisplayListDetails
End Sub

Private Sub cmdMasAdd_Click()
    lblADDEDIT.Caption = "ADD"

    cmdMasAdd.Enabled = False
    cmdMasEdit.Enabled = False
    cmdMasDelete.Enabled = False
    lsvMAS.Enabled = False

    cmdMasSave.Enabled = True
    cmdMasCancel.Enabled = True

    Call CleanMasText
    fmeMAS.Enabled = True
    txtMASMajor.SetFocus
End Sub

Private Sub cmdMasCancel_Click()
    cmdMasCancel.Enabled = False
    cmdMasSave.Enabled = False

    cmdMasDelete.Enabled = True
    cmdMasAdd.Enabled = True
    cmdMasEdit.Enabled = True
    lsvMAS.Enabled = True

    fmeMAS.Enabled = False
End Sub

Private Sub cmdMasDelete_Click()
    If Not lblMASID.Caption = "" Then
        If MsgBox("Delete This Entry", vbQuestion + vbYesNo + vbDefaultButton2, "Are You Sure") = vbYes Then
            gconDMIS.Execute ("Delete  From PAS_FORMA_MASTER Where Empno = '" & lblEMPNO.Caption & _
                              "' And ID = '" & Trim(lblMASID.Caption) & "'")
            gconDMIS.Execute ("Delete From PAS_FORMA_CATEGORY Where Empno = '" & lblEMPNO.Caption & _
                              "' And ID = '" & Trim(lblMASID.Caption) & "'")
            gconDMIS.Execute ("Delete From PAS_FORMA_DETAILS Where Empno = '" & lblEMPNO.Caption & _
                              "' And ID = '" & Trim(lblMASID.Caption) & "'")

            Call CleanMasText
            Call ListViewMas

            fmeMAS.Enabled = False
        End If
    Else
        MsgBox "Choose What To Delete", vbInformation, "Information"
    End If
End Sub

Private Sub cmdMasEdit_Click()
    If Not lblMASID.Caption = "" Then
        lblADDEDIT.Caption = "EDIT"

        cmdMasAdd.Enabled = False
        cmdMasDelete.Enabled = False
        cmdMasEdit.Enabled = False
        lsvMAS.Enabled = False

        cmdMasSave.Enabled = True
        cmdMasCancel.Enabled = True

        fmeMAS.Enabled = True
        txtMASMajor.SetFocus
    Else
        MsgBox "Choose What Entry to Edit", vbInformation, "Information"
    End If
End Sub

Private Sub cmdMasExit_Click()
    picMAS.Visible = False
    picMAS.ZOrder 1
    Call DisabledFrames(True)

    Call FillcboMajor
    cboMajorRes.SetFocus
End Sub

Private Sub cmdMasSave_Click()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ID                                                            As Integer
    Dim MAJORREASON                                                   As String
    Dim PERIODCON                                                     As String
    Dim JOBSUMM                                                       As String

    If MsgBox("Save Entry", vbQuestion + vbYesNo + vbDefaultButton1, "Are You Sure") = vbYes Then
        MAJORREASON = N2Str2Null(txtMASMajor.Text)
        PERIODCON = N2Str2Null(txtMASperiod.Text)
        JOBSUMM = N2Str2Null(txtMASJOB.Text)

        If lblADDEDIT.Caption = "ADD" Then
            Set rsTmp = gconDMIS.Execute("Select ID From PAS_FORMA_MASTER Where EmpNo = '" & Trim(lblEMPNO.Caption) & "' Order By ID DESC")
            If Not (rsTmp.BOF And rsTmp.EOF) Then
                ID = rsTmp!ID
            End If
            ID = ID + 1

            gconDMIS.Execute ("Insert Into PAS_FORMA_MASTER Values('" & Trim(lblEMPNO.Caption) & _
                              "'," & MAJORREASON & "," & PERIODCON & "," & JOBSUMM & ",'" & ID & "')")
        Else
            ID = lblMASID.Caption
            gconDMIS.Execute ("Update PAS_FORMA_MASTER Set MajorRespon = " & MAJORREASON & ",PerCov = " & PERIODCON & ",JobSumm = " & JOBSUMM & " Where EmpNO = '" & Trim(lblEMPNO.Caption) & _
                              "' And ID = '" & ID & "'")

        End If

        Call ListViewMas
        Call cmdMasCancel_Click
    End If
End Sub



Private Sub Command2_Click()
    lblAdd_EDIT_DET.Caption = "ADD"

    cmdDetAdd.Enabled = False
    cmdMasEdit.Enabled = False
    cmdMasDelete.Enabled = False
    lsvMAS.Enabled = False

    cmdMasSave.Enabled = True
    cmdMasCancel.Enabled = True

    Call CleanMasText
    fmeMAS.Enabled = True
    txtMASMajor.SetFocus
End Sub

Private Sub Command1_Click()
Unload Me
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape:

    End Select
End Sub

Private Sub Form_Load()

    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub lsvCATEGORY_Click()
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim INDEX                                                         As Double
    Dim ITEM                                                          As ListItem

    If Not lsvCATEGORY.ListItems.count = 0 Then
        INDEX = lsvCATEGORY.SelectedItem.INDEX

        With lsvCATEGORY
            lblCATEID.Caption = .ListItems(INDEX).SubItems(1)
            Call DisplayListDetails
        End With
    End If
End Sub

Private Sub lsvCHOOSE_DblClick()
    Dim INDEX                                                         As Double

    Dim ZEROS                                                         As String

    If Not lsvCHOOSE.ListItems.count = 0 Then
        INDEX = lsvCHOOSE.SelectedItem.INDEX
        lblTotal.Caption = ""

        With lsvCHOOSE
            lblEmpName.Caption = Null2String(.ListItems(INDEX).Text)
            lblEMPNO.Caption = Null2String(.ListItems(INDEX).SubItems(1))

            Call DisabledFrames(True)
            picCHOOSE.Visible = False
            picCHOOSE.ZOrder 1

            txtPerCover.Text = ""
            txtJSummary.Text = ""

            Call FillcboMajor

            lsvCATEGORY.ListItems.Clear
            lsvDetails.ListItems.Clear

            If Not cboMajorRes.Text = "" Then Call cboMajorRes_Click
            If Not cboMajorRes.Text = "" Then Call lsvCATEGORY_Click
            cboMajorRes.SetFocus
        End With
    End If
End Sub

Private Sub lsvFields_DblClick()
    Dim INDEX                                                         As Double

    If Not lsvFields.ListItems.count = 0 Then

        With lsvFields


        End With
    End If
End Sub

Private Sub lsvDET_Click()
    Dim INDEX                                                         As Double

    If Not lsvDET.ListItems.count = 0 Then
        INDEX = lsvDET.SelectedItem.INDEX
        With lsvDET
            lblDETID.Caption = .ListItems(INDEX).SubItems(4)
            txtDetProc.Text = .ListItems(INDEX).Text
            txtDetStd.Text = .ListItems(INDEX).SubItems(1)
            txtdetMon.Text = .ListItems(INDEX).SubItems(2)
            txtDetScore.Text = .ListItems(INDEX).SubItems(3)
        End With
    End If
End Sub

Private Sub lsvDET_DblClick()
    If Not lsvDET.ListItems.count = 0 Then
        Call cmdDetEdit_Click
    End If
End Sub

Private Sub lsvMAS_Click()
    Dim INDEX                                                         As Double

    If Not lsvMAS.ListItems.count = 0 Then
        INDEX = lsvMAS.SelectedItem.INDEX
        With lsvMAS
            txtMASMajor.Text = .ListItems(INDEX).Text
            txtMASperiod.Text = .ListItems(INDEX).SubItems(1)
            txtMASJOB.Text = .ListItems(INDEX).SubItems(2)
            lblMASID.Caption = .ListItems(INDEX).SubItems(3)
        End With
    End If
End Sub

Private Sub lsvMAS_DblClick()
    Call lsvMAS_Click
    Call cmdMasEdit_Click
End Sub

Private Sub optFNAME_Click()
    txtCHOOSE.SetFocus
End Sub

Private Sub optLNAME_Click()
    txtCHOOSE.SetFocus
End Sub

Private Sub Timer1_Timer()
    If lblCAP(0).ForeColor = &HFFFF& Then
        lblCAP(0).ForeColor = &H8000&
    Else
        lblCAP(0).ForeColor = &HFFFF&
    End If
End Sub

Private Sub txtCHOOSE_Change()
    Dim SQL                                                           As String
    Dim Keyword                                                       As String
    Dim rsTmp                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Keyword = txtCHOOSE.Text
    If Not Null2String(Keyword) = "" Then
        If optLNAME.Value = True Then
            Set rsTmp = gconDMIS.Execute("Select EmpNO,FirstName,LastName,MiddleName From HRMS_EMPINFO Where LastName Like '" & Keyword & "%' ORder By LastName ASC")
        Else
            Set rsTmp = gconDMIS.Execute("Select EmpNO,FirstName,LastName,MiddleName From HRMS_EMPINFO Where FirstName Like '" & Keyword & "%' ORder By FirstName ASC")
        End If

        If Not (rsTmp.BOF And rsTmp.EOF) Then
            Do While Not rsTmp.EOF
                Set ITEM = lsvCHOOSE.ListItems.Add(, , Null2String(rsTmp!lastname & "," & rsTmp!FIRSTNAME & " " & rsTmp!MIDDLENAME))
                ITEM.SubItems(1) = Null2String(rsTmp!EMPNO)

                rsTmp.MoveNext
            Loop
        Else
            lsvCHOOSE.ListItems.Clear
        End If
    Else
        lsvCHOOSE.ListItems.Clear
    End If

    Set rsTmp = Nothing
End Sub

