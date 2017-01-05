VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmCSMSClockINOUT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Clock / Job Clock Log-In"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmClockINOUT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14370
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   60
      ScaleHeight     =   555
      ScaleWidth      =   6825
      TabIndex        =   85
      Top             =   6000
      Width           =   6855
      Begin VB.CommandButton Command3 
         Caption         =   "View DTR"
         Height          =   375
         Left            =   4200
         TabIndex        =   89
         Top             =   60
         Width           =   1155
      End
      Begin VB.ComboBox cboFilter_Yearly 
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
         IntegralHeight  =   0   'False
         Left            =   2910
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   88
         ToolTipText     =   "Select This [OPTION] To Filter DTR By Year - Or Select Option [ ALL ] to View For All the Year"
         Top             =   90
         Width           =   1245
      End
      Begin VB.ComboBox cboFilter_month 
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
         IntegralHeight  =   0   'False
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   86
         ToolTipText     =   "Select This [OPTION] To Filter DTR By Month - Or Select Option [ ALL ] to View For All the Month"
         Top             =   90
         Width           =   1635
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month / Year"
         Height          =   225
         Left            =   90
         TabIndex        =   87
         ToolTipText     =   "Select  [OPTION] To Filter DTR By Month Or By Year - Or Option [ ALL ] to View For All the Month Or For All The Year"
         Top             =   165
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2865
      Left            =   180
      TabIndex        =   61
      Top             =   9510
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5054
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
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TechCode"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   45
      ImageHeight     =   46
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClockINOUT.frx":1082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClockINOUT.frx":150F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClockINOUT.frx":1F42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbojStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   4545
   End
   Begin MSComctlLib.ListView lblTech 
      Height          =   5295
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmClockINOUT.frx":29B7
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Technician"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "R. Order"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Techname"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "In/Out"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "techcode"
         Object.Width           =   2
      EndProperty
   End
   Begin MSComctlLib.ListView lstDTR 
      Height          =   2205
      Left            =   60
      TabIndex        =   3
      Top             =   6630
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   3889
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
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmClockINOUT.frx":2B19
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "RO Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Time In Am"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time Out PM"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Clock-In"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Clock-Out"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Hours"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Job Code"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Clock Out - Status"
         Object.Width           =   3616
      EndProperty
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   6930
      ScaleHeight     =   555
      ScaleWidth      =   7365
      TabIndex        =   68
      Top             =   6000
      Width           =   7395
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FINISH JOB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   5880
         TabIndex        =   71
         Top             =   135
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   5340
         Picture         =   "FrmClockINOUT.frx":2C7B
         Stretch         =   -1  'True
         Top             =   -30
         Width           =   570
      End
      Begin VB.Image img1 
         Height          =   510
         Left            =   210
         Picture         =   "FrmClockINOUT.frx":30F8
         Stretch         =   -1  'True
         Top             =   30
         Width           =   600
      End
      Begin VB.Label lblIPD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOT STARTED"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   240
         Left            =   840
         TabIndex        =   70
         Top             =   150
         Width           =   1365
      End
      Begin VB.Label lblOPD 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WORK IN PROCESS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   3165
         TabIndex        =   69
         Top             =   135
         Width           =   1875
      End
      Begin VB.Image img2 
         Height          =   555
         Left            =   2565
         Picture         =   "FrmClockINOUT.frx":3B1B
         Stretch         =   -1  'True
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.PictureBox picClockIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   6930
      ScaleHeight     =   5835
      ScaleWidth      =   7365
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4590
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         Height          =   585
         Left            =   540
         Picture         =   "FrmClockINOUT.frx":4580
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print Time Card"
         Top             =   7350
         Width           =   1575
      End
      Begin VB.CommandButton cmdInCancel 
         Height          =   585
         Left            =   5610
         Picture         =   "FrmClockINOUT.frx":8016
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   5220
         Width           =   1575
      End
      Begin VB.CommandButton cmdIn 
         Height          =   585
         Left            =   4050
         Picture         =   "FrmClockINOUT.frx":B58C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Clock In Technician"
         Top             =   5220
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   7500
         TabIndex        =   37
         Top             =   870
         Visible         =   0   'False
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
         Format          =   84279299
         CurrentDate     =   39021
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3585
         Left            =   90
         ScaleHeight     =   3555
         ScaleWidth      =   7185
         TabIndex        =   8
         Top             =   1590
         Width           =   7215
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Time  :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            TabIndex        =   67
            Top             =   2700
            Width           =   4245
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Date  :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            TabIndex        =   66
            Top             =   2190
            Width           =   4245
         End
         Begin VB.Label labJobDesc 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   675
            Left            =   120
            TabIndex        =   63
            Top             =   1020
            Width           =   6885
         End
         Begin VB.Label labJobCode 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   630
            Width           =   2655
         End
         Begin VB.Label labtime 
            BackStyle       =   0  'Transparent
            Caption         =   "01:32:25 PM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   4710
            TabIndex        =   19
            Top             =   2700
            Width           =   1755
         End
         Begin VB.Label labDate 
            BackStyle       =   0  'Transparent
            Caption         =   "10/31/2006"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4770
            TabIndex        =   18
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "You will be Clocked In for this Job  :"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   150
            TabIndex        =   12
            Top             =   1740
            Width           =   4545
         End
         Begin VB.Label labEmployeein 
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1710
            TabIndex        =   10
            Top             =   60
            Width           =   5085
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee :"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   150
            TabIndex        =   9
            Top             =   90
            Width           =   1575
         End
         Begin VB.Label labJobItemNo 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   2790
            TabIndex        =   64
            Top             =   630
            Visible         =   0   'False
            Width           =   705
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   72
         Top             =   0
         Width           =   7425
         _Version        =   655364
         _ExtentX        =   13097
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   8388608
      End
      Begin VB.Image Image1 
         Height          =   765
         Left            =   210
         Picture         =   "FrmClockINOUT.frx":EB02
         Stretch         =   -1  'True
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Clock"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   645
         Left            =   1890
         TabIndex        =   7
         Top             =   -1140
         Width           =   2835
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1590
         TabIndex        =   6
         Top             =   -420
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblIS_TECH_STATUS 
         BackStyle       =   0  'Transparent
         Caption         =   "Clock In"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   645
         Left            =   2610
         TabIndex        =   5
         Top             =   540
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   6930
      ScaleHeight     =   5835
      ScaleWidth      =   7365
      TabIndex        =   39
      Top             =   90
      Width           =   7395
      Begin TabDlg.SSTab lblJob4ServicePast 
         Height          =   4200
         Left            =   60
         TabIndex        =   74
         Top             =   1590
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   7408
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "FrmClockINOUT.frx":F4FD
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblJob4Service"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSComctlLib.ListView lblJob4Service 
            Height          =   4050
            Left            =   60
            TabIndex        =   75
            Top             =   60
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   7144
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   0
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmClockINOUT.frx":F519
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Job Code"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   " Jobs Description"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Std.Rate"
               Object.Width           =   1676
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Line No"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Note"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "STATUS"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.PictureBox theIdlePic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   4200
         Left            =   60
         ScaleHeight     =   4170
         ScaleWidth      =   7215
         TabIndex        =   76
         Top             =   1590
         Width           =   7245
         Begin VB.TextBox txtidle 
            Height          =   1845
            Left            =   150
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   81
            Top             =   1020
            Width           =   6915
         End
         Begin VB.CommandButton cmdCancelidle 
            Caption         =   "Cancel"
            Height          =   525
            Left            =   6270
            TabIndex        =   79
            Top             =   3390
            Width           =   795
         End
         Begin VB.ComboBox cboIdleTemplate 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   78
            Top             =   450
            Width           =   5805
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Refresh"
            Height          =   495
            Left            =   6000
            TabIndex        =   77
            Top             =   420
            Width           =   1065
         End
         Begin VB.CommandButton CmdIdle 
            Caption         =   "Ok"
            Height          =   525
            Left            =   5460
            TabIndex        =   80
            Top             =   3390
            Width           =   825
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Reason for Idle time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   150
            TabIndex        =   84
            Top             =   60
            Width           =   2325
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Note:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   210
            TabIndex        =   83
            Top             =   2910
            Width           =   555
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Please Input your reason\s for clocking out,and be specific.."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   210
            TabIndex        =   82
            Top             =   3180
            Width           =   6135
         End
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Promise"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4890
         TabIndex        =   58
         Top             =   90
         Width           =   705
      End
      Begin VB.Label labPromise 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5700
         TabIndex        =   57
         Top             =   90
         Width           =   1605
      End
      Begin VB.Label labVinNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3120
         TabIndex        =   56
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VIN No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2370
         TabIndex        =   55
         Top             =   1230
         Width           =   570
      End
      Begin VB.Label labSection 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         TabIndex        =   54
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Section "
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
         Height          =   210
         Left            =   315
         TabIndex        =   53
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label labkmRdg 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5700
         TabIndex        =   52
         Top             =   1170
         Width           =   1605
      End
      Begin VB.Label labActNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         TabIndex        =   51
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label labPlateNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5700
         TabIndex        =   50
         Top             =   810
         Width           =   1605
      End
      Begin VB.Label labVehicle 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         TabIndex        =   49
         Top             =   810
         Width           =   3795
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4920
         TabIndex        =   48
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle"
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
         Height          =   210
         Left            =   330
         TabIndex        =   47
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "KM Rdg."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4995
         TabIndex        =   46
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label labApptDate 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3390
         TabIndex        =   45
         Top             =   90
         Width           =   1425
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2925
         TabIndex        =   44
         Top             =   90
         Width           =   360
      End
      Begin VB.Label labRO 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         TabIndex        =   43
         Top             =   90
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " R/O No."
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
         Height          =   210
         Left            =   330
         TabIndex        =   42
         Top             =   120
         Width           =   615
      End
      Begin VB.Label labCustomer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2400
         TabIndex        =   41
         Top             =   450
         Width           =   4905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Height          =   210
         Left            =   105
         TabIndex        =   40
         Top             =   480
         Width           =   840
      End
   End
   Begin VB.PictureBox picClockOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5865
      Left            =   6930
      ScaleHeight     =   5835
      ScaleWidth      =   7365
      TabIndex        =   20
      Top             =   90
      Visible         =   0   'False
      Width           =   7395
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4710
         TabIndex        =   36
         Top             =   1170
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
         Format          =   84279299
         CurrentDate     =   39021
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   1170
         Top             =   480
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   150
         ScaleHeight     =   855
         ScaleWidth      =   7035
         TabIndex        =   31
         Top             =   1590
         Width           =   7065
         Begin Crystal.CrystalReport rptClock 
            Left            =   90
            Top             =   30
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.Label labOut 
            BackStyle       =   0  'Transparent
            Caption         =   "10/31/2006   02:32:25 AM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   2580
            TabIndex        =   35
            Top             =   420
            Width           =   4215
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Clock Out Time  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   480
            TabIndex        =   34
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label labIn 
            BackStyle       =   0  'Transparent
            Caption         =   "10/31/2006   02:32:25 AM"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   345
            Left            =   2580
            TabIndex        =   33
            Top             =   30
            Width           =   4215
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Clock In Time  :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   660
            TabIndex        =   32
            Top             =   60
            Width           =   1965
         End
      End
      Begin VB.CommandButton cmdOutCancel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5610
         Picture         =   "FrmClockINOUT.frx":F67B
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Cancel "
         Top             =   5100
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Height          =   585
         Left            =   3750
         Picture         =   "FrmClockINOUT.frx":12BF1
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Print Time Card"
         Top             =   7140
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTClockIn 
         Height          =   315
         Left            =   4560
         TabIndex        =   38
         Top             =   -750
         Visible         =   0   'False
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
         Format          =   84279299
         CurrentDate     =   39021
      End
      Begin VB.CommandButton cmdOut 
         Height          =   585
         Left            =   4050
         Picture         =   "FrmClockINOUT.frx":16687
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Clock Out Technician"
         Top             =   5100
         Width           =   1575
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   150
         ScaleHeight     =   1635
         ScaleWidth      =   7035
         TabIndex        =   21
         Top             =   2700
         Width           =   7065
         Begin VB.ComboBox cboReasonClockingout 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   3060
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1050
            Width           =   3855
         End
         Begin VB.Label labIdleReason 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   285
            Left            =   6090
            TabIndex        =   59
            Top             =   180
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label labHours 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "12.56"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   5970
            TabIndex        =   30
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Reason for Clocking Out"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   2985
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee :"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   180
            TabIndex        =   24
            Top             =   90
            Width           =   1575
         End
         Begin VB.Label labEmployeeOut 
            BackStyle       =   0  'Transparent
            Caption         =   "Status:"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   15.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1740
            TabIndex        =   23
            Top             =   30
            Width           =   5085
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "You are Clocking Out, your hour for this period are"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   22
            Top             =   540
            Width           =   5865
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   345
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   7395
         _Version        =   655364
         _ExtentX        =   13044
         _ExtentY        =   609
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   16711680
         GradientColorDark=   8388608
      End
      Begin VB.Label labOutItemNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   150
         TabIndex        =   65
         Top             =   4410
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Clock Out"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   2520
         TabIndex        =   60
         Top             =   510
         Width           =   2205
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Clock Out"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   2790
         TabIndex        =   27
         Top             =   -600
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1770
         TabIndex        =   26
         Top             =   -600
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Clock"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   2520
         TabIndex        =   25
         Top             =   -360
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Image Image2 
         Height          =   765
         Left            =   90
         Picture         =   "FrmClockINOUT.frx":1A11D
         Stretch         =   -1  'True
         Top             =   450
         Width           =   765
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TECH. JOB STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   2175
   End
End
Attribute VB_Name = "frmCSMSClockINOUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheLineNo                                          As String
Dim TechCode                                           As String
Dim theRo                                              As String
Dim TheEmpNO                                           As String
Dim RSUPLOAD                                           As ADODB.Recordset
Dim xJOBCODE                                           As String
Dim EVENTRO                                            As String
Event JOBCLOCKED()
Event FORMCLOSED()

Function CheckAllJobsISDone(vRO) As Boolean
    Dim RS                                             As New ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE LIVIL = '1' AND (DONE = 'N' OR DONE ='W' OR DONE IS NULL) and REP_OR = " & vRO & "")
    If RS.EOF And RS.BOF Then
        CheckAllJobsISDone = True
    Else
        CheckAllJobsISDone = False
    End If
    Set RS = Nothing
End Function

Function GetJobStatus(XXX As String) As String
    Dim rsJOBSTATUS                                    As New ADODB.Recordset
    Set rsJOBSTATUS = gconDMIS.Execute("Select * from CSMS_JobStatus Where Description = '" & XXX & "'")
    If Not rsJOBSTATUS.EOF And Not rsJOBSTATUS.BOF Then
        GetJobStatus = Null2String(rsJOBSTATUS!Code)
    End If
End Function

Function IfROIsFinish(xRO_NO As Variant) As Boolean
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT Status FROM CSMS_repairOrder Where Status = 'Finish Job' AND RO_No = '" & theRo & "'")
    If Not RS.EOF And Not RS.BOF Then
        IfROIsFinish = True
        MsgBox "This Job Is Already Finished!", vbExclamation, "Information"
    Else
        IfROIsFinish = False
    End If
    Set RS = Nothing
End Function

Sub FillDTR()
    Dim strSearch                                      As String

    If cboFilter_month <> "ALL" And cboFilter_month <> "" Then
        strSearch = " AND MONTH(TRANDATE) =" & What_month(cboFilter_month)
    End If

    If cboFilter_Yearly <> "ALL" And cboFilter_Yearly <> "" Then
        strSearch = strSearch & " AND YEAR(TRANDATE) =" & cboFilter_Yearly
    End If

    Dim Item                                           As ListItem
    Dim xMIN                                           As Double
    Dim xHOUR                                          As Integer
    Dim xtime                                          As String
    Dim xREM                                           As Double

    lstDTR.Sorted = False: lstDTR.ListItems.Clear
    'Set rsUpload = gconDMIS.Execute("select  Trandate,RO_NO,Time_in_Am,Time_Out_Am,Time_In_Pm,Time_Out_Pm,CONVERT(VarChar,CAST(clockin AS SMALLDATETIME),8) AS CLOCKIN,CONVERT(Varchar,CAST(clockout AS SMALLDATETIME),8) AS CLOCKOUT,hrsWorked,DetCde,reasonforclockout from CSMS_JobClock where technician = '" & lblTech.SelectedItem & "' order by trandate desc,RO_NO desc")
    Set RSUPLOAD = gconDMIS.Execute("select  Trandate,RO_NO,Time_in_Am,Time_Out_Pm,CONVERT(VarChar,CAST(clockin AS SMALLDATETIME),8) AS CLOCKIN,CONVERT(Varchar,CAST(clockout AS SMALLDATETIME),8) AS CLOCKOUT,hrsWorked,DetCde,reasonforclockout from CSMS_JobClock where technician = '" & LTrim(RTrim(lbltech.SelectedItem)) & "'" & strSearch & " order by trandate desc,RO_NO desc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        'Listview_Loadval Me.lstDTR.ListItems, rsUpload
        Do While Not RSUPLOAD.EOF
            Set Item = lstDTR.ListItems.Add(, , RSUPLOAD!trandate)
            Item.SubItems(1) = Null2String(RSUPLOAD!RO_NO)
            Item.SubItems(2) = Null2String(RSUPLOAD!TIME_IN_AM)
            Item.SubItems(3) = Null2String(RSUPLOAD!TIME_OUT_PM)
            Item.SubItems(4) = Null2String(RSUPLOAD!CLOCKIN)
            Item.SubItems(5) = Null2String(RSUPLOAD!CLOCKOUT)

            'If (Null2String(rsUpload!clockout) = "" Or Null2String(rsUpload!clockin) = "") Then
            '    Item.SubItems(6) = Null2String("")
            'Else
            'xHOUR = DateDiff("M", Null2String(rsUpload!CLOCKIN), Null2String(rsUpload!CLOCKOUT))  'Null2String(rsUpload!hrsWorked)
            '    xMIN = DateDiff("N", Null2String(rsUpload!clockin), Null2String(rsUpload!clockout))   'Null2String(rsUpload!hrsWorked)
            '    xHOUR = 0
            '    If xMIN >= 60 Then
            '        xHOUR = xMIN / 60
            '        xREM = xMIN \ 60
            '
            '        xtime = xHOUR & "." & Format(xREM, "00")
            '    Else
            '        xtime = xHOUR & "." & Format(xMIN, "00")
            '    End If
            '    Item.SubItems(6) = xtime
            'End If

            Item.SubItems(6) = Null2String(RSUPLOAD!hrsWorked)
            Item.SubItems(7) = Null2String(RSUPLOAD!DETCDE)
            Item.SubItems(8) = Null2String(RSUPLOAD!reasonforclockout)

            RSUPLOAD.MoveNext
        Loop
    End If
End Sub

Sub FillReasonClockOut()
    Dim rsReasonClockOut                               As New ADODB.Recordset
    Set rsReasonClockOut = gconDMIS.Execute("Select description,Inout,Out from CSMS_JobStatus Where [Out] = 'O' order by description asc")
    If Not rsReasonClockOut.EOF And Not rsReasonClockOut.BOF Then
        rsReasonClockOut.MoveFirst: cboReasonClockingout.Clear
        Do While Not rsReasonClockOut.EOF
            If Null2String(rsReasonClockOut![out]) = "O" Then cboReasonClockingout.AddItem RTrim(LTrim(Null2String(rsReasonClockOut![Description])))
            rsReasonClockOut.MoveNext
        Loop
        cboReasonClockingout.ListIndex = 0
    End If
End Sub

Sub RefreshTechnicianStatus()
    Dim EVENTRO As String
    EVENTRO = LTrim(RTrim(N2Str2Null(lbltech.SelectedItem.SubItems(2))))
    lbltech.Sorted = False: lbltech.ListItems.Clear
    Set RSUPLOAD = New ADODB.Recordset
    If cbojStatus.Text = "All" Then
        Set RSUPLOAD = gconDMIS.Execute("select  distinct empno,firstname,assignedro,status,tech_name,InOut,techcode from CSMS_vw_TechnicianAvailability order by tech_name asc")
    Else
        Set RSUPLOAD = gconDMIS.Execute("select  distinct empno,firstname,assignedro,status,tech_name,InOut,techcode from CSMS_vw_TechnicianAvailability where status = '" & cbojStatus & "' and [code] <> 'A' order by tech_name asc")
    End If
    Listview_Loadval lbltech.ListItems, RSUPLOAD
End Sub

Sub ViewRoDetail()
    lblJob4Service.ListItems.Clear
    Dim rsFind                                         As New ADODB.Recordset
    Dim rsFind2                                        As New ADODB.Recordset
    Dim IndexList                                      As Integer
    Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where RO_NO = '" & lbltech.SelectedItem.SubItems(2) & "'")
    If Not rsFind.EOF And Not rsFind.BOF Then
        labRO.Caption = lbltech.SelectedItem.SubItems(2)
        labActNo.Caption = Null2String(rsFind![ACCT_NO])
        labCustomer.Caption = Null2String(rsFind![Customer])
        labApptDate.Caption = Null2String(rsFind![AppointmentDate])
        labPromise.Caption = Null2String(rsFind![PromiseDate])
        labPlateNo.Caption = Null2String(rsFind![PLATE_NO])

        Dim rsVehicleKo                                As New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where Cuscde = '" & labActNo.Caption & "' and plate_no = '" & labPlateNo.Caption & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            labVinNo.Caption = Null2String(rsVehicleKo![Vin])
            labVehicle.Caption = Trim(Null2String(rsVehicleKo![YER])) & " " & Trim(Null2String(rsVehicleKo![Make])) & " " & Trim(Null2String(rsVehicleKo![Model]))
        End If

        Dim Item                                       As ListItem

        Set rsFind2 = New ADODB.Recordset
        Dim RSMAKOY                                    As New ADODB.Recordset

        Set rsFind2 = gconDMIS.Execute("Select * from CSMS_vw_EditRO where LIVIL = '1' AND rep_OR = '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(2))) & "' AND TECHCODE = '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(6))) & "' order by line_no asc")
        If Not rsFind2.EOF And Not rsFind2.BOF Then
            labkmRdg.Caption = Null2String(rsFind2![km_rdg])
            labSection.Caption = Null2String(rsFind2![sektion])
            '                With lblJob4Service
            '                    If Null2String(rsFind2![done]) = "Y" Then
            '                        .ListItems.Add , , Null2String(rsFind2![DetCDE]), , 1
            '                    ElseIf Null2String(rsFind2![done]) = "N" Then
            '                        .ListItems.Add , , Null2String(rsFind2![DetCDE]), , 2
            '                    ElseIf Null2String(rsFind2![done]) = "W" Then
            '                        .ListItems.Add , , Null2String(rsFind2![DetCDE]), , 3
            '                    Else
            '                        .ListItems.Add , , Null2String(rsFind2![DetCDE]), , 2
            '                    End If
            '                    .ListItems(.ListItems.Count).ListSubItems.Add 1, , Null2String(rsFind2![Detdsc])
            '                    .ListItems(.ListItems.Count).ListSubItems.Add 2, , NumericVal(rsFind2![DET_HRS])
            '                    .ListItems(.ListItems.Count).ListSubItems.Add 3, , Null2String(rsFind2![Line_No])
            '
            '                End With
            Do While Not rsFind2.EOF
                If Null2String(rsFind2!DONE) = "Y" Then
                    Set Item = lblJob4Service.ListItems.Add(, , Null2String(rsFind2!DETCDE), , 1)
                ElseIf Null2String(rsFind2!DONE) = "N" Then
                    Set Item = lblJob4Service.ListItems.Add(, , Null2String(rsFind2!DETCDE), , 2)
                ElseIf Null2String(rsFind2!DONE) = "W" Then
                    Set Item = lblJob4Service.ListItems.Add(, , Null2String(rsFind2!DETCDE), , 3)
                Else
                    Set Item = lblJob4Service.ListItems.Add(, , Null2String(rsFind2!DETCDE), , 2)
                End If

                Item.SubItems(1) = Null2String(rsFind2!DETDSC)
                Item.SubItems(2) = NumericVal(rsFind2!DET_HRS)
                Item.SubItems(3) = LTrim(RTrim(Null2String(rsFind2!LINE_NO)))
                'ITEM.SubItems(4) = ""
                Item.SubItems(5) = LTrim(RTrim(Null2String(rsFind2!DONE)))

                rsFind2.MoveNext
            Loop
        Else
            'AXP 05082008845 FOR AUTOMATION OF ASSIGNED TECHNICIAN
            '            If labRO.Caption <> "" Then
            '                Screen.MousePointer = 0
            '
            '                MsgBox "Technicain Status Incorrect " & vbCrLf & " System Will Set Technician Status to Available", vbInformation, "Invalid Entry"
            '                gconDMIS.Execute "update HRMS_EmpInfo set JStatus = 'A' ,ASSIGNEDRO = NULL where empno = '" & lblTech.SelectedItem.Text & "'"
            '                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = 'A' , ASSIGNEDRO = NULL where empno = '" & lblTech.SelectedItem.Text & "'"
            '                If CheckAllJobsISDone(labRO) = True Then
            '                    gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = '" & LOGDATE & "', STATUS = 'Finish Job', JStatus = 'F' where RO_No = '" & labRO & "'"
            '                End If
            '                cbojStatus_Click
            '
            '                Exit Sub
            '            End If
        End If
    Else

        labRO.Caption = ""
        labActNo.Caption = ""
        labCustomer.Caption = ""
        labApptDate.Caption = ""
        labPromise.Caption = ""
        labPlateNo.Caption = ""
        labVinNo.Caption = ""
        labVehicle.Caption = ""

        labkmRdg.Caption = ""
        labSection.Caption = ""
    End If
End Sub

Private Sub cboFilter_Ok_Click()

End Sub

Private Sub cboIdleTemplate_Click()
    txtidle.Enabled = True
    cboIdleTemplate.Enabled = False
    txtidle.Text = UCase(cboIdleTemplate.Text) & vbCrLf
End Sub

Private Sub cbojStatus_Click()
    Dim Item                                           As ListItem

    'UPDATED BY: JUN-----------------------------------------------------------------------------------
    'DATE UPDATED: 03-09-2009
    'DESCRIPTION: UPDATE THE JSTATUS TO AVAILABLE  IF THE THE ASSIGNEDRO IS NULL
    '             UPDATE FOR TICKET NO: HGC-12727
     'HRMS_EMPINFO
     'gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET JSTATUS = 'A' WHERE ASSIGNEDRO IS NULL AND IS_TECHNICIAN = '1'")
         
     'CSMS_EMPINFO
     'gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS = 'A' WHERE ASSIGNEDRO IS NULL AND IS_TECHNICIAN = '1'")
    'UPDATED BY: JUN-----------------------------------------------------------------------------------

    lbltech.Sorted = False
    lbltech.ListItems.Clear

    If UCase(cbojStatus.Text) = "ALL" Then
        Set RSUPLOAD = gconDMIS.Execute("SELECT  EMPNO,FIRSTNAME,ASSIGNEDRO,STATUS,TECH_NAME,INOUT,TECHCODE FROM CSMS_VW_TECHNICIANAVAILABILITY ORDER BY TECH_NAME ASC")
    Else
        Set RSUPLOAD = gconDMIS.Execute("SELECT  EMPNO,FIRSTNAME,ASSIGNEDRO,STATUS,TECH_NAME,INOUT,TECHCODE FROM CSMS_VW_TECHNICIANAVAILABILITY WHERE STATUS = '" & cbojStatus & "' ORDER BY  TECH_NAME ASC")
    End If
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lbltech.ListItems.Add(, , Null2String(RSUPLOAD!EMPNO))
            Item.ListSubItems.Add , , Null2String(RSUPLOAD!FIRSTNAME)
            If LTrim(RTrim(Null2String(RSUPLOAD!assignedro))) = "R/O" Then
                Item.ListSubItems.Add , , Null2String("")
                Item.ListSubItems.Add , , Null2String("Available")
            Else
                Item.ListSubItems.Add , , LTrim(RTrim(Null2String(RSUPLOAD!assignedro)))
                Item.ListSubItems.Add , , Null2String(RSUPLOAD!Status)
            End If
            Item.ListSubItems.Add , , Null2String(RSUPLOAD!TECH_NAME)
            Item.ListSubItems.Add , , Null2String(RSUPLOAD!inout)
            Item.ListSubItems.Add , , LTrim(RTrim(Null2String(RSUPLOAD!TechCode)))
            RSUPLOAD.MoveNext
        Loop
    End If

    If lbltech.ListItems.Count = 0 Then
        lblJob4Service.ListItems.Clear
    End If
End Sub

Private Sub cboReasonClockingout_Click()
    If cboReasonClockingout.Text = "Idle Time" Then
        theIdlePic.ZOrder 0
        cboIdleTemplate.Clear
        Dim rsIdleTemplate                             As ADODB.Recordset
        Set rsIdleTemplate = New ADODB.Recordset
        Set rsIdleTemplate = gconDMIS.Execute("Select DISTINCT rtrim(ltrim(REPLACE(REPLACE(REPLACE(TECH2, CHAR(10), ''), CHAR(13), ''), CHAR(9), ''))) as tech2 from CSMS_RepairOrder where isnull(tech2,'')<>''  Order by tech2 asc")
        If Not rsIdleTemplate.EOF And Not rsIdleTemplate.BOF Then
            rsIdleTemplate.MoveFirst
            Do While Not rsIdleTemplate.EOF
                cboIdleTemplate.AddItem UCase(Null2String(rsIdleTemplate!tech2))
                rsIdleTemplate.MoveNext
            Loop
        End If
        theIdlePic.Visible = True
        picClockOut.Visible = False
    End If
End Sub

Private Sub cmdCancelidle_Click()
    cbojStatus.Enabled = True
    lbltech.Enabled = True

    picClockOut.Visible = True
    theIdlePic.Visible = False
    cmdOutCancel.Value = True
End Sub

Private Sub CmdIdle_Click()
    If txtidle.Text = "" Then
        'MsgBox "Pls Input A Reason!", vbExclamation, "Warning!"
        ShowIsRequiredMsg ("Reason cannot be Blank")
        txtidle.SetFocus
        Exit Sub
    End If
    'picClockOut.Visible = True
    cmdOut_Click
End Sub

Function CheckIfBackJob(MLINE As String, XRONO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT BACKJOB_COUNT FROM CSMS_RO_DET WHERE LIVIL = '1' AND LINE_NO = '" & MLINE & "' AND REP_OR = '" & XRONO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If NumericVal(rstmp!BACKJOB_COUNT) > 0 Then
            CheckIfBackJob = True
        Else
            CheckIfBackJob = False
        End If
    End If

    Set rstmp = Nothing
End Function

Private Sub cmdIn_Click()
    Call set_date_time
    Call Sleep(2000)

    labDate.Caption = Format(Now, "MM/dd/yyyy")
    labtime.Caption = Format(Now, "hh:mm:ss ampm")
    DTPicker2.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
        
    If MsgBox("Clock IN this Job", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then
        Exit Sub
    End If

    Dim xTranDate, xTechnician, xTech_Name, xRO_NO, xClockIn, xJStatus, xStatus, xReasonForClockOut As String
    TheLineNo = labJobItemNo.Caption

    xTranDate = N2Str2Null(Format(DTPicker2, "MM/dd/yyyy"))
    xTechnician = LTrim(RTrim(N2Str2Null(lbltech.SelectedItem)))
    xTech_Name = LTrim(RTrim(N2Str2Null(lbltech.SelectedItem.SubItems(4))))
    xRO_NO = N2Str2Null(lbltech.SelectedItem.SubItems(2))
    xClockIn = N2Str2Null(DTPicker2)
    xJStatus = "'W'"
    xStatus = "'IN'"

    'updated by: IEBV 032120130230pm
    'description: more validation
    If N2Str2IntZero(gconDMIS.Execute("Select count(*) as bilang from csms_ro_det where status = 'W' AND LIVIL = '1' AND REP_or =  '" & lbltech.SelectedItem.SubItems(2) & "' and LINE_NO = '" & TheLineNo & "' AND DETCDE ='" & xJOBCODE & "'").Fields(0).Value) > 0 Then
        MsgBox "Job Is Already In Working Status", vbInformation, "CSMS"
        Exit Sub
    End If
    '--------------


    Dim PATS_Time_IN_AM
    Dim PATS_TIME_OUT_AM
    Dim PATS_TIME_IN_PM
    Dim PATS_TIME_OUT_PM
    Dim PATS_NumHrs
    Dim PATS_STDHRS
    Dim temprs                                         As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * FROM HRMS_ATTEND WHERE EMPNO = " & xTechnician & " AND DATETODAY = " & xTranDate)

    If Not (temprs.EOF Or temprs.BOF) Then
        If Null2String(temprs!INAM) = "" Then PATS_Time_IN_AM = "NULL"
        If Not Null2String(temprs!INAM) = "" Then PATS_Time_IN_AM = N2Str2Null(TimeValue(Null2String(temprs!INAM)))

        If Null2String(temprs!OUTAM) = "" Then PATS_TIME_OUT_AM = "NULL"
        If Not Null2String(temprs!OUTAM) = "" Then PATS_TIME_OUT_AM = N2Str2Null(TimeValue(Null2String(temprs!OUTAM)))

        If Null2String(temprs!InPM) = "" Then PATS_TIME_IN_PM = "NULL"
        If Not Null2String(temprs!InPM) = "" Then PATS_TIME_IN_PM = N2Str2Null(TimeValue(Null2String(temprs!InPM)))

        If Null2String(temprs!InPM) = "" Then PATS_TIME_OUT_PM = "NULL"
        If Not Null2String(temprs!InPM) = "" Then PATS_TIME_OUT_PM = N2Str2Null(TimeValue(Null2String(temprs!InPM)))

        PATS_NumHrs = NumericVal(temprs!TOTALHRSAM) + NumericVal(temprs!TOTALHRSPM)
        PATS_STDHRS = NumericVal(temprs!ActualHrsAM) + NumericVal(temprs!ActualHrsPM)

        SQL_STATEMENT = "Insert into CSMS_JobClock " & _
            " ( TranDate, Technician, Tech_Name, RO_No, ClockIn, JStatus, Status, line_no,ITEMNO,DETCDE, techcode, Time_IN_AM , TIME_OUT_AM , TIME_IN_PM ,TIME_OUT_PM, NumHrs , STDHRS ) " & _
            "   values ( " & xTranDate & _
            ", " & xTechnician & _
            ", " & xTech_Name & _
            ", " & xRO_NO & _
            ", " & xClockIn & _
            ", " & xJStatus & _
            ", " & xStatus & _
            ", '" & labJobItemNo.Caption & _
            "', '" & labJobItemNo.Caption & _
            "', '" & labJobCode & _
            "', '" & TechCode & _
            "', " & PATS_Time_IN_AM & _
            ", " & PATS_TIME_OUT_AM & _
            ", " & PATS_TIME_IN_PM & _
            ", " & PATS_TIME_OUT_PM & _
            ", " & PATS_NumHrs & _
            ", " & PATS_STDHRS & ")"
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = "Insert into CSMS_JobClock " & _
            " ( TranDate, Technician, Tech_Name, RO_No, ClockIn, JStatus, Status, line_no, ITEMNO, DETCDE, techcode) " & _
            " values ( " & xTranDate & _
            ", " & xTechnician & _
            ", " & xTech_Name & _
            ", " & xRO_NO & _
            ", " & xClockIn & _
            ", " & xJStatus & _
            ", " & xStatus & _
            ", '" & labJobItemNo.Caption & _
            "', '" & labJobItemNo.Caption & _
            "', '" & labJobCode & _
            "', '" & TechCode & "')"
        gconDMIS.Execute SQL_STATEMENT
    End If

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("JI", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(Null2String(xTechnician), "EMPNO", "HRMS_EMPINFO"), "", "JOB CODE: " & labJobCode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    MessagePop InfoFriend, "Job Information Updated", "Technician Sucessfully Clock In", 1000

    SQL_STATEMENT = "update HRMS_EmpInfo set  JStatus = 'W' where empno = " & xTechnician & ""
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT------------------------------------------------------------------------------------
        Call NEW_LogAudit("E", "EMPLOYEE INFORMATION", SQL_STATEMENT, FindTransactionID(N2Str2Null(Null2String(xTechnician)), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & Null2String(xRO_NO), "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------------

    gconDMIS.Execute "update CSMS_EmpInfo set  JStatus = 'W' where empno = " & xTechnician & ""
    gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Working',JStatus='W' where RO_No = " & xRO_NO & ""

    SQL_STATEMENT = "Update CSMS_Ro_DEt set status = 'W', DONE = 'W' where LIVIL = '1' AND REP_or = " & xRO_NO & " and LINE_NO = '" & labJobItemNo.Caption & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("JI", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(Null2String(xRO_NO), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & xJOBCODE, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call RefreshTechnicianStatus
    lbltech.Enabled = True
    lblJob4Service.ListItems.Clear
    cmdInCancel.Value = True


    'UPDATED BY: JUN---------------------------------------------------------------
    'DATE UPDATED: 12-06-2008
    'DESCRIPTION: AUTOMATIC REFRESH THE SERVICE COUNTER AFTER THE USER CLOCKING OUT
    'PRE COMMENTED BY PIYA DUE TO MODAL ISSUE AND
'    EVENTRO = LTrim(RTrim(N2Str2Null(lbltech.SelectedItem.SubItems(2))))
    RaiseEvent JOBCLOCKED
    ' YOU CAN SEE CODE IN frmCSMSServiceCounter IF ITS A CALLE
    ' MAKE EVERY FORM INDENPENDET SO THAT YOU CAN UTILIZE IN SERVICE COUNTER OR IN DASH BOARD
    ' PLUS YOU KNOW NA MAN WHEN ACCURRATY TO REFRESH THE JOB
    ' frmCSMSServiceCounter.ViewActiveRO
    'UPDATED BY: JUN---------------------------------------------------------------
    Exit Sub

Error:
    ShowVBError
End Sub

Private Sub cmdInCancel_Click()
    cbojStatus.Enabled = True
    lbltech.Enabled = True

    picClockIn.Visible = False

    'UPDATE BY : MJP 05 15 2008
    lblJob4Service.Enabled = True
    lbltech.Enabled = True
    'UPDATE BY : MJP 05 15 2008
End Sub

Sub BackJobOUT()
    Dim rstmp                                          As New ADODB.Recordset
    Dim rsRO_DET                                       As ADODB.Recordset
    Dim vRONO                                          As String

    If cboReasonClockingout.Text = "Idle Time" Then
        If cboIdleTemplate.Text = "" Then
            ShowIsRequiredMsg ("Reason for Clocking Out")
            cboIdleTemplate.SetFocus
            Exit Sub
        End If
    End If

    If MsgBox("Continue Clock Out", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then
        'lblTech.Enabled = True
        Exit Sub
    End If

    lbltech.Enabled = True
    cbojStatus.Enabled = True

    Dim mdone                                           As String
    Dim xhrsWorked                                      As Double
    Dim xClockOut                                       As String
    Dim xJStatus                                        As String
    Dim xStatus                                         As String
    Dim xReasonForClockOut                              As String
    Dim xRO_NO                                          As String
    Dim xTechnician                                     As String

    xClockOut = N2Str2Null(DTPicker1)
    xhrsWorked = NumericVal(labHours.Caption)
    xRO_NO = N2Str2Null(LTrim(RTrim(lbltech.SelectedItem.SubItems(2))))
    vRONO = LTrim(RTrim(lbltech.SelectedItem.SubItems(2)))
    xTechnician = N2Str2Null(LTrim(RTrim(lbltech.SelectedItem)))

    Dim HRMS_SQL_STATEMENT                             As String

    If cboReasonClockingout.Text = "Lunch Break" Then
        labIdleReason.Caption = "": xJStatus = "'L'"
        gconDMIS.Execute "update CSMS_RepairOrder set STATUS = 'Lunch Break', JStatus = 'L' where RO_No = " & xRO_NO & ""

        SQL_STATEMENT = "update CSMS_ro_det set " & _
                      " STATUS = 'L' where LIVIL = '1' AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                        "' and LINE_NO = '" & labOutItemNo.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT

        HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        gconDMIS.Execute HRMS_SQL_STATEMENT

        gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
    ElseIf cboReasonClockingout.Text = "Break Time" Then
        labIdleReason.Caption = "": xJStatus = "'B'"
        gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Break Time', JStatus = 'B' where RO_No = " & xRO_NO & ""

        SQL_STATEMENT = "update CSMS_ro_det set " & _
                      " STATUS = 'B' where LIVIL = '1' AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                        "' and LINE_NO = '" & labOutItemNo & "'"
        gconDMIS.Execute SQL_STATEMENT

        HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        gconDMIS.Execute HRMS_SQL_STATEMENT

        gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
    ElseIf cboReasonClockingout.Text = "Going Home" Then
        labIdleReason.Caption = "": xJStatus = "'G'"
        gconDMIS.Execute "update CSMS_RepairOrder set STATUS = 'Going Home', JStatus = 'G' where RO_No = " & xRO_NO & ""

        SQL_STATEMENT = "update CSMS_ro_det set " & _
                      " STATUS = 'G' where LIVIL = '1' AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                        "' and LINE_NO = '" & labOutItemNo & "'"
        gconDMIS.Execute SQL_STATEMENT

        HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        gconDMIS.Execute HRMS_SQL_STATEMENT

        gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
    ElseIf cboReasonClockingout.Text = "Idle Time" Then
        gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Idle Time', JStatus = 'I' where RO_No = " & xRO_NO & ""
        xJStatus = "'A'"
        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            xJStatus = "'S'"
            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT
            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        Else
            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
            gconDMIS.Execute HRMS_SQL_STATEMENT

            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
            gconDMIS.Execute "update CSMS_RepairOrder set  tech2 = '" & Repleys(txtidle) & "' where RO_No = " & xRO_NO & ""
        End If

        'Remove Technician from Being Assigned
        SQL_STATEMENT = "update CSMS_ro_det set " & _
            " STATUS = 'I' " & _
            ", Technician = NULL " & _
            ", TechCode = NULL " & _
            ", Done = NULL " & _
            " where LIVIL = '1' " & _
            " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
            "' and LINE_NO = '" & labOutItemNo.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT

        theIdlePic.Visible = False
        lblJob4Service.ListItems.Clear
    ElseIf cboReasonClockingout.Text = "Finish Job" Then
        labIdleReason.Caption = "": xJStatus = "'F'"
        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            xJStatus = "'S'"
            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT

            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        Else
            xJStatus = "'A'"
            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = NULL where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT

            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = NULL where empno = " & xTechnician
        End If

        SQL_STATEMENT = "update CSMS_ro_det set " & _
            " STATUS = 'Q' " & _
            ", Done = 'Y' " & _
            " where LIVIL = '1' " & _
            " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
            "' and LINE_NO = '" & labOutItemNo.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT


        If CheckAllJobsISDone(xRO_NO) = True Then
            gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = '" & labDate.Caption & "', STATUS = 'Finish Job', JStatus = 'F' where RO_No = " & xRO_NO & ""
        Else
            'Status is retained since other Jobs are still on going
        End If
    Else
        gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Idle Time', JStatus = 'I' where RO_No = " & xRO_NO & ""
        xJStatus = "'A'"
        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            xJStatus = "'S'"
            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            gconDMIS.Execute SQL_STATEMENT
            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        Else
            SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
            gconDMIS.Execute SQL_STATEMENT
            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
            gconDMIS.Execute "update CSMS_RepairOrder set  tech2 = '" & txtidle.Text & "' where RO_No = " & xRO_NO & ""
        End If

        'Remove Technician from Being Assigned
        SQL_STATEMENT = "update CSMS_ro_det set " & _
            " STATUS = 'I' " & _
            ", Technician = NULL " & _
            ", TechCode = NULL " & _
            ", Done = NULL" & _
            " where LIVIL = '1' " & _
            " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
            "' and LINE_NO = '" & labOutItemNo.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT

        theIdlePic.Visible = False
        lblJob4Service.ListItems.Clear
    End If
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("JO", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lbltech.SelectedItem.SubItems(2)), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & xJOBCODE & " - " & cboReasonClockingout, "", "")
        Call NEW_LogAudit("JO", "EMPLOYEE INFO", HRMS_SQL_STATEMENT, FindTransactionID(Null2String(xTechnician), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & lbltech.SelectedItem.SubItems(2) & " - " & xJOBCODE, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    xStatus = "'OUT'"
    xReasonForClockOut = N2Str2Null(Trim(cboReasonClockingout.Text) & " - " & Trim(labIdleReason.Caption))
    gconDMIS.Execute "update CSMS_JobClock set " & _
                   " ClockOut = " & xClockOut & "," & _
                   " hrsWorked = " & xhrsWorked & "," & _
                   " JStatus = " & xJStatus & "," & _
                   " Status = " & xStatus & "," & _
                   " ReasonForClockOut = " & xReasonForClockOut & "" & _
                   " where ITEMNO = '" & labOutItemNo.Caption & "' AND RO_No = '" & Trim(lbltech.SelectedItem.SubItems(2)) & "' and Technician = '" & LTrim(RTrim(lbltech.SelectedItem)) & "' and status = 'IN'"

    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND LINE_NO = '" & labOutItemNo & "' AND REP_OR = '" & lbltech.SelectedItem.SubItems(2) & "'")
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        Dim TotalLaborCost                             As Double

        Dim rsTechSalary                               As ADODB.Recordset
        Dim TechCost                                   As Double
        Set rsTechSalary = New ADODB.Recordset
        Set rsTechSalary = gconDMIS.Execute("Select * from CSMS_vw_Technician where empno = " & xTechnician & "")
        If Not rsTechSalary.EOF And Not rsTechSalary.BOF Then
            TechCost = (N2Str2Zero(rsRO_DET!HRSWRK) + NumericVal(xhrsWorked)) * (Round(((N2Str2Zero(rsTechSalary!Monthly) * 12) / 314) / 8, 2))
        Else
            TechCost = 0
        End If

        SQL_STATEMENT = "update CSMS_ro_det set " & _
                      " BACKJOB_HOURS = " & N2Str2Zero(rsRO_DET!BACKJOB_HOURS) & " + " & NumericVal(xhrsWorked) & _
                      " Where LIVIL = '1' AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                        "' and LINE_NO = '" & labOutItemNo.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------------
        Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lbltech.SelectedItem.SubItems(2)), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & xJOBCODE, "", "")
        'NEW LOG AUDIT-----------------------------------------------------------
    End If

    RefreshTechnicianStatus
    MessagePop InfoFriend, "Job Information Updated", "Technician Sucessfully Clock out", 1000

    cmdOutCancel.Value = True

    lblJob4Service.ListItems.Clear
End Sub

Sub set_date_time()
    On Error GoTo ErrorCode
    Dim RS                                 As ADODB.Recordset
    
    Set RS = gconDMIS.Execute("Select getdate() as DateNow, host_name() as PCName")
    If RS!PCNAME <> "SERVER" Or RS!PCNAME <> "DMISSERVER" Then
        Date = RS!DateNow
        Time = RS!DateNow
        LOGTIME = Time
        LOGDATE = Date
    End If
    Set RS = Nothing
     
    Exit Sub
ErrorCode:                    '
  Err.Clear
End Sub

Private Sub cmdOut_Click()
    Call set_date_time
    Call Sleep(2000)
    
    labDate.Caption = Format(Now, "MM/dd/yyyy")
    labtime.Caption = Format(Now, "hh:mm:ss ampm")
    DTPicker2.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    labOut.Caption = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    DTPicker1.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    
    xJOBCODE = RTrim(LTrim(lblJob4Service.ListItems(lblJob4Service.SelectedItem.Index).Text))
    
    'updated by: IEBV 032120130230pm
    'description: more validation
    If N2Str2IntZero(gconDMIS.Execute("Select count(*) as bilang from csms_ro_det where status IN ('L','B','G','I') AND LIVIL = '1' AND REP_or =  '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(2))) & "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'").Fields(0).Value) > 0 Then
        MsgBox "Job Is Already Clock Out", vbInformation, "CSMS"
        Exit Sub
    End If
    If N2Str2IntZero(gconDMIS.Execute("Select count(*) as bilang from csms_ro_det where status = 'Y' AND LIVIL = '1' AND REP_or =  '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(2))) & "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'").Fields(0).Value) > 0 Then
        MsgBox "Job Is Already Finish", vbInformation, "CSMS"
        Exit Sub
    End If
    '--------------
    
    Dim rstmp                                          As New ADODB.Recordset
    Dim rsRO_DET                                       As ADODB.Recordset
    Dim vRONO                                          As String

    If QC_MODULE_ON = "ON" Then
        If CheckIfBackJob(labOutItemNo, LTrim(RTrim(lbltech.SelectedItem.SubItems(2)))) = True Then
            Call BackJobOUT
        Else
            GoTo NORMAL_CLOCKING_OUT
        End If
    Else
NORMAL_CLOCKING_OUT:

        If cboReasonClockingout.Text = "Idle Time" Then
            If cboIdleTemplate.Text = "" Then
                ShowIsRequiredMsg ("Reason for Clocking Out")
                On Error Resume Next
                cboIdleTemplate.SetFocus
                Exit Sub
            End If
        End If

        If MsgBox("Continue Clock Out", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then
            'lblTech.Enabled = True
            Exit Sub
        End If

        lbltech.Enabled = True
        cbojStatus.Enabled = True

        Dim mdone                                       As String
        Dim xhrsWorked                                  As Double
        Dim xClockOut                                   As String
        Dim xJStatus                                    As String
        Dim xStatus                                     As String
        Dim xReasonForClockOut                          As String
        Dim xRO_NO                                      As String
        Dim xTechnician                                 As String
        Dim RSJOBCLOCK                                  As New ADODB.Recordset
        Dim jMIN_WRK                                    As Double
        Dim JHRS_WKR                                    As Double
        Dim rsOUTCHECKER                                As New ADODB.Recordset
        
        
        Set rsOUTCHECKER = gconDMIS.Execute("SELECT CLOCKIN, CLOCKOUT FROM CSMS_JOBCLOCK " & _
            " WHERE RO_NO = '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(2))) & _
            "' AND LINE_NO = '" & labOutItemNo & _
            "' AND CLOCKOUT IS NOT NULL AND DETCDE = '" & xJOBCODE & "'")
        If Not (rsOUTCHECKER.BOF And rsOUTCHECKER.EOF) Then
            If DateDiff("N", Null2String(rsOUTCHECKER!CLOCKIN), DTPicker1) < 0 Then
                MsgBox "Clock out time cannot be less than the clock in time. " & vbCrLf & "kindly check your computer time and date if its correct.", vbExclamation, "Back date not allowed"
                Exit Sub
            End If
        End If
        Set rsOUTCHECKER = Nothing
        
        
        Set RSJOBCLOCK = gconDMIS.Execute("SELECT CLOCKIN,CLOCKOUT FROM CSMS_JOBCLOCK " & _
            " WHERE RO_NO = '" & LTrim(RTrim(lbltech.SelectedItem.SubItems(2))) & _
            "' AND LINE_NO = '" & labOutItemNo & "'AND DETCDE = '" & xJOBCODE & "'")
        If Not (RSJOBCLOCK.BOF And RSJOBCLOCK.EOF) Then
            Do While Not RSJOBCLOCK.EOF
                If Not Null2String(RSJOBCLOCK!CLOCKIN) = "" And Not Null2String(RSJOBCLOCK!CLOCKOUT) = "" Then
                    jMIN_WRK = DateDiff("N", Null2String(RSJOBCLOCK!CLOCKIN), Null2String(RSJOBCLOCK!CLOCKOUT))

                    JHRS_WKR = JHRS_WKR + (jMIN_WRK / 60)
                End If

                RSJOBCLOCK.MoveNext
            Loop
        End If
        Set RSJOBCLOCK = Nothing


        xClockOut = N2Str2Null(DTPicker1)
        xhrsWorked = NumericVal(labHours.Caption) + Format(JHRS_WKR, MAXIMUM_DIGIT)

        xRO_NO = N2Str2Null(LTrim(RTrim(lbltech.SelectedItem.SubItems(2))))
        vRONO = LTrim(RTrim(lbltech.SelectedItem.SubItems(2)))
        xTechnician = N2Str2Null(LTrim(RTrim(lbltech.SelectedItem)))

        Dim HRMS_SQL_STATEMENT                         As String

        If cboReasonClockingout.Text = "Lunch Break" Then
            labIdleReason.Caption = "": xJStatus = "'L'"
            gconDMIS.Execute "update CSMS_RepairOrder set STATUS = 'Lunch Break', JStatus = 'L' where RO_No = " & xRO_NO & ""

            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " STATUS = 'L' where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"
            gconDMIS.Execute SQL_STATEMENT


            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set " & _
                " JStatus = " & xJStatus & _
                ", ASSIGNEDRO = " & xRO_NO & _
                " where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT
            
            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        ElseIf cboReasonClockingout.Text = "Break Time" Then
            labIdleReason.Caption = "": xJStatus = "'B'"
            gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Break Time', JStatus = 'B' where RO_No = " & xRO_NO & ""

            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " STATUS = 'B' " & _
                " where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' AND LINE_NO = '" & labOutItemNo & "' AND DETCDE = '" & xJOBCODE & "'"
            gconDMIS.Execute SQL_STATEMENT

            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT

            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        ElseIf cboReasonClockingout.Text = "Going Home" Then
            labIdleReason.Caption = "": xJStatus = "'G'"
            gconDMIS.Execute "update CSMS_RepairOrder set STATUS = 'Going Home', JStatus = 'G' where RO_No = " & xRO_NO & ""

            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " STATUS = 'G' " & _
                " where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' and LINE_NO = '" & labOutItemNo & "' AND DETCDE = '" & xJOBCODE & "'"
            gconDMIS.Execute SQL_STATEMENT

            HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            gconDMIS.Execute HRMS_SQL_STATEMENT

            gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
        ElseIf cboReasonClockingout.Text = "Idle Time" Then
            gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Idle Time', JStatus = 'I' where RO_No = " & xRO_NO & ""
            xJStatus = "'A'"
            Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")

            If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
                xJStatus = "'S'"
                HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
                gconDMIS.Execute HRMS_SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            Else
                HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
                gconDMIS.Execute HRMS_SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
                gconDMIS.Execute "update CSMS_RepairOrder set  tech2 = '" & Repleys(txtidle) & "' where RO_No = " & xRO_NO & ""
            End If

            'Remove Technician from Being Assigned
            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " STATUS = 'I' " & _
                ", Technician = NULL " & _
                ", TechCode = NULL " & _
                ", Done = NULL" & _
                " where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' AND LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"

            gconDMIS.Execute SQL_STATEMENT


            theIdlePic.Visible = False
            lblJob4Service.ListItems.Clear

        ElseIf cboReasonClockingout.Text = "Finish Job" Then
            If CheckProductiveTimeAboveEqualMinimum(labIn, labOut, LTrim(RTrim(lbltech.SelectedItem.SubItems(2))), labOutItemNo) = False Then
                MsgBox "Hours Work is Less than the 25% completion of the Job. " & vbCrLf & "System Will not allowed this Clock Out", vbExclamation, "CSMS"
                lbltech.Enabled = False
                cbojStatus.Enabled = False
                '<TO DO TASK >
                Exit Sub
            End If

            labIdleReason.Caption = "": xJStatus = "'F'"
            Set rsRO_DET = New ADODB.Recordset
            Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")
            If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
                xJStatus = "'S'"
                HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
                gconDMIS.Execute HRMS_SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            
            Else
                xJStatus = "'A'"
                HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = NULL where empno = " & xTechnician
                gconDMIS.Execute HRMS_SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = NULL where empno = " & xTechnician
            End If
               
            If QC_MODULE_ON = "ON" Then
                SQL_STATEMENT = "update CSMS_ro_det set " & _
                    " STATUS = 'Q' " & _
                    ", Done = 'Y' " & _
                    " where LIVIL = '1' " & _
                    " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                    "' AND LINE_NO = '" & labOutItemNo.Caption & "'"
                gconDMIS.Execute SQL_STATEMENT
            Else
                SQL_STATEMENT = "update CSMS_ro_det set " & _
                    " STATUS = 'Y' " & _
                    ", Done = 'Y' " & _
                    " where LIVIL = '1' " & _
                    " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                    "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"
                gconDMIS.Execute SQL_STATEMENT
                 ' update BTT : to kill CICO
                 
                Dim nard As New ADODB.Recordset
                Dim Again As New ADODB.Recordset
                Set Again = gconDMIS.Execute("Select empno,jstatus from HRMS_empinfo where jstatus='A' and empno =" & xTechnician & "")
                If Not (Again.EOF And Again.BOF) Then
                   'expected that is already out
                   Set nard = gconDMIS.Execute("Select rep_or,techcode,done,status from CSMS_RO_DET where livil = '1' and rep_or='" & lbltech.SelectedItem.SubItems(2) & "' and done = 'W' and techcode='" & TechCode & "' and line_no='" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'")
                   If Not (nard.EOF And nard.BOF) Then
                        SQL_STATEMENT = "update CSMS_ro_det set " & _
                           " STATUS = 'Y' " & _
                           ", Done = 'Y' " & _
                           " where LIVIL = '1' " & _
                           " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                           "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"
                       gconDMIS.Execute SQL_STATEMENT
                   End If
                End If
                Set nard = Nothing
                Set Again = Nothing
            End If
            
            If CheckAllJobsISDone(xRO_NO) = True Then
                gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = '" & labDate.Caption & "', STATUS = 'Finish Job', JStatus = 'F' where RO_No = " & xRO_NO & ""
            Else
                'Status is retained since other Jobs are still on going
            End If
        Else
            gconDMIS.Execute "update CSMS_RepairOrder set  STATUS = 'Idle Time', JStatus = 'I' where RO_No = " & xRO_NO & ""
            xJStatus = "'A'"
            Set rsRO_DET = New ADODB.Recordset
            Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND TECHCODE = '" & RTrim(LTrim(lbltech.SelectedItem.SubItems(6))) & "' AND REP_OR = " & xRO_NO & " AND (DONE = 'N' OR DONE IS NULL)")
            If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
                xJStatus = "'S'"
                HRMS_SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
                gconDMIS.Execute SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ",ASSIGNEDRO = " & xRO_NO & " where empno = " & xTechnician
            Else
                SQL_STATEMENT = "update HRMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
                gconDMIS.Execute SQL_STATEMENT
                gconDMIS.Execute "update CSMS_EmpInfo set JStatus = " & xJStatus & ", ASSIGNEDRO = NULL where empno = " & xTechnician & ""
                gconDMIS.Execute "update CSMS_RepairOrder set  tech2 = '" & txtidle.Text & "' where RO_No = " & xRO_NO & ""
            End If

            'Remove Technician from Being Assigned
            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " STATUS = 'I' " & _
                ", Technician = NULL " & _
                ", TechCode = NULL " & _
                ", Done = NULL" & _
                " where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' AND LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"
            gconDMIS.Execute SQL_STATEMENT

            theIdlePic.Visible = False
            lblJob4Service.ListItems.Clear
        End If
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("JO", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lbltech.SelectedItem.SubItems(2)), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & xJOBCODE & " - " & cboReasonClockingout, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("JO", "EMPLOYEE INFO", HRMS_SQL_STATEMENT, FindTransactionID(Null2String(xTechnician), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & lbltech.SelectedItem.SubItems(2) & " - " & xJOBCODE, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        xStatus = "'OUT'"
        xReasonForClockOut = N2Str2Null(Trim(cboReasonClockingout.Text) & " - " & Trim(labIdleReason.Caption))
        gconDMIS.Execute "update CSMS_JobClock set " & _
            " ClockOut = " & xClockOut & "," & _
            " hrsWorked = " & labHours & "," & _
            " JStatus = " & xJStatus & "," & _
            " Status = " & xStatus & "," & _
            " ReasonForClockOut = " & xReasonForClockOut & "" & _
            " where ITEMNO = '" & labOutItemNo.Caption & _
            "' AND RO_No = '" & Trim(lbltech.SelectedItem.SubItems(2)) & "' and Technician = '" & LTrim(RTrim(lbltech.SelectedItem)) & "' and status = 'IN' AND DETCDE = '" & xJOBCODE & "'"

        'gconDMIS.Execute "update CSMS_JobClock set " & _
         '               " ClockOut = " & xClockOut & "," & _
         '               " hrsWorked = " & xhrsWorked & "," & _
         '               " JStatus = " & xJStatus & "," & _
         '               " Status = " & xStatus & "," & _
         '               " ReasonForClockOut = " & xReasonForClockOut & "" & _
         '               " where ITEMNO = '" & labOutItemNo & "' AND RO_No = '" & Trim(lblTech.SelectedItem.SubItems(2)) & "' and Technician = '" & lblTech.SelectedItem & "' and status='IN'"

        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_DET where LIVIL = '1' AND LINE_NO = '" & labOutItemNo & "' AND REP_OR = '" & lbltech.SelectedItem.SubItems(2) & "' AND DETCDE = '" & xJOBCODE & "'")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            Dim TotalLaborCost                         As Double

            Dim rsTechSalary                           As New ADODB.Recordset
            Dim TechCost                               As Double
            Set rsTechSalary = gconDMIS.Execute("Select * from CSMS_vw_Technician where empno = " & xTechnician & "")
            If Not rsTechSalary.EOF And Not rsTechSalary.BOF Then
                'Updated by JJE 08/16/2012 10:06AM
                'Fixing DETCOST Computation
                'TechCost = (N2Str2Zero(rsRO_DET!HRSWRK) + NumericVal(xhrsWorked)) * (Round(((N2Str2Zero(rsTechSalary!Monthly) * 12) / 314) / 8, 2))
                TechCost = (NumericVal(xhrsWorked)) * (Round(((N2Str2Zero(rsTechSalary!Monthly) * 12) / 314) / 8, 2))
            Else
                TechCost = 0
            End If
            
            'Updated by JJE 08/16/2012 10:06AM
            'Fixing HRSWRK Computation
            '"HRSWRK = " & N2Str2Zero(rsRO_DET!HRSWRK) & " + " & NumericVal(labHours) & ","
            SQL_STATEMENT = "update CSMS_ro_det set " & _
                " HRSWRK = " & NumericVal(xhrsWorked) & "," & _
                " DETCOST = " & TechCost & _
                " where LIVIL = '1' " & _
                " AND Rep_or = '" & lbltech.SelectedItem.SubItems(2) & _
                "' and LINE_NO = '" & labOutItemNo.Caption & "' AND DETCDE = '" & xJOBCODE & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------------
                Call NEW_LogAudit("E", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(lbltech.SelectedItem.SubItems(2)), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & xJOBCODE, "", "")
            'NEW LOG AUDIT-----------------------------------------------------------
        End If

        'COMMENT BY : MJP 011908 0200PM
            RefreshTechnicianStatus
        'COMMENT BY : MJP 011908 0200PM
        
        'UPDATE BY : MJP 011908 0200PM
            Call cbojStatus_Click
        'UPDATE BY : MJP 011908 0200PM
        MessagePop InfoFriend, "Job Information Updated", "Technician Sucessfully Clock out", 1000
        
        cmdOutCancel.Value = True

        lblJob4Service.ListItems.Clear
        EVENTRO = lbltech.SelectedItem.SubItems(2)
        'UPDATED BY: JUN---------------------------------------------------------------
        'DATE UPDATED: 12-06-2008
        'DESCRIPTION: AUTOMATIC REFRESH THE SERVICE COUNTER AFTER THE USER CLOCKING OUT
        'PRE COMMENTED BY PIYA DUE TO MODAL ISSUE AND
        'RaiseEvent JOBCHANGED(lblTech.SelectedItem.SubItems(2))
        RaiseEvent JOBCLOCKED
        ' YOU CAN SEE CODE IN frmCSMSServiceCounter IF ITS A CALLE
        ' MAKE EVERY FORM INDENPENDET SO THAT YOU CAN UTILIZE IN SERVICE COUNTER OR IN DASH BOARD
        ' PLUS YOU KNOW NA MAN WHEN ACCURRATY TO REFRESH THE JOB
        ' frmCSMSServiceCounter.ViewActiveRO
        'UPDATED BY: JUN---------------------------------------------------------------
    End If
End Sub

Function CheckProductiveTimeAboveEqualMinimum(vCLOCKIN_TIME As Date, vCLOCKOUT_TIME As Date, vRONO As String, vlineNo As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Dim RSJOBCLOCK                                     As New ADODB.Recordset
    Dim kMIN_WRK                                       As Double
    Dim kHRS_WRK                                       As Double
    Dim jMIN_WRK                                       As Double
    Dim JHRS_WKR                                       As Double

    Set RSJOBCLOCK = gconDMIS.Execute("SELECT * FROM CSMS_JOBCLOCK WHERE RO_NO = '" & vRONO & "' AND LINE_NO = '" & vlineNo & "'")
    If Not (RSJOBCLOCK.BOF And RSJOBCLOCK.EOF) Then
        Do While Not RSJOBCLOCK.EOF
            If Not Null2String(RSJOBCLOCK!CLOCKIN) = "" And Not Null2String(RSJOBCLOCK!CLOCKOUT) = "" Then
                jMIN_WRK = DateDiff("N", Null2String(RSJOBCLOCK!CLOCKIN), Null2String(RSJOBCLOCK!CLOCKOUT))

                JHRS_WKR = JHRS_WKR + (jMIN_WRK / 60)
            End If

            RSJOBCLOCK.MoveNext
        Loop
    End If
    Set RSJOBCLOCK = Nothing

    Set rstmp = gconDMIS.Execute("SELECT DET_HRS FROM CSMS_RO_DET WHERE LIVIL = '1' AND REP_OR = '" & vRONO & "' AND LINE_NO = '" & vlineNo & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        kMIN_WRK = DateDiff("N", vCLOCKIN_TIME, vCLOCKOUT_TIME)
        kHRS_WRK = kMIN_WRK / 60

        If (NumericVal(rstmp!DET_HRS) * 0.25) <= (kHRS_WRK + Format(JHRS_WKR, MAXIMUM_DIGIT)) Then
            CheckProductiveTimeAboveEqualMinimum = True
        Else
            CheckProductiveTimeAboveEqualMinimum = False
        End If
    End If
    Set rstmp = Nothing
End Function

Private Sub cmdOutCancel_Click()
    cbojStatus.Enabled = True
    lbltech.Enabled = True
    picClockOut.Visible = False

    'UPDATE BY : MJP 05 15 2008
    lblJob4Service.Enabled = True
    'UPDATE BY : MJP 05 15 2008
End Sub

Private Sub Command2_Click()
    If MsgBox("Refresh Selected Idle Status?", vbQuestion + vbYesNo, "Re-Select Status") = vbYes Then
        cboIdleTemplate.Enabled = True
        txtidle.Text = ""
        'txtidle.Enabled = False
        VBComBoBoxDroppedDown cboIdleTemplate
    End If
End Sub

Private Sub Command3_Click()
    Call FillDTR
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    EVENTRO = ""
    labEmployeein.Caption = ""
    labEmployeeOut.Caption = ""
    'Set rsUpload = New ADODB.Recordset
    'DON'T INISITATE AN OBJECT WHEN YOUR DIRECT CASTING IN NEXT LINE
    'GCON.EXECUTE ALWAYS RETURNS RECORDSET' TIPS PRESS ALT + I IN GCON.EXECUTE IT WILL SAY RECORDSET
    'Set rsUpload = gconDMIS.Execute("select  description,Inout,Out from CSMS_JobStatus Where ([Out] <> 'O' OR [OUT] IS NULL) order by description asc")
    
    'UPDATED BY: JUN-----------------------------------------------------------------------------------
    'DATE UPDATED: 03-09-2009
    'DESCRIPTION: UPDATE THE JSTATUS TO AVAILABLE  IF THE THE ASSIGNEDRO IS NULL
    '             UPDATE FOR TICKET NO: HGC-12727
     'HRMS_EMPINFO
     gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET JSTATUS = 'A' WHERE ASSIGNEDRO IS NULL AND IS_TECHNICIAN = '1'")
         
     'CSMS_EMPINFO
     gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS = 'A' WHERE ASSIGNEDRO IS NULL AND IS_TECHNICIAN = '1'")
    'UPDATED BY: JUN-----------------------------------------------------------------------------------
    
    'UPDATED BY: JUN-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DATE UPDATED: 06-01-2009
    'DESCRIPTION: CHECK IF THE STATUS OF THE ASSIGNED RO IF BILLED/RELEASE/ OR ALREADY FINISHED JOB
    '             TCN: 12873-(HGC)
     Dim rsBILLED As ADODB.Recordset
     Dim rsREPAIR_ORDER As ADODB.Recordset
     Dim rsHRMS_EMPNO As ADODB.Recordset
     Set rsBILLED = gconDMIS.Execute("SELECT ASSIGNEDRO,EMPNO FROM CSMS_vw_Technician WHERE ASSIGNEDRO IS NOT NULL")
     If Not rsBILLED.EOF And Not rsBILLED.BOF Then
        Do While Not rsBILLED.EOF
            Set rsREPAIR_ORDER = gconDMIS.Execute("SELECT STATUS FROM CSMS_REPAIRORDER WHERE RO_NO = '" & Null2String(rsBILLED!assignedro) & "' AND STATUS IN('Billed','Finish Job','Released')")
            If Not rsREPAIR_ORDER.EOF And Not rsREPAIR_ORDER.BOF Then
                Set rsHRMS_EMPNO = gconDMIS.Execute("SELECT * FROM HRMS_EMPINFO WHERE EMPNO = '" & Null2String(rsBILLED!EMPNO) & "' AND IS_TECHNICIAN = '1'")
                If Not rsHRMS_EMPNO.EOF And Not rsHRMS_EMPNO.BOF Then
                    gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = NULL , JSTATUS = 'A' WHERE EMPNO = '" & Null2String(rsBILLED!EMPNO) & "' AND IS_TECHNICIAN = '1'")
                Else
                    gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET ASSIGNEDRO = NULL , JSTATUS = 'A' WHERE EMPNO = '" & Null2String(rsBILLED!EMPNO) & "' AND IS_TECHNICIAN = '1'")
                End If
            Else
                'RETAIN STATUS
            End If
            rsBILLED.MoveNext
        Loop
     End If
     Set rsBILLED = Nothing
     Set rsHRMS_EMPNO = Nothing
     Set rsREPAIR_ORDER = Nothing
    'UPDATED BY: JUN-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    
    
    Set RSUPLOAD = gconDMIS.Execute("SELECT  DESCRIPTION FROM CSMS_JOBSTATUS WHERE ([OUT] <> 'O' OR [OUT] IS NULL) ORDER BY DESCRIPTION ASC")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        cbojStatus.Clear
        cbojStatus.AddItem "All"
        Do Until RSUPLOAD.EOF
            cbojStatus.AddItem Null2String(RSUPLOAD![Description])
            RSUPLOAD.MoveNext
        Loop
    End If
    cbojStatus.Text = "All"
    fillcbomonth cboFilter_month: cboFilter_month.AddItem "ALL", 0
    'FillCboMoreYear cboFilter_Yearly: cboFilter_Yearly.AddItem "ALL"

    FillCboMoreYear cboFilter_Yearly: cboFilter_Yearly.AddItem "ALL"

    'LOAD SAVED SETTINGS: IF THERE IS OTHER WISE DEFAULT CURRENT MONTH AND CURRENT YEAR
    cboFilter_month.Text = GetSetting("DMIS 2.0", "CSMS-CLOCKINCLOCKOUT", "DTR-FILTER-MONTH", MonthName(Month(Now)))
    cboFilter_Yearly.Text = GetSetting("DMIS 2.0", "CSMS-CLOCKINCLOCKOUT", "DTR-FILTER-YEAR", Year(Now))

    Call FillReasonClockOut
    Call RefreshTechnicianStatus
    theIdlePic.Visible = False

    If lbltech.ListItems.Count > 0 Then
        lblTech_ItemClick lbltech.SelectedItem
        lbltech.SelectedItem.EnsureVisible

    End If
    '
    Call cbojStatus_Click
    'Call checkIFFinish
    'Call lblTech_Click:NO NEED
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'If EVENTRO <> "" Then
    '    RaiseEvent FORMCLOSED
    'End If

    'SAVE SETTINGS: IF THERE IS OTHER WISE DEFAULT CURRENT MONTH AND CURRENT YEAR
    SaveSetting "DMIS 2.0", "CSMS-CLOCKINCLOCKOUT", "DTR-FILTER-MONTH", cboFilter_month
    SaveSetting "DMIS 2.0", "CSMS-CLOCKINCLOCKOUT", "DTR-FILTER-YEAR", cboFilter_Yearly
End Sub

Private Sub lblJob4Service_DblClick()
    If lblJob4Service.ListItems.Count = 0 Then Exit Sub

    Dim Index                                          As Integer
    Index = lblJob4Service.SelectedItem.Index
    If lblJob4Service.ListItems(Index).ListSubItems(5).Text = "Y" Then
        MsgBox "Job Is Already Finish", vbInformation, "CSMS"
        Exit Sub
    End If

    labJobItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    labOutItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    TheLineNo = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    cbojStatus.Enabled = False
    lbltech.Enabled = False
    Call lblJob4Service_KeyPress(13)
End Sub

Private Sub lblJob4Service_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim Index                                          As Integer
    Index = lblJob4Service.SelectedItem.Index

    labJobItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    labOutItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))

    TheLineNo = RTrim(LTrim(lblJob4Service.SelectedItem.SubItems(3)))
End Sub

Private Sub lblJob4Service_KeyPress(KeyAscii As Integer)
    Dim Index                                          As Integer

    If lblJob4Service.ListItems.Count = 0 Then Exit Sub

    Index = lblJob4Service.SelectedItem.Index

    labJobItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    labOutItemNo.Caption = RTrim(LTrim(lblJob4Service.ListItems(Index).ListSubItems(3)))
    xJOBCODE = RTrim(LTrim(lblJob4Service.ListItems(Index).Text))
    TheLineNo = labJobItemNo.Caption

    If KeyAscii = 13 And Me.ActiveControl.Name = "lblJob4Service" Then
        If lblJob4Service.ListItems.Count = 0 Then Exit Sub

        If StrComp(Trim(theRo), "") = 0 Then
            MsgBox "No Ro Available!", vbExclamation, "Warning"
            Exit Sub
        End If

        If IfROIsFinish(theRo) = True Then Exit Sub

        If lbltech.SelectedItem.SubItems(3) = "Available" Then
            lbltech.SelectedItem.SubItems(3) = "Finish Job"
            lbltech.SelectedItem.SubItems(5) = "I"
        End If

        lbltech.Enabled = False
        Dim rsRO_DET, rsRO_DET2                        As ADODB.Recordset

        Set rsRO_DET = New ADODB.Recordset
        Set rsRO_DET = gconDMIS.Execute("Select * from CSMS_RO_Det where LIVIL = '1' AND REP_OR = " & N2Str2Null(labRO.Caption) & " and ltrim(rtrim(LINE_NO))= '" & labJobItemNo.Caption & "'")
        If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
            'labJobItemNo.Caption = Null2String(rsRO_DET!LINE_NO)
            labJobCode.Caption = Null2String(rsRO_DET!DETCDE)
            labJobDesc.Caption = Null2String(rsRO_DET!DETDSC)
            If Null2String(rsRO_DET!DONE) = "Y" Then Exit Sub
            If Null2String(rsRO_DET!DONE) = "N" Then
                Set rsRO_DET2 = New ADODB.Recordset
                Set rsRO_DET2 = gconDMIS.Execute("Select * from CSMS_RO_Det where LIVIL = '1' AND REP_OR = " & N2Str2Null(labRO.Caption) & " AND DONE = 'W' AND ID <> " & rsRO_DET!ID)
                If Not rsRO_DET2.EOF And Not rsRO_DET2.BOF Then
                    rsRO_DET2.MoveFirst
                    Do While Not rsRO_DET2.EOF
                        If LTrim(RTrim(Null2String(rsRO_DET2!TechCode))) = LTrim(RTrim(lbltech.SelectedItem.SubItems(6))) Then
                            'If Null2String(rsRO_DET2!TechCode) = LTrim(RTrim(lblTech.SelectedItem.SubItems(6))) Then
                            MsgBox "Another Job assigned to this Technician is still in progress", vbInformation, "CSMS"
                            lbltech.Enabled = True
                            Exit Sub
                        End If
                        rsRO_DET2.MoveNext
                    Loop
                End If
            End If
        End If

        If lbltech.SelectedItem.SubItems(3) <> "Available" Then
            If lbltech.SelectedItem.SubItems(5) = "I" Then
                labEmployeein.Caption = LTrim(RTrim(lbltech.SelectedItem.SubItems(4)))
                'UPDATE BY : MJP 05 15 2008
                lblJob4Service.Enabled = False
                'UPDATE BY : MJP 05 15 2008
                picClockIn.ZOrder 0
                picClockIn.Visible = True
                picClockOut.Visible = False
                If cmdIn.Enabled = True Then
                    cmdIn.SetFocus
                End If
            Else
                'labOutItemNo = labJobItemNo.Caption
                Set RSUPLOAD = New ADODB.Recordset
                Set RSUPLOAD = gconDMIS.Execute("select * from CSMS_JobClock where rtrim(ltrim(itemNO)) = '" & labJobItemNo.Caption & "' AND RO_No = '" & lbltech.SelectedItem.SubItems(2) & "' and Technician = '" & LTrim(RTrim(lbltech.SelectedItem)) & "' and status = 'IN' order by ID desc")
                'Set rsUpload = gconDMIS.Execute("select * from CSMS_JobClock where itemNO = '" & lblJob4Service.SelectedItem.SubItems(3) & "' AND RO_No = '" & lblTech.SelectedItem.SubItems(2) & "' and Technician = '" & lblTech.SelectedItem & "' and status = 'IN' order by ID desc")
                'Set rsUpload = gconDMIS.Execute("select * from CSMS_JobClock where line_NO = '" & lblJob4Service.SelectedItem.SubItems(3) & "' AND RO_No = '" & lblTech.SelectedItem.SubItems(2) & "' and Technician = '" & lblTech.SelectedItem & "' and status='IN' order by ID desc")
                If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
                    labIn.Caption = Trim(Null2String(RSUPLOAD![CLOCKIN]))
                    DTClockIn.Value = Trim(Null2String(RSUPLOAD![CLOCKIN]))
                    labHours.Caption = Round(DateDiff("n", DTClockIn, Now) / 60, 2)
                    'labOutItemNo = RTrim(LTrim((Null2String(rsUpload![itemno]))))
                End If

                labEmployeeOut.Caption = lbltech.SelectedItem.SubItems(4)
                picClockIn.Visible = False
                picClockOut.Visible = True
                picClockOut.ZOrder 0

                If cmdOut.Enabled = True Then
                    Call FillReasonClockOut
                    'cboReasonClockingout.SetFocus
                    VBComBoBoxDroppedDown cboReasonClockingout
                End If

            End If
            '            'UPDATE BY : MJP 05 15 2008
            '                lblJob4Service.Enabled = False
            '            'UPDATE BY : MJP 05 15 2008
        End If

    End If
End Sub

Private Sub lblTech_DblClick()
    On Error Resume Next
    lblJob4Service.SetFocus
    If lblJob4Service.ListItems.Count > 0 Then
        lblJob4Service.ListItems(1).EnsureVisible
        'lblJob4Service.ListItems(1).Selected = True
    End If
End Sub

Private Sub lblTech_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'axp when not items are there then it will create bug
    If lbltech.SelectedItem Is Nothing Then Exit Sub
    TechCode = LTrim(RTrim(lbltech.ListItems.Item(lbltech.SelectedItem.Index)))
    TechCode = LTrim(RTrim(lbltech.SelectedItem.SubItems(6)))
    theRo = Trim(lbltech.SelectedItem.SubItems(2))
    TheEmpNO = lbltech.SelectedItem.SubItems(1)
    '   not needed
    '    If StrComp(Trim(theRo), "") = 0 Then
    '        TrapNoRO.Visible = True
    '        lblJob4Service.Enabled = False
    '    Else
    '        TrapNoRO.Visible = False
    '        lblJob4Service.Enabled = True
    '    End If

    Call ViewRoDetail
    Call FillDTR
End Sub

Private Sub lblTech_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        lblTech_DblClick
    End If
End Sub

Private Sub lstDTR_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstDTR
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

Private Sub Timer1_Timer()
    If picClockIn.Visible = False Then
        labDate.Caption = Format(Now, "MM/dd/yyyy")
        labtime.Caption = Format(Now, "hh:mm:ss ampm")
        DTPicker2.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    End If
End Sub

Private Sub Timer2_Timer()
    labOut.Caption = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    DTPicker1.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    'If sw <= 0 Then
    '    DTClockIn.Value = Format(Now, "MM/dd/yyyy  hh:mm:ss ampm")
    'End If
End Sub

Private Sub txtidle_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub
