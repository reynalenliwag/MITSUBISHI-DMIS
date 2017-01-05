VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "COBB8E~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSViewRO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Repair Order Details"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   HasDC           =   0   'False
   Icon            =   "FrmViewRO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   2865
      TabIndex        =   59
      Top             =   2430
      Width           =   6045
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8E3E3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   6015
         TabIndex        =   60
         Top             =   3630
         Width           =   6045
      End
      Begin MSComctlLib.ListView techlist 
         Height          =   2715
         Left            =   90
         TabIndex        =   64
         Top             =   390
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   4789
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
         MousePointer    =   99
         MouseIcon       =   "FrmViewRO.frx":05CA
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmoNo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Technician"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Assigned RO"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Name"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "inout"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Techcode"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton Change 
         Caption         =   "Change Technician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         MouseIcon       =   "FrmViewRO.frx":072C
         MousePointer    =   99  'Custom
         TabIndex        =   63
         Top             =   3180
         Width           =   1875
      End
      Begin VB.CommandButton Assigned 
         Caption         =   "Assign Technician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         MouseIcon       =   "FrmViewRO.frx":0A36
         MousePointer    =   99  'Custom
         TabIndex        =   62
         ToolTipText     =   "Assign Technician"
         Top             =   3180
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4770
         MouseIcon       =   "FrmViewRO.frx":0D40
         MousePointer    =   99  'Custom
         TabIndex        =   61
         ToolTipText     =   "Close "
         Top             =   3180
         Width           =   1155
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   6045
         _Version        =   655364
         _ExtentX        =   10663
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "SELECT TECHNICIAN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   1620
      ScaleHeight     =   525
      ScaleWidth      =   9915
      TabIndex        =   77
      Top             =   8010
      Width           =   9915
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8520
         MouseIcon       =   "FrmViewRO.frx":104A
         MousePointer    =   99  'Custom
         TabIndex        =   78
         ToolTipText     =   "Close Window"
         Top             =   30
         Width           =   1335
      End
      Begin VB.CommandButton cmdChangestat 
         Caption         =   "Print RO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7080
         MouseIcon       =   "FrmViewRO.frx":1354
         MousePointer    =   99  'Custom
         TabIndex        =   79
         ToolTipText     =   "Change Status"
         Top             =   30
         Width           =   1455
      End
      Begin VB.CommandButton cmdJobClock 
         Caption         =   "&Job Clock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5640
         MouseIcon       =   "FrmViewRO.frx":165E
         MousePointer    =   99  'Custom
         TabIndex        =   80
         ToolTipText     =   "View Job Clock"
         Top             =   30
         Width           =   1455
      End
      Begin VB.CommandButton cmdChangeTech 
         Caption         =   "Change &Technician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3900
         MouseIcon       =   "FrmViewRO.frx":1968
         MousePointer    =   99  'Custom
         TabIndex        =   81
         ToolTipText     =   "Change Technician"
         Top             =   30
         Width           =   1755
      End
      Begin VB.CommandButton cmdAssignedTech 
         Caption         =   "&Assign Technician"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2220
         MouseIcon       =   "FrmViewRO.frx":1C72
         MousePointer    =   99  'Custom
         TabIndex        =   82
         ToolTipText     =   "Assign Technician"
         Top             =   30
         Width           =   1695
      End
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1605
      Left            =   3330
      TabIndex        =   76
      Top             =   8820
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   2831
      View            =   3
      Arrange         =   2
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
         Text            =   "NO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ClockIn"
         Object.Width           =   2540
      EndProperty
   End
   Begin Crystal.CrystalReport rptRepairOrder 
      Left            =   210
      Top             =   8070
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10110
      Top             =   480
   End
   Begin VB.TextBox txtApptDate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   6000
      Width           =   2085
   End
   Begin VB.TextBox txtAdvisor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   7200
      Width           =   2085
   End
   Begin VB.TextBox txtPromise 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   6570
      Width           =   2085
   End
   Begin VB.TextBox txtTech3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1350
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   9180
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtTech2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1350
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   8970
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtTech1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   7980
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Frame Frame4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2505
      Left            =   2310
      TabIndex        =   6
      Top             =   5490
      Width           =   9195
      Begin VB.TextBox txtVatMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtDiscMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtRateMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0.00"
         Top             =   1290
         Width           =   975
      End
      Begin VB.TextBox txtWarMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtSalesMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtCompMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtEstMat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0.00"
         Top             =   1290
         Width           =   1125
      End
      Begin VB.TextBox txtVatTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtVatAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtVatParts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtVatLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtDiscTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtDiscAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtDiscParts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtDiscLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtWarLaborTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtWarLaborAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtWarParts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtWarLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4740
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtSalesTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtSalesAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtSalesParts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtSalesLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtCompTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtCompAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtCompPart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtCompLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2070
         Width           =   1125
      End
      Begin VB.TextBox txtEstAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1680
         Width           =   1125
      End
      Begin VB.TextBox txtEstParts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   900
         Width           =   1125
      End
      Begin VB.TextBox txtEstLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   510
         Width           =   1125
      End
      Begin VB.TextBox txtRateLabor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   510
         Width           =   975
      End
      Begin VB.TextBox txtRateparts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtRateAces 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5910
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Materials"
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
         Left            =   405
         TabIndex        =   66
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT"
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
         Left            =   8730
         TabIndex        =   44
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount"
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
         Left            =   7260
         TabIndex        =   43
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty"
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
         Left            =   5130
         TabIndex        =   42
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales"
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
         Left            =   4230
         TabIndex        =   41
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company"
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
         Left            =   2760
         TabIndex        =   40
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   1680
         TabIndex        =   39
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
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
         Left            =   615
         TabIndex        =   38
         Top             =   2100
         Width           =   555
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1740
         Width           =   1050
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parts"
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
         Left            =   735
         TabIndex        =   36
         Top             =   960
         Width           =   435
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Labor"
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
         Left            =   690
         TabIndex        =   35
         Top             =   570
         Width           =   480
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.Rate"
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
         Left            =   6090
         TabIndex        =   34
         Top             =   270
         Width           =   750
      End
   End
   Begin XtremeSuiteControls.TabControl SSTab1 
      Height          =   4095
      Left            =   60
      TabIndex        =   83
      Top             =   1320
      Width           =   11445
      _Version        =   655364
      _ExtentX        =   20188
      _ExtentY        =   7223
      _StockProps     =   64
      Appearance      =   3
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.FixedTabWidth=   130
      ItemCount       =   5
      Item(0).Caption =   "RO Jobs"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lblJob4Service"
      Item(1).Caption =   "PMS Jobs"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lstPMSDet"
      Item(2).Caption =   "Issued Parts"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "ListParts"
      Item(3).Caption =   "Issued Materials"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "ListMaterial"
      Item(4).Caption =   "Issued Accessories"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "ListAccessories"
      Begin MSComctlLib.ListView lblJob4Service 
         Height          =   3555
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6271
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
         MouseIcon       =   "FrmViewRO.frx":1F7C
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "NO"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Jobtype"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "FlatRate"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Std.Rate"
            Object.Width           =   1766
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Technician"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Hours Work"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Techcode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "DETCDE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Charged To"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "DETCDE"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "LINENO"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView lstPMSDet 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   85
         Top             =   480
         Visible         =   0   'False
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6271
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
         MouseIcon       =   "FrmViewRO.frx":20DE
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   15522
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Model"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView ListParts 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6271
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
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmViewRO.frx":2240
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Parts No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Parts Description"
            Object.Width           =   13053
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "SRP"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListMaterial 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   87
         Top             =   480
         Visible         =   0   'False
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6271
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
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmViewRO.frx":23A2
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material Description"
            Object.Width           =   13053
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "SRP"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView ListAccessories 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   88
         Top             =   480
         Visible         =   0   'False
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   6271
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
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmViewRO.frx":2504
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material Description"
            Object.Width           =   13053
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "SRP"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2BDB6&
      BackStyle       =   0  'Transparent
      Caption         =   " Plate no."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   91
      Top             =   930
      Width           =   2025
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2BDB6&
      BackStyle       =   0  'Transparent
      Caption         =   " Vehicle Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   90
      Top             =   645
      Width           =   2025
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2BDB6&
      BackStyle       =   0  'Transparent
      Caption         =   " Customer Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   89
      Top             =   375
      Width           =   2025
   End
   Begin VB.Label LBLtechcode 
      Caption         =   "TECHCODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5490
      TabIndex        =   75
      Top             =   8970
      Width           =   2145
   End
   Begin VB.Label lblLineNO 
      Caption         =   "LINENO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3540
      TabIndex        =   74
      Top             =   8970
      Width           =   1665
   End
   Begin VB.Label labRO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "labRO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2340
      TabIndex        =   58
      Top             =   90
      Width           =   4875
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00D2BDB6&
      BackStyle       =   0  'Transparent
      Caption         =   " Repair Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   57
      Top             =   90
      Width           =   2025
   End
   Begin VB.Label labCustomer 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "labCustomer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2340
      TabIndex        =   56
      Top             =   375
      Width           =   4875
   End
   Begin VB.Label labVehicle 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "labVehicle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2340
      TabIndex        =   55
      Top             =   645
      Width           =   4875
   End
   Begin VB.Label labActNo 
      BackColor       =   &H00E4FEF1&
      BackStyle       =   0  'Transparent
      Caption         =   "labActNo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8070
      TabIndex        =   54
      Top             =   240
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label labStatus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E4FEF1&
      BackStyle       =   0  'Transparent
      Caption         =   "labStatus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8310
      TabIndex        =   53
      Top             =   120
      Width           =   3165
   End
   Begin VB.Label labPlateNo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "labPlateNo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2340
      TabIndex        =   52
      Top             =   930
      Width           =   4875
   End
   Begin VB.Label labSourceLed 
      BackColor       =   &H00E4FEF1&
      BackStyle       =   0  'Transparent
      Caption         =   "labSourceLed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8130
      TabIndex        =   51
      Top             =   930
      Width           =   3075
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Repair Order Date:"
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
      Height          =   225
      Left            =   120
      TabIndex        =   49
      Top             =   5760
      Width           =   1545
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Advisor:"
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
      Height          =   285
      Left            =   90
      TabIndex        =   48
      Top             =   6990
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Promise Date:"
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
      Height          =   225
      Left            =   120
      TabIndex        =   45
      Top             =   6390
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Technician &3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   9240
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Technician &2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   2
      Top             =   9000
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Technician:"
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
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   7740
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "frmCSMSViewRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thetechcode                                         As String
Dim thetechnician                                       As String
Dim TheEmpNO                                            As String
Dim THEDETCDE                                           As String
Dim theBlank                                            As Integer
Attribute theBlank.VB_VarUserMemId = 1073938436
Dim theRo                                               As String
Attribute theRo.VB_VarUserMemId = 1073938437
Dim Tech_mode                                           As Boolean
Attribute Tech_mode.VB_VarUserMemId = 1073938438
Dim Assingned_mode                                      As Boolean
Dim change_flag                                         As Boolean
Attribute change_flag.VB_VarUserMemId = 1073938440
Dim Newtechcode                                         As String
Attribute Newtechcode.VB_VarUserMemId = 1073938441
Dim NewTechnician                                       As String    'this used for updating the technician
Dim rsFind                                              As ADODB.Recordset
Attribute rsFind.VB_VarUserMemId = 1073938443

Function CheckIfThereAPMS(vRO As String) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT JOBTYPE FROM CSMS_RO_DET WHERE REP_OR = '" & vRO & "' AND LIVIL = '1' AND JOBTYPE = 'PMS'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        CheckIfThereAPMS = True
    Else
        CheckIfThereAPMS = False
    End If

    Set RSTMP = Nothing
End Function

Function GetTaym()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim X                                              As Integer
    Dim cnt                                            As Integer
    cnt = 0
    
    Set RSTMP = gconDMIS.Execute("Select PromiseDate From CSMS_RepairOrder Where RO_no = '" & labRO.Caption & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        For X = 1 To Len(RSTMP!PromiseDate)
            If Mid(RSTMP!PromiseDate, X, 1) = "/" Then cnt = cnt + 1
            If cnt = 2 Then
                GetTaym = Mid(RSTMP!PromiseDate, X + 6, Len(RSTMP!PromiseDate) - X)
                Exit For
            End If
        Next
    End If
    Set RSTMP = Nothing
End Function

Function FindTechName(SACODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Dim RSCON                                          As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & SACODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindTechName = Null2String(RSTMP!TECH_NAME)
    Else
        Set RSCON = gconDMIS.Execute("SELECT COMPANYNAME FROM CSMS_CONTRACTOR WHERE CODE = '" & SACODE & "'")
        If Not (RSCON.BOF And RSCON.EOF) Then
            FindTechName = Null2String(RSCON!CompanyName)
        Else
            Set RSCON = gconDMIS.Execute("SELECT NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = " & N2Str2Null(SACODE) & "")
            If Not (RSCON.BOF And RSCON.EOF) Then
                FindTechName = Null2String(RSCON!nameofvendor)
            Else
                FindTechName = ""
            End If
        End If
        Set RSCON = Nothing
    End If

    Set RSTMP = Nothing
End Function

Sub CheckIfBlank()
    If StrComp(thetechnician, "") = 0 Then
        Tech_mode = False
    Else
        Tech_mode = True
    End If
End Sub

Sub IfTechISClockIn()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim cnt                                            As Integer

    ListView2.Enabled = False
    'COMMENT    : MJP RSC 021209 0522PM
    'DECRIPTION : INCOMPLETE QUERY
        'SQL = "SELECT * FROM CSMS_JobClock WHERE techcode = '" & LTrim(RTrim(thetechcode)) & _
        '        "' and RO_NO = '" & labRO.Caption & "'"
    'COMMENT    : MJP RSC 021209 0522PM
    
    'UPDATE BY  : MJP RSC 021209 0522PM
    'DECRIPTION : ADD AN FILTER FOR JOB
        SQL = "SELECT * FROM CSMS_JobClock WHERE techcode = '" & LTrim(RTrim(thetechcode)) & _
            "' and RO_NO = '" & labRO.Caption & _
            "' AND DETCDE = '" & THEDETCDE & "'"
    'UPDATE BY  : MJP RSC 021209 0522PM
    Set RS = gconDMIS.Execute(SQL)

    ListView2.Enabled = False
    change_flag = False

    ListView2.ListItems.Clear
    cnt = 0

    If Not RS.EOF And Not RS.BOF Then
        ListView2.Enabled = True
    End If

    With RS
        If .EOF And .BOF Then
            Frame1.Visible = True
            'DATE UPDATED: 02-10-2009
             SSTab1.Enabled = False
            Call FillTheTech
        End If
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = ListView2.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!CLOCKIN)
            .MoveNext

            If ITEM.SubItems(1) = "" Then
                Call FillTheTech
                Frame1.Visible = True
            Else
                change_flag = True
            End If
        Loop
    End With
    Set RS = Nothing
End Sub

Sub displayJobservice()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem
    Dim theRo                                          As String
    Dim cnt                                            As Integer
    theRo = Trim(labRO.Caption)

    Set RS = gconDMIS.Execute("SELECT jobtype,detail,flatrate,detprc,DET_HRS,technician,HRSWRK,techcode,DETCDE,livil,wcode,taxval,DISVAL,LINE_NO FROM CSMS_RO_det WHERE rep_OR = '" & theRo & "' and livil='1'")
    lblJob4Service.ListItems.Clear
    cnt = 0

    With RS
        Do While Not .EOF
            cnt = cnt + 1
            Set ITEM = lblJob4Service.ListItems.Add(, , cnt)
            ITEM.SubItems(1) = Null2String(!JOBTYPE)
            ITEM.SubItems(2) = Null2String(!Detail)
            ITEM.SubItems(3) = Null2String(!FLATRATE)
            ITEM.SubItems(4) = Null2String(!DET_HRS)
            'Item.SubItems(5) = FindTechName(Null2String(!Technician))
            ITEM.SubItems(5) = FindTechName(Null2String(!TechCode))
            ITEM.SubItems(6) = NumericVal(!HRSWRK)
            ITEM.SubItems(7) = Null2String(!TechCode)
            ITEM.SubItems(8) = Null2String(!DETCDE)
            ITEM.SubItems(11) = Null2String(!LINE_NO)

            If (RS!wCode) = "W" Then
                txtWarLabor = NumericVal(txtWarLabor) + (N2Str2Zero(RS!DetPrc))
                txtVatLabor = NumericVal(txtVatLabor) + (N2Str2Zero(RS!TAXVAL))
            ElseIf (RS!wCode) = "C" Then
                txtCompLabor = NumericVal(txtCompLabor) + (N2Str2Zero(RS!DetPrc))
            ElseIf (RS!wCode) = "S" Then
                txtSalesLabor = NumericVal(txtSalesLabor) + (N2Str2Zero(RS!DetPrc))
            Else
                txtEstLabor = NumericVal(txtEstLabor) + (N2Str2Zero(RS!DetPrc))
                txtVatLabor = NumericVal(txtVatLabor) + (N2Str2Zero(RS!TAXVAL))
            End If
            txtDiscLabor = NumericVal(txtDiscLabor) + (N2Str2Zero(RS!disval))
            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub DisplayComPute()
    On Error Resume Next

    Dim SQL                                            As String
    Dim ITEM                                           As ListItem
    Dim RS                                             As New ADODB.Recordset
    Dim theRo                                          As String

    theRo = Trim(labRO.Caption)
    Set RS = gconDMIS.Execute(SQL = "SELECT * FROM CSMS_repor WHERE Rep_OR ='" & theRo & "'")

    With RS
        '        txtEstLabor.Text = ToDoubleNumber(N2Str2Zero(!labor))
        '        txtEstParts.Text = ToDoubleNumber(N2Str2Zero(!parts))
        '        txtEstMat.Text = ToDoubleNumber(N2Str2Zero(!Materials))
        '        txtEstAces.Text = ToDoubleNumber(N2Str2Zero(!Accessories))
        '
        '        txtVatLabor.Text = ToDoubleNumber(N2Str2Zero(!l_taxval))
        '        txtVatParts.Text = ToDoubleNumber(N2Str2Zero(!p_taxval))
        '        txtVatMat.Text = ToDoubleNumber(N2Str2Zero(!m_taxval))
        '        txtVatAces.Text = ToDoubleNumber(N2Str2IntZero(!a_taxval))
        '        txtVatTotal.Text = ToDoubleNumber(N2Str2Zero(!l_taxval) + N2Str2Zero(!p_taxval) + N2Str2Zero(!m_taxval) + N2Str2Zero(!a_taxval))
        '
        '        txtDiscLabor = ToDoubleNumber(N2Str2Zero(!l_discount))
        '        txtDiscParts = ToDoubleNumber(N2Str2Zero(!p_discount))
        '        txtDiscMat = ToDoubleNumber(N2Str2Zero(!m_discount))
        '        txtDiscAces = ToDoubleNumber(N2Str2Zero(!a_discount))
        '        txtDiscTotal = ToDoubleNumber(N2Str2Zero(!l_discount) + N2Str2Zero(!p_discount) + N2Str2Zero(!m_discount) + N2Str2Zero(!a_discount))
    End With
    Set RS = Nothing
End Sub

Sub FillTheTech()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    techlist.Enabled = False
    Set RS = gconDMIS.Execute("SELECT * FROM CSMS_vw_technicianAvailability WHERE Code = 'A'")

    techlist.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        techlist.Enabled = True
    End If

    With RS
        Do While Not .EOF
            Set ITEM = techlist.ListItems.Add(, , !EMPNO)
            ITEM.SubItems(1) = Null2String(!TECH_NAME)
            ITEM.SubItems(2) = Null2String(!Status)
            ITEM.SubItems(3) = Null2String(!assignedro)
            ITEM.SubItems(4) = Null2String(!Firstname)
            ITEM.SubItems(5) = Null2String(!Code)
            ITEM.SubItems(6) = Null2String(!inout)
            ITEM.SubItems(7) = Null2String(!TechCode)

            .MoveNext
        Loop
    End With
    Set RS = Nothing
End Sub

Sub CheckIfFill()
    If StrComp(thetechnician, "") = 0 Then
        Assingned_mode = False
    Else
        Assingned_mode = True
    End If
End Sub

Sub FillPMSJob()
    '    Dim SQL                                                           As String
    '    Dim rs                                                            As New ADODB.Recordset
    '    Dim Item                                                          As ListItem
    '    Dim cnt                                                           As Integer
    '    SQL = "SELECT DETCDE,jobtype,DETDSC FROM CSMS_Ro_Det Where Rep_or='" & labRO.Caption & "' and jobtype='PMS '"
    '
    '    lstPMSDet.Enabled = False
    '
    '    Set rs = New ADODB.Recordset
    '    Set rs = gconDMIS.Execute(SQL)
    '
    '    lstPMSDet.ListItems.Clear
    '    cnt = 0
    '
    '    If Not rs.EOF And Not rs.BOF Then
    '        lstPMSDet.Enabled = True
    '    End If
    '
    '    With rs
    '        Do While Not .EOF
    '            Set Item = lstPMSDet.ListItems.Add(, , !DetCDE)
    '            Item.SubItems(1) = Null2String(!JOBTYPE)
    '            Item.SubItems(2) = Null2String(!Detdsc)
    '
    '            .MoveNext
    '
    '        Loop
    '    End With
    '    Set rs = Nothing
End Sub

Sub FillParts()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RS = gconDMIS.Execute("SELECT DETCDE,DETDSC,detvol,DETPRC,det_hrs,rep_or,wcode,taxval,DISVAL from CSMS_ro_det Where livil ='2' and rep_or='" & labRO.Caption & "' Order by [LINE_NO] Asc")

    ListParts.ListItems.Clear
    txtWarParts.Text = 0
    txtCompPart.Text = 0
    txtSalesParts.Text = 0
    txtEstParts.Text = 0

    Do While Not RS.EOF
        Set ITEM = ListParts.ListItems.Add(, , RS!DETCDE)
        'Item.SubItems(1) = N2Str2Zero(rs!DetCDE) 'comment by JUN 01/16/2008
        ITEM.SubItems(1) = Null2String(RS!DETCDE)
        'Item.SubItems(2) = N2Str2Zero(rs!Detdsc) 'comment by JUN 01/16/2008
        ITEM.SubItems(2) = Null2String(RS!DETDSC)
        ITEM.SubItems(3) = N2Str2Zero(RS!detvol)
        ITEM.SubItems(4) = N2Str2Zero(RS!DetPrc)

        If (RS!wCode) = "W" Then
            txtWarParts = NumericVal(txtWarParts) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatParts = NumericVal(txtVatParts) + (N2Str2Zero(RS!TAXVAL))
        ElseIf (RS!wCode) = "C" Then
            txtCompPart = NumericVal(txtCompPart) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        ElseIf (RS!wCode) = "S" Then
            txtSalesParts = NumericVal(txtSalesParts) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        Else
            txtEstParts = NumericVal(txtEstParts) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatParts = NumericVal(txtVatParts) + (N2Str2Zero(RS!TAXVAL))
        End If
        txtDiscParts = NumericVal(txtDiscParts) + (N2Str2Zero(RS!disval))

        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub FillMaterial()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    ListMaterial.Enabled = False
    Set RS = gconDMIS.Execute("SELECT DETCDE,DETDSC,detvol,DETPRC,det_hrs,rep_or,wcode,taxval,DISVAL from CSMS_ro_det Where livil ='3' and rep_or='" & labRO.Caption & "' Order by [LINE_NO] Asc")

    ListMaterial.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        ListMaterial.Enabled = True
    End If
    Do While Not RS.EOF
        'Set Item = ListParts.ListItems.Add(, , rs!DetCDE) 'comment by JUN 01/16/2008
        Set ITEM = ListMaterial.ListItems.Add(, , RS!DETCDE)
        ITEM.SubItems(1) = Null2String(RS!DETCDE)
        ITEM.SubItems(2) = Null2String(RS!DETDSC)
        ITEM.SubItems(3) = Null2String(RS!detvol)
        ITEM.SubItems(4) = Null2String(RS!DetPrc)

        If (RS!wCode) = "W" Then
            txtWarMat = NumericVal(txtWarMat) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatMat = NumericVal(txtVatMat) + (N2Str2Zero(RS!TAXVAL))
        ElseIf (RS!wCode) = "C" Then
            txtCompMat = NumericVal(txtCompMat) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        ElseIf (RS!wCode) = "S" Then
            txtSalesMat = NumericVal(txtSalesMat) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        Else
            txtEstMat = NumericVal(txtEstMat) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatMat = NumericVal(txtVatMat) + (N2Str2Zero(RS!TAXVAL))
        End If
        txtDiscMat = NumericVal(txtDiscMat) + (N2Str2Zero(RS!disval))

        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub FillAccessories()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    ListAccessories.Enabled = False
    Set RS = gconDMIS.Execute("SELECT DETCDE,DETDSC,detvol,DETPRC,det_hrs,rep_or,wcode,taxval,DISVAL from CSMS_ro_det Where livil ='4' and rep_or='" & labRO.Caption & "' Order by [LINE_NO] Asc")

    ListAccessories.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        ListAccessories.Enabled = True
    End If
    Do While Not RS.EOF
        'Set Item = ListParts.ListItems.Add(, , rs!DetCDE) 'comment by JUN 01/16/2008
        Set ITEM = ListAccessories.ListItems.Add(, , RS!DETCDE)
        ITEM.SubItems(1) = Null2String(RS!DETCDE)
        ITEM.SubItems(2) = Null2String(RS!DETDSC)
        ITEM.SubItems(3) = Null2String(RS!detvol)
        ITEM.SubItems(4) = Null2String(RS!DetPrc)

        If (RS!wCode) = "W" Then
            txtWarLaborAces = NumericVal(txtWarLaborAces) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatAces = NumericVal(txtVatAces) + (N2Str2Zero(RS!TAXVAL))
        ElseIf (RS!wCode) = "C" Then
            txtCompAces = NumericVal(txtCompAces) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        ElseIf (RS!wCode) = "S" Then
            txtSalesAces = NumericVal(txtSalesAces) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
        Else
            txtEstAces = NumericVal(txtEstAces) + (N2Str2Zero(RS!DetPrc) * N2Str2Zero(RS!detvol))
            txtVatAces = NumericVal(txtVatAces) + (N2Str2Zero(RS!TAXVAL))
        End If
        txtDiscAces = NumericVal(txtDiscAces) + (N2Str2Zero(RS!disval))

        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Sub TheSum()


    txtWarLabor = ToDoubleNumber(txtWarLabor)
    txtVatLabor = ToDoubleNumber(txtVatLabor)
    txtCompLabor = ToDoubleNumber(txtCompLabor)
    txtSalesLabor = ToDoubleNumber(txtSalesLabor)
    txtEstLabor = ToDoubleNumber(txtEstLabor)
    txtVatLabor = ToDoubleNumber(txtVatLabor)
    txtDiscLabor = ToDoubleNumber(txtDiscLabor)


    txtWarParts = ToDoubleNumber(txtWarParts)
    txtVatParts = ToDoubleNumber(txtVatParts)
    txtCompPart = ToDoubleNumber(txtCompPart)
    txtSalesParts = ToDoubleNumber(txtSalesParts)
    txtEstParts = ToDoubleNumber(txtEstParts)
    txtVatParts = ToDoubleNumber(txtVatParts)
    txtDiscParts = ToDoubleNumber(txtDiscParts)


    txtWarMat = ToDoubleNumber(txtWarMat)
    txtVatMat = ToDoubleNumber(txtVatMat)
    txtCompMat = ToDoubleNumber(txtCompMat)
    txtSalesMat = ToDoubleNumber(txtSalesMat)
    txtEstMat = ToDoubleNumber(txtEstMat)
    txtVatMat = ToDoubleNumber(txtVatMat)
    txtDiscMat = ToDoubleNumber(txtDiscMat)


    txtWarLaborAces = ToDoubleNumber(txtWarLaborAces)
    txtVatAces = ToDoubleNumber(txtVatAces)
    txtCompAces = ToDoubleNumber(txtCompAces)
    txtSalesAces = ToDoubleNumber(txtSalesAces)
    txtEstAces = ToDoubleNumber(txtEstAces)
    txtVatAces = ToDoubleNumber(txtVatAces)
    txtDiscAces = ToDoubleNumber(txtDiscAces)

    txtTotalAmt = ToDoubleNumber(NumericVal(txtEstLabor.Text) + NumericVal(txtEstParts.Text) + NumericVal(txtEstMat.Text) + NumericVal(txtEstAces.Text))

    txtWarLaborTotal = ToDoubleNumber(NumericVal(txtWarLabor) + NumericVal(txtWarParts) + NumericVal(txtWarMat) + NumericVal(txtWarLaborAces))

    txtSalesTotal = ToDoubleNumber(NumericVal(txtSalesLabor) + NumericVal(txtSalesParts) + NumericVal(txtSalesMat) + NumericVal(txtSalesAces))

    txtCompTotal = ToDoubleNumber(NumericVal(txtCompLabor) + NumericVal(txtCompPart) + NumericVal(txtCompMat) + NumericVal(txtCompAces))
    txtVatTotal.Text = ToDoubleNumber(NumericVal(txtVatLabor) + NumericVal(txtVatParts) + NumericVal(txtVatMat) + NumericVal(txtVatAces))
    txtDiscTotal.Text = ToDoubleNumber(NumericVal(txtDiscLabor) + NumericVal(txtDiscParts) + NumericVal(txtDiscMat) + NumericVal(txtDiscAces))

End Sub

Private Sub Assigned_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim Index                                          As Integer
    Dim vEMPNO                                         As String

    Index = techlist.SelectedItem.Index
    vEMPNO = techlist.ListItems(Index).Text

    
    'UPDATED BY: JUN-----------------------------------------------------
    'DATE UPDATED: 02-10-2009
    'DESCRIPTION: VALIDATE THE NEW TECHNICIAN CODE IF IT HAS A VALUE
    If TheEmpNO = "" Then
        MsgBox "Please select a technician.", vbInformation, "INFORMATION"
        Exit Sub
    End If
    'UPDATED BY: JUN-----------------------------------------------------
    
    
    SQL_STATEMENT = "update CSMS_Ro_det set techcode = '" & Newtechcode & "', TECHNICIAN = '" & NewTechnician & "' WHERE Rep_or = '" & labRO.Caption & "' and DETCDE = '" & THEDETCDE & "'"
    gconDMIS.Execute (SQL_STATEMENT)
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & THEDETCDE & " TECH CODE: " & Newtechcode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Set RSTMP = gconDMIS.Execute("SELECT empno from hrms_empinfo where empno = '" & vEMPNO & "' ")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET ASSIGNEDRO = '" & labRO.Caption & "',JSTATUS = 'S' WHERE EMPNO = '" & vEMPNO & "'"
        gconDMIS.Execute (SQL_STATEMENT)
    Else
        SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET ASSIGNEDRO = '" & labRO.Caption & "',JSTATUS = 'S' WHERE EMPNO = '" & vEMPNO & "'"
        gconDMIS.Execute (SQL_STATEMENT)
    End If
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AS", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    'COMMENT BY : MJP
    'gconDMIS.Execute "update CSMS_vw_Technician set AssignedRO = '" & labRO.Caption & "',JStatus = 'S' where Technician = '" & Newtechcode & "'"
    'COMMENT BY : MJP

    MsgBox "Technician Has Been Assigned!", vbInformation, "Confirm"
    picButton.Enabled = True
    theBlank = 0
    displayJobservice
    Frame1.Visible = False
    
    'UPDATED BY: JUN
    'DATE UPDATED: 02-09-2009
    'DESCRIPTION: ENABLE THE SSTAB AFTER THE USER ASSIGNED A TECHNICIAN
     SSTab1.Enabled = True
     TheEmpNO = ""
     thetechcode = ""
     thetechnician = ""
    'UPDATED BY: JUN
End Sub

Private Sub Change_Click()
    Dim NewStatus                                       As String
    Dim RSTMP                                           As New ADODB.Recordset
    Dim RSTMP1                                          As New ADODB.Recordset
    Dim rsEmpNo                                         As New ADODB.Recordset
    Dim OLDEMPNO                                        As String
    Dim NewEmpNo                                        As String

    'UPDATED BY: JUN-----------------------------------------------------
    'DATE UPDATED: 12-16-2008
    'DESCRIPTION: VALIDATE THE NEW TECHNICIAN CODE IF IT HAS A VALUE
    If Newtechcode = "" Then
        MsgBox "Please select a technician.", vbInformation, "INFORMATION"
        Exit Sub
    End If
    'UPDATED BY: JUN-----------------------------------------------------

    'UPDATED BY: JUN
    'DATE UPDATED: 02-09-2009
    'DESCRIPTION: DO NOT ALLOW THE USER TO CHANGE THE TECHNICIAN IF THE JOB IS NOT YET ASSIGN TO A PARTICULAR TECHNICIAN
     If thetechnician = "" Then
        MsgBox ("You Cannot Change technician" & vbCrLf & "This Job is not yet been assigned to technician" & vbCrLf & "Ruther Select Assign technician"), vbInformation, "INFORMATION"
        Exit Sub
     End If
    'UPDATED BY: JUN

    Set rsEmpNo = gconDMIS.Execute("SELECT EMPNO FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & thetechcode & "'")
    If Not (rsEmpNo.EOF And rsEmpNo.BOF) Then
        OLDEMPNO = rsEmpNo!EMPNO
    End If
    
    Set rsEmpNo = gconDMIS.Execute("SELECT EMPNO FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & Newtechcode & "'")
    If Not (rsEmpNo.EOF And rsEmpNo.BOF) Then
        NewEmpNo = rsEmpNo!EMPNO
    End If

    SQL_STATEMENT = "update CSMS_Ro_det set techcode = '" & Newtechcode & _
        "', TECHNICIAN = '" & NewTechnician & _
        "' WHERE Rep_or = '" & labRO.Caption & _
        "' and techcode = '" & thetechcode & _
        "' AND DETCDE = '" & THEDETCDE & "'"
    gconDMIS.Execute (SQL_STATEMENT)
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("AS", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(labRO.Caption), "REP_OR", "CSMS_REPOR"), "", "JOB CODE: " & THEDETCDE & " TECH CODE: " & thetechcode, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    NewStatus = "A"
    'Exit Sub
    
    'UPDATE BY : MJP
    Set RSTMP = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & OLDEMPNO & "'")
    If Not (RSTMP.EOF And RSTMP.BOF) Then
        'UPDATED BY: JUN--------------------------------------------------------------------------------------------------
        'DATE UPDATE: 03-10-2009
        'DESCRIPTION: CHECK IF THE TECHNICIAN IS CURRENLTY WORKING OR HAS A JOB ASSIGNED
        If CHECK_JOB_STATUS_TO_CHANGE(LTrim(RTrim(labRO)), LTrim(RTrim(thetechcode))) = True Then
            'technician must not be set ASSIGNEDRO NULL and STATUS to AVAILABLE
        Else
            SQL_STATEMENT = "update HRMS_EMPINFO set AssignedRO = NULL " & _
                ", JStatus = '" & NewStatus & _
                "' where EMPNO = '" & OLDEMPNO & "'"
            gconDMIS.Execute (SQL_STATEMENT)
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(OLDEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
        'UPDATED BY: JUN--------------------------------------------------------------------------------------------------
        
        Set RSTMP1 = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & NewEmpNo & "'")
        If Not (RSTMP1.BOF And RSTMP1.EOF) Then
            SQL_STATEMENT = "update HRMS_EMPINFO set AssignedRO = '" & labRO.Caption & _
                "', JStatus = 'S' " & _
                " where EMPNO = '" & NewEmpNo & "'"
            gconDMIS.Execute (SQL_STATEMENT)
        Else
            SQL_STATEMENT = "update CSMS_EMPINFO set AssignedRO = '" & labRO.Caption & _
                "', JStatus = 'S' " & _
                " where EMPNO = '" & NewEmpNo & "'"
            gconDMIS.Execute (SQL_STATEMENT)
        End If
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("AS", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(NewEmpNo), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    Else
        'UPDATED BY: JUN--------------------------------------------------------------------------------------------------
        'DATE UPDATE: 03-10-2009
        'DESCRIPTION: CHECK IF THE TECHNICIAN IS CURRENLTY WORKING OR HAS A JOB ASSIGNED
        If CHECK_JOB_STATUS_TO_CHANGE(LTrim(RTrim(labRO)), LTrim(RTrim(thetechcode))) = True Then
            'technician must not be set ASSIGNEDRO NULL and STATUS to AVAILABLE
        Else
            SQL_STATEMENT = "update csms_EMPINFO set AssignedRO = NULL " & _
                ", JStatus = '" & NewStatus & _
                "' where EMPNO = '" & OLDEMPNO & "'"
            gconDMIS.Execute (SQL_STATEMENT)
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(OLDEMPNO), "EMPNO", "csms_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
            'NEW LOG AUDIT-----------------------------------------------------
        End If
        'UPDATED BY: JUN--------------------------------------------------------------------------------------------------
        
        Set RSTMP1 = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & NewEmpNo & "'")
        If Not (RSTMP1.BOF And RSTMP1.EOF) Then
            SQL_STATEMENT = "update HRMS_EMPINFO set AssignedRO = '" & labRO.Caption & _
                "', JStatus = 'S' " & _
                " where EMPNO = '" & NewEmpNo & "'"
            gconDMIS.Execute (SQL_STATEMENT)
        Else
            SQL_STATEMENT = "update CSMS_EMPINFO set AssignedRO = '" & labRO.Caption & _
                "', JStatus = 'S' " & _
                " where EMPNO = '" & NewEmpNo & "'"
            gconDMIS.Execute (SQL_STATEMENT)
        End If
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("AS", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(NewEmpNo), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & labRO.Caption, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
    End If
    'UPDATE BY : MJP

    Call displayJobservice

    MsgBox "Repair order Information Has Been Change!", vbInformation, "Info"
    picButton.Enabled = True
    theBlank = 0
    Frame1.Visible = False

    'UPDATED BY: JUN---------------------------------------------------
    'DATE UPDATED: 12-16-2008
    'DESCRIPTION: CLEAR THE OLDTECHNICIAN CODE AND NEW TECHNICIAN CODE
        thetechcode = ""
        Newtechcode = ""
    'UPDATED BY: JUN---------------------------------------------------

    'UPDATED BY: JUN
    'DATE UPDATED: 02-09-2009
    'DESCRIPTION: ENABLE THE SSTAB AFTER THE USER ASSIGNED A TECHNICIAN
         SSTab1.Enabled = True
    'UPDATED BY: JUN
End Sub

Function CHECK_JOB_STATUS_TO_CHANGE(xLAB_RO As String, xTECH_CODE As String) As Boolean
    'UPDATED BY: JUN
    'DATE UPDATED: 03092009
    'DESCRIPTION:
    
    Dim rsChecK_TechJob                                             As New ADODB.Recordset
    
    Set rsChecK_TechJob = gconDMIS.Execute("Select * FROM CSMS_RO_DET WHERE TECHCODE = '" & xTECH_CODE & "' AND REP_OR = '" & xLAB_RO & "' and (DONE = NULL or DONE <> 'Y')")
    If Not rsChecK_TechJob.EOF And Not rsChecK_TechJob.BOF Then
        CHECK_JOB_STATUS_TO_CHANGE = True
    Else
        CHECK_JOB_STATUS_TO_CHANGE = False
    End If
    Set rsChecK_TechJob = Nothing
End Function

Private Sub cmdAssignedTech_Click()
    'BTT - 05242007
    If theBlank <> 0 Then
        picButton.Enabled = False
        Assigned.Visible = True
        Change.Visible = False
        Call CheckIfFill
        If Assingned_mode = True Then
            MsgBox "Job Has Already A Technician!", vbInformation, "Information"
            'DATE UPDATED: 02-09-2009
            SSTab1.Enabled = True
            picButton.Enabled = True
            Exit Sub
        End If
        
        'UPDATE BY   : MJP 021609 0550PM
        'DESCRIPTION :
            If CheckifJobisAlreadyFinish(labRO, THEDETCDE) = True Then
                MsgBox "Job is already Finish", vbInformation, "CSMS"
                picButton.Enabled = True
                Exit Sub
            End If
        'UPDATE BY   : MJP 021609 0550PM
        
        Frame1.Visible = True
        Call FillTheTech
        'UPDATED BY: JUN
        'DATE UPDATED: 02-09-2009
        'DESCRIPTION: DIS ALLOW THE USER TO SELECT THE JOB DETAILS BECAUSE IT MIGHT CAUSE WRONG UPDATE OF TECHNICIAN TO OTHER JOB
         SSTab1.Enabled = False
        'UPDATED BY: JUN
    Else
        MsgBox "Please Select A Job!", vbExclamation, "Information"
        'DATE UPDATED: 02-09-2009
         SSTab1.Enabled = True
    End If
End Sub

Function CheckifJobisAlreadyFinish(XRONO As String, xJOBCODE As String) As Boolean
    Dim RSTMP                                           As New ADODB.Recordset
    
    Set RSTMP = gconDMIS.Execute("SELECT DONE FROM CSMS_RO_DET WHERE DETCDE = " & N2Str2Null(xJOBCODE) & _
        " AND REP_OR = " & N2Str2Null(XRONO) & "")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!DONE) = "Y" Then
            CheckifJobisAlreadyFinish = True
        Else
            CheckifJobisAlreadyFinish = False
        End If
    End If
    Set RSTMP = Nothing
End Function

Private Sub cmdChangestat_Click()
    Screen.MousePointer = 11
    rptRepairOrder.WindowShowPrintSetupBtn = True
    rptRepairOrder.WindowTitle = "Repair Order Print Out"
    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HPI" Then
        rptRepairOrder.Formulas(3) = "TAYM = '" & GetTaym & "'"
    End If

    If COMPANY_CODE = "HAS" Then
        If CheckIfThereAPMS(labRO) = True Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & labRO & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder_NOPMS.rpt", "{repor.rep_or} = '" & labRO & "'", CSMS_REPORT_CONNECTION, 1
        End If
    Else
        If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HEI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HCI" Then
            Dim RSTMP                                  As New ADODB.Recordset
            Dim FJOB                                   As String
            Dim SJOB                                   As String
            Dim TJOB                                   As String
            Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & labPlateNo.Caption & "' AND TRANSTYPE = 'R' AND DTE_RECD < '" & Date & "' ORDER BY DTE_RECD ASC ")
            If Not (RSTMP.BOF And RSTMP.EOF) Then
                If Not RSTMP.BOF Then
                    RSTMP.MoveFirst
                    
                    If COMPANY_CODE = "HCI" Then
                        FJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
                        FJOB = FJOB & "    " & GetJobList(RSTMP!REP_OR)
                    Else
                        FJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
 
                    End If
                    RSTMP.MoveNext

                    If Not RSTMP.EOF Then
                        If COMPANY_CODE = "HCI" Then
                            SJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
                            SJOB = SJOB & "    " & GetJobList(RSTMP!REP_OR)
                        Else
                            SJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
                        End If
                        RSTMP.MoveNext

                        If Not RSTMP.EOF Then
                            If COMPANY_CODE = "HCI" Then
                                TJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
                            Else
                                TJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!dte_recd) & "    " & Null2String(RSTMP!km_rdg)
                                TJOB = TJOB & "    " & GetJobList(RSTMP!REP_OR)
                            End If
                        End If
                    End If
                End If
            End If
            Set RSTMP = Nothing
            
            rptRepairOrder.Formulas(0) = "RO1 = '" & FJOB & "'"
            rptRepairOrder.Formulas(1) = "RO2 = '" & SJOB & "'"
            rptRepairOrder.Formulas(2) = "RO3 = '" & TJOB & "'"
            
            If COMPANY_CODE = "HCI" Then rptRepairOrder.Formulas(2) = "TAYM = '" & GetTaym & "'"
            If COMPANY_CODE = "HCI" Then rptRepairOrder.Formulas(3) = "ATTENDED = '" & GetAttendedTaym & "'"
            
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & labRO & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & labRO & "'", CSMS_REPORT_CONNECTION, 1
        End If
    End If

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "BILLING SYSTEM", "", FindTransactionID(N2Str2Null(labRO), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & labRO & " - VIEW RO DETAILS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    'LogAudit "V", "PRINT RO"
    Screen.MousePointer = 0
    
End Sub
Function GetAttendedTaym()
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select SAVETIME From CSMS_RepOR Where REP_OR = '" & labRO.Caption & "' AND TRANSTYPE = 'R'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GetAttendedTaym = Null2String(RSTMP!savetime)
        If Hour(GetAttendedTaym) < 12 Then
            GetAttendedTaym = Mid(GetAttendedTaym, 1, 5) & " AM"
        Else
            GetAttendedTaym = Mid(GetAttendedTaym, 1, 5) & " PM"
        End If
    End If

    Set RSTMP = Nothing
End Function


Function GetJobList(xxxRO_NO As String) As String
    Dim RSTMP As New ADODB.Recordset
    Dim XXX As String
    Set RSTMP = gconDMIS.Execute("SELECT TOP 5 DETCDE FROM CSMS_RO_DET WHERE REP_OR = '" & xxxRO_NO & "' AND LIVIL = '1'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            XXX = XXX & "     " & Null2String(RSTMP!DETCDE)
            RSTMP.MoveNext
        Loop
    End If
    GetJobList = XXX
    Set RSTMP = Nothing
End Function

Private Sub cmdChangeTech_Click()

    'UPDATED BY: JUN----------------------------------------------------------------------------------
    'DATE UPDATED: 12-16-2008
    'DESCRIPTION: VALIDATE THE TECHNICIAN CODE
    If thetechcode = "" Then
        MsgBox "You can't change technician" & vbCrLf & "This job is not yet been assign to a technician" & vbCrLf & "Rather select Assign technician", vbInformation, "INFORMATION"
        'DATE UPDATED: 02-09-2009
        SSTab1.Enabled = True
        picButton.Enabled = True
        Exit Sub
    End If
    'UPDATED BY: JUN----------------------------------------------------------------------------------
    
    'BTT - 05242007
    If theBlank <> 0 Then
        picButton.Enabled = False
        Assigned.Visible = False
        Change.Visible = True
        Change.ZOrder 0
        Call CheckIfBlank
        
        If Tech_mode = False Then
            MsgBox "Cannot Change Pls Assigned Technician!", vbExclamation, "Information"
            picButton.Enabled = True
            Exit Sub
        End If
        Call IfTechISClockIn
        If change_flag = True Then
            If Trim(labStatus) = "Released" Then
                MsgBox "Cannot Change Technician.Job Is Aready Released!", vbInformation, "Information"
                'DATE UPDATED: 02-09-2009
                picButton.Enabled = True
                SSTab1.Enabled = True
            Else
                MsgBox "Cannot Change Technician. Job Is Aready Starting!", vbInformation, "Information"
                'DATE UPDATED: 02-09-2009
                picButton.Enabled = True
                SSTab1.Enabled = True
            End If
        End If
    Else
        MsgBox "Please Select A Technician...", vbExclamation, "Information"
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdJobClock_Click()
    If Module_Access(LOGID, "JOB CLOCK", "TRANSACTION") = False Then Exit Sub
    frmCSMSClockINOUT.Show 1
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
    'UPDATED BY: JUN
    'DATE UPDATED: 02-09-2009
    'DESCRIPTION: ENABLE THE SSTAB IF THE USER CANCEL THE USE OF CHANGE TECHNICIAN OR ASSIGNED OF TECHNICIAN
        theBlank = 0
        picButton.Enabled = True
        SSTab1.Enabled = True
        TheEmpNO = ""
        thetechcode = ""
        thetechnician = ""
    'UPDATED BY: JUN
End Sub

Sub initMemvars()
    labCustomer.Caption = ""
    labVehicle.Caption = ""
    labPlateNo.Caption = ""
End Sub

Private Sub Form_Activate()
    Change.Visible = False
    Assigned.Visible = False
    
    Call initMemvars

    Set rsFind = New ADODB.Recordset
    Set rsFind = gconDMIS.Execute("select * from CSMS_vw_REPAIRORDER where RO_NO = '" & Trim(labRO.Caption) & "'")
    If Not rsFind.EOF And Not rsFind.BOF Then
        labActNo.Caption = Null2String(rsFind![ACCT_NO])
        labCustomer.Caption = UCase(Null2String(rsFind![Customer]))
        labStatus.Caption = Null2String(rsFind![Status])
        labPlateNo.Caption = Null2String(rsFind![PLATE_NO])
        txtApptDate = Null2String(rsFind![AppointmentDate])
        txtPromise = Null2String(rsFind![PromiseDate])
        txtTech1 = Null2String(rsFind![tech1])
        txtTech2 = Null2String(rsFind![tech2])
        txtTech3 = Null2String(rsFind![tech3])
        txtAdvisor = Null2String(rsFind![writer])
        labSourceLed.Caption = Null2String(rsFind![customersourcelead])

        Dim rsVehicleKo                                As New ADODB.Recordset
        Set rsVehicleKo = gconDMIS.Execute("select * from CSMS_Cusveh where plate_no = '" & labPlateNo.Caption & "'")
        If Not (rsVehicleKo.EOF And rsVehicleKo.BOF) Then
            labVehicle.Caption = Null2String(Trim(Null2String(rsVehicleKo![YER]))) & "  " & Null2String(Trim(Null2String(rsVehicleKo![Make]))) & "  " & Null2String(Trim(Null2String(rsVehicleKo![MODEL])))
        End If

        Call displayJobservice
        lstPMSDet.Sorted = False: lstPMSDet.ListItems.Clear

        Set rsFind = New ADODB.Recordset
        Set rsFind = gconDMIS.Execute("select DETCDE,jobtype,detdsc,PMS_Model from CSMS_PMS_Job_Det where rep_OR = '" & Trim(labRO.Caption) & "' order by pms_model,detcde asc")
        If Not rsFind.EOF And Not rsFind.BOF Then
            Listview_Loadval Me.lstPMSDet.ListItems, rsFind
        End If
    End If

    Frame1.Visible = False

    Call DisplayComPute
    Call FillPMSJob
    Call FillParts
    Call FillMaterial
    Call FillAccessories
    Call TheSum
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    SSTab1.TabIndex = 0
End Sub

Private Sub lblJob4Service_Click()
'    On Error Resume Next
'
'    theBlank = lblJob4Service.ListItems.Item(lblJob4Service.SelectedItem.INDEX)
'    thetechcode = lblJob4Service.SelectedItem.SubItems(7)
'    LBLtechcode.Caption = lblJob4Service.SelectedItem.SubItems(7)
'    thetechnician = lblJob4Service.SelectedItem.SubItems(5)
'    THEDETCDE = lblJob4Service.SelectedItem.SubItems(8)
'    lblLineNO.Caption = lblJob4Service.SelectedItem.SubItems(9)
'    txtTech1.Text = thetechnician
End Sub

Private Sub lblJob4Service_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next

    theBlank = lblJob4Service.ListItems.ITEM(lblJob4Service.SelectedItem.Index)
    thetechcode = lblJob4Service.SelectedItem.SubItems(7)
    LBLtechcode.Caption = lblJob4Service.SelectedItem.SubItems(7)
    thetechnician = lblJob4Service.SelectedItem.SubItems(5)
    THEDETCDE = lblJob4Service.SelectedItem.SubItems(8)
    lblLineNO.Caption = lblJob4Service.SelectedItem.SubItems(9)
    txtTech1.Text = thetechnician
End Sub

Private Sub techlist_Click()
    On Error Resume Next

    TheEmpNO = techlist.ListItems.ITEM(techlist.SelectedItem.Index)
    Newtechcode = techlist.SelectedItem.SubItems(7)
    NewTechnician = techlist.SelectedItem.SubItems(4)
    TheEmpNO = techlist.SelectedItem.SubItems(0)
End Sub

Private Sub Timer1_Timer()
    If labStatus.ForeColor = &HC0& Then
        labStatus.ForeColor = &HC0C0&
    Else
        labStatus.ForeColor = &HC0&
    End If
End Sub
