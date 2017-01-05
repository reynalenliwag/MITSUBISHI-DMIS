VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO301B~1.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_ServiceCounter 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Counter"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   15165
   FillColor       =   &H00808080&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "FrmCSMS_ServiceCounter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15165
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   30
      ScaleHeight     =   675
      ScaleWidth      =   2295
      TabIndex        =   19
      Top             =   9270
      Width           =   2325
      Begin VB.TextBox txtTLhrs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   " 0.00"
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblUR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1140
         TabIndex        =   48
         Top             =   300
         Width           =   1155
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   -30
         Width           =   2295
         _Version        =   655364
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "  SCHD HRS              UR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2400
      TabIndex        =   17
      Top             =   360
      Width           =   4545
   End
   Begin VB.Timer tmr_CICO 
      Interval        =   1000
      Left            =   1020
      Top             =   5610
   End
   Begin VB.CheckBox Check2 
      Caption         =   "All Ro for Selected Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2355
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All Open and Current R/O's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   2355
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   5610
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2145
      Left            =   60
      TabIndex        =   1
      Top             =   7020
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   3784
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   98172929
      TitleBackColor  =   8388608
      TitleForeColor  =   16777215
      TrailingForeColor=   13932144
      CurrentDate     =   38458
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
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
      ForeColor       =   &H80000008&
      Height          =   6885
      Left            =   30
      ScaleHeight     =   6885
      ScaleWidth      =   2310
      TabIndex        =   0
      Top             =   60
      Width           =   2310
      Begin Crystal.CrystalReport rptRepairOrder 
         Left            =   120
         Top             =   5550
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport rptNard 
         Left            =   1830
         Top             =   5550
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         Picture         =   "FrmCSMS_ServiceCounter.frx":20D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4200
         Width           =   2265
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5 - Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         MouseIcon       =   "FrmCSMS_ServiceCounter.frx":3154
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMS_ServiceCounter.frx":345E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Refresh"
         Top             =   3360
         Width           =   2265
      End
      Begin VB.CommandButton cmdPartsInquiry 
         Caption         =   "Parts Inquiry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         Picture         =   "FrmCSMS_ServiceCounter.frx":44E0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Inquire for parts avialability"
         Top             =   2520
         Width           =   2265
      End
      Begin VB.CommandButton cmdViewRODetails 
         Caption         =   "View RO Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1380
         Picture         =   "FrmCSMS_ServiceCounter.frx":5562
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5970
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View Job Clock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         MouseIcon       =   "FrmCSMS_ServiceCounter.frx":65E4
         MousePointer    =   99  'Custom
         Picture         =   "FrmCSMS_ServiceCounter.frx":68EE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Time Clock/Job Clock Log-In"
         Top             =   1680
         Width           =   2265
      End
      Begin VB.CommandButton cmdWriteEstimate 
         Caption         =   "Write Estimate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         Picture         =   "FrmCSMS_ServiceCounter.frx":7970
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Create an estimate of Repair"
         Top             =   840
         Width           =   2265
      End
      Begin VB.CommandButton cmdWriteRO 
         Caption         =   "Create Repair Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   30
         Picture         =   "FrmCSMS_ServiceCounter.frx":89F2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Create a reapir order"
         Top             =   0
         Width           =   2265
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   570
         Top             =   5550
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lblHRSCNT 
         BackColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   75
         Top             =   6420
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblURCNT 
         BackColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   74
         Top             =   6210
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lblAppCnt 
         BackColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   73
         Top             =   6000
         Visible         =   0   'False
         Width           =   1305
      End
   End
   Begin VB.Frame frmJobs 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
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
      Height          =   2625
      Left            =   2430
      TabIndex        =   63
      Top             =   6630
      Width           =   12645
      Begin XtremeSuiteControls.TabControl tabDet 
         Height          =   2565
         Left            =   0
         TabIndex        =   64
         Top             =   30
         Width           =   12585
         _Version        =   655364
         _ExtentX        =   22199
         _ExtentY        =   4524
         _StockProps     =   64
         Appearance      =   6
         Color           =   4
         PaintManager.Layout=   2
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.FixedTabWidth=   160
         ItemCount       =   5
         Item(0).Caption =   "View Jobs"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lstJob4Service"
         Item(1).Caption =   "View PMS Jobs"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lstPMSJobs"
         Item(2).Caption =   "View Issued Parts"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "lstParts"
         Item(3).Caption =   "View Issued Materials"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "lstMaterials"
         Item(4).Caption =   "View Issued Accessories"
         Item(4).ControlCount=   1
         Item(4).Control(0)=   "lstAccessories"
         Begin MSComctlLib.ListView lstJob4Service 
            Height          =   2115
            Left            =   90
            TabIndex        =   65
            Top             =   390
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3731
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":9A74
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   3000
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Jobs Description"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Flat Rate"
               Object.Width           =   1886
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Std. Time"
               Object.Width           =   1868
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Technician"
               Object.Width           =   3969
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Hrs. Work"
               Object.Width           =   1956
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "RO"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "TCOde"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "line no"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "Status"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lstPMSJobs 
            Height          =   2115
            Left            =   -69910
            TabIndex        =   66
            Top             =   390
            Visible         =   0   'False
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3731
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":9BD6
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "  Jobs Description"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Model"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstParts 
            Height          =   2115
            Left            =   -69910
            TabIndex        =   67
            Top             =   390
            Visible         =   0   'False
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3731
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":9D38
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Parts Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Price"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Total Amount"
               Object.Width           =   4057
            EndProperty
         End
         Begin MSComctlLib.ListView lstMaterials 
            Height          =   2115
            Left            =   -69910
            TabIndex        =   68
            Top             =   390
            Visible         =   0   'False
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3731
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":9E9A
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Materials Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "QTY"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Price"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Total Amount"
               Object.Width           =   4057
            EndProperty
         End
         Begin MSComctlLib.ListView lstAccessories 
            Height          =   2115
            Left            =   -69910
            TabIndex        =   69
            Top             =   390
            Visible         =   0   'False
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3731
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":9FFC
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Accessories Description"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "QTY"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Price"
               Object.Width           =   4057
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Total Amount"
               Object.Width           =   4057
            EndProperty
         End
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   9315
      Left            =   2370
      TabIndex        =   15
      Top             =   720
      Width           =   12765
      _Version        =   655364
      _ExtentX        =   22516
      _ExtentY        =   16431
      _StockProps     =   64
      Appearance      =   2
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   100
      ItemCount       =   3
      Item(0).Caption =   "Repair Order"
      Item(0).ImageIndex=   0
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "rptRO"
      Item(0).Control(1)=   "Frame1"
      Item(0).Control(2)=   "Prg"
      Item(1).Caption =   "Estimate"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "rptEST"
      Item(1).Control(1)=   "Picture2"
      Item(2).Caption =   "Appointment"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "rptAPP"
      Item(2).Control(1)=   "picAppDet"
      Item(2).Control(2)=   "Picture4"
      Begin XtremeReportControl.ReportControl rptEST 
         Height          =   5175
         Left            =   -69910
         TabIndex        =   23
         Top             =   750
         Visible         =   0   'False
         Width           =   12555
         _Version        =   655364
         _ExtentX        =   22146
         _ExtentY        =   9128
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnReorder=   0   'False
         MultipleSelection=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptRO 
         Height          =   5175
         Left            =   90
         TabIndex        =   18
         Top             =   750
         Width           =   12555
         _Version        =   655364
         _ExtentX        =   22146
         _ExtentY        =   9128
         _StockProps     =   64
         BorderStyle     =   4
         AllowColumnReorder=   0   'False
         MultipleSelection=   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptAPP 
         Height          =   7755
         Left            =   -69910
         TabIndex        =   16
         Top             =   750
         Visible         =   0   'False
         Width           =   9255
         _Version        =   655364
         _ExtentX        =   16325
         _ExtentY        =   13679
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin wizProgBar.Prg Prg 
         Height          =   285
         Left            =   3420
         TabIndex        =   77
         Top             =   3210
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   503
         Picture         =   "FrmCSMS_ServiceCounter.frx":A15E
         ForeColor       =   0
         BarForeColor    =   8454016
         BarPicture      =   "FrmCSMS_ServiceCounter.frx":A17A
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.PictureBox picAppDet 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   7740
         Left            =   -60640
         ScaleHeight     =   7710
         ScaleWidth      =   3315
         TabIndex        =   24
         Top             =   750
         Visible         =   0   'False
         Width           =   3350
         Begin VB.TextBox txtnote 
            BorderStyle     =   0  'None
            ForeColor       =   &H00004000&
            Height          =   2805
            Left            =   15
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Top             =   2340
            Width           =   3285
         End
         Begin VB.Label cboRecd_by 
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
            ForeColor       =   &H00004000&
            Height          =   795
            Left            =   1305
            TabIndex        =   47
            Top             =   5160
            Width           =   1995
         End
         Begin VB.Label txtMake 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   46
            Top             =   930
            Width           =   1995
         End
         Begin VB.Label txtModel 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   45
            Top             =   630
            Width           =   1995
         End
         Begin VB.Label txtApptno 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   44
            Top             =   330
            Width           =   1995
         End
         Begin VB.Label lblLogLoan 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   43
            Top             =   2040
            Width           =   1995
         End
         Begin VB.Label txtDescription 
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
            ForeColor       =   &H00004000&
            Height          =   495
            Left            =   15
            TabIndex        =   42
            Top             =   1530
            Width           =   3285
         End
         Begin VB.Label Label2 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Recieved By"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   795
            Left            =   15
            TabIndex        =   41
            Top             =   5160
            Width           =   1275
         End
         Begin XtremeShortcutBar.ShortcutCaption captionInformation 
            Height          =   315
            Left            =   0
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   0
            Width           =   3345
            _Version        =   655364
            _ExtentX        =   5900
            _ExtentY        =   556
            _StockProps     =   14
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label lblEmails 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Description"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   39
            Top             =   1230
            Width           =   3285
         End
         Begin VB.Label lblLoan 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Notes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   38
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label lblVisits 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Appt No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   37
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label lblLetters 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Make"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   36
            Top             =   930
            Width           =   1275
         End
         Begin VB.Label lblCalls 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Model"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   35
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label txtDte_recd 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   34
            Top             =   5970
            Width           =   1995
         End
         Begin VB.Label Label3 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Promise Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   33
            Top             =   5970
            Width           =   1275
         End
         Begin VB.Label txtKm_rdg 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   32
            Top             =   6570
            Width           =   1995
         End
         Begin VB.Label Label7 
            BackColor       =   &H00D2BDB6&
            Caption         =   " KM Rdg"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   31
            Top             =   6570
            Width           =   1275
         End
         Begin VB.Label txtVIN 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   30
            Top             =   6270
            Width           =   1995
         End
         Begin VB.Label Label9 
            BackColor       =   &H00D2BDB6&
            Caption         =   " VIN No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   29
            Top             =   6270
            Width           =   1275
         End
         Begin VB.Label lblCN2 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   1305
            TabIndex        =   28
            Top             =   6870
            Width           =   1995
         End
         Begin VB.Label Label11 
            BackColor       =   &H00D2BDB6&
            Caption         =   " Contacts"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Left            =   15
            TabIndex        =   27
            Top             =   6870
            Width           =   1275
         End
         Begin VB.Label lblCN1 
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
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   15
            TabIndex        =   26
            Top             =   9330
            Width           =   3285
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   -69940
         ScaleHeight     =   675
         ScaleWidth      =   12645
         TabIndex        =   61
         Top             =   8580
         Visible         =   0   'False
         Width           =   12675
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Uploaded to RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   450
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A196
            MousePointer    =   99  'Custom
            TabIndex        =   71
            ToolTipText     =   "Click to view Repair Order Filter by Parked Status"
            Top             =   360
            Width           =   1440
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   285
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   15045
            _Version        =   655364
            _ExtentX        =   26538
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "Appointment Legend"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Shape Shape9 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   90
            Top             =   360
            Width           =   225
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   -69940
         ScaleHeight     =   675
         ScaleWidth      =   12645
         TabIndex        =   59
         Top             =   8580
         Visible         =   0   'False
         Width           =   12675
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Uploaded to RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   450
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A2E8
            MousePointer    =   99  'Custom
            TabIndex        =   70
            ToolTipText     =   "Click to view Repair Order Filter by Parked Status"
            Top             =   360
            Width           =   1440
         End
         Begin VB.Shape Shape8 
            BorderColor     =   &H00000000&
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   90
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   285
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   15045
            _Version        =   655364
            _ExtentX        =   26538
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "Estimate Legend"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   60
         ScaleHeight     =   675
         ScaleWidth      =   12645
         TabIndex        =   49
         Top             =   8580
         Width           =   12675
         Begin VB.Shape shpvoid 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   11400
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Void"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   11640
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A43A
            MousePointer    =   99  'Custom
            TabIndex        =   76
            ToolTipText     =   "Click to view Repair Order that are already Released"
            Top             =   360
            Width           =   495
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
            Height          =   285
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   15045
            _Version        =   655364
            _ExtentX        =   26538
            _ExtentY        =   503
            _StockProps     =   14
            Caption         =   "Legend (Click on legend to Filter by Status)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00000000&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   90
            Shape           =   1  'Square
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00C0C000&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   2520
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00000000&
            FillColor       =   &H000000C0&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   6960
            Top             =   360
            Width           =   225
         End
         Begin VB.Label labPark 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Park"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   420
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A58C
            MousePointer    =   99  'Custom
            TabIndex        =   57
            ToolTipText     =   "Click to view Repair Order Filter by Parked Status"
            Top             =   360
            Width           =   510
         End
         Begin VB.Label labWork 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Working"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   2790
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A6DE
            MousePointer    =   99  'Custom
            TabIndex        =   56
            ToolTipText     =   "Click to view Repair Order Filter by Working Status"
            Top             =   360
            Width           =   825
         End
         Begin VB.Shape Shape5 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   8460
            Top             =   360
            Width           =   225
         End
         Begin VB.Label labBilled 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Billed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   8730
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A830
            MousePointer    =   99  'Custom
            TabIndex        =   55
            ToolTipText     =   "Click to view Repair Order that are already Billed"
            Top             =   360
            Width           =   570
         End
         Begin VB.Label labOver 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Over"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   7230
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":A982
            MousePointer    =   99  'Custom
            TabIndex        =   54
            ToolTipText     =   "Click to view Repair Order Filter by Due by promised Date"
            Top             =   360
            Width           =   525
         End
         Begin VB.Shape Shape6 
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   5490
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape3 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   9900
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape7 
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   1140
            Top             =   360
            Width           =   225
         End
         Begin VB.Shape Shape 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   3960
            Top             =   360
            Width           =   225
         End
         Begin VB.Label labIdleTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Idle Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   4230
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":AAD4
            MousePointer    =   99  'Custom
            TabIndex        =   53
            ToolTipText     =   "Click to view Repair Order Filter by Ideal Status"
            Top             =   360
            Width           =   870
         End
         Begin VB.Label labBackJob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Back Job"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1425
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":AC26
            MousePointer    =   99  'Custom
            TabIndex        =   52
            ToolTipText     =   "Click to view Repair Order Filter by Back Job Status"
            Top             =   360
            Width           =   885
         End
         Begin VB.Label labFinish 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Finish"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   5745
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":AD78
            MousePointer    =   99  'Custom
            TabIndex        =   51
            ToolTipText     =   "Click to view Repair Order Filter by Finished Job Status"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Released"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   10170
            MouseIcon       =   "FrmCSMS_ServiceCounter.frx":AECA
            MousePointer    =   99  'Custom
            TabIndex        =   50
            ToolTipText     =   "Click to view Repair Order that are already Released"
            Top             =   360
            Width           =   915
         End
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption capHEAD 
      Height          =   330
      Left            =   2370
      TabIndex        =   72
      Top             =   15
      Width           =   12735
      _Version        =   655364
      _ExtentX        =   22463
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Type Keyword Here (F3 - to search)"
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      ForeColor       =   4194304
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Counter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Index           =   1
      Left            =   12810
      TabIndex        =   12
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Counter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   345
      Index           =   0
      Left            =   12780
      TabIndex        =   13
      Top             =   360
      Width           =   2265
   End
   Begin VB.Label LABTIMEIN 
      Caption         =   "Label2"
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   4110
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label LABTIMEOUT 
      Caption         =   "Label2"
      Height          =   225
      Left            =   90
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblROChange 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   60
      TabIndex        =   4
      Top             =   3540
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Service Option"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintRO 
         Caption         =   "Print Repair Order"
      End
      Begin VB.Menu mnuOption1 
         Caption         =   "Add General Job(s)"
      End
      Begin VB.Menu mnuOtherJobs 
         Caption         =   "Add Other Jobs"
      End
      Begin VB.Menu mnuOption13 
         Caption         =   "Add PMS Jobs"
      End
      Begin VB.Menu mnuCanedJob 
         Caption         =   "Add Canned Labor"
      End
      Begin VB.Menu mnuOption2 
         Caption         =   "Edit Repair Order (R/O)"
      End
      Begin VB.Menu mnuViewRODet 
         Caption         =   "View RO Details"
      End
      Begin VB.Menu mnuChangeVehicle 
         Caption         =   "Change Vehicle"
      End
      Begin VB.Menu mnuAsgnedbay 
         Caption         =   "Assign to Bay"
      End
      Begin VB.Menu mnuremovebay 
         Caption         =   "Remove to bay"
      End
      Begin VB.Menu mnuBilledRO 
         Caption         =   "Bill Repair Order"
      End
      Begin VB.Menu mnuOption3 
         Caption         =   "Show Clock In/Out"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBACKJOB 
         Caption         =   "Tag Repair Order as Back Job"
      End
      Begin VB.Menu mnuprintur 
         Caption         =   "Print All UR"
      End
   End
   Begin VB.Menu mnuOption4 
      Caption         =   "Job Option"
      Visible         =   0   'False
      Begin VB.Menu mnuAsgnedTech 
         Caption         =   "Assign Technician"
      End
      Begin VB.Menu mnuChangeTech 
         Caption         =   "Change Technician"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemoveTech 
         Caption         =   "Remove Technician"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuAsscontractor 
         Caption         =   "Assign &Contractor"
      End
      Begin VB.Menu mnuJobDone 
         Caption         =   "Tag Jobs Done"
      End
      Begin VB.Menu mnuOption2_2 
         Caption         =   "Remove Job(s)"
      End
   End
   Begin VB.Menu mnuAppointment 
      Caption         =   "Appointment Option"
      Visible         =   0   'False
      Begin VB.Menu mnuCreateApp 
         Caption         =   "Create Appointment"
      End
      Begin VB.Menu mnuEditAppoitnment 
         Caption         =   "Edit Selected Appointment"
      End
      Begin VB.Menu mnuUplodAppointment 
         Caption         =   "Upload Selected Appointment to RO"
      End
      Begin VB.Menu mnuDeleteAppointment 
         Caption         =   "Delete Appointment"
      End
      Begin VB.Menu mnuPrintAppointment 
         Caption         =   "Print Selected Appointment"
      End
   End
   Begin VB.Menu mnuEstimate 
      Caption         =   "Estimate Option"
      Visible         =   0   'False
      Begin VB.Menu mnuPEstimate 
         Caption         =   "Print Estimate"
      End
      Begin VB.Menu mnuAddEstJob 
         Caption         =   "Add Job"
      End
      Begin VB.Menu mnuAddEstDet 
         Caption         =   "Add Part/ Materials/ Accessories"
      End
      Begin VB.Menu mnuUEstimate 
         Caption         =   "Upload Estimate"
      End
      Begin VB.Menu mnuDeleteEst 
         Caption         =   "Delete Estimate"
      End
      Begin VB.Menu mnudelete_estjob 
         Caption         =   "Delete Estimate Job"
      End
   End
End
Attribute VB_Name = "frmCSMS_ServiceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCicoCount                                         As ADODB.Recordset
Dim AUDIT_SQL                                           As String
Dim thestatus                                           As String
Dim theRo                                               As String
Dim Thedate                                             As Date
Dim tlHrs                                               As Double
Dim tlFR                                                As Double
Dim bevvy                                               As Long
Dim CHKSTATUS                                           As String
Dim zRONO                                               As String
Dim thetechcode                                         As String
Dim ISLOGIN                                             As Boolean
Dim theflatrate                                         As Double
Dim THESTDRATE                                          As Double
Dim THEJOBCODE                                          As String
Dim THEJOBDEST                                          As String
Dim PERJOBSTATUS                                        As String
Dim vlineNo                                             As String
Dim XREPAIRORDER                                        As String
Dim XPREVIOUS_DATE                                      As String
Dim WithEvents JOBCLOCKFORM                             As frmCSMSClockINOUT
Attribute JOBCLOCKFORM.VB_VarHelpID = -1
Dim WithEvents frm                                      As frmCSMSEditRO
Attribute frm.VB_VarHelpID = -1
Dim WithEvents FRMx                                     As frmCSMS_MasterStockInquiry
Attribute FRMx.VB_VarHelpID = -1
Dim WithEvents frmApp                                   As frmCSMS_UploadEstimate
Attribute frmApp.VB_VarHelpID = -1
Dim WithEvents FRM_EST                                  As frmCSMS_MasterEstimateDet
Attribute FRM_EST.VB_VarHelpID = -1
Dim public_ESTNO                                        As String
Dim xESTNO                                              As String

Function Click_ScheduleGrid(Optional xRO As String)
    Call DisplayRODetails(xRO)
    'grdCounter_RowColChange grdCounter.ActiveCell.Row, grdCounter.ActiveCell.Col
End Function

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check2.Value = 0
        CHKSTATUS = "All"
        Call cmdRefresh_Click
    End If
    If Check1.Value = 0 Then
        Check2.Value = 1
        CHKSTATUS = "All"
        Call cmdRefresh_Click
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 0
        CHKSTATUS = "All"
    End If
    If Check2.Value = 0 Then
        Check1.Value = 1
        CHKSTATUS = "All"
    End If
End Sub

Function CheckAllJobsISDone(XXX As Variant) As Boolean
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DONE  FROM CSMS_RO_DET WHERE LIVIL = '1' AND (DONE ='N' OR DONE='W' OR DONE IS NULL) and REP_OR = '" & XXX & "'")
    If RS.EOF And RS.BOF Then
        CheckAllJobsISDone = True
    Else
        CheckAllJobsISDone = False
    End If
    Set RS = Nothing
End Function

Function CheckIfJobIsFinish(vLINE_NO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT DONE FROM CSMS_RO_dET WHERE REP_OR = '" & theRo & "' and line_no = '" & vLINE_NO & "' AND LIVIL = '1'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If Null2String(rstmp!DONE) = "Y" Then
            CheckIfJobIsFinish = True
        Else
            CheckIfJobIsFinish = False
        End If
    End If
    Set rstmp = Nothing
End Function

Sub CleanListViewDetails()
    lstJob4Service.ListItems.Clear
    lstPMSJobs.ListItems.Clear
    lstParts.ListItems.Clear
    lstMaterials.ListItems.Clear
    lstAccessories.ListItems.Clear
End Sub

Sub ClearOrStayTechnician(vEMPNO As String, vRONO As String, VTECHCODE As String)
    Dim rsHRMS                                         As New ADODB.Recordset
    Dim rsDet                                          As New ADODB.Recordset
    Dim rstmp                                          As New ADODB.Recordset
    Dim X                                              As Integer

    Set rsHRMS = gconDMIS.Execute("SELECT EMPNO FROM HRMS_EMPINFO WHERE EMPNO = '" & vEMPNO & "'")
    If Not (rsHRMS.BOF And rsHRMS.EOF) Then
        Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND (DONE IS NULL OR DONE <> 'Y')")
        If Not (rsDet.BOF And rsDet.EOF) Then
            Do While Not rsDet.EOF
                X = X + 1
                rsDet.MoveNext
            Loop
            If X > 1 Then
                Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (rstmp.BOF And rstmp.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set rstmp = Nothing
            Else
                gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET JSTATUS = 'A', ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'")
            End If
        End If
    Else
        Set rsDet = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & VTECHCODE & "' AND (DONE IS NULL OR DONE <> 'Y')")
        If Not (rsDet.BOF And rsDet.EOF) Then
            Do While Not rsDet.EOF
                X = X + 1
                rsDet.MoveNext
            Loop
            If X > 1 Then
                Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (rstmp.BOF And rstmp.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    'NEW LOG AUDIT-----------------------------------------------------
                        Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set rstmp = Nothing
            Else
                gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS = 'A', ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'")
            End If
        End If
    End If

    Set rsDet = Nothing
End Sub

Private Sub cmdDateBack_Click()
    MonthView1 = MonthView1 - 1
    Check2.Value = 1
    cmdRefresh.Value = True
End Sub

Private Sub cmdDateForward_Click()
    MonthView1 = MonthView1 + 1
    Check2.Value = 1
    cmdRefresh.Value = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPartsInquiry_Click()
    If Module_Access(LOGID, "PARTS INQUIRY", "INQUIRY") = False Then Exit Sub
    
    Call FRMx.SetType("P", "Parts Stock Inquiry")
    FRMx.Show
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = 11
    txtSearch.Text = ""
    
    Call CleanListViewDetails
    
    theRo = "":    CHKSTATUS = "All"
    DoEvents
    Call ProcesUpdate
    Call ViewActiveRO
    Call ViewEstimate
    Call ViewAppointment
    DoEvents
    
    Call ComputeMeCTR
    Screen.MousePointer = 0
End Sub

Private Sub cmdToday_Click()
    MonthView1.Value = Format(Now, "MM/dd/yyyy")
    Check2.Value = 1
    cmdRefresh.Value = True
End Sub

Private Sub cmdsms_Click()
 'If Module_Access(LOGID, "SMS", "TRANSACTION") = False Then Exit Sub
' If COMPANY_CODE = "HCA" Then
'   FrmCSMS_SMS.Show 1
'End If
End Sub

Private Sub cmdViewRODetails_Click()
    Dim XRONO                                           As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to View", vbInformation, "Info."
        Exit Sub
    End If
    
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    frmCSMSViewRO.labRO.Caption = XRONO
    frmCSMSViewRO.Show 1
End Sub

Private Sub cmdWriteEstimate_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_ADD", "JOB ESTIMATE") = False Then Exit Sub
    
    frmCSMSNewAppointment.labType(0).Caption = "Estimate"
    frmCSMSNewAppointment.labType(1).Caption = "Estimate"
    frmCSMSNewAppointment.GetDefaultTransactionType
    Timer1.Enabled = False
    frmCSMSNewAppointment.Show 1
    Timer1.Enabled = True
End Sub

Private Sub cmdWriteRO_Click()
    'UPDATE BY   : MJP 11082009 0200PM
    'DESCRIPTION : CRF 108
'        If COMPANY_CODE = "HGC" Then
'            If CheckIfUserIsAnServiceAdviser = False Thengoo
'                MsgBox "Your User Account dont have the access to create an Repair Order", vbInformation, "Info."
'                Exit Sub
'            End If
'        End If
    'UPDATE BY   : MJP 11082009 0200PM
    
    frmCSMSNewAppointment.labType(0) = "Repair Order"
    frmCSMSNewAppointment.labType(1) = "Repair Order"
    frmCSMSNewAppointment.GetDefaultTransactionType
    frmCSMSNewAppointment.Show 1
End Sub


Private Sub Command1_Click()
    JOBCLOCKFORM.Show 1
End Sub

Private Sub Command10_Click()
    frmCSMSShowTechnician.Show 1
End Sub

Private Sub Command12_Click()
    Call ViewActiveRO
End Sub

Private Sub Command5_Click()
    frmCSMSReqJobs.Show 1
End Sub

Private Sub Command7_Click()
    frmCSMSPMS.Show 1
End Sub

Sub ComputeMeCTR()
    Dim rstmp                                   As New ADODB.Recordset
    If Check1.Value = 1 Then
        If CHKSTATUS = "All" Then
            Set rstmp = gconDMIS.Execute("Select sum(isnull([Hours],0)) as TOT_HRS " & _
                " from CSMS_vw_RepairOrder " & _
                " where (TransType = 'R' " & _
                " and status <> 'Released') ")
        Else
            Set rstmp = gconDMIS.Execute("Select sum(isnull([Hours],0)) as TOT_HRS " & _
                " from CSMS_vw_RepairOrder " & _
                " where TransType = 'R' " & _
                " and Status = '" & CHKSTATUS & "'")
        End If
    ElseIf Check2.Value = 1 Then
        If CHKSTATUS = "All" Then
            Set rstmp = gconDMIS.Execute("Select sum(isnull([Hours],0)) as TOT_HRS " & _
                " from CSMS_vw_RepairOrder " & _
                " where TransType = 'R' " & _
                " and (AppointmentDate = '" & DateValue(MonthView1) & "')")
        Else
            Set rstmp = gconDMIS.Execute("Select sum(isnull([Hours],0)) as TOT_HRS " & _
                " from CSMS_vw_RepairOrder " & _
                " where TransType = 'R' " & _
                " and AppointmentDate = '" & DateValue(MonthView1) & _
                "' and Status = '" & CHKSTATUS & "'")
        End If
    End If
    If Not (rstmp.BOF And rstmp.EOF) Then
        If NumericVal(rstmp!tot_hrs) = 0 Then
            txtTLhrs.Text = "0.00"
        Else
            txtTLhrs.Text = Format(rstmp!tot_hrs, MAXIMUM_DIGIT)
        End If
    Else
        txtTLhrs.Text = "0.00"
    End If
    lblHRSCNT.Caption = txtTLhrs
    Set rstmp = Nothing
End Sub

Sub ComputeMePo()
    tlHrs = 0: tlFR = 0
    For bevvy = 1 To Me.lstJob4Service.ListItems.Count
        tlHrs = tlHrs + NumericVal(lstJob4Service.ListItems(bevvy).SubItems(3))
        tlFR = tlFR + NumericVal(lstJob4Service.ListItems(bevvy).SubItems(2))
    Next bevvy

    gconDMIS.Execute "update CSMS_RepairOrder set hours = " & tlHrs & " where Ro_no = '" & zRONO & "'"
End Sub

Sub EnableDropDownMenu(COND As Boolean)
    mnuOption1.Enabled = COND
    mnuOtherJobs.Enabled = COND
    mnuOption13.Enabled = COND
    mnuCanedJob.Enabled = COND
    mnuOption2.Enabled = COND
    mnuAsgnedbay.Enabled = COND
    mnuremovebay.Enabled = COND
    mnuBilledRO.Enabled = COND
    mnuViewRODet.Enabled = COND
    mnuChangeVehicle.Enabled = COND
    mnuBACKJOB.Enabled = Not COND
End Sub

Function FindTechName(SACODE As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Dim RSCON                                          As New ADODB.Recordset
    Dim rsVEN                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & SACODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindTechName = Null2String(rstmp!TECH_NAME)
    Else
        Set RSCON = gconDMIS.Execute("SELECT COMPANYNAME FROM CSMS_CONTRACTOR WHERE CODE = '" & SACODE & "'")
        If Not (RSCON.BOF And RSCON.EOF) Then
            FindTechName = Null2String(RSCON!CompanyName)
        Else
            Set rsVEN = gconDMIS.Execute("SELECT CODE,NAMEOFVENDOR FROM ALL_VENDOR_TABLE WHERE CODE = '" & SACODE & "'")
            If Not (rsVEN.BOF And rsVEN.EOF) Then
                FindTechName = Null2String(rsVEN!nameofvendor)
            Else
                FindTechName = ""
            End If
        End If
        Set RSCON = Nothing
    End If

    Set rstmp = Nothing
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyPageUp Then
            Me.Top = Me.Top - 200
        End If

        If KeyCode = vbKeyPageDown Then
            Me.Top = Me.Top + 200
        End If
    End If

    'UPDATE BY   : MJP 02122009 0938AM
    'DESCRIPTION : TO LIMIT THE VB MODAL ERROR (TCN 12711)
        If frmCSMSNewAppointment.Visible = True Then Exit Sub
        If frmCSMSClockINOUT.Visible = True Then Exit Sub
    'UPDATE BY   : MJP 02122009 0938AM
        
    If KeyCode = vbKeyF5 Then
        cmdRefresh_Click
    ElseIf KeyCode = vbKeyF3 Then
        txtSearch.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    
    Call CenterMe(frmMain, Me, 0)
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set JOBCLOCKFORM = New frmCSMSClockINOUT
    Set FRMx = New frmCSMS_MasterStockInquiry
    Set frm = New frmCSMSEditRO
    Set frmApp = New frmCSMS_UploadEstimate
    Set FRM_EST = New frmCSMS_MasterEstimateDet
    
    MonthView1.Value = Format(Now, "MM/dd/yyyy")
    
    'UPDATE BY   : MJP09222009 0430PM
    'DESCRIPTION : TO INILIZE REPORT CONTROL
        Call InitializeReportControl
    'UPDATE BY   : MJP09222009 0430PM
    
    'Call InitGrid
    Check1.Value = 1:               CHKSTATUS = "All":
    Screen.MousePointer = 0
    
'    Dim rsCico                                         As new ADODB.Recordset
'    Set rsCico = gconDMIS.Execute("SELECT MAX(CLOCKIN) KEYIN, MAX(CLOCKOUT) KEYCOUT FROM CSMS_JOBCLOCK")
'    If Not rsCico.EOF Or rsCico.BOF Then
'        LABTIMEIN = rsCico!KEYIN
'        LABTIMEOUT = rsCico!KEYCOUT
'        tmr_CICO.Enabled = True
'    End If

    mnuOption.Visible = False
    mnuOption4.Visible = False
    'mnuEstimate.Visible = False
    
'    If COMPANY_CODE = "HCA" Then
'        cmdsms.Enabled = True
'    Else
'        cmdsms.Enabled = False
'    End If
End Sub

Private Sub FRM_EST_AddDetails(xESTNO As String, xCHANGE As Integer, xID As Long)
    Unload FRM_EST
    
    If xCHANGE = 1 Then
        Screen.MousePointer = 11
        Dim rsEsti_Det                              As New ADODB.Recordset
        Dim PartsComTotal                           As Double
        Dim TOTPARTSAMT                             As Double
        Dim PartsSalesTotal                         As Double
        Dim PartsWarTotal                           As Double
        Dim TOTPARTSDISC                            As Double
        Dim TOTPARTSDISCVAL                         As Double
        Dim TOTPARTSTAX                             As Double
        Dim JobComTotal                             As Double
        Dim JobSalesTotal                           As Double
        Dim JobWarTotal                             As Double
        Dim TOTJOBAMT                               As Double
        Dim TOTJOBDISC                              As Double
        Dim TOTJOBDISCVAL                           As Double
        Dim TOTJOBTAX                               As Double
        Dim MatComTotal                             As Double
        Dim MatSalesTotal                           As Double
        Dim MatWarTotal                             As Double
        Dim TOTMATAMT                               As Double
        Dim TOTMATDISC                              As Double
        Dim TOTMATDISCVAL                           As Double
        Dim TOTMATTAX                               As Double
        Dim ACCComTotal                             As Double
        Dim ACCSalesTotal                           As Double
        Dim ACCWarTotal                             As Double
        Dim TOTACCAMT                               As Double
        Dim TOTACCDISC                              As Double
        Dim TOTACCDISCVAL                           As Double
        Dim TOTACCTAX                               As Double
        Dim ROTotal                                 As Double
        
        Set rsEsti_Det = gconDMIS.Execute("select discount_2, det_amt, wcode, disval, taxval from CSMS_EstDETAILS " & _
            " where estimateno = '" & xESTNO & "' and livil = '1' " & " order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            Do While Not rsEsti_Det.EOF
                If Null2String(rsEsti_Det!wCode) = "C" Then
                    JobComTotal = JobComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "S" Then JobSalesTotal = JobSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "W" Then JobWarTotal = JobWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                Else
                    TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                    TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                    TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                    TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
                End If
                
                rsEsti_Det.MoveNext
            Loop
        End If
            
        Set rsEsti_Det = gconDMIS.Execute("select det_amt, wcode, discount_2, disval, taxval from CSMS_EstDETAILS where " & _
            " estimateno = '" & xESTNO & "' And livil = '2' order by LINE_NO asc")
        If Not (rsEsti_Det.EOF And rsEsti_Det.BOF) Then
            Do While Not rsEsti_Det.EOF
                If Null2String(rsEsti_Det!wCode) = "C" Then
                    PartsComTotal = PartsComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "S" Then PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "W" Then PartsWarTotal = PartsWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                Else
                    TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                    TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                    TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                    TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
                End If
                rsEsti_Det.MoveNext
            Loop
        End If
        
        Set rsEsti_Det = gconDMIS.Execute("select det_amt, wcode, discount_2, disval, taxval from CSMS_EstDETAILS where " & _
            " estimateno = '" & xESTNO & "' and livil = '3' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            Do While Not rsEsti_Det.EOF
                If Null2String(rsEsti_Det!wCode) = "C" Then
                    MatComTotal = MatComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "S" Then MatSalesTotal = MatSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "W" Then MatWarTotal = MatWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                Else
                    TOTMATAMT = TOTMATAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                    TOTMATDISC = TOTMATDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                    TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                    TOTMATTAX = TOTMATTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
                End If
                rsEsti_Det.MoveNext
            Loop
        End If
        
        Set rsEsti_Det = gconDMIS.Execute("select det_amt, wcode, discount_2, disval, taxval from CSMS_EstDETAILS where " & _
            " estimateno = '" & xESTNO & "' and livil = '4' order by LINE_NO asc")
        If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
            Do While Not rsEsti_Det.EOF
                If Null2String(rsEsti_Det!wCode) = "C" Then
                    ACCComTotal = ACCComTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "S" Then ACCSalesTotal = ACCSalesTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                ElseIf Null2String(rsEsti_Det!wCode) = "W" Then ACCWarTotal = ACCWarTotal + N2Str2Zero(rsEsti_Det!DET_AMT)
                Else
                    TOTACCAMT = TOTACCAMT + N2Str2Zero(rsEsti_Det!DET_AMT)
                    TOTACCDISC = TOTACCDISC + N2Str2Zero(rsEsti_Det!Discount_2)
                    TOTACCDISCVAL = TOTACCDISCVAL + N2Str2Zero(rsEsti_Det!disval)
                    TOTACCTAX = TOTACCTAX + N2Str2Zero(rsEsti_Det!TAXVAL)
                End If
                rsEsti_Det.MoveNext
            Loop
        End If
        
        ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT + TOTACCAMT
        SQL_STATEMENT = "update CSMS_EstHd set" & _
            " labor = " & TOTJOBAMT - TOTJOBTAX & _
            ", l_amtvalue = " & TOTJOBAMT & ", l_disc = " & TOTJOBDISCVAL & _
            ", l_disc2 = " & TOTJOBDISC * (VAT_RATE / 100) & ", l_taxval = " & TOTJOBTAX & _
            ", l_discount = " & TOTJOBDISC & _
            ", PARTS = " & TOTPARTSAMT - TOTPARTSTAX & _
            ", P_amtvalue = " & TOTPARTSAMT & ", P_disc = " & TOTPARTSDISCVAL & _
            ", P_disc2 = " & TOTPARTSDISC * (VAT_RATE / 100) & ", P_taxval = " & TOTPARTSTAX & _
            ", P_discount = " & TOTPARTSDISC & _
            ", MATERIAL = " & TOTMATAMT - TOTPARTSTAX & _
            ", M_amtvalue = " & TOTMATAMT & ", M_disc = " & TOTMATDISCVAL & _
            ", M_disc2 = " & TOTMATDISC * (VAT_RATE / 100) & ", M_taxval = " & TOTMATTAX & _
            ", M_discount = " & TOTMATDISC & _
            ", ACCESSORIES = " & TOTACCAMT - TOTPARTSTAX & _
            ", A_amtvalue = " & TOTACCAMT & ", A_disc = " & TOTACCDISCVAL & _
            ", A_disc2 = " & TOTACCDISC * (VAT_RATE / 100) & ", A_taxval = " & TOTACCTAX & _
            ", A_discount = " & TOTACCDISC & _
            ", AMOUNT = " & ROTotal - (TOTJOBDISC + TOTPARTSDISC + TOTMATDISC + TOTACCDISC) & _
            ", ROVAT = " & (TOTJOBTAX + TOTPARTSTAX + TOTMATTAX + TOTACCTAX) & _
            ", WL_AMT = " & 0 & _
            ", RO_AMOUNT = " & ROTotal - (TOTJOBDISC + TOTPARTSDISC + TOTMATDISC + TOTACCDISC) & _
            " where ID = " & xID
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT----------------------------------------------------------
            Call NEW_LogAudit("E", "JOB ESTIMATE", SQL_STATEMENT, CStr(xID), "", "EST NO: " & xESTNO, "", "")
        'NEW LOG AUDIT----------------------------------------------------------
        Call FillEstimateDetails(public_ESTNO)
        
        Screen.MousePointer = 0
    End If
End Sub

Private Sub frm_SaveEditRO()
    Unload frm
    Call cmdRefresh_Click
    
    MessagePop InfoFriend, "RO Information Updated", "Repair Order Information Sucessfully Updated!", 1000
End Sub

Private Sub grdCounter_DblClick()
    Call cmdViewRODetails_Click
End Sub

Sub IsTechLogIn()
    Dim RS                                             As New ADODB.Recordset

    Set RS = gconDMIS.Execute("SELECT * FROM CSMS_vw_technicianAvailability WHERE Techcode = '" & thetechcode & "' and AssignedRo='" & theRo & "'")
    Do While Not RS.EOF
        If Null2String(RS!Code) = "W" Then
            ISLOGIN = True
        End If
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub frmApp_PutEstimateNo(ByVal xESTNO As String, FromForm As String)
    If FromForm = "SERVICE COUNTER" Then
        Call cmdRefresh_Click
        
        Unload frmApp
    End If
End Sub

Private Sub JOBCLOCKFORM_FORMCLOSED()
    cmdRefresh.Value = True
End Sub

Private Sub JOBCLOCKFORM_JOBCLOCKED()
    Call ProcesUpdate
End Sub

Private Sub labBackJob_Click()
    CHKSTATUS = "Back Job"
    Call ViewActiveRO
End Sub

Private Sub labBilled_Click()
    CHKSTATUS = "Billed"
    Call ViewActiveRO
End Sub

Private Sub Label1_Click()
    CHKSTATUS = "Released"
    Call ViewActiveRO
End Sub

Private Sub Label6_Click()
    CHKSTATUS = "Voided"
    Call ViewActiveRO
End Sub

Private Sub labFinish_Click()
    CHKSTATUS = "Finish job"
    Call ViewActiveRO
End Sub

Private Sub labIdleTime_Click()
    CHKSTATUS = "Idle Time"
     Call ViewActiveRO
End Sub

Private Sub labOver_Click()
    CHKSTATUS = "Over"
     Call ViewActiveRO
End Sub

Private Sub labPark_Click()
    CHKSTATUS = "Park"
    Call ViewActiveRO
End Sub

Private Sub labWork_Click()
    CHKSTATUS = "Working"
    Call ViewActiveRO
End Sub

Private Sub lstCounter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOption
    End If
End Sub

Private Sub lblSearch_Click()

End Sub

Private Sub lstJob4Service_Click()
    If TabControl.SelectedItem <> 0 Then Exit Sub
    'Update by BTT/RSC
    If lstJob4Service.SelectedItem Is Nothing Then Exit Sub
    
    THEJOBCODE = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index)
    THEJOBDEST = lstJob4Service.SelectedItem.SubItems(1)
    theflatrate = NumericVal(lstJob4Service.SelectedItem.SubItems(2))
    THESTDRATE = NumericVal(lstJob4Service.SelectedItem.SubItems(3))
    vlineNo = lstJob4Service.SelectedItem.SubItems(8)

    PERJOBSTATUS = lstJob4Service.SelectedItem.SubItems(9)
End Sub

'UPDATE BY   : MJP 11082009 0254 PM
'DESCRIPTION : CRF 121
    Private Sub lstJob4Service_DblClick()
'        If Not lstJob4Service.ListItems.Count = 0 Then
'            Dim Index                               As Integer
'
'            Index = lstJob4Service.SelectedItem.Index
'            If CheckJobStatusIfIdle(Null2String(rptRO.SelectedRows(0).Record(4).Value), lstJob4Service.ListItems(Index).Text) = "I" Then
'                Dim RSTMP                           As New ADODB.Recordset
'                Dim rsCLOCK                         As New ADODB.Recordset
'                Dim xPREV_TECH                      As String
'                Null2String (rptRO.SelectedRows(0).Record(4).Value)
'                Set RSTMP = gconDMIS.Execute("SELECT TECHNICIAN FROM CSMS_RO_DET WHERE " & _
'                    " REP_OR = '" & Null2String(rptRO.SelectedRows(0).Record(4).Value) & _
'                    "' AND LIVIL = 1 " & _
'                    " AND DETCDE = " & N2Str2Null(lstJob4Service.ListItems(Index).Text) & "")
'                If Not (RSTMP.BOF And RSTMP.EOF) Then
'                    Set rsCLOCK = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_JOBCLOCK WHERE " & _
'                        " RO_NO = '" & Null2String(rptRO.SelectedRows(0).Record(4).Value) & _
'                        "' AND DETCDE = " & N2Str2Null(lstJob4Service.ListItems(Index).Text) & _
'                        " ORDER BY ID DESC")
'                    If Not (rsCLOCK.BOF And rsCLOCK.EOF) Then
'                        xPREV_TECH = Null2String(rsCLOCK!TECH_NAME)
'                    End If
'                    Set rsCLOCK = Nothing
'                End If
'                Set RSTMP = Nothing
'
'                MessagePop InfoFriend, "Previous Techncian", Null2String(xPREV_TECH), 1000
'            End If
'        End If
    End Sub
'UPDATE BY   : MJP 11082009 0254 PM

Private Sub lstJob4Service_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If TabControl.SelectedItem <> 0 Then Exit Sub
    If lstJob4Service.ListItems.Count = 0 Then Exit Sub

    lstJob4Service.ToolTipText = Null2String(Item.ListSubItems(1))
End Sub

Private Sub lstJob4Service_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If TabControl.SelectedItem = 0 Or TabControl.SelectedItem = 1 Then
    Else
        Exit Sub
    End If
    Dim i As String
    i = vbRightButton
    
    If Button = vbRightButton Then
        If TabControl.SelectedItem = 0 Then
            If CheckIfRoIsAlreadyInvoice(Null2String(rptRO.SelectedRows(0).Record(4).Value)) = True Then
                If Module_Access(LOGID, "EDIT JOB DESCRIPTIONS", "SYSTEM") = False Then Exit Sub
    
                mnuAsgnedTech.Enabled = False
                MnuAsscontractor.Enabled = False
                mnuOption2_2.Enabled = False
                'UPDATED BY: JUN-----------------------------------
                'DATE UPDATED: 1-29-2008
                'DESCRIPTION: DISABLED THE JOBDONE IN ORDER IF IT IS ALREADY BILLED OR RELEASE SO THAT IT WILL NOT HAVE AN UPDATE OF RO AMOUNT WHICH IS ALREADY BILLED
                    If COMPANY_CODE = "HPI" Then
                        mnuJobDone.Enabled = True
                    Else
                        mnuJobDone.Enabled = False
                    End If
                'UPDATED BY: JUN-----------------------------------
                
                PopupMenu mnuOption4
            Else
                If lstJob4Service.ListItems.Count = 0 Or lstJob4Service.SelectedItem Is Nothing Then: Exit Sub
    
                If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HCI" Then
                    MnuAsscontractor.Visible = False
                Else
                    MnuAsscontractor.Visible = True
                End If
    
                mnuAsgnedTech.Enabled = True
                MnuAsscontractor.Enabled = True
                mnuOption2_2.Enabled = True
                
                'UPDATED BY: JUN-----------------------------------
                'DATE UPDATED: 1-29-2008
                'DESCRIPTION: ENABLE JOB DONE IF REPAIR ORDER IS NOT YET INVOICE OR RELEASED.
                 mnuJobDone.Enabled = True
                'UPDATED BY: JUN-----------------------------------
                
                Call lstJob4Service_Click
                thetechcode = lstJob4Service.SelectedItem.SubItems(7)
                IsTechLogIn
                PopupMenu mnuOption4
            End If
        ElseIf TabControl.SelectedItem = 1 And Null2String(rptEST.SelectedRows(0).Record(5).Value) > "" Then
            If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then Exit Sub
            If lstJob4Service.SelectedItem.SubItems(1) <> "" Then
                mnuPEstimate.Enabled = False
                mnuAddEstJob.Enabled = False
                mnuAddEstDet.Enabled = False
                mnuUEstimate.Enabled = False
                mnuDeleteEst.Enabled = False
                mnudelete_estjob.Enabled = True
                PopupMenu mnuEstimate
            End If
        End If

    End If
End Sub

Private Sub mnuAddEstDet_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_add", "JOB ESTIMATE") = False Then Exit Sub
    
    If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then
        MsgBox "You cannot add details to a already uploaded Estimate", vbInformation, "Info."
    Else
        If CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "UPLOADED" Then
            MsgBox "Estimate already uploaded to repair order." & vbCrLf & "kindly refresh your service counter to display fresh data", vbInformation, "Info."
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT FOUND" Then
            MsgBox "Estimate Record not found. kindly refresh your service counter to display fresh data", vbCritical, "Info"
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT UPLOADED" Then
            Call FRM_EST.SetType(Null2String(rptEST.SelectedRows(0).Record(5).Value), Null2String(rptEST.SelectedRows(0).Record(10).Value))
            FRM_EST.Show 1
        End If
    End If
End Sub
Private Sub mnudelete_estjob_click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_DELETE", "JOB ESTIMATE") = False Then Exit Sub
Dim sqlcommand                                    As String
    If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then
        MsgBox "You cannot add details to a already uploaded Estimate", vbInformation, "Info."
    Else
        If CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "UPLOADED" Then
            MsgBox "Estimate already uploaded to repair order." & vbCrLf & "kindly refresh your service counter to display fresh data", vbInformation, "Info."
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT FOUND" Then
            MsgBox "Estimate Record not found. kindly refresh your service counter to display fresh data", vbCritical, "Info"
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT UPLOADED" Then
'            Call FRM_EST.SetType(Null2String(rptEST.SelectedRows(0).Record(5).Value), Null2String(rptEST.SelectedRows(0).Record(10).Value))
'            FRM_EST.Show 1
        sqlcommand = ""
        sqlcommand = "DELETE FROM CSMS_ESTDETAILS WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "' and Line_no = '" & Null2String(lstJob4Service.SelectedItem.SubItems(8)) & "' and livil = '1' and detcde = '" & Null2String(lstJob4Service.SelectedItem) & "'"
        gconDMIS.Execute (sqlcommand)
        
        sqlcommand = ""
        sqlcommand = "Delete from csms_ro_det where rep_or ='" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "' and livil = '1' and line_no = '" & Null2String(lstJob4Service.SelectedItem.SubItems(8)) & "' and detcde ='" & Null2String(lstJob4Service.SelectedItem) & "'"
        gconDMIS.Execute (sqlcommand)
        Call ShowDeletedMsg
        Call cmdRefresh_Click
        Call FillEstimateDetails(xESTNO)
        End If
    End If
End Sub

Private Sub mnuAddEstJob_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_add", "JOB ESTIMATE") = False Then Exit Sub
    
    If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then
        MsgBox "You cannot add details to a already uploaded Estimate", vbInformation, "Info."
    Else
        If CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "UPLOADED" Then
            MsgBox "Estimate already uploaded to repair order." & vbCrLf & "kindly refresh your service counter to display fresh data", vbInformation, "Info."
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT FOUND" Then
            MsgBox "Estimate Record not found. kindly refresh your service counter to display fresh data", vbCritical, "Info"
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT UPLOADED" Then
'            Call FRM_EST.SetType(Null2String(rptEST.SelectedRows(0).Record(5).Value), Null2String(rptEST.SelectedRows(0).Record(10).Value))
'            FRM_EST.Show 1
    If xESTNO = "" Then
        MsgBox "Choose a Estimate to add Other job", vbInformation, "Info."
        Exit Sub
    End If

    Dim xCUSTOMERNAME                                           As String
    Dim xACCTNO                                                 As String
    Dim xEST                                                   As String

    xCUSTOMERNAME = Null2String(rptEST.SelectedRows(0).Record(1).Value)
    xACCTNO = Null2String(rptEST.SelectedRows(0).Record(9).Value)
    xEST = Null2String(rptEST.SelectedRows(0).Record(5).Value)
    
    With frmCSMS_EstimateAddJob
        .txtCustomer.Text = xCUSTOMERNAME
        .txtActNo.Text = xACCTNO
        .txtROno.Text = xEST
        .txtCheckMe.Text = "est"
    End With
    
    frmCSMS_EstimateAddJob.Show 1
    Call DisplayRODetails(xEST)
        End If
    End If
End Sub

Private Sub mnuAsgnedbay_Click()
    Dim XRONO                                           As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Please Select a Repair Order to be Assigned with technician.", vbInformation, "Info."
        Exit Sub
    End If
    
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)

    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    With frmCSMSUpdatebayInfo
        .labRO.Caption = XRONO
    End With
    frmCSMSUpdatebayInfo.Show 1
End Sub

Private Sub mnuAsgnedTech_Click()
    Dim Index                                          As Integer

    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Please Select a Repair Order to be Assigned with technician.", vbInformation, "Info."
        Exit Sub
    End If

    Index = lstJob4Service.SelectedItem.Index
    If Null2String(lstJob4Service.SelectedItem.SubItems(9)) = "Finish Job" Then
        MsgBox "Job Is Already Finish", vbInformation, "Info"
        Exit Sub
    End If

    Dim xCUSTOMERNAME                                   As String
    Dim xACCTNO                                         As String
    Dim XRONO                                           As String
    
    xCUSTOMERNAME = Null2String(rptRO.SelectedRows(0).Record(1).Value)
    xACCTNO = Null2String(rptRO.SelectedRows(0).Record(13).Value)
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
   
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided Cannot Assign Technician", vbInformation + vbOKOnly
        Exit Sub
    End If
    If Null2String(lstJob4Service.SelectedItem.SubItems(4)) = "" Then
        With frmCSMSUpdateCustomerInfo
            .labRO.Caption = XRONO
            .lblJobCode.Caption = Null2String(lstJob4Service.SelectedItem.Text)
            .labCust.Caption = xACCTNO
            .LABITEMNO.Caption = Null2String(lstJob4Service.SelectedItem)
        End With
        frmCSMSUpdateCustomerInfo.Show 1
        Call DisplayRODetails(XRONO)
    Else
        If CheckIfJobIsFinish(lstJob4Service.ListItems(Index).ListSubItems(8)) = False Then
            If MsgBox("Technician is already assign to this job, do you Want to change", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            Call cmdViewRODetails_Click
        Else
            MsgBox "Job Already Finish", vbInformation, "Info"
        End If
    End If
End Sub

Private Sub MnuAsscontractor_Click()
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Please Select a Repair Order Job to be Assigned with Contractor", vbInformation, "Info."
        Exit Sub
    End If
    
    Dim xCUSTOMERNAME                                   As String
    Dim xACCTNO                                         As String
    Dim XRONO                                           As String
    Dim XPLATENO                                        As String
    Dim xMODEL                                          As String
    Dim Index                                           As Integer
    
    xCUSTOMERNAME = Null2String(rptRO.SelectedRows(0).Record(1).Value)
    xACCTNO = Null2String(rptRO.SelectedRows(0).Record(13).Value)
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    XPLATENO = Null2String(rptRO.SelectedRows(0).Record(3).Value)
    xMODEL = Null2String(rptRO.SelectedRows(0).Record(2).Value)
    
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    Index = lstJob4Service.SelectedItem.Index
    If CheckIfJobIsFinish(lstJob4Service.ListItems(Index).SubItems(8)) = False Then
        With frmCSMSSelectContractor
            .labCust.Caption = xACCTNO
            .LABITEMNO.Caption = lstJob4Service.SelectedItem

            .lblro.Caption = XRONO
            .lblCustomer = xCUSTOMERNAME
            .lblplate = XPLATENO
            .LblModel = xMODEL

            .lblCONCODE.Caption = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index).ListSubItems(7)
            .lblLINENO.Caption = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index).ListSubItems(8)
        End With
        frmCSMSSelectContractor.Show 1
        frmCSMSSelectContractor.ZOrder 0
    Else
        MsgBox "Job Already Finish", vbInformation, "Info."
    End If
End Sub

Private Sub mnuBACKJOB_Click()
    If QC_MODULE_ON = "ON" Then
        If MsgBox("Tag This Repair Order as Back Job", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            gconDMIS.Execute "UPDATE CSMS_REPOR SET BACK_JOB = 'Y',BACKJOB_COUNT = BACKJOB_COUNT + 1 WHERE REP_OR = '" & theRo & "'"
        End If
    Else
        MsgBox "Quality Control Inpection Module is not yet ON", vbInformation, "CSMS"
        Exit Sub
    End If
End Sub

Private Sub mnuBilledRO_Click()
    Dim XRONO                                       As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add General job", vbInformation, "Info"
        Exit Sub
    End If

    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    If Module_Access(LOGID, "BILLING SYSTEM", "TRANSACTION") = False Then Exit Sub
    frmCSMSDataEntry.Show
    Call frmCSMSDataEntry.SearchRoToBilled(XRONO)
End Sub

Private Sub mnuCanedJob_Click()
    Dim XRONO                                       As String
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add Canned job", vbInformation, "Info."
        Exit Sub
    End If

    Screen.MousePointer = 11
    
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    frmCSMSGetCannedLabor.txtCheckMe = "MAIN"
    frmCSMSGetCannedLabor.lblro.Caption = XRONO
    frmCSMSGetCannedLabor.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub mnuChangeTech_Click()
    Dim Index                                          As Integer

    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Please Select a Repair Order to be Assigned with technician.", vbInformation, "Info."
        Exit Sub
    End If

    Index = lstJob4Service.SelectedItem.Index
End Sub

Private Sub mnuChangeVehicle_Click()
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "You have not select Repair Order", vbInformation, "Info."
        Exit Sub
    End If

    If CheckIfRoIsAlreadyInvoice(theRo) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If

    
    lblROChange.Caption = theRo
    frmCSMS_ChangeVehicle.lblRONO = theRo
    frmCSMS_ChangeVehicle.Show 1
    lblROChange.Caption = ""
    Call ViewActiveRO
End Sub

Private Sub mnuCreateApp_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_ADD", "APPOINTMENT") = False Then Exit Sub
    
    Dim XCustomer                                                   As String
    FROM_APPOINTMENT = "MAIN"
    
    txtApptno.Caption = MakeApptNo
    If txtApptno = "" Then
        MsgBox "Please select appointment schedule time...", vbInformation, "Info"
        Exit Sub
    End If
    
    XCustomer = CheckIfAppointmentTimeIsAvailable(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    'XCustomer = Null2String(rptAPP.SelectedRows(0).Record(2).Value)
    If XCustomer <> "" Then
        MsgBox "This time already schedule to: " & XCustomer & "", vbInformation, "Info."
        Exit Sub
    End If
    
    
    frmCSMSNewAppointment.txtTranNo = txtApptno
    frmCSMSNewAppointment.labType(0) = "Appointment"
    frmCSMSNewAppointment.labType(1) = "Appointment"
    frmCSMSNewAppointment.lblTime.Caption = Null2String(rptAPP.SelectedRows(0).Record(1).Value)
    frmCSMSNewAppointment.lblDate.Caption = MonthView1.Value
    frmCSMSNewAppointment.GetDefaultTransactionType
    frmCSMSNewAppointment.Show 1
    Call cmdRefresh_Click
End Sub

Private Sub mnuDeleteAppointment_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_delete", "APPOINTMENT") = False Then Exit Sub

    Dim xStatus                                         As String
    Dim xAPPNO                                          As String
    Dim XCustomer                                       As String
    Dim xCuscde                                         As String
    Dim xMODEL                                          As String
    Dim XPLATENO                                        As String
    
    If txtApptno = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    
    'XCustomer = Null2String(rptAPP.SelectedRows(0).Record(2).Value)
    XCustomer = CheckIfAppointmentTimeIsAvailable(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If XCustomer = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    
    'xStatus = Null2String(rptAPP.SelectedRows(0).Record(5).Value)
    xStatus = CheckAppointmentStatus(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If xStatus = "Served" Then
        MessagePop InfoFriend, "Appointment Information", "Appointment Already served", 1000
        Exit Sub
    End If
    
    xAPPNO = Null2String(rptAPP.SelectedRows(0).Record(0).Value)
    xCuscde = Null2String(rptAPP.SelectedRows(0).Record(7).Value)
    xMODEL = Null2String(rptAPP.SelectedRows(0).Record(3).Value)
    XPLATENO = Null2String(rptAPP.SelectedRows(0).Record(4).Value)
    
    If MsgBox("delete this appointment, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    'VERY WRONG LOGIC MADE! - FML 03032008
    'UNCOMMENT BY : MJP 09042009 0544PM
    'DESCRIPTION  : CHANGE IN LOGIC, DONT SAVE IN CSMS_APPOINTMENT IF ITS NOT A FINAL APPOINTMENT TO SAVE TABLE SPACE
    SQL_STATEMENT = "DELETE FROM CSMS_APPOINTMENT WHERE APPTNO = '" & txtApptno & "'"
    gconDMIS.Execute (SQL_STATEMENT)
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_Appointment"), "", "APPT NO: " & txtApptno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    
    SQL_STATEMENT = "Delete from CSMS_RepairOrder Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno & " - SERVICE COUNTER", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Delete from CSMS_Repor Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Delete from CSMS_RO_Det Where TransType = 'A' AND ApptNo = '" & txtApptno & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtApptno), "APPTNO", "CSMS_REPOR"), "", "APPT NO: " & txtApptno & " - JOBS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call ShowDeletedMsg
    Call cmdRefresh_Click
End Sub

Private Sub mnuDeleteEst_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_DELETE", "JOB ESTIMATE") = False Then Exit Sub
    
    If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then
        MsgBox "You cannot delete this estimate. Estimate already uploaded to Repair order.", vbInformation, "Info."
    Else
        If CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "UPLOADED" Then
            MsgBox "You cannot delete this estimate. Estimate already uploaded to Repair order. " & vbCrLf & " kindly refresh your service counter to display fresh data.", vbInformation, "Info."
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT FOUND" Then
            MsgBox "Estimate Record not found. kindly refresh your service counter to display fresh data", vbCritical, "Info"
            Exit Sub
        ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "NOT UPLOADED" Then
            If MsgBox("Delete this estimate record", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
            
            gconDMIS.Execute ("DELETE FROM CSMS_ESTHD WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_ESTDETAILS WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_REPOR WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_REPAIRORDER WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_RO_DET WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            gconDMIS.Execute ("DELETE FROM CSMS_PMS_JOB_DET WHERE ESTIMATENO = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'")
            
            Call ShowDeletedMsg
            Call cmdRefresh_Click
        End If
    End If
End Sub

Private Sub mnuEditAppoitnment_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_EDIT", "APPOINTMENT") = False Then Exit Sub
    
    Dim xStatus                                         As String
    Dim XCustomer                                       As String
    
    If txtApptno = "" Then
        MsgBox "Please select appointment schedule time...", vbInformation, "Info."
        Exit Sub
    End If
    
    'XCustomer = Null2String(rptAPP.SelectedRows(0).Record(2).Value)
    XCustomer = CheckIfAppointmentTimeIsAvailable(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If XCustomer = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    
    'xStatus = Null2String(rptAPP.SelectedRows(0).Record(5).Value)
    xStatus = CheckAppointmentStatus(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If xStatus = "Served" Then
        MessagePop InfoFriend, "Appointment Information", "Cannot edit Appointment, it has been already served", 1000
        Exit Sub
    End If

    frmCSMSEditAppointment.StoreApptTme
    frmCSMSEditAppointment.lblOLDAPPTNO.Caption = txtApptno.Caption
    frmCSMSEditAppointment.StoreAppInfo txtApptno.Caption
    frmCSMSEditAppointment.Show 1
    Call cmdRefresh_Click
End Sub

Private Sub mnujobdone_Click()
    Dim xCUSTOMERNAME                                   As String
    Dim xACCTNO                                         As String
    Dim XRONO                                           As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to View", vbInformation, "Info"
        Exit Sub
    End If
    
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
     
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    'UPDATE BY: JUN
    'DATE UPDATED: 03-18-2008
    'DESCRIPTION: FOR EDITING OF THE JOB COST IF THE SERVICE ADVISER WAS FORGOT TO INPUT THE JOB COST UPON CREATING OF REPAIR ORDER
        Dim rsJOBCOST                                           As New ADODB.Recordset
        Dim xJOB_COST  As Double
        Dim xJType     As String
            Set rsJOBCOST = gconDMIS.Execute("Select DETCOST,JOBTYPE from CSMS_RO_DET where rep_or = '" & theRo & "' and DETCDE = '" & THEJOBCODE & "' and livil = '1'")
            If Not rsJOBCOST.EOF And Not rsJOBCOST.BOF Then
                xJOB_COST = NumericVal(rsJOBCOST!DetCost)
                xJType = Null2String(rsJOBCOST!JOBTYPE)
            End If
        Set rsJOBCOST = Nothing
    'DATE UPDATED: 03-18-2008
    'UPDATE BY: JUN
    
    
    Dim Index                                      As Integer
    
    If COMPANY_CODE = "HPI" Then
        With frmCSMS_Jobdone_HPI
            Index = lstJob4Service.SelectedItem.Index
    
            .lblJobCode = THEJOBCODE
            .lbljobdesc = THEJOBDEST
            .lblro = theRo
    
            'UPDATED BY: JUN 03-18-2009
                If LTrim(RTrim(xJType)) <> "BP" Then
                    .txtJOBCOST.Visible = False
                    .lblSTIME.Visible = False
                    .Label10.Visible = False
                    .lblJobCost.Visible = False
                Else
                    .txtJOBCOST.Visible = True
                    .lblSTIME.Visible = True
                    .Label10.Visible = True
                    .lblJobCost.Visible = True
                End If
            'UPDATED BY: JUN---------------
    
    
            If StrComp(Trim(thestatus), "Billed") = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then
                .txtflatrate.Enabled = False
                'UPDATED BY: JUN 03-18-2009
                    .txtJOBCOST.Enabled = False
                    .txtstdrate.Enabled = False
                'UPDATED BY: JUN-----------
            Else
                .txtflatrate.Enabled = True
                .txtstdrate.Enabled = True
            End If
    
            .LABITEMNO.Caption = RTrim(LTrim(lstJob4Service.ListItems(Index).ListSubItems(8)))
            .lblflatrate = theflatrate
            .lblstdrate = THESTDRATE
            .txtflatrate = theflatrate
            .txtstdrate = THESTDRATE
            .txtJobDesc = THEJOBDEST
            'UPDATED BY: JUN 03-18-2009
                .lblJobCost = xJOB_COST
                .txtJOBCOST = xJOB_COST
            'UPDATED BY: JUN-----------
        End With
    
        frmCSMS_Jobdone_HPI.Show 1, frmCSMS_ServiceCounter
    Else
        With frmCSMS_Jobdone
            Index = lstJob4Service.SelectedItem.Index
    
            .lblJobCode = THEJOBCODE
            .lbljobdesc = THEJOBDEST
            .lblro = XRONO
    
            If StrComp(Trim(thestatus), "Billed") = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then
                .txtflatrate.Enabled = False
            Else
                .txtflatrate.Enabled = True
            End If
    
            .LABITEMNO.Caption = RTrim(LTrim(lstJob4Service.ListItems(Index).ListSubItems(8)))
            .lblflatrate = theflatrate
            .lblstdrate = THESTDRATE
            .txtflatrate = theflatrate
            .txtstdrate = THESTDRATE
            .txtJobDesc = THEJOBDEST
        End With
    
        frmCSMS_Jobdone.Show 1, frmCSMS_ServiceCounter
    End If
    
    Call DisplayRODetails(XRONO)
End Sub

Private Sub mnuOption1_Click()
    Dim xCUSTOMERNAME                               As String
    Dim xACCTNO                                     As String
    Dim XRONO                                       As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add General job", vbInformation, "Info"
        Exit Sub
    End If
    
    
    xCUSTOMERNAME = Null2String(rptRO.SelectedRows(0).Record(1).Value)
    xACCTNO = Null2String(rptRO.SelectedRows(0).Record(13).Value)
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    
    With frmCSMSReqJobs
        .txtCustomer.Text = xCUSTOMERNAME
        .txtActNo.Text = xACCTNO
        .txtROno.Text = XRONO
        .txtCheckMe = "main"
    End With
    
    If StrComp(Trim(thestatus), "Finish Job") = 0 Then
        If MsgBox("This Job Is Already Finish!,Do You Want To Add New Job?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            frmCSMSReqJobs.Show 1
            Exit Sub
        Else
            Exit Sub
        End If
    ElseIf StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    frmCSMSReqJobs.Show 1
End Sub

Private Sub mnuOption13_Click()
    Dim xPROMISE                                                As String
    Dim XRONO                                                   As String

    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add PMS job", vbInformation, "Info."
        Exit Sub
    End If

    xPROMISE = Null2String(rptRO.SelectedRows(0).Record(9).Value)
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)

    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    
    With frmCSMSPMS
        .txtRO.Text = XRONO
        .dtpromise.Value = xPROMISE
    End With

    If StrComp(Trim(thestatus), "Finish Job") = 0 Then
        If MsgBox("This Repair Order Is Already Finish!,Do You Want To Add New Job?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            If frmCSMSPMS.txtRO.Text = "" Then
                Exit Sub
            Else
                frmCSMSPMS.Show 1
            End If
        End If
    ElseIf StrComp(Trim(thestatus), "Billed") = 0 Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice", 1000
        Exit Sub
    ElseIf StrComp(Trim(thestatus), "Released") = 0 Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Released", 1000
        Exit Sub
    ElseIf StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    Else
        frmCSMSPMS.Show 1
    End If
End Sub

Private Sub mnuOption2_2_Click()
    Dim rsJOBSTATUS                                    As New ADODB.Recordset

    Set rsJOBSTATUS = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & theRo & "' AND LIVIL = '1' AND LINE_NO = '" & vlineNo & "'")
    If Not (rsJOBSTATUS.BOF And rsJOBSTATUS.EOF) Then
        If Null2String(rsJOBSTATUS!Status) = "Y" Then
            PERJOBSTATUS = "Finish Job"
        ElseIf Null2String(rsJOBSTATUS!Status) = "I" Then
            PERJOBSTATUS = "Idle Time"
        ElseIf Null2String(rsJOBSTATUS!Status) = "B" Then
            PERJOBSTATUS = "Break Time"
        ElseIf Null2String(rsJOBSTATUS!Status) = "G" Then
            PERJOBSTATUS = "Going Home"
        ElseIf Null2String(rsJOBSTATUS!Status) = "L" Then
            PERJOBSTATUS = "Lunch Break"
        ElseIf Null2String(rsJOBSTATUS!Status) = "W" Then
            PERJOBSTATUS = "Working"
        Else
            PERJOBSTATUS = ""
        End If
    End If

    If PERJOBSTATUS = "Finish Job" Then
        MsgBox "Cannot Remove Job. Job Already Finish", vbExclamation, "Info."
        Exit Sub
    ElseIf PERJOBSTATUS = "Idle Time" Or PERJOBSTATUS = "Going Home" Or PERJOBSTATUS = "Break Time" Or PERJOBSTATUS = "Lunch Break" Or PERJOBSTATUS = "Working" Then
        MsgBox "Cannot Remove Job. Job Already Started", vbExclamation, "Info."
        Exit Sub
    Else

    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
JUMP1:
    If MsgBox("Delete Job : " & lstJob4Service.SelectedItem.SubItems(1), vbYesNo + vbQuestion + vbDefaultButton1, "Are You Sure") = vbNo Then
        Exit Sub
    End If

    Dim rsEmpNo                                        As New ADODB.Recordset
    Dim vEMPNO                                         As String
    Dim vRONO                                          As String
    Dim VTECHCODE                                      As String

    VTECHCODE = LTrim(RTrim(lstJob4Service.SelectedItem.ListSubItems(7).Text))
    vRONO = LTrim(RTrim(lstJob4Service.SelectedItem.ListSubItems(6).Text))
    Set rsEmpNo = gconDMIS.Execute("SELECT EMPNO FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & LTrim(RTrim(lstJob4Service.SelectedItem.ListSubItems(7).Text)) & "'")
    If Not rsEmpNo.EOF Or Not rsEmpNo.BOF Then
        vEMPNO = LTrim(RTrim(Null2String(rsEmpNo!EMPNO)))
    End If

    Call ClearOrStayTechnician(vEMPNO, vRONO, VTECHCODE)

    Dim Index                                          As Integer
    Index = lstJob4Service.SelectedItem.Index

    AUDIT_SQL = "delete from CSMS_Ro_Det where REP_OR = '" & lstJob4Service.SelectedItem.SubItems(6) & "' and line_no = '" & lstJob4Service.ListItems(Index).ListSubItems(8) & "' and livil = '1'"
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        Dim VTRANID                                        As String
        VTRANID = FindTransactionID(N2Str2Null(lstJob4Service.SelectedItem.SubItems(6)), "REP_OR", "CSMS_REPOR")
        Call NEW_LogAudit("XX", "BILLING SYSTEM", AUDIT_SQL, VTRANID, "J", "JOB CODE: " & lstJob4Service.ListItems(Index).Text, "", "")
    'NEW LOG AUDIT ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    SQL_STATEMENT = "delete from CSMS_JobClock where ro_nO = '" & lstJob4Service.SelectedItem.SubItems(6) & "' and detcde = '" & lstJob4Service.SelectedItem & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT ------------------------------------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", AUDIT_SQL, VTRANID, "J", "JOB CODE: " & lstJob4Service.ListItems(Index).Text & " - CLOCK IN RECORD", "", "")
    'NEW LOG AUDIT ------------------------------------------------------------------------------


    SQL_STATEMENT = "delete from CSMS_PMS_Job_det where REP_OR = '" & lstJob4Service.SelectedItem.SubItems(6) & "' AND PMS_MODEL = '" & lstJob4Service.SelectedItem & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT ------------------------------------------------------------------------------
        Call NEW_LogAudit("XX", "BILLING SYSTEM", AUDIT_SQL, VTRANID, "J", "JOB CODE: " & lstJob4Service.ListItems(Index).Text & " - PMS DETAILS", "", "")
    'NEW LOG AUDIT ------------------------------------------------------------------------------

    Me.lstJob4Service.ListItems.Remove Me.lstJob4Service.SelectedItem.Index
    MessagePop InfoFriend, "RO Information Updated", "Job Succesfully Remove", 1000

    If CheckAllJobsISDone(vRONO) = True Then
        gconDMIS.Execute "update CSMS_RepairOrder set dateFinish = '" & LOGDATE & "', STATUS = 'Finish Job', JStatus = 'F' where RO_No = '" & vRONO & "'"
    End If
    
    'cmdRefresh.Value = True
End Sub

'Edit RO
Private Sub mnuOption2_Click()
    Dim XRONO                                       As String
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Please Select a Repair Order to be edit.", vbInformation, "Info"
        Exit Sub
    End If

    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    
    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Cannot edit Repair Order. ", vbInformation + vbOKOnly, "Repair Order Already Voided"
        Exit Sub
    End If
    Call frm.PassRepairOrderNo(XRONO, GetReporID(XRONO))
    frm.Show 1
End Sub

Function GetReporID(XRONO As String)
    Dim rstmp As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(XRONO) & "")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetReporID = rstmp!ID
    Else
        GetReporID = 0
    End If
    Set rstmp = Nothing
End Function

Private Sub mnuOtherJobs_Click()
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add Other job", vbInformation, "Info."
        Exit Sub
    End If

    Dim xCUSTOMERNAME                                           As String
    Dim xACCTNO                                                 As String
    Dim XRONO                                                   As String

    xCUSTOMERNAME = Null2String(rptRO.SelectedRows(0).Record(1).Value)
    xACCTNO = Null2String(rptRO.SelectedRows(0).Record(13).Value)
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)

    If CheckIfRoIsAlreadyInvoice(XRONO) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    
    With frmCSMSOtherJobs
        .txtCustomer.Text = xCUSTOMERNAME
        .txtActNo.Text = xACCTNO
        .txtROno.Text = XRONO
        .txtCheckMe.Text = "main"
    End With
    
    If StrComp(Trim(thestatus), "Finish Job") = 0 Then
        If MsgBox("This Repair Order Is Already Finish!,Do You Want To Add New Job?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            frmCSMSOtherJobs.Show 1
            Exit Sub
        Else
            Exit Sub
        End If
    ElseIf StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    frmCSMSOtherJobs.Show 1
    Call DisplayRODetails(XRONO)
End Sub

Private Sub mnuPEstimate_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_PRINT", "JOB ESTIMATE") = False Then Exit Sub
    
    CrystalReport1.WindowTitle = "Estimate Print Out"
    CrystalReport1.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    CrystalReport1.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport CrystalReport1, CSMS_REPORT_PATH & "PrintEstimate.rpt", "{repor.ESTIMATENO} = '" & Null2String(rptEST.SelectedRows(0).Record(5).Value) & "'", CSMS_REPORT_CONNECTION, 1
End Sub

Private Sub mnuPrintAppointment_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "ACESS_PRINT", "APPOINTMENT") = False Then Exit Sub
    
    Dim XCustomer                                   As String
    
    If txtApptno = "" Then
        MsgBox "Please select appointment schedule time...", vbInformation, "Info."
        Exit Sub
    End If

    XCustomer = Null2String(rptAPP.SelectedRows(0).Record(2).Value)
    If XCustomer = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    
    If MsgBox("Print this Appointment, Are You Sure", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                PrintSQLReport rptNard, CSMS_REPORT_PATH & "appointment.rpt", "{CSMS_appointment.apptno} = '" & txtApptno.Caption & "'", CSMS_REPORT_CONNECTION, 1
    End If
End Sub

Function GetTaym(XXX As String)
    Dim rstmp                                          As New ADODB.Recordset
    Dim X                                              As Integer
    Dim cnt                                            As Integer
    cnt = 0
    Set rstmp = gconDMIS.Execute("Select PromiseDate From CSMS_RepairOrder Where RO_no = '" & XXX & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        For X = 1 To Len(rstmp!PromiseDate)
            If Mid(rstmp!PromiseDate, X, 1) = "/" Then cnt = cnt + 1
            If cnt = 2 Then
                GetTaym = Mid(rstmp!PromiseDate, X + 6, Len(rstmp!PromiseDate) - X)
                Exit For
            End If
        Next
    End If

    Set rstmp = Nothing
End Function

Function CheckIfThereAPMS(vRO As String) As Boolean
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT JOBTYPE FROM CSMS_RO_DET WHERE REP_OR = '" & vRO & "' AND LIVIL = '1' AND JOBTYPE = 'PMS'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        CheckIfThereAPMS = True
    Else
        CheckIfThereAPMS = False
    End If

    Set rstmp = Nothing
End Function

Private Sub mnuPrintRO_Click()
    Dim XRONO                                       As String
    Dim XPLATENO                                    As String
    Dim rsEndUser                                   As ADODB.Recordset
    Dim rsrocuscode                                 As ADODB.Recordset
    Dim xend                                        As String
    
    If theRo = "" Or theRo = "R/O" Then
        MsgBox "Choose a Repair Order to add General job", vbInformation, "Info"
        Exit Sub
    End If
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Cannot Print Repair Order." & vbCrLf & "Printing has been disabled.", vbInformation + vbOKOnly
        Exit Sub
    End If

    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    XPLATENO = Null2String(rptRO.SelectedRows(0).Record(3).Value)
    
    Screen.MousePointer = 11
    rptRepairOrder.WindowShowPrintSetupBtn = True
    rptRepairOrder.WindowTitle = "Repair Order Print Out"
    rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
    Set rsrocuscode = New ADODB.Recordset
    Set rsrocuscode = gconDMIS.Execute("Select enduser as xendser from csms_cusveh where plate_no = '" & XPLATENO & "'")
    If Not (rsrocuscode.EOF And rsrocuscode.BOF) Then
        xend = Null2String(rsrocuscode!xendser)
        Set rsrocuscode = Nothing
    End If
    
    Set rsEndUser = New ADODB.Recordset
    Set rsEndUser = gconDMIS.Execute("Select acctname from ALL_Customer_Table where cuscde ='" & xend & "' ")
    
    If Not (rsEndUser.EOF And rsEndUser.BOF) Then
         rptRepairOrder.Formulas(3) = "niymandenduser = '" & Null2String(rsEndUser!AcctName) & "'"
          
    Else
         rptRepairOrder.Formulas(3) = "niymandenduser = '" & "" & "'"
    End If
    Set rsEndUser = Nothing

    
    If VALID_COMPANY_CODE_FORHAI = True Or COMPANY_CODE = "HSR" Then
        rptRepairOrder.Formulas(3) = "TAYM = '" & GetTaym(XRONO) & "'"
    End If

    If COMPANY_CODE = "HAS" Then
        If CheckIfThereAPMS(XRONO) = True Then
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder_NOPMS.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
        End If
    Else
    
        If COMPANY_CODE = "HSR" Or COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "HEI" Or COMPANY_CODE = "HPI" Or COMPANY_CODE = "HCI" Or COMPANY_CODE = "HLU" Then
            Dim rstmp                                  As New ADODB.Recordset
            Dim FJOB                                   As String
            Dim SJOB                                   As String
            Dim TJOB                                   As String
            Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & XPLATENO & "' AND TRANSTYPE = 'R' AND DTE_RECD < '" & Date & "' ORDER BY DTE_RECD ASC ")
            If Not (rstmp.BOF And rstmp.EOF) Then
                If Not rstmp.BOF Then
                    rstmp.MoveFirst
                    
                    If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HSR" Then
                        FJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                        FJOB = FJOB & "    " & GetJobList(rstmp!REP_OR)
                    Else
                        FJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                    End If
                    rstmp.MoveNext

                    If Not rstmp.EOF Then
                        If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HSR" Then
                            SJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                            SJOB = SJOB & "    " & GetJobList(rstmp!REP_OR)
                        Else
                            SJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                        End If
                        rstmp.MoveNext

                        If Not rstmp.EOF Then
                            If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HSR" Then
                                TJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                            Else
                                TJOB = Null2String(rstmp!REP_OR) & "     " & Null2String(rstmp!DTE_RECD) & "    " & Null2String(rstmp!km_rdg)
                                TJOB = TJOB & "    " & GetJobList(rstmp!REP_OR)
                            End If
                        End If
                    End If
                End If
            End If
            Set rstmp = Nothing
            
            rptRepairOrder.Formulas(0) = "RO1 = '" & FJOB & "'"
            rptRepairOrder.Formulas(1) = "RO2 = '" & SJOB & "'"
            rptRepairOrder.Formulas(2) = "RO3 = '" & TJOB & "'"
            
            If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HLU" Then
                rptRepairOrder.Formulas(2) = "TAYM = '" & GetTaym2(XRONO) & "'"
            End If
            If COMPANY_CODE = "HCI" Or COMPANY_CODE = "HLU" Then
                rptRepairOrder.Formulas(3) = "ATTENDED = '" & GetAttendedTaym(XRONO) & "'"
            End If
            
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
        Else
            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
        End If
'        If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HEI" Then
'            Dim RSTMP                                  As New ADODB.Recordset
'            Dim FJOB                                   As String
'            Dim SJOB                                   As String
'            Dim TJOB                                   As String
'            Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_REPOR WHERE PLATE_NO = '" & XPLATENO & "' AND TRANSTYPE = 'R' AND DTE_RECD < '" & Date & "' ORDER BY DTE_RECD ASC ")
'            If Not (RSTMP.BOF And RSTMP.EOF) Then
'                If Not RSTMP.BOF Then
'                    RSTMP.MoveFirst
'                    FJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
'                    RSTMP.MoveNext
'
'                    If Not RSTMP.EOF Then
'                        SJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
'                        RSTMP.MoveNext
'
'                        If Not RSTMP.EOF Then
'                            TJOB = Null2String(RSTMP!REP_OR) & "     " & Null2String(RSTMP!DTE_RECD) & "    " & Null2String(RSTMP!km_rdg)
'                        End If
'                    End If
'                End If
'            End If
'            Set RSTMP = Nothing
'            rptRepairOrder.Formulas(0) = "RO1 = '" & FJOB & "'"
'            rptRepairOrder.Formulas(1) = "RO2 = '" & SJOB & "'"
'            rptRepairOrder.Formulas(2) = "RO3 = '" & TJOB & "'"
'            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
'        Else
'            PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printRepairOrder.rpt", "{repor.rep_or} = '" & XRONO & "'", CSMS_REPORT_CONNECTION, 1
'        End If
    End If

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "BILLING SYSTEM", "", FindTransactionID(N2Str2Null(XRONO), "REP_OR", "CSMS_REPOR"), "", "RO NO: " & XRONO & " - VIEW RO DETAILS", "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    'LogAudit "V", "PRINT RO"
    Screen.MousePointer = 0
End Sub

Function GetTaym2(ro As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Dim X                                              As Integer
    Dim cnt                                            As Integer
    cnt = 0
    
    Set rstmp = gconDMIS.Execute("Select PromiseDate From CSMS_RepairOrder Where RO_no = '" & ro & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        For X = 1 To Len(rstmp!PromiseDate)
            If Mid(rstmp!PromiseDate, X, 1) = "/" Then cnt = cnt + 1
            If cnt = 2 Then
                GetTaym2 = Mid(rstmp!PromiseDate, X + 6, Len(rstmp!PromiseDate) - X)
                Exit For
            End If
        Next
    End If
    Set rstmp = Nothing
End Function
Function GetAttendedTaym(ro As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("Select SAVETIME From CSMS_RepOR Where REP_OR = '" & ro & "' AND TRANSTYPE = 'R'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GetAttendedTaym = Null2String(rstmp!savetime)
        If Hour(GetAttendedTaym) < 12 Then
            GetAttendedTaym = Mid(GetAttendedTaym, 1, 5) & " AM"
        Else
            GetAttendedTaym = Mid(GetAttendedTaym, 1, 5) & " PM"
        End If
    End If

    Set rstmp = Nothing
End Function

Function GetJobList(xxxRO_NO As String) As String
    Dim rstmp As New ADODB.Recordset
    Dim XXX As String
    Set rstmp = gconDMIS.Execute("SELECT TOP 5 DETCDE FROM CSMS_RO_DET WHERE REP_OR = '" & xxxRO_NO & "' AND LIVIL = '1'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            XXX = XXX & "     " & Null2String(rstmp!DETCDE)
            rstmp.MoveNext
        Loop
    End If
    GetJobList = XXX
    Set rstmp = Nothing
End Function

Private Sub mnuprintur_Click()
'
'    If rptRO.Rows.Count = 0 Then
'            MsgBox "No Record(s) to Print", vbInformation
'            Exit Sub
'        End If
'        On Error GoTo ErrorCode:
'        Dim xlSheet1
'        Dim RSHD                                           As ADODB.Recordset
'        Dim objXL                                          As New Excel.Application
'        Dim wbXL                                           As New Excel.Workbook
'        Dim wsXL                                           As New Excel.Worksheet
'        Dim intRow                                         As Long    ' counter
'        Dim intCol                                         As Long    ' counter
'        If Not IsObject(objXL) Then
'            MsgBox "You need Microsoft Excel to use this function", _
'                   vbExclamation, "Print to Excel"
'            Exit Sub
'        End If
'        On Error Resume Next
'        Set wbXL = objXL.Workbooks.Add
'        Set wsXL = objXL.ActiveSheet
'        wsXL.Name = "PARTS QUERY"
'        Prg.Max = rptRO.Rows.Count
'        Prg.Value = 0
'        wsXL.Cells(2, 1) = "DEALER NAME : " & COMPANY_NAME
'        wsXL.Cells(3, 1) = "Unit Received : " & MonthView1.Value & " "
'
'
'        Screen.MousePointer = 11
'        For intCol = 0 To rptRO.Columns.Count - 1
'            wsXL.Cells(6, intCol + 1).Value = "" & CStr(rptRO.Columns(intCol).Caption) & "  "
'        Next
'        Prg.ZOrder 0
'        For intRow = 0 To rptRO.Rows.Count - 1
'            For intCol = 0 To rptRO.Columns.Count - 1
'                wsXL.Cells(intRow + 7, intCol + 1).Value = "" & CStr(rptRO.Rows(intRow).Record(intCol).Value) & "  "
'            Next
'            Prg.Value = Prg.Value + 1
'            Prg.Text = Round((Prg.Value / Prg.Max) * 100, 0) & "%"
'        Next
'        wsXL.Cells("A2").CopyFromRecordset RSHD
'        For intCol = 1 To rptRO.Columns.Count
'            wsXL.Columns(intCol).AutoFit
'        Next
'        wsXL.Range("A6", Right(wsXL.Columns(rptRO.Columns.Count).AddressLocal, 1) & rptRO.Rows.Count + 6).AutoFormat 2
'        'wsXL.Cells("A1").CopyFromRecordset RSHD
'        objXL.Visible = True
'        Screen.MousePointer = 0
'          Prg.ZOrder 1
'        Exit Sub
'
'ErrorCode:
'        MsgBox Err.Description
'        Err.Clear
 Dim XRONO                                       As String
    Dim XPLATENO                                    As String
    Dim rsEndUser                                   As ADODB.Recordset
    Dim xend                                        As String
    Set rsEndUser = New ADODB.Recordset
    Set rsEndUser = gconDMIS.Execute("Select acctname from ALL_Customer_Table where cuscde ='" & xend & "' ")
    If Not (rsEndUser.EOF And rsEndUser.BOF) Then
             rptRepairOrder.Formulas(3) = "niymandenduser = '" & Null2String(rsEndUser!AcctName) & "'"
              
        Else
             rptRepairOrder.Formulas(3) = "niymandenduser = '" & "" & "'"
    End If
        
        Screen.MousePointer = 11
        rptRepairOrder.WindowShowPrintSetupBtn = True
        rptRepairOrder.WindowTitle = "Repair Order Print Out"
        rptRepairOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptRepairOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    If Check1.Value = 1 Then
       PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printAll_Unit_Received.rpt", "", CSMS_REPORT_CONNECTION, 1
    Else
    
       PrintSQLReport rptRepairOrder, CSMS_REPORT_PATH & "printAll_Unit_Received.rpt", "{repor.dte_recd}= Date(" & Year(MonthView1) & "," & Month(MonthView1) & "," & Day(MonthView1) & ")", CSMS_REPORT_CONNECTION, 1
       
    End If
End Sub

Private Sub mnuremovebay_Click()
    Dim RS                                             As New ADODB.Recordset

    If CheckIfRoIsAlreadyInvoice(theRo) = True Then
        MessagePop InfoFriend, "Repair order Information", "Repair order already Invoice, please refresh your Service Counter", 1000
        Exit Sub
    End If
    
    Set RS = gconDMIS.Execute("SELECT Ro from CSMS_baymonitoring where RO = '" & theRo & "'")
    If Not (RS.EOF And RS.BOF) Then
        If MsgBox("Are you sure do you want to remove this Repair Order to Bay", vbQuestion + vbYesNo) = vbYes Then
            gconDMIS.Execute "Update CSMS_Baymonitoring set " & _
                " ro = null " & _
                ", bay_status = 'Available' " & _
                " where ro='" & theRo & "'"
            
            Call ShowSuccessFullyUpdated
        End If
    Else
        MsgBox "Repair Order not in the bay.", vbInformation, "Infomartion"
    End If
End Sub

Private Sub mnuRemoveTech_Click()
'    Dim INDEX                                          As Integer
'
'    If theRo = "" Or theRo = "R/O" Then
'        MsgBox "Please Select a Repair Order to be Assigned with technician.", vbInformation, "Info."
'        Exit Sub
'    End If
'
'    INDEX = lstJob4Service.SelectedItem.INDEX
'
'    If CheckIfTheJobIsFinish(theRo, Null2String(lstJob4Service.ListItems(INDEX).Text)) = "Finish" Then
'        MsgBox "Job already finish", vbInformation, "Info"
'        Exit Sub
'    ElseIf CheckIfTheJobIsFinish(theRo, Null2String(lstJob4Service.ListItems(0).Text)) = "Not Finish" Then
'        MsgBox "Technician cannot be remove when Job is already Starting", vbInformation, "Info"
'        Exit Sub
'    Else
'        If MsgBox("Remove this technician from this job", vbQuestion + vbYesNo, "COnfirm") = vbYes Then
'            gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = NULL, JSTATUS = 'A' WHERE EMPNO = " & xEMPNO & "")
'
'            gconDMIS.Execute ("UPDATE CSMS_RO_DET WHERE TECHCODE = NULL " & _
'                " AND REP_OR = " & N2Str2Null(theRo) & "")
'
'            Call ShowSuccessFullyUpdated
'        End If
'    End If
End Sub

Private Sub mnuUEstimate_Click()
    If Module_Access(LOGID, "JOB ESTIMATE", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "acess_post", "JOB ESTIMATE") = False Then Exit Sub
    
    If Null2String(rptEST.SelectedRows(0).Record(6).Value) = "Uploaded" Then
        MsgBox "Estimate Already uploaded to Repair Order", vbInformation, "Info"
        Exit Sub
    ElseIf CheckEstimateStatus(Null2String(rptEST.SelectedRows(0).Record(5).Value)) = "UPLOADED" Then
        MsgBox "Estimate already uploaded to Repair order. " & vbCrLf & "kindly refresh your service counter to display fresh data.", vbInformation, "Info."
        Exit Sub
    End If
    
    Call frmApp.FillEstimateno(Null2String(rptEST.SelectedRows(0).Record(5).Value), "SERVICE COUNTER")
    frmApp.ZOrder 0
    frmApp.Show 1
    
End Sub

Private Sub mnuUplodAppointment_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub
    If Function_Access(LOGID, "ACESS_POST", "APPOINTMENT") = False Then Exit Sub
    Dim xStatus                                         As String
    Dim xAPPNO                                          As String
    Dim XCustomer                                       As String
    Dim xCuscde                                         As String
    Dim xMODEL                                          As String
    Dim XPLATENO                                        As String
    
    xStatus = CheckAppointmentStatus(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If xStatus = "Served" Then
        MessagePop InfoFriend, "Appointment Information", "Appointment Already served!", 1000
        Exit Sub
    End If
    
    xAPPNO = Null2String(rptAPP.SelectedRows(0).Record(0).Value)
    XCustomer = Null2String(rptAPP.SelectedRows(0).Record(2).Value)
    xCuscde = Null2String(rptAPP.SelectedRows(0).Record(7).Value)
    xMODEL = Null2String(rptAPP.SelectedRows(0).Record(3).Value)
    XPLATENO = Null2String(rptAPP.SelectedRows(0).Record(4).Value)
    
    XCustomer = CheckIfAppointmentTimeIsAvailable(MonthView1.Value, Null2String(rptAPP.SelectedRows(0).Record(1).Value))
    If XCustomer = "" Then
        Call ShowNoRecord
        Exit Sub
    End If
    
    With frmCSMSLoadApointmentToRO
        .txtAppt = xAPPNO
        .txtAcct_No = xCuscde
        .txtCustomer = XCustomer
        .txtModel = xMODEL
        .txtPlanteNo = XPLATENO
        .txtROno = GetNewROno(txtApptno)
    End With
    frmCSMSLoadApointmentToRO.Show 1
    Call cmdRefresh_Click
End Sub

Function GetNewROno(XXX As Variant)
    Dim rsNewRO                                        As New ADODB.Recordset
    If FOR_J = True Then
        Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' AND left(rep_or,1) = 'J' order by rep_or desc")
    Else
        Set rsNewRO = gconDMIS.Execute("select id,rep_or from CSMS_RepOr where TransType='R' order by rep_or desc")
    End If
    If Not rsNewRO.EOF And Not rsNewRO.BOF Then
        GetNewROno = Format(NumericVal(Mid$(rsNewRO!REP_OR, 3, 8)) + 1, "00000000")
    Else
        GetNewROno = "00000001"
    End If
    Set rsNewRO = Nothing
End Function

Private Sub mnuViewRODet_Click()
    If StrComp(Trim(thestatus), "Voided") = 0 Then
        MsgBox "Repair Order Already Voided!", vbInformation + vbOKOnly
        Exit Sub
    End If
    Call cmdViewRODetails_Click
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Thedate = Format(Now, "MM/dd/yyyy")
    Check1.Value = 0
    Check2.Value = 1
    cmdRefresh.Value = True
End Sub

Sub ProcesUpdate()
    Screen.MousePointer = 11
    '''''''''AXP PROCESS UPDATE REVISED
    gconDMIS.Execute ("UPDATE CSMS_JOBCLOCKOPENRO SET " & _
        " HRSWORKED = ROUND(DATEDIFF(MINUTE, CASE WHEN ISDATE(CLOCKIN) = 0 THEN GETDATE() ELSE CLOCKIN END, CASE WHEN ISDATE(CLOCKOUT) = 0 THEN GETDATE() ELSE CLOCKOUT END) / 60.00, 2)")
    
    gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET TODAY = GETDATE(), XHRSWORK = T.TOTAL_HRS " & vbCrLf _
        & " FROM  (SELECT RO_NO, SUM(ISNULL(HRS,0)) AS TOTAL_HRS  FROM CSMS_VW_JOBCLOCKTORO GROUP BY RO_NO) T" & vbCrLf _
        & " INNER JOIN  CSMS_REPAIRORDER  ON  CSMS_REPAIRORDER.RO_NO = T.RO_NO ")
    
    Screen.MousePointer = 0
    '''''''''PROCESS UPDATE REVISED

    Exit Sub

    Dim rsProces                                       As New ADODB.Recordset
    Set rsProces = gconDMIS.Execute("Select ID, ClockIn, ClockOut, HrsWorked from CSMS_JobClockOpenRO")
    If Not (rsProces.EOF And rsProces.BOF) Then
        Dim xTimeInAM                                   As String
        Dim xTimeOutAM                                  As String
        Dim xMam                                        As Double
        Dim hrAM                                        As Double
        Dim tlHrs                                       As Double

        Do Until rsProces.EOF
            If IsNull(rsProces![CLOCKIN]) = True Then
                xTimeInAM = Now
            Else
                xTimeInAM = rsProces![CLOCKIN]
            End If

            If IsNull(rsProces![CLOCKOUT]) = True Then
                xTimeOutAM = Now
            Else
                xTimeOutAM = rsProces![CLOCKOUT]
            End If

            xMam = DateDiff("N", xTimeInAM, xTimeOutAM)

            hrAM = Round(xMam / 60, 2)

            tlHrs = hrAM

            gconDMIS.Execute "UPDATE CSMS_JOBCLOCK SET HRSWORKED = " & tlHrs & " WHERE ID = " & rsProces![ID]

            rsProces.MoveNext
        Loop
    End If

    Dim XHRS                                           As Double
    Dim XRONO                                          As String
    Dim PREVNO                                         As String
    Dim NEWRO                                          As String
    Set rsProces = gconDMIS.Execute("SELECT SUM(ISNULL(HRS,0)) AS TOTAL_HRS,RO_NO FROM CSMS_VW_JOBCLOCKTORO GROUP BY RO_NO ORDER BY RO_NO")
    If Not rsProces.EOF And Not rsProces.BOF Then
        Do Until rsProces.EOF
            DoEvents
            gconDMIS.Execute "update CSMS_RepairOrder set [today]= '" & Now & "', xHrsWork = " & N2Str2Zero(rsProces![TOTAL_hrs]) & " where RO_No = " & N2Str2Null(rsProces![RO_NO])
            rsProces.MoveNext
        Loop
    End If
    Set rsProces = Nothing
    Screen.MousePointer = 0
End Sub

Function SetCusVehCondNo(XXX As String, yyy As String) As String
    Dim rsCusVeh                                       As New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select VCOND_NO from CSMS_CUSVEH where CUSCDE = '" & yyy & "' AND PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        SetCusVehCondNo = Null2String(rsCusVeh!VCOND_NO)
    End If
    Set rsCusVeh = Nothing
End Function

Function SetCusVehDesc(XXX As String, yyy As String) As String
    Dim rsCusVeh                                       As New ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select Description from CSMS_CUSVEH where CUSCDE = '" & yyy & "' AND PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        SetCusVehDesc = Null2String(rsCusVeh!Description)
    End If
    Set rsCusVeh = Nothing
End Function

Function SetCusVehDetail(XXX As String, yyy As String) As String()
    Dim rsCusVeh                                       As New ADODB.Recordset
    Dim Veh_detail(1)                                  As String
    Veh_detail(0) = ""
    Veh_detail(1) = ""

    Set rsCusVeh = gconDMIS.Execute("Select Description,VCOND_NO from CSMS_CUSVEH where CUSCDE = '" & yyy & "' AND PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        Veh_detail(0) = Null2String(rsCusVeh!Description)
        Veh_detail(1) = Null2String(rsCusVeh!VCOND_NO)
    End If
    Set rsCusVeh = Nothing
    SetCusVehDetail = Veh_detail
End Function

Private Sub rptAPP_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(5).Value = "Served" Then
        Metrics.ForeColor = Shape9.FillColor
    Else
        Metrics.ForeColor = vbBlack
    End If
End Sub

Private Sub rptAPP_MouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If Button = vbRightButton Then
        PopupMenu mnuAppointment
    End If
End Sub

Private Sub rptAPP_SelectionChanged()
    Call InitViewInfo
    If Not Null2String(rptAPP.SelectedRows(0).Record(0).Value) = "" Then
        Call StoreAppInfo(Null2String(rptAPP.SelectedRows(0).Record(0).Value), Null2String(rptAPP.SelectedRows(0).Record(4).Value), Null2String(rptAPP.SelectedRows(0).Record(7).Value))
    End If
End Sub

Sub InitViewInfo()
    txtnote = "":           txtnote = ""
    cboRecd_by = ""
    txtKm_rdg = "":         txtDte_recd = "":   txtVIN = "":
    txtModel = "":          txtMake = "":       txtDescription = "":    txtDte_recd = ""
    lblCN1 = "":            lblCN2 = ""

    'lstJob4Service.ListItems.Clear
End Sub

Sub StoreAppInfo(XXX As String, XPLATENO As String, xCuscde As String)
    Dim rsCusVeh                                       As New ADODB.Recordset
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim rsAppointment                                  As New ADODB.Recordset
    Dim rstmp                                          As New ADODB.Recordset
    
    txtApptno.Caption = XXX
    Set RSUPLOAD = gconDMIS.Execute("select * from csms_repairorder where apptno = '" & XXX & "'")
    If Not RSUPLOAD.EOF Or Not RSUPLOAD.BOF Then
        cboRecd_by = Null2String(RSUPLOAD!writer)
        txtnote = Null2String(RSUPLOAD!RECOMMENDATION)
        If IsDate(RSUPLOAD!PromiseDate) = True Then
            txtDte_recd = DateValue(RSUPLOAD!PromiseDate)
        End If
    End If

    Set rsAppointment = gconDMIS.Execute("Select * from CSMS_Appointment Where ApptNo = '" & XXX & "'")
    If Not rsAppointment.EOF And Not rsAppointment.BOF Then
        txtnote = txtnote & " " & Null2String(rsAppointment!NOTE)
        txtKm_rdg = Null2String(rsAppointment!km_rdg)
    End If

    Set rsCusVeh = gconDMIS.Execute("Select * from CSMS_CUSVEH where PLATE_NO = '" & XPLATENO & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        txtModel = Null2String(rsCusVeh!Model)
        txtMake = Null2String(rsCusVeh!Make)
        txtDescription = Null2String(rsCusVeh!Description)
        txtVIN = UCase(Null2String(rsCusVeh!Vin))
    End If

    Set rstmp = gconDMIS.Execute("Select HomePhone,TelephoneNo , Mobile From All_Customer Where CusCde = '" & xCuscde & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        lblCN2 = Null2String(rstmp!HomePhone)
        lblCN1 = Null2String(rstmp!TelephoneNo) & " " & Null2String(rstmp!Mobile)
    End If

    Set rstmp = Nothing
End Sub

Private Sub rptEST_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(6).Value = "Uploaded" Then
        Metrics.ForeColor = Shape8.FillColor
    Else
        Metrics.ForeColor = vbBlack
    End If
End Sub

Private Sub rptEST_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)

    public_ESTNO = Null2String(rptEST.SelectedRows(0).Record(5).Value)
    xESTNO = Null2String(rptEST.SelectedRows(0).Record(5).Value)
    
    Call FillEstimateDetails(xESTNO)
    
    tabDet.SelectedItem = 0
End Sub

Private Sub rptEST_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    mnuAddEstJob.Enabled = True
    mnuAddEstDet.Enabled = True
    mnuUEstimate.Enabled = True
    mnuDeleteEst.Enabled = True
    mnuPEstimate.Enabled = True
    mnudelete_estjob.Enabled = False
    PopupMenu mnuEstimate
End Sub

Private Sub rptEST_SelectionChanged()
    
    public_ESTNO = Null2String(rptEST.SelectedRows(0).Record(5).Value)
    xESTNO = Null2String(rptEST.SelectedRows(0).Record(5).Value)
    
    Call FillEstimateDetails(xESTNO)
    
    tabDet.SelectedItem = 0
End Sub

Sub FillEstimateDetails(xESTNO As String)
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim Item                                           As ListItem

    Call CleanListViewDetails

    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETAIL,FLATRATE,det_hrs,TECHNICIAN,HRSWRK,REP_OR,TechCode,LINE_NO,status from CSMS_ESTDETAILS where LIVIL='1' AND ESTIMATENO = '" & xESTNO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstJob4Service.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))

            Item.SubItems(1) = Replace(Null2String(RSUPLOAD!Detail), vbCrLf, " ")
            Item.SubItems(2) = Format(NumericVal(RSUPLOAD!FLATRATE), MAXIMUM_DIGIT)
            Item.SubItems(3) = Null2String(RSUPLOAD!DET_HRS)
            Item.SubItems(4) = FindTechName(LTrim(RTrim(Null2String(RSUPLOAD!TechCode))))
            Item.SubItems(5) = Null2String(RSUPLOAD!HRSWRK)
            Item.SubItems(6) = Null2String(RSUPLOAD!REP_OR)
            Item.SubItems(7) = Null2String(RSUPLOAD!TechCode)
            Item.SubItems(8) = Null2String(RSUPLOAD!LINE_NO)

            If Null2String(RSUPLOAD!Status) = "W" Then Item.SubItems(9) = "Working": Call CHECK_IN_OUT(RSUPLOAD!REP_OR, RSUPLOAD!DETCDE, RSUPLOAD!TechCode, RSUPLOAD!Status) 'UPDATED BY JUN: 03-23-2009
            If Null2String(RSUPLOAD!Status) = "I" Then Item.SubItems(9) = "Idle Time"
            If Null2String(RSUPLOAD!Status) = "L" Then Item.SubItems(9) = "Lunch Break"
            If Null2String(RSUPLOAD!Status) = "G" Then Item.SubItems(9) = "Going Home"
            If Null2String(RSUPLOAD!Status) = "B" Then Item.SubItems(9) = "Break Time"
            If Null2String(RSUPLOAD!Status) = "Y" Or Null2String(RSUPLOAD!Status) = "R" Then Item.SubItems(9) = "Finish Job"

            If Null2String(RSUPLOAD!Status) = "J" Then Item.SubItems(9) = "Back Job"
            If Null2String(RSUPLOAD!Status) = "Q" Then Item.SubItems(9) = "Waiting for QC"
            RSUPLOAD.MoveNext
        Loop
    End If

    'PMS JOBS
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,PMS_MODEL from CSMS_PMS_Job_det where ESTIMATENO = '" & xESTNO & "' and rep_or = '" & xESTNO & "'Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstPMSJobs.ListItems, RSUPLOAD
    End If
    Set RSUPLOAD = Nothing

    'PARTS
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_ESTDETAILS where LIVIL='2' AND ESTIMATENO = '" & xESTNO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstParts.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)
            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing

    'MATERIALS
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_ESTDETAILS where LIVIL='3' AND ESTIMATENO = '" & xESTNO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            DoEvents
            Set Item = lstMaterials.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)

            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing

    'ACCESSORIES
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_ESTDETAILS where LIVIL='4' AND ESTIMATENO = '" & xESTNO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            DoEvents
            Set Item = lstAccessories.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)

            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing
End Sub

Private Sub rptRO_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    If Row.Record(8).Value = "Released" Then
        Metrics.ForeColor = Shape3.FillColor
    ElseIf Row.Record(8).Value = "Billed" Then
        Metrics.ForeColor = Shape5.FillColor
    ElseIf Row.Record(8).Value = "Park" Then
        Metrics.ForeColor = Shape1.FillColor
    ElseIf Row.Record(8).Value = "Finish Job" Then
        Metrics.ForeColor = Shape6.FillColor
    ElseIf Row.Record(8).Value = "Working" Then
        If DateDiff("s", Null2String(Row.Record(9).Value), Now) < 0 Then
            Metrics.ForeColor = Shape2.FillColor
        Else
            Metrics.ForeColor = Shape4.FillColor
        End If
    ElseIf Row.Record(8).Value = "Idle Time" Then
        Metrics.ForeColor = Shape.FillColor
    ElseIf Row.Record(8).Value = "Back Job" Then
        Metrics.ForeColor = Shape7.FillColor
    ElseIf Row.Record(8).Value = "Voided" Then
        Metrics.ForeColor = shpvoid.FillColor
    Else
        Metrics.ForeColor = vbRed
    End If
End Sub

Private Sub rptRO_RowRClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim test                                        As String
    Dim XRONO                                       As String
    Dim xStatus                                     As String
    
    XRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    'xStatus = Null2String(rptRO.SelectedRows(0).Record(8).Value)
    xStatus = GetFreshServiceCounterStatus(XRONO)
    
    'test = "Billed"
    'If StrComp(Trim(thestatus), test) = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then
    
    If StrComp(Trim(xStatus), "Billed") = 0 Or StrComp(Trim(xStatus), "Released") = 0 Then
        If QC_MODULE_ON = "ON" Then
            Call EnableDropDownMenu(False)
            Call PopupMenu(mnuOption)
        Else
            Call EnableDropDownMenu(False)
            mnuBACKJOB.Enabled = False
            Call PopupMenu(mnuOption)
        End If
    Else
        Call EnableDropDownMenu(True)
        If xStatus = "Finish Job" Then
            mnuBilledRO.Visible = True
        Else
            mnuBilledRO.Visible = False
        End If
        PopupMenu mnuOption
    End If
End Sub

Private Sub rptRO_SelectionChanged()
    thestatus = Trim((Null2String(rptRO.SelectedRows(0).Record(8).Value)))
    theRo = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    zRONO = Null2String(rptRO.SelectedRows(0).Record(4).Value)
    
    Call DisplayRODetails(Null2String(rptRO.SelectedRows(0).Record(4).Value))
    Call ComputeMePo
    
    tabDet.SelectedItem = 0
End Sub

Public Sub DisplayRODetails(xREP_OR As String)
    Call ViewJobs(xREP_OR)
End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If TabControl.SelectedItem = 0 Then
        frmJobs.Enabled = True:     frmJobs.Visible = True
        ShortcutCaption1.Caption = "  SCHD HRS              UR"
        lblUR.Caption = lblURCNT:   txtTLhrs.Text = lblHRSCNT
    ElseIf TabControl.SelectedItem = 1 Then
        Call CleanListViewDetails
        frmJobs.Enabled = True:     frmJobs.Visible = True
        ShortcutCaption1.Caption = "ESTIMATE COUNTER"
        lblUR.Caption = 0:          txtTLhrs.Text = 0
    Else
        Call CleanListViewDetails
        frmJobs.Visible = False
        ShortcutCaption1.Caption = "SCHEDULE APPOINTMENT"
        txtTLhrs.Text = "0.00":     lblUR.Caption = lblAppCnt
    End If
End Sub

Private Sub tmr_CICO_Timer()
    'COMMENT BY : MJP012609 0223PM
        'Set rsCicoCount = gconDMIS.Execute("SELECT MAX(CLOCKIN) CLOCKIN, MAX(CLOCKOUT) CLOCKOUT FROM CSMS_JOBCLOCK ")
    'COMMENT BY : MJP012609 0223PM
    
    'UPDATE BY   : MJP012609 0223PM
    'DESCRIPTION : TO CHECK ONLY IN THIS DATE
    '    Set rsCicoCount = gconDMIS.Execute("SELECT MAX(CLOCKIN) CLOCKIN, MAX(CLOCKOUT) CLOCKOUT FROM CSMS_JOBCLOCK " & _
    '        " WHERE MONTH(TRANDATE) = " & MonthView1.Month & _
    '        " AND YEAR(TRANDATE) = " & MonthView1.Year & _
    '        " AND DAY(TRANDATE) = " & MonthView1.Day & "")
    'UPDATE BY   : MJP012609 0223PM
    'If rsCicoCount(0) <> LABTIMEIN Or rsCicoCount(1) <> LABTIMEOUT Then
    '    LABTIMEIN = Null2String(rsCicoCount!CLOCKIN)
    '    LABTIMEOUT = Null2String(rsCicoCount!CLOCKOUT)
    '
    '    MessagePop InfoFriend, "Record Update", "New clockin updated Please Update Your record", 3000
    'Else
    '    'LABTIMEIN = Null2String(rsCicoCount!CLOCKIN)
    '    'LABTIMEOUT = Null2String(rsCicoCount!CLOCKOUT)
    'End If
End Sub

Private Sub txtSearch_Change()
    If TabControl.SelectedItem = 0 Then
        rptRO.FilterText = txtSearch.Text
        rptRO.Populate
    ElseIf TabControl.SelectedItem = 1 Then
        rptEST.FilterText = txtSearch.Text
        rptEST.Populate
    Else
'        rptAPP.FilterText = txtSearch.Text
'        rptAPP.Populate
    End If
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Sub ViewActiveRO()
    Screen.MousePointer = 11
    Dim lng                                            As Long
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim rsrotypeload                                   As New ADODB.Recordset

    Call CleanListViewDetails
    
    'COMMENT BY  : MJP07242009 1137AM
    'DESCRIPTION : TO CHECK IF CICO IS REALLY DEAD
            Set RSUPLOAD = gconDMIS.Execute("Select RO_NO from CSMS_RepairOrder where (TransType = 'R' and (status = 'Working' OR STATUS = 'Idle Time' or status = 'Lunch Break' or status = 'Going Home' or status = 'Break Time')) order by AppointmentDate asc")
            If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
                RSUPLOAD.MoveFirst
        
                Do While Not RSUPLOAD.EOF
                    If CheckAllJobsISDone(Null2String(RSUPLOAD!RO_NO)) = True Then
                        gconDMIS.Execute ("Update CSMS_RepairOrder Set JStatus = 'F',Status = 'Finish Job' Where RO_NO = " & N2Str2Null(RSUPLOAD!RO_NO))
                        Call gconDMIS.Execute("UPDATE HRMS_EMPINFO SET JSTATUS='A' , ASSIGNEDRO=NULL  WHERE ASSIGNEDRO=" & N2Str2Null(RSUPLOAD!RO_NO), lng)
                        If lng = 0 Then
                            gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS='A' , ASSIGNEDRO=NULL  WHERE ASSIGNEDRO=" & N2Str2Null(RSUPLOAD!RO_NO))
                        End If
                    End If
                    RSUPLOAD.MoveNext
                Loop
            End If
            Set RSUPLOAD = New ADODB.Recordset
    'COMMENT BY  : MJP07242009 1137AM

    If Check1.Value = 1 Then
        If CHKSTATUS = "All" Then
            Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate], Customer, Model, PLATE_NO, RO_No, [Hours], [xHrsWork], " & _
                " CASE ISNULL([HOURS],0) " & _
                " WHEN 0 THEN 0 " & _
                " ELSE cast(([xHrsWork] / [Hours]) * 100 as DECIMAL(18,2)) " & _
                " END AS PERC, PromiseDate, [Today], status, Writer, TECH1, TECH2, TECH3, ACCT_NO, datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status not in('Released','Voided')) order by AppointmentDate asc")  ' AND AppointmentDate <= '" & DateValue(MonthView1) & "' order by AppointmentDate asc")    'BTT - 05242007
        Else
            Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate], Customer, Model, PLATE_NO, RO_No, [Hours], [xHrsWork], " & _
                " CASE ISNULL([HOURS],0) " & _
                " WHEN 0 THEN 0 " & _
                " ELSE cast(([xHrsWork] / [Hours]) * 100 as DECIMAL(18,2)) END AS PERC, PromiseDate, [Today], status, Writer, TECH1, TECH2, TECH3, ACCT_NO, datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' Order by RO_No asc")    'BTT - 05242007
        End If
    ElseIf Check2.Value = 1 Then
        If CHKSTATUS = "All" Then
            Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate], Customer, model, PLATE_NO, RO_No, [Hours], [xHrsWork], " & _
                " CASE ISNULL([HOURS],0) " & _
                " WHEN 0 THEN 0 " & _
                "  ELSE cast(([xHrsWork] / [Hours]) * 100 as DECIMAL(18,2)) END AS PERC, PromiseDate, [Today], status, Writer, TECH1, TECH2, TECH3, ACCT_NO, Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and (AppointmentDate = '" & DateValue(MonthView1) & "') order by RO_No asc")
        Else
            Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate], Customer, model, PLATE_NO, RO_No, [Hours], [xHrsWork], " & _
                " CASE ISNULL([HOURS],0) " & _
                " WHEN 0 THEN 0 " & _
                " ELSE cast(([xHrsWork] / [Hours]) * 100 as DECIMAL(18,2)) END AS PERC, PromiseDate, [Today], status, Writer, TECH1, TECH2, TECH3, ACCT_NO, Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' order by RO_No asc")    'MJP - 07172007
        End If
    End If

    Dim REC                                            As XtremeReportControl.ReportRecord
    lblUR.Caption = ""
    rptRO.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptRO.Records.Add
        REC.AddItem (Trim(RSUPLOAD![AppointmentDate]))
        REC.AddItem (Trim(RSUPLOAD!Customer))
        REC.AddItem (Trim(RSUPLOAD!Model))
        REC.AddItem (Trim(RSUPLOAD!PLATE_NO))
        REC.AddItem (Trim(RSUPLOAD!RO_NO))
        REC.AddItem (FormatNumber(RSUPLOAD![HOURS]))
        REC.AddItem (FormatNumber(RSUPLOAD![xHrsWork]))
        REC.AddItem (FormatNumber(RSUPLOAD!Perc))
        REC.AddItem (Trim(RSUPLOAD!Status))
        REC.AddItem (Trim(RSUPLOAD!PromiseDate))
        REC.AddItem (Trim(RSUPLOAD!writer))
        REC.AddItem (Trim(RSUPLOAD!datefinish))
        REC.AddItem (Trim(Replace(Null2String(RSUPLOAD!tech2), vbCrLf, " ")))
        REC.AddItem (Trim(RSUPLOAD!ACCT_NO))
        
        'Updated by:    IEBV 06282010 1037AM
        'Description:   To display Rotype for the HCI
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        If COMPANY_CODE = "HCI" Then
            Dim XXX         As String
            Set rsrotypeload = gconDMIS.Execute("select ROTYPE from csms_repor where rep_or = '" & Null2String(RSUPLOAD!RO_NO) & "'")
            If Not (rsrotypeload.BOF And rsrotypeload.EOF) Then
                XXX = ShowROTYPEdetails(Null2String(rsrotypeload!ROTYPE))
                REC.AddItem (Trim(XXX))
            End If
            Set rsrotypeload = Nothing
        End If
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptRO.Populate
    lblUR.Caption = rptRO.Records.Count
    lblURCNT.Caption = lblUR.Caption
    Screen.MousePointer = 0
    Set RSUPLOAD = Nothing
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub ViewEstimate()
    Screen.MousePointer = 11
    Dim RSUPLOAD                                       As New ADODB.Recordset

    Call CleanListViewDetails

    Set RSUPLOAD = gconDMIS.Execute("Select ID, SERVICE_ADVISER, ISNULL(UPLOAD_STATUS,'N') as UPLOAD_STATUS, DTE_RECD, NIYM, Model, PLATE_NO, RO_NO, ESTIMATENO, STATUS, " & _
        " RECD_BY, DATE_UPLOAD, ACCT_NO from CSMS_ESTHD " & _
        " where DTE_RECD = '" & DateValue(MonthView1) & _
        "' order by ESTIMATENO asc")
    Dim REC                                            As XtremeReportControl.ReportRecord
    rptEST.Records.DeleteAll
    While Not RSUPLOAD.EOF
        Set REC = rptEST.Records.Add
        REC.AddItem (Trim(Null2String(RSUPLOAD!DTE_RECD)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!NIYM)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!Model)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!PLATE_NO)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!RO_NO)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!EstimateNo)))
        If Null2String(RSUPLOAD!UPLOAD_Status) = "N" Then
            REC.AddItem "Not Yet Uploaded"
        Else
            REC.AddItem "Uploaded"
        End If
        REC.AddItem (Trim(Null2String(RSUPLOAD!SERVICE_ADVISER)))
        REC.AddItem (Trim(Null2String(RSUPLOAD![DATE_UPLOAD])))
        REC.AddItem (Trim(Null2String(RSUPLOAD!ACCT_NO)))
        REC.AddItem (Trim(Null2String(RSUPLOAD!ID)))

        RSUPLOAD.MoveNext
        Set REC = Nothing
    Wend
    rptEST.Populate
    Set RSUPLOAD = Nothing

    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub ViewAppointment()
    Screen.MousePointer = 11
    Dim RSUPLOAD                                        As New ADODB.Recordset

    Call CleanListViewDetails

    Dim xTranDate                                       As String
    Dim RSVIEWGRID                                      As New ADODB.Recordset
    Dim RSAPPT                                          As New ADODB.Recordset
    Dim REC                                             As XtremeReportControl.ReportRecord
    Dim rstmp                                           As New ADODB.Recordset
    
    Set rstmp = gconDMIS.Execute("SELECT COUNT(ID) FROM CSMS_APPOINTMENT WHERE TRANDATE = " & N2Str2Null(MonthView1.Value) & "")
    lblAppCnt.Caption = rstmp.Fields(0)
    Set rstmp = Nothing
    
    rptAPP.Records.DeleteAll
    
    RSVIEWGRID.Open "select  TimeInterval from CSMS_ApptSchedule order by ID asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSVIEWGRID.EOF And Not RSVIEWGRID.BOF Then
        
        xTranDate = N2Str2Null(Format(MonthView1, "MM/dd/yyyy"))
        Do While Not RSVIEWGRID.EOF
            Set RSAPPT = gconDMIS.Execute("select * from CSMS_APPOINTMENT where " & _
                " ApptTime = '" & RSVIEWGRID!TimeInterval & _
                "' and trandate = " & xTranDate & "")
            If Not (RSAPPT.EOF And RSAPPT.BOF) Then
                Set REC = rptAPP.Records.Add
                REC.AddItem (Trim(Null2String(RSAPPT!APPTNO)))
                REC.AddItem (Trim(Null2String(RSVIEWGRID!TimeInterval)))
                REC.AddItem (Trim(Null2String(RSAPPT!CUSNAM)))
                REC.AddItem (Trim(Null2String(RSAPPT!Model)))
                REC.AddItem (Trim(Null2String(RSAPPT!PLATE_NO)))
                REC.AddItem (Trim(Null2String(RSAPPT!Status)))
                REC.AddItem (Trim(Null2String(RSAPPT!ID)))
                REC.AddItem (Trim(Null2String(RSAPPT!CUSCDE)))
            Else
                Set REC = rptAPP.Records.Add
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String(RSVIEWGRID!TimeInterval)))
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String("")))
                REC.AddItem (Trim(Null2String("")))
            End If
            
            RSVIEWGRID.MoveNext
        Loop
    End If
    rptAPP.Populate
    Set RSVIEWGRID = Nothing
    
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Function MakeApptNo() As String
    Dim rsMakeAptNo                                    As New ADODB.Recordset
    rsMakeAptNo.Open "select [ApptNo] from CSMS_Appointment order by ApptNo desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMakeAptNo.EOF And Not rsMakeAptNo.BOF Then
        MakeApptNo = Format(Val(rsMakeAptNo![APPTNO]) + 1, "000000000")
    Else
        MakeApptNo = Format(1, "000000000")
    End If
End Function

Sub ViewJobs(XRONO As String)
    Dim RSUPLOAD                                       As New ADODB.Recordset
    Dim Item                                           As ListItem

    Call CleanListViewDetails
    DoEvents
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETAIL,FLATRATE,det_hrs,TECHNICIAN,HRSWRK,REP_OR,TechCode,LINE_NO,status from CSMS_Ro_Det where LIVIL = '1' " & _
        " AND REP_OR = '" & XRONO & _
        "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstJob4Service.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))

            Item.SubItems(1) = Null2String(RSUPLOAD!Detail)
            Item.SubItems(2) = Format(NumericVal(RSUPLOAD!FLATRATE), MAXIMUM_DIGIT)
            Item.SubItems(3) = Null2String(RSUPLOAD!DET_HRS)
            Item.SubItems(4) = FindTechName(LTrim(RTrim(Null2String(RSUPLOAD!TechCode))))
            Item.SubItems(5) = Null2String(RSUPLOAD!HRSWRK)
            Item.SubItems(6) = Null2String(RSUPLOAD!REP_OR)
            Item.SubItems(7) = Null2String(RSUPLOAD!TechCode)
            Item.SubItems(8) = Null2String(RSUPLOAD!LINE_NO)

            If Null2String(RSUPLOAD!Status) = "W" Then Item.SubItems(9) = "Working": Call CHECK_IN_OUT(RSUPLOAD!REP_OR, RSUPLOAD!DETCDE, RSUPLOAD!TechCode, RSUPLOAD!Status) 'UPDATED BY JUN: 03-23-2009
            If Null2String(RSUPLOAD!Status) = "I" Then Item.SubItems(9) = "Idle Time"
            If Null2String(RSUPLOAD!Status) = "L" Then Item.SubItems(9) = "Lunch Break"
            If Null2String(RSUPLOAD!Status) = "G" Then Item.SubItems(9) = "Going Home"
            If Null2String(RSUPLOAD!Status) = "B" Then Item.SubItems(9) = "Break Time"
            If Null2String(RSUPLOAD!Status) = "Y" Or Null2String(RSUPLOAD!Status) = "R" Then Item.SubItems(9) = "Finish Job"

            If Null2String(RSUPLOAD!Status) = "J" Then Item.SubItems(9) = "Back Job"
            If Null2String(RSUPLOAD!Status) = "Q" Then Item.SubItems(9) = "Waiting for QC"
            RSUPLOAD.MoveNext
        Loop
    End If

    'PMS JOBS
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,PMS_MODEL from CSMS_PMS_Job_det " & _
        " where REP_OR = '" & XRONO & _
        "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstPMSJobs.ListItems, RSUPLOAD
    End If
    Set RSUPLOAD = Nothing

    'PARTS
    DoEvents
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det " & _
        " where LIVIL = '2' " & _
        " AND REP_OR = '" & XRONO & _
        "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstParts.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)
            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing

    'MATERIALS
    'Set RSUPLOAD = New ADODB.Recordset
    DoEvents
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det " & _
        " where LIVIL = '3' " & _
        " AND REP_OR = '" & XRONO & _
        "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            DoEvents
            Set Item = lstMaterials.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)

            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing

    'ACCESSORIES
    'Set RSUPLOAD = New ADODB.Recordset
    DoEvents
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det " & _
        " where LIVIL = '4' " & _
        " AND REP_OR = '" & XRONO & _
        "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            DoEvents
            Set Item = lstAccessories.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))
            Item.SubItems(1) = Null2String(RSUPLOAD!DETDSC)
            Item.SubItems(2) = Null2String(RSUPLOAD!detvol)
            Item.SubItems(3) = Format(Null2String(RSUPLOAD!DetPrc), MAXIMUM_DIGIT)
            Item.SubItems(4) = Format(Null2String(RSUPLOAD!DET_AMT), MAXIMUM_DIGIT)

            RSUPLOAD.MoveNext
        Loop
    End If
    Set RSUPLOAD = Nothing
End Sub

Sub CHECK_IN_OUT(xRO_NO As String, xDET_CODE As String, xTECHCODE As String, xJStatus As String)
    'UPDATED BY: JUN
    'DATE UPDATED: 03-23-2009
    'DESCRIPTION: TCN: 12759 CONCERN BY HGC AND HCI
    'CHECK IN CSMS_JOBCLOCK TABLE THE STATUS OF THE TECHNICIAN IF THE TECHNICIAN HAS REALLY LOG IN OR OUT
    
    Dim rsGET_EMPNO As ADODB.Recordset
    Dim RSJOBCLOCK As ADODB.Recordset
    Dim rsHRMS As ADODB.Recordset
    Dim xEMPNO As String
            
    Set rsGET_EMPNO = gconDMIS.Execute("Select EMPNO from CSMS_vw_technician where TECHNICIAN  = '" & RTrim(LTrim(xTECHCODE)) & "'")
        If Not rsGET_EMPNO.EOF And Not rsGET_EMPNO.BOF Then
            xEMPNO = Null2String(rsGET_EMPNO!EMPNO)
        End If
        
    Set RSJOBCLOCK = gconDMIS.Execute("Select * from CSMS_JOBCLOCK where RO_NO = '" & LTrim(RTrim(xRO_NO)) & "' and DETCDE = '" & LTrim(RTrim(xDET_CODE)) & "' and JSTATUS ='" & RTrim(LTrim(xJStatus)) & "' and TECHNICIAN = '" & LTrim(RTrim(xEMPNO)) & "'")
    If Not RSJOBCLOCK.EOF And Not RSJOBCLOCK.BOF Then
        Set rsHRMS = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where EMPNO = '" & LTrim(RTrim(xEMPNO)) & "' and IS_TECHNICIAN = '1'")
        If Not rsHRMS.EOF And Not rsHRMS.BOF Then
           gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = '" & xRO_NO & "', JSTATUS = '" & xJStatus & "' where EMPNO = '" & xEMPNO & "' AND IS_TECHNICIAN = '1'")
        Else
           gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET ASSIGNEDRO = '" & xRO_NO & "', JSTATUS = '" & xJStatus & "' where EMPNO = '" & xEMPNO & "'")
        End If
    Else
        Dim rsRECHECH_STATUS As ADODB.Recordset
                
        Set rsRECHECH_STATUS = gconDMIS.Execute("Select REASONFORCLOCKOUT FROM CSMS_JOBCLOCK WHERE REASONFORCLOCKOUT = 'Finish Job -' AND RO_NO = '" & LTrim(RTrim(xRO_NO)) & "' AND DETCDE = '" & LTrim(RTrim(xDET_CODE)) & "' and TECHNICIAN = '" & LTrim(RTrim(xEMPNO)) & "'")
            If Not rsRECHECH_STATUS.EOF And Not rsRECHECH_STATUS.BOF Then
                gconDMIS.Execute "update CSMS_ro_det set " & _
                    " STATUS = 'Y' " & _
                    ", Done = 'Y' " & _
                    " where LIVIL = '1' " & _
                    " AND RO_NO = '" & xRO_NO & _
                    "' and DETCDE = '" & RTrim(LTrim(xDET_CODE)) & "'"
            Else
                Set rsHRMS = gconDMIS.Execute("Select EMPNO from HRMS_EMPINFO where EMPNO = '" & LTrim(RTrim(xEMPNO)) & "' and IS_TECHNICIAN = '1'")
                If Not rsHRMS.EOF And Not rsHRMS.BOF Then
                   gconDMIS.Execute ("UPDATE HRMS_EMPINFO SET ASSIGNEDRO = '" & xRO_NO & "', JSTATUS = '" & xJStatus & "' where EMPNO = '" & xEMPNO & "' AND IS_TECHNICIAN = '1'")
                Else
                   gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET ASSIGNEDRO = '" & xRO_NO & "', JSTATUS = '" & xJStatus & "' where EMPNO = '" & xEMPNO & "'")
                End If
            End If
        Set rsRECHECH_STATUS = Nothing
    End If
    
    Set rsGET_EMPNO = Nothing
    Set RSJOBCLOCK = Nothing
End Sub

Sub InitializeReportControl()
    Screen.MousePointer = 11
    
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "Date", 75, True::              .Columns(0).Resizable = False:                 .Columns(0).AllowRemove = False
        .Columns.Add 1, "Customer", 175, True:          .Columns(1).AllowRemove = False
        .Columns.Add 2, "Vehicle", 100, True:           .Columns(2).AllowRemove = False
        .Columns.Add 3, "Plate no", 60, True:           .Columns(3).Alignment = xtpAlignmentCenter:     .Columns(3).AllowRemove = False
        .Columns.Add 4, "RO NO", 75, True:              .Columns(4).Alignment = xtpAlignmentLeft:       .Columns(4).AllowRemove = False
        .Columns.Add 5, "Std Hrs", 60, True:            .Columns(5).Alignment = xtpAlignmentCenter:     .Columns(5).AllowRemove = False
        .Columns.Add 6, "Hrs Work", 60, True:           .Columns(6).Alignment = xtpAlignmentCenter:     .Columns(6).AllowRemove = False
        .Columns.Add 7, "(%)", 60, True:                .Columns(7).Alignment = xtpAlignmentCenter:     .Columns(7).AllowRemove = False
        .Columns.Add 8, "Status", 80, True:             .Columns(8).Alignment = xtpAlignmentLeft:       .Columns(8).AllowRemove = False
        .Columns.Add 9, "Promise", 140, True:           .Columns(9).Alignment = xtpAlignmentLeft:       .Columns(9).AllowRemove = False
        .Columns.Add 10, "Service Advisor", 130, True:  .Columns(10).Alignment = xtpAlignmentLeft:      .Columns(10).AllowRemove = False
        .Columns.Add 11, "Date Finish", 80, True:       .Columns(11).Alignment = xtpAlignmentLeft:      .Columns(11).AllowRemove = False
        .Columns.Add 12, "Remarks", 120, True:          .Columns(12).Alignment = xtpAlignmentLeft:      .Columns(12).AllowRemove = False
        .Columns.Add 13, "Cust Code", 0, False:         .Columns(13).Alignment = xtpAlignmentLeft:      .Columns(13).AllowRemove = False
'Updated by:    IEBV 06282010 1034AM
'Description:   To Add ROTYPE column if dealer is HCI
    If COMPANY_CODE = "HCI" Then
        .Columns.Add 14, "RO Type", 140, True:          .Columns(14).Alignment = xtpAlignmentLeft:      .Columns(14).AllowRemove = False
    End If
'Updated by:    IEBV 06282010 1034AM
'Description:   To Add ROTYPE column if dealer is HCI
        
        '.GroupsOrder.Add .Columns(8)
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.TextFont.Weight = 540
        .PaintManager.CaptionFont.Bold = True
        '.PaintManager.TextFont.Bold = True
    End With
    
    With rptEST
        .Columns.DeleteAll
        .Columns.Add 0, "Date", 75, True:               .Columns(0).Resizable = False:              .Columns(0).AllowRemove = False
        .Columns.Add 1, "Customer", 175, True:          .Columns(1).AllowRemove = False
        .Columns.Add 2, "Vehicle", 100, True:           .Columns(2).AllowRemove = False
        .Columns.Add 3, "Plate no", 60, True:           .Columns(3).Alignment = xtpAlignmentCenter: .Columns(3).AllowRemove = False
        .Columns.Add 4, "Repair Order no.", 100, True:   .Columns(4).Alignment = xtpAlignmentLeft:   .Columns(4).AllowRemove = False
        .Columns.Add 5, "Estimate no", 85, True:        .Columns(5).Alignment = xtpAlignmentLeft:   .Columns(5).AllowRemove = False
        .Columns.Add 6, "Status", 100, True:            .Columns(6).Alignment = xtpAlignmentLeft:   .Columns(6).AllowRemove = False
        .Columns.Add 7, "Service Advisor", 130, True:   .Columns(7).Alignment = xtpAlignmentLeft:   .Columns(7).AllowRemove = False
        .Columns.Add 8, "Date Uploaded", 110, True:     .Columns(8).Alignment = xtpAlignmentLeft:   .Columns(8).AllowRemove = False
        .Columns.Add 9, "Cust Code", 0, False:          .Columns(9).Alignment = xtpAlignmentCenter: .Columns(9).AllowRemove = False
        .Columns.Add 10, "ID", 0, False:                .Columns(10).Alignment = xtpAlignmentCenter: .Columns(10).AllowRemove = False
        
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        '.PaintManager.TextFont.Bold = True
    End With
    
    With rptAPP
        .Columns.DeleteAll
        .Columns.Add 0, "APPT NO.", 50, True:           .Columns(0).Resizable = False:              .Columns(0).AllowRemove = False
        .Columns.Add 1, "TIME", 45, True:               .Columns(1).AllowRemove = False
        .Columns.Add 2, "CUSTOMER NAME", 220, True:     .Columns(2).AllowRemove = False
        .Columns.Add 3, "MODEL", 80, True:              .Columns(3).Alignment = xtpAlignmentLeft:   .Columns(3).AllowRemove = False
        .Columns.Add 4, "PLATE NO", 50, True:           .Columns(4).Alignment = xtpAlignmentLeft:   .Columns(4).AllowRemove = False
        .Columns.Add 5, "STATUS", 60, True:             .Columns(5).Alignment = xtpAlignmentLeft:   .Columns(5).AllowRemove = False
        .Columns.Add 6, "ID", 0, True:                  .Columns(6).Alignment = xtpAlignmentLeft:   .Columns(6).AllowRemove = False
        .Columns.Add 7, "cuscde", 0, True:              .Columns(7).Alignment = xtpAlignmentLeft:   .Columns(6).AllowRemove = False

        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        '.PaintManager.TextFont.Bold = True
    End With
    
    Screen.MousePointer = 0
End Sub
Function ShowROTYPEdetails(XXX As String) As String
    If (XXX) = "WTY" Then
       ShowROTYPEdetails = "Warranty"
    ElseIf (XXX) = "BRP" Then
       ShowROTYPEdetails = "Body Repair and Painting"
    ElseIf (XXX) = "JET" Then
       ShowROTYPEdetails = "Jet Service"
    ElseIf (XXX) = "GJ" Then
       ShowROTYPEdetails = "General Job"
    ElseIf (XXX) = "O/H" Then
       ShowROTYPEdetails = "Overhauling"
    ElseIf (XXX) = "PP" Then
       ShowROTYPEdetails = "Painting Protection"
    ElseIf (XXX) = "RF" Then
       ShowROTYPEdetails = "Rustproofing"
    ElseIf (XXX) = "QS" Then
       ShowROTYPEdetails = "Quick Service"
    ElseIf (XXX) = "AC" Then
       ShowROTYPEdetails = "Aircon/Electrical"
    ElseIf (XXX) = "PDI" Then
       ShowROTYPEdetails = "Pre-delivery Inspection"
    ElseIf (XXX) = "QC" Then
       ShowROTYPEdetails = "Quality Control"
    ElseIf (XXX) = "FI" Then
       ShowROTYPEdetails = "Final Inspection"
    ElseIf (XXX) = "DET" Then
       ShowROTYPEdetails = "Detailing"
    Else
       ShowROTYPEdetails = ""
    End If
End Function


