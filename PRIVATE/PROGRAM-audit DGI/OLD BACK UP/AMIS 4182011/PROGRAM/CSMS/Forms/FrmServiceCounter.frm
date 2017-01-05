VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Object = "{A9046457-E246-455F-A58F-D670C44E8BEA}#2.0#0"; "wizFlexCracker.ocx"
Begin VB.Form frmCSMSServiceCounter 
   BackColor       =   &H00F5F5F5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Service Counter"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15180
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
   Icon            =   "FrmServiceCounter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15180
   Begin VB.Timer tmr_CICO 
      Interval        =   1000
      Left            =   900
      Top             =   1170
   End
   Begin VB.CheckBox Check2 
      Caption         =   "All Ro for Selected Date"
      Height          =   375
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   750
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
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
      Height          =   705
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   15180
      TabIndex        =   13
      Top             =   0
      Width           =   15180
      Begin VB.CommandButton Command8 
         Height          =   315
         Left            =   6930
         MouseIcon       =   "FrmServiceCounter.frx":05CA
         MousePointer    =   99  'Custom
         Picture         =   "FrmServiceCounter.frx":08D4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   60
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00F5F5F5&
         Caption         =   "Legend: (Click on Legend to Filter By Status)"
         Height          =   645
         Left            =   7500
         TabIndex        =   34
         Top             =   30
         Width           =   7635
         Begin VB.Shape Shape1 
            BorderColor     =   &H00000000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   60
            Top             =   263
            Width           =   195
         End
         Begin VB.Shape Shape2 
            FillColor       =   &H00C0C000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   1890
            Top             =   263
            Width           =   195
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H00000000&
            FillColor       =   &H000000C0&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   4950
            Top             =   263
            Width           =   195
         End
         Begin VB.Label labPark 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Park"
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
            Left            =   210
            MouseIcon       =   "FrmServiceCounter.frx":110E
            MousePointer    =   99  'Custom
            TabIndex        =   41
            ToolTipText     =   "Click to view Repair Order Filter by Parked Status"
            Top             =   255
            Width           =   480
         End
         Begin VB.Label labWork 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Working"
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
            Left            =   2040
            MouseIcon       =   "FrmServiceCounter.frx":1260
            MousePointer    =   99  'Custom
            TabIndex        =   40
            ToolTipText     =   "Click to view Repair Order Filter by Working Status"
            Top             =   255
            Width           =   795
         End
         Begin VB.Shape Shape5 
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   5730
            Top             =   263
            Width           =   195
         End
         Begin VB.Label labBilled 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Billed"
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
            Left            =   5880
            MouseIcon       =   "FrmServiceCounter.frx":13B2
            MousePointer    =   99  'Custom
            TabIndex        =   38
            ToolTipText     =   "Click to view Repair Order that are already Billed"
            Top             =   255
            Width           =   555
         End
         Begin VB.Label labOver 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Over"
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
            Left            =   5100
            MouseIcon       =   "FrmServiceCounter.frx":1504
            MousePointer    =   99  'Custom
            TabIndex        =   37
            ToolTipText     =   "Click to view Repair Order Filter by Due by promised Date"
            Top             =   255
            Width           =   495
         End
         Begin VB.Shape Shape6 
            FillColor       =   &H00C00000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   4080
            Top             =   263
            Width           =   195
         End
         Begin VB.Shape Shape3 
            FillColor       =   &H00008000&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   6510
            Top             =   263
            Width           =   195
         End
         Begin VB.Shape Shape7 
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   840
            Top             =   263
            Width           =   195
         End
         Begin VB.Shape Shape 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   195
            Left            =   2940
            Top             =   263
            Width           =   195
         End
         Begin VB.Label labIdleTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Idle Time"
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
            Left            =   3090
            MouseIcon       =   "FrmServiceCounter.frx":1656
            MousePointer    =   99  'Custom
            TabIndex        =   35
            ToolTipText     =   "Click to view Repair Order Filter by Ideal Status"
            Top             =   255
            Width           =   870
         End
         Begin VB.Label labBackJob 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Back Job"
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
            Left            =   990
            MouseIcon       =   "FrmServiceCounter.frx":17A8
            MousePointer    =   99  'Custom
            TabIndex        =   42
            ToolTipText     =   "Click to view Repair Order Filter by Back Job Status"
            Top             =   255
            Width           =   840
         End
         Begin VB.Label labFinish 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Finish"
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
            Left            =   4230
            MouseIcon       =   "FrmServiceCounter.frx":18FA
            MousePointer    =   99  'Custom
            TabIndex        =   39
            ToolTipText     =   "Click to view Repair Order Filter by Finished Job Status"
            Top             =   255
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Released"
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
            Left            =   6660
            MouseIcon       =   "FrmServiceCounter.frx":1A4C
            MousePointer    =   99  'Custom
            TabIndex        =   36
            ToolTipText     =   "Click to view Repair Order that are already Released"
            Top             =   255
            Width           =   870
         End
      End
      Begin VB.CommandButton Command5 
         Height          =   675
         Left            =   2190
         MouseIcon       =   "FrmServiceCounter.frx":1B9E
         MousePointer    =   99  'Custom
         Picture         =   "FrmServiceCounter.frx":1EA8
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Job Data Entry"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Height          =   675
         Left            =   1470
         MouseIcon       =   "FrmServiceCounter.frx":2673
         MousePointer    =   99  'Custom
         Picture         =   "FrmServiceCounter.frx":297D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   675
         Left            =   750
         MouseIcon       =   "FrmServiceCounter.frx":3158
         MousePointer    =   99  'Custom
         Picture         =   "FrmServiceCounter.frx":3462
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Time Clock/Job Clock Log-In"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Height          =   675
         Left            =   30
         MouseIcon       =   "FrmServiceCounter.frx":3DD5
         MousePointer    =   99  'Custom
         Picture         =   "FrmServiceCounter.frx":40DF
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   30
         Width           =   735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   32
         Top             =   90
         Width           =   3975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Index           =   0
         Left            =   3030
         TabIndex        =   33
         Top             =   150
         Width           =   3975
      End
   End
   Begin VB.Frame frmJobs 
      BackColor       =   &H00F5F5F5&
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
      Height          =   2745
      Left            =   2370
      TabIndex        =   24
      Top             =   7350
      Width           =   13005
      Begin TabDlg.SSTab SSTab1 
         Height          =   2655
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   12765
         _ExtentX        =   22516
         _ExtentY        =   4683
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   617
         ShowFocusRect   =   0   'False
         BackColor       =   16119285
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "View Jobs"
         TabPicture(0)   =   "FrmServiceCounter.frx":49D4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstJob4Service"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "View PMS Jobs "
         TabPicture(1)   =   "FrmServiceCounter.frx":49F0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstPMSJobs"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "View Issued Parts"
         TabPicture(2)   =   "FrmServiceCounter.frx":4A0C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstParts"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "View Issued Materials"
         TabPicture(3)   =   "FrmServiceCounter.frx":4A28
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lstMaterials"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "View Issued Accessories"
         TabPicture(4)   =   "FrmServiceCounter.frx":4A44
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lstAccessories"
         Tab(4).ControlCount=   1
         Begin MSComctlLib.ListView lstJob4Service 
            Height          =   2115
            Left            =   60
            TabIndex        =   26
            Top             =   420
            Width           =   12615
            _ExtentX        =   22251
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmServiceCounter.frx":4A60
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
         Begin MSComctlLib.ListView lstMaterials 
            Height          =   2115
            Left            =   -74940
            TabIndex        =   27
            Top             =   420
            Width           =   12615
            _ExtentX        =   22251
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmServiceCounter.frx":4BC2
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
         Begin MSComctlLib.ListView lstParts 
            Height          =   2115
            Left            =   -74940
            TabIndex        =   28
            Top             =   420
            Width           =   12645
            _ExtentX        =   22304
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmServiceCounter.frx":4D24
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
         Begin MSComctlLib.ListView lstPMSJobs 
            Height          =   2115
            Left            =   -74940
            TabIndex        =   29
            Top             =   420
            Width           =   12615
            _ExtentX        =   22251
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmServiceCounter.frx":4E86
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Code"
               Object.Width           =   1587
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
         Begin MSComctlLib.ListView lstAccessories 
            Height          =   2115
            Left            =   -74940
            TabIndex        =   30
            Top             =   420
            Width           =   12615
            _ExtentX        =   22251
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "FrmServiceCounter.frx":4FE8
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
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FrmServiceCounter.frx":514A
      Left            =   2370
      List            =   "FrmServiceCounter.frx":5157
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   1200
      Width           =   3225
   End
   Begin VB.CheckBox chkWithColors 
      BackColor       =   &H00F5F5F5&
      Caption         =   "Display With Colors"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7830
      TabIndex        =   22
      Top             =   840
      Width           =   2205
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Clear"
      Height          =   375
      Left            =   13860
      MouseIcon       =   "FrmServiceCounter.frx":51A0
      MousePointer    =   99  'Custom
      TabIndex        =   21
      ToolTipText     =   "Next Date "
      Top             =   1193
      Width           =   1245
   End
   Begin VB.TextBox txtSearchName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   19
      Top             =   1208
      Width           =   6945
   End
   Begin VB.CommandButton cmdDateForward 
      Caption         =   "Tomorrow"
      Height          =   375
      Left            =   13860
      MouseIcon       =   "FrmServiceCounter.frx":54AA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Next Date "
      Top             =   750
      Width           =   1245
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All Open and Current R/O's"
      Height          =   375
      Left            =   2370
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   750
      Width           =   2475
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9090
      Top             =   270
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
      Height          =   4905
      Left            =   30
      ScaleHeight     =   4905
      ScaleWidth      =   2310
      TabIndex        =   5
      Top             =   1620
      Width           =   2310
      Begin VB.CommandButton cmdExit 
         Caption         =   "EXIT"
         Height          =   675
         Left            =   30
         TabIndex        =   49
         Top             =   3960
         Width           =   2235
      End
      Begin VB.CommandButton cmdPartsInquiry 
         Caption         =   "PART INQUIRY"
         Height          =   675
         Left            =   30
         TabIndex        =   48
         Top             =   3300
         Width           =   2235
      End
      Begin VB.CommandButton cmdViewRODetails 
         Caption         =   "VIEW R/O DETAILS"
         Height          =   675
         Left            =   30
         TabIndex        =   47
         Top             =   2640
         Width           =   2235
      End
      Begin VB.CommandButton cmdWriteEstimate 
         Caption         =   "WRITE ESTIMATE"
         Height          =   675
         Left            =   30
         TabIndex        =   46
         Top             =   1980
         Width           =   2235
      End
      Begin VB.CommandButton cmdEditRO 
         Caption         =   "EDIT R/O"
         Height          =   675
         Left            =   30
         TabIndex        =   45
         Top             =   1320
         Width           =   2235
      End
      Begin VB.CommandButton cmdWriteRO 
         Caption         =   "CREATE R/O"
         Height          =   675
         Left            =   30
         TabIndex        =   44
         Top             =   660
         Width           =   2235
      End
      Begin VB.CommandButton cmdWriteAppointment 
         Caption         =   "  APPOINTMENT"
         Height          =   675
         Left            =   30
         TabIndex        =   43
         Top             =   0
         Width           =   2235
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2145
      Left            =   60
      TabIndex        =   6
      Top             =   7770
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
      StartOfWeek     =   54788097
      TitleBackColor  =   8388608
      TitleForeColor  =   16777215
      TrailingForeColor=   13932144
      CurrentDate     =   38458
   End
   Begin VB.TextBox txtTLhrs 
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "24 Schedule Hrs"
      Top             =   825
      Width           =   4305
   End
   Begin VB.CommandButton cmdToday 
      Caption         =   "Today"
      Height          =   375
      Left            =   12630
      MouseIcon       =   "FrmServiceCounter.frx":57B4
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Present Date"
      Top             =   750
      Width           =   1245
   End
   Begin VB.CommandButton cmdDateBack 
      Caption         =   "Yesterday"
      Height          =   375
      Left            =   11400
      MouseIcon       =   "FrmServiceCounter.frx":5ABE
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "Previous Date"
      Top             =   750
      Width           =   1245
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "F5 - Refresh"
      Height          =   375
      Left            =   10110
      MouseIcon       =   "FrmServiceCounter.frx":5DC8
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Refresh"
      Top             =   750
      Width           =   1305
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Search"
      Height          =   375
      Left            =   12630
      MouseIcon       =   "FrmServiceCounter.frx":60D2
      MousePointer    =   99  'Custom
      TabIndex        =   20
      ToolTipText     =   "Present Date"
      Top             =   1193
      Width           =   1245
   End
   Begin FlexCell.Grid grdCounter 
      Height          =   5685
      Left            =   2400
      TabIndex        =   8
      Top             =   1620
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10028
      Appearance      =   0
      BackColorBkg    =   -2147483645
      Cols            =   5
      DefaultFontName =   "Arial"
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      GridColor       =   -2147483634
      Rows            =   30
      SelectionMode   =   1
      EnterKeyMoveTo  =   1
   End
   Begin MSComctlLib.ListView lstCounter 
      Height          =   4155
      Left            =   2550
      TabIndex        =   2
      Top             =   2100
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   7329
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
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmServiceCounter.frx":63DC
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Appt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Vehicle"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Plate NO."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "R/O"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Std Hrs"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Hr Wrk"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Percentage"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Promise Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Today"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Status"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Service Adviser"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Account No"
         Object.Width           =   0
      EndProperty
   End
   Begin wizFlexCracker.wizFlexCrack wizFlexCrack1 
      Height          =   3765
      Left            =   2610
      TabIndex        =   7
      Top             =   2100
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   6641
   End
   Begin VB.Label LABTIMEIN 
      Caption         =   "Label2"
      Height          =   225
      Left            =   60
      TabIndex        =   51
      Top             =   7170
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label LABTIMEOUT 
      Caption         =   "Label2"
      Height          =   225
      Left            =   60
      TabIndex        =   50
      Top             =   6900
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.Label lblROChange 
      BackColor       =   &H000000FF&
      Height          =   405
      Left            =   30
      TabIndex        =   31
      Top             =   6420
      Visible         =   0   'False
      Width           =   2265
   End
   Begin MSForms.Label Label13 
      Height          =   570
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   615
      ForeColor       =   192
      VariousPropertyBits=   8388627
      PicturePosition =   262148
      Size            =   "1085;1005"
      FontName        =   "Arial Narrow"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuOption1 
         Caption         =   "Add &General Job(s)"
      End
      Begin VB.Menu mnuOtherJobs 
         Caption         =   "Add &Other Jobs"
      End
      Begin VB.Menu mnuOption13 
         Caption         =   "Add &PMS Jobs"
      End
      Begin VB.Menu mnuCanedJob 
         Caption         =   "Add &Canned Labor"
      End
      Begin VB.Menu mnuOption2 
         Caption         =   "&Edit Repair Order (R/O)"
      End
      Begin VB.Menu mnuChangeVehicle 
         Caption         =   "Change Vehicle"
      End
      Begin VB.Menu mnuAsgnedbay 
         Caption         =   "&Assign to Bay"
      End
      Begin VB.Menu mnuremovebay 
         Caption         =   "&Remove to bay"
      End
      Begin VB.Menu mnuBilledRO 
         Caption         =   "Bill Repair Order"
      End
      Begin VB.Menu mnuOption3 
         Caption         =   "&Show Clock In/Out"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBACKJOB 
         Caption         =   "Tag Repair Order as Back Job"
      End
   End
   Begin VB.Menu mnuOption4 
      Caption         =   "Option4"
      Begin VB.Menu mnuAsgnedTech 
         Caption         =   "Assign Technician"
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
   Begin VB.Menu mnuEstimate 
      Caption         =   "Estimate"
      Begin VB.Menu mnuEstimateAdd 
         Caption         =   "&Add New Estimate"
      End
      Begin VB.Menu mnuUpdateEstimate 
         Caption         =   "&Upload Estimate to Repair Order"
      End
   End
End
Attribute VB_Name = "frmCSMSServiceCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCicoCount                                        As ADODB.Recordset
Dim AUDIT_SQL                                          As String
Dim thestatus                                          As String
Dim theRo                                              As String
Dim Thedate                                            As Date
Dim tlHrs                                              As Double
Dim tlFR                                               As Double
Dim bevvy                                              As Long
Dim CHKSTATUS                                          As String
Dim zRONO                                              As String
Dim thetechcode                                        As String
Dim ISLOGIN                                            As Boolean
Dim theflatrate                                        As Double
Dim THESTDRATE                                         As Double
Dim THEJOBCODE                                         As String
Dim THEJOBDEST                                         As String
Dim PERJOBSTATUS                                       As String
Dim vlineNo                                            As String
Dim XREPAIRORDER                                       As String
Dim XPREVIOUS_DATE                                     As String
Dim WithEvents JOBCLOCKFORM                            As frmCSMSClockINOUT
Attribute JOBCLOCKFORM.VB_VarHelpID = -1

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check2.Value = 0
        CHKSTATUS = "All"
        cmdRefresh_Click
    End If
    If Check1.Value = 0 Then
        Check2.Value = 1
        CHKSTATUS = "All"
        cmdRefresh_Click
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
    Dim RS                                             As ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DONE  FROM CSMS_RO_DET WHERE LIVIL = '1' AND (DONE ='N' OR DONE='W' OR DONE IS NULL) and REP_OR = '" & XXX & "'")
    'If XXX = "R-00001515" Then Stop
    If RS.EOF And RS.BOF Then
        CheckAllJobsISDone = True
    Else
        CheckAllJobsISDone = False
    End If
    Set RS = Nothing
End Function

Function CheckIfJobIsFinish(vLINE_NO As String) As Boolean
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT DONE FROM CSMS_RO_dET WHERE REP_OR = '" & theRo & "' and line_no = '" & vLINE_NO & "' AND LIVIL = '1'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        If Null2String(RSTMP!DONE) = "Y" Then
            CheckIfJobIsFinish = True
        Else
            CheckIfJobIsFinish = False
        End If
    End If
    Set RSTMP = Nothing
End Function

Private Sub chkWithColors_Click()
    Call SaveSetting("DMIS 2.0", "CSMS", "SERVICE COUNTER WITHCOLORS", chkWithColors.Value)
End Sub

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
    Dim RSTMP                                          As New ADODB.Recordset
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
                Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (RSTMP.BOF And RSTMP.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set RSTMP = Nothing

                'SQL_STATEMENT = "UPDATE HRMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
                'gconDMIS.Execute SQL_STATEMENT
                'NEW LOG AUDIT-----------------------------------------------------
                '    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "HRMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
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
                Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & vRONO & "' AND TECHCODE = '" & Trim(VTECHCODE) & "' AND DONE = 'W'")
                If Not (RSTMP.BOF And RSTMP.EOF) Then

                Else
                    SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'S',ASSIGNEDRO = '" & vRONO & "' WHERE EMPNO = '" & vEMPNO & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW LOG AUDIT-----------------------------------------------------
                    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                    'NEW LOG AUDIT-----------------------------------------------------
                End If
                Set RSTMP = Nothing
                'SQL_STATEMENT = "UPDATE CSMS_EMPINFO SET JSTATUS = 'A',ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'"
                'gconDMIS.Execute SQL_STATEMENT
                'NEW LOG AUDIT-----------------------------------------------------
                '    Call NEW_LogAudit("RE", "EMPLOYEE INFO", SQL_STATEMENT, FindTransactionID(N2Str2Null(vEMPNO), "EMPNO", "CSMS_EMPINFO"), "", "RO NO: " & vRONO, "", "")
                'NEW LOG AUDIT-----------------------------------------------------
            Else
                gconDMIS.Execute ("UPDATE CSMS_EMPINFO SET JSTATUS = 'A', ASSIGNEDRO = NULL WHERE EMPNO = '" & vEMPNO & "'")
            End If
        End If
    End If

    Set rsDet = Nothing
End Sub

Function Click_ScheduleGrid()
    grdCounter_RowColChange grdCounter.ActiveCell.Row, grdCounter.ActiveCell.Col
End Function

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

Private Sub cmdEditRO_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to Edit", vbInformation, "CSMS"
        Exit Sub
    End If
    Load frmCSMSEditRO
    frmCSMSEditRO.txtRep_Or = grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text
    frmCSMSEditRO.lblOLDRO.Caption = grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text
    frmCSMSEditRO.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPartsInquiry_Click()
    'frmCSMS_PartsInquiry.Show
    'frmCSMS_PartsInquiry.ZOrder 0
End Sub

Private Sub cmdRefresh_Click()
    txtSearchName.Text = ""
    Screen.MousePointer = 11
    
    Call CleanListViewDetails
    theRo = ""
    CHKSTATUS = "All"
    cmdDateBack.Caption = " " & Format(MonthView1, "MMM") & " " & Format(MonthView1 - 1, "dd")
    cmdDateForward.Caption = " " & Format(MonthView1, "MMM") & " " & Format(MonthView1 + 1, "dd")
    frmSplash.Show
    frmSplash.labCon.Caption = "Updating Job Status... Please wait..."
    DoEvents
    ProcesUpdate
    frmSplash.Show
    frmSplash.labCon.Caption = "Refreshing View for Active RO... Please wait..."
    DoEvents
    ViewActiveRO
    frmSplash.Show: frmSplash.labCon.Caption = "Finalizing View... ": DoEvents
    ComputeMeCTR
    Unload frmSplash
    Screen.MousePointer = 0
End Sub

Private Sub cmdToday_Click()
    MonthView1.Value = Format(Now, "MM/dd/yyyy")
    Check2.Value = 1
    cmdRefresh.Value = True
End Sub

Private Sub cmdViewRODetails_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to View", vbInformation, "CSMS"
        Exit Sub
    End If
    frmCSMSViewRO.labRO.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    frmCSMSViewRO.Show 1
End Sub

Private Sub cmdWriteAppointment_Click()
    If Module_Access(LOGID, "APPOINTMENT", "TRANSACTION") = False Then Exit Sub

    FROM_APPOINTMENT = "MAIN"
    frmCSMSAppointment.Show 1
    Call ViewActiveRO
End Sub

Private Sub cmdWriteEstimate_Click()
    PopupMenu mnuEstimate
End Sub

Private Sub cmdWriteEstimate_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEstimate
    End If
End Sub

Private Sub cmdWriteRO_Click()
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

Private Sub Command11_Click()
    txtSearchName.Text = ""
    ViewActiveRO
End Sub

Private Sub Command12_Click()
    ViewActiveRO
End Sub

Private Sub Command5_Click()
    frmCSMSReqJobs.Show 1
End Sub

Private Sub Command7_Click()
    frmCSMSPMS.Show 1
End Sub

Sub ComputeMeCTR()
    tlHrs = 0: tlFR = 0
    For bevvy = 1 To grdCounter.Rows - 1
        tlHrs = tlHrs + NumericVal(grdCounter.Cell(bevvy, 7).Text)
    Next bevvy
    txtTLhrs = Format(Trim(STR(tlHrs)), MAXIMUM_DIGIT) & " " & "Scheduled Hrs"
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
    mnuBACKJOB.Enabled = Not COND
End Sub

Function FindTechName(SACODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Dim RSCON                                          As New ADODB.Recordset
    Dim rsVEN                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT TECH_NAME FROM CSMS_VW_TECHNICIAN WHERE TECHNICIAN = '" & SACODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindTechName = Null2String(RSTMP!TECH_NAME)
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

    Set RSTMP = Nothing
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
        
    If KeyCode = vbKeyF5 Then cmdRefresh_Click
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11

    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set JOBCLOCKFORM = New frmCSMSClockINOUT
    chkWithColors.Value = GetSetting("DMIS 2.0", "CSMS", "SERVICE COUNTER WITHCOLORS", 1)
    MonthView1.Value = Format(Now, "MM/dd/yyyy")
    InitGrid
    Check1.Value = 1
    CHKSTATUS = "All":
    CenterMe frmMain, Me, 0
    Screen.MousePointer = 0
    cboSearch.ListIndex = 0
'
'    Dim rsCico                                         As ADODB.Recordset
'    Set rsCico = gconDMIS.Execute("SELECT MAX(CLOCKIN) KEYIN, MAX(CLOCKOUT) KEYCOUT FROM CSMS_JOBCLOCK")
'    If Not rsCico.EOF Or rsCico.BOF Then
'        LABTIMEIN = rsCico!KEYIN
'        LABTIMEOUT = rsCico!KEYCOUT
'        tmr_CICO.Enabled = True
'    End If

    mnuOption.Visible = False
    mnuOption4.Visible = False
    mnuEstimate.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labBilled.FontBold = False
    labBilled.ForeColor = &H0&
    labFinish.FontBold = False
    labFinish.ForeColor = &H0&
    labOver.FontBold = False
    labOver.ForeColor = &H0&
    labWork.FontBold = False
    labWork.ForeColor = &H0&
    labPark.FontBold = False
    labPark.ForeColor = &H0&
    Label1.FontBold = False
    Label1.ForeColor = &H0&
    labIdleTime.FontBold = False
    labIdleTime.ForeColor = &H0&

    Shape1.BorderColor = &H0&
    Shape2.BorderColor = &H0&
    Shape4.BorderColor = &H0&
    Shape5.BorderColor = &H0&
    Shape6.BorderColor = &H0&
    Shape3.BorderColor = &H0&

    labBackJob.FontBold = False
    labBackJob.ForeColor = &H0&
    Shape7.BorderColor = &H0&
End Sub



Private Sub grdCounter_Click()
    thestatus = Trim(grdCounter.Cell(grdCounter.ActiveCell.Row, 10).Text)
    theRo = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    zRONO = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    ViewJobs
    ComputeMePo
    SSTab1.Tab = 0

End Sub

Private Sub grdCounter_DblClick()
    cmdViewRODetails_Click
End Sub

'Private Sub grdCounter_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
'    '    'UPDATED BY: JUN
'    '    'DATE UPDATED: 01-06-2008
'    '    'DESCRIPTION: GET THE RO INFO WHEN KEY DOWN IS PRESS
'    '    thestatus = Trim(grdCounter.Cell(grdCounter.ActiveCell.Row, 10).Text)
'    '    theRo = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
'    '    zRONO = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
'
'    '    ViewJobs
'    '    ComputeMePo
'    '    SSTab1.Tab = 0
'End Sub
'
'Private Sub grdCounter_KeyUp(KeyCode As Integer, Shift As Integer)
'    'UPDATED BY: JUN
'    'DATE UPDATED: 01-06-2008
'    'DESCRIPTION: GET THE RO INFO WHEN KEY UP IS PRESS
'    '    thestatus = Trim(grdCounter.Cell(grdCounter.ActiveCell.Row, 10).Text)
'    '    theRo = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
'    '    zRONO = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
'    '
'    '    ViewJobs
'    '    ComputeMePo
'    '
'    '    SSTab1.Tab = 0
'End Sub

Private Sub grdCounter_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If grdCounter.ActiveCell.Row = 0 Then Exit Sub
        Dim test                                       As String
        test = "Billed"
        If StrComp(Trim(thestatus), test) = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then
            If QC_MODULE_ON = "ON" Then
                Call EnableDropDownMenu(False)
                PopupMenu mnuOption
            End If
        Else
            Call EnableDropDownMenu(True)
            If Trim(grdCounter.Cell(grdCounter.ActiveCell.Row, 9).Text) = "Finish Job" Then
                mnuBilledRO.Visible = True
            Else
                mnuBilledRO.Visible = False
            End If
            PopupMenu mnuOption
        End If
    End If
End Sub

Private Sub grdCounter_RowColChange(ByVal Row As Long, ByVal Col As Long)
    thestatus = Trim(grdCounter.Cell(grdCounter.ActiveCell.Row, 10).Text)
    theRo = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    zRONO = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    grdCounter.Range(Row, 0, Row, 15).Selected
    ViewJobs
    ComputeMePo
End Sub

Sub InitGrid()
    With grdCounter
        .Cols = 18: .Rows = 2
        .DisplayFocusRect = False: .AllowUserResizing = True
        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Date"
        .Cell(0, 2).Text = "Customer"
        .Cell(0, 3).Text = "Vehicle"
        .Cell(0, 4).Text = "Plate No."
        .Cell(0, 5).Text = "CS No."
        .Cell(0, 6).Text = "R/O"
        .Cell(0, 7).Text = "Std.Hrs"
        .Cell(0, 8).Text = "Hrs.Work"
        .Cell(0, 9).Text = "(%)"
        .Cell(0, 10).Text = "Status"
        .Cell(0, 11).Text = "Promise"
        .Cell(0, 12).Text = "TODAY"
        .Cell(0, 13).Text = "Service Adviser"
        .Cell(0, 14).Text = "Date Finish"

        .Cell(0, 15).Text = "Remarks"
        .Cell(0, 16).Text = "Tech-3"
        .Cell(0, 17).Text = "Account No"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:
        .Column(3).CellType = cellTextBox:
        .Column(4).CellType = cellTextBox
        .Column(5).CellType = cellTextBox
        .Column(6).CellType = cellTextBox
        .Column(7).CellType = cellTextBox
        .Column(8).CellType = cellTextBox
        .Column(9).CellType = cellTextBox
        .Column(10).CellType = cellTextBox
        .Column(11).CellType = cellTextBox
        .Column(12).CellType = cellTextBox
        .Column(13).CellType = cellTextBox
        .Column(14).CellType = cellTextBox
        .Column(15).CellType = cellTextBox
        .Column(16).CellType = cellTextBox
        .Column(17).CellType = cellTextBox

        .Column(0).Width = 18
        .Column(1).Width = 60: .Column(1).Locked = True
        .Column(2).Width = 150: .Column(2).Locked = True
        .Column(3).Width = 200: .Column(3).Locked = True
        .Column(4).Width = 60: .Column(4).Locked = True
        .Column(5).Width = 60: .Column(5).Locked = True
        .Column(6).Width = 65: .Column(6).Locked = True
        .Column(7).Width = 50: .Column(7).Locked = True
        .Column(8).Width = 60: .Column(8).Locked = True
        .Column(9).Width = 50: .Column(9).Locked = True
        .Column(10).Width = 60: .Column(10).Locked = True
        .Column(11).Width = 125: .Column(11).Locked = True
        .Column(12).Width = 125: .Column(12).Locked = True
        .Column(13).Width = 100: .Column(13).Locked = True
        .Column(14).Width = 100: .Column(14).Locked = True

        .Column(15).Width = 200: .Column(15).Locked = True
        .Column(16).Width = 0: .Column(16).Locked = True
        .Column(17).Width = 0: .Column(17).Locked = True

        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 17, .Rows - 1, 17).ForeColor = RGB(0, 0, 128)
    End With

    If Not COMPANY_CODE = "HGC" Then
        mnuremovebay.Enabled = False
        mnuAsgnedbay.Enabled = False
    Else
        mnuremovebay.Enabled = True
        mnuAsgnedbay.Enabled = True
    End If
End Sub

Sub IsTechLogIn()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset

    SQL = "SELECT * FROM CSMS_vw_technicianAvailability WHERE Techcode ='" & thetechcode & "' and AssignedRo='" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    Do While Not RS.EOF
        If Null2String(RS!Code) = "W" Then
            ISLOGIN = True
        End If
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub JOBCLOCKFORM_FORMCLOSED()
    cmdRefresh.Value = True
End Sub

Private Sub JOBCLOCKFORM_JOBCLOCKED()
    ProcesUpdate
End Sub

Private Sub labBackJob_Click()
    CHKSTATUS = "Back Job"
    ViewActiveRO
End Sub

Private Sub labBackJob_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labBackJob.FontBold = True
    labBackJob.ForeColor = &HFF0000
    Shape7.BorderColor = &HFFFF&
End Sub

Private Sub labBilled_Click()
    CHKSTATUS = "Billed"
    ViewActiveRO
End Sub

Private Sub labBilled_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labBilled.FontBold = True
    labBilled.ForeColor = &HFF0000
    Shape5.BorderColor = &HFFFF&
End Sub

Private Sub Label1_Click()
    CHKSTATUS = "Released"
    ViewActiveRO
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Label1.FontBold = True
    Label1.ForeColor = &HFF0000
    Shape3.BorderColor = &HFFFF&
End Sub

Private Sub labFinish_Click()
    CHKSTATUS = "Finish job"
    ViewActiveRO
End Sub

Private Sub labFinish_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labFinish.FontBold = True
    labFinish.ForeColor = &HFF0000
    Shape4.BorderColor = &HFFFF&
End Sub

Private Sub labIdleTime_Click()
    CHKSTATUS = "Idle Time"
    ViewActiveRO
End Sub

Private Sub labIdleTime_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labIdleTime.FontBold = True
    labIdleTime.ForeColor = &HFF0000
    Shape.BorderColor = &HFFFF&
End Sub

Private Sub labOver_Click()
    CHKSTATUS = "Over"
    ViewActiveRO
End Sub

Private Sub labOver_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labOver.FontBold = True
    labOver.ForeColor = &HFF0000
    Shape6.BorderColor = &HFFFF&
End Sub

Private Sub labPark_Click()
    CHKSTATUS = "Park"
    ViewActiveRO
End Sub

Private Sub labPark_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labPark.FontBold = True
    labPark.ForeColor = &HFF0000
    Shape1.BorderColor = &HFFFF&
End Sub


Private Sub labWork_Click()
    CHKSTATUS = "Working"
    ViewActiveRO
End Sub

Private Sub labWork_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labWork.FontBold = True
    labWork.ForeColor = &HFF0000
    Shape2.BorderColor = &HFFFF&
End Sub

Private Sub lstCounter_ItemClick(ByVal Item As MSComctlLib.ListItem)
    zRONO = lstCounter.SelectedItem.SubItems(4)
    ViewJobs
    ComputeMePo
End Sub

Private Sub lstCounter_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuOption
    End If
End Sub

Private Sub lstJob4Service_Click()
    'Update by BTT/RSC
    On Error Resume Next

    If lstJob4Service.SelectedItem Is Nothing Then Exit Sub
    THEJOBCODE = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index)
    THEJOBDEST = lstJob4Service.SelectedItem.SubItems(1)
    theflatrate = lstJob4Service.SelectedItem.SubItems(2)
    THESTDRATE = lstJob4Service.SelectedItem.SubItems(3)
    vlineNo = lstJob4Service.SelectedItem.SubItems(8)

    PERJOBSTATUS = lstJob4Service.SelectedItem.SubItems(9)
End Sub

Private Sub lstJob4Service_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lstJob4Service.ListItems.Count = 0 Then Exit Sub

    lstJob4Service.ToolTipText = Null2String(Item.ListSubItems(1))
End Sub

Private Sub lstJob4Service_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If StrComp(Trim(thestatus), "Billed") = 0 Or StrComp(Trim(thestatus), "Released") = 0 Then
            '            'Do Nothing
            If Module_Access(LOGID, "EDIT JOB DESCRIPTIONS", "SYSTEM") = False Then Exit Sub

            mnuAsgnedTech.Enabled = False
            MnuAsscontractor.Enabled = False
            mnuOption2_2.Enabled = False
            'UPDATED BY: JUN-----------------------------------
            'DATE UPDATED: 1-29-2008
            'DESCRIPTION: DISABLED THE JOBDONE IN ORDER IF IT IS ALREADY BILLED OR RELEASE SO THAT IT WILL NOT HAVE AN UPDATE OF RO AMOUNT WHICH IS ALREADY BILLED
             mnuJobDone.Enabled = False
            'UPDATED BY: JUN-----------------------------------
            
            PopupMenu mnuOption4
        Else
            If lstJob4Service.ListItems.Count = 0 Or lstJob4Service.SelectedItem Is Nothing Then: Exit Sub

            If COMPANY_CODE = "HGC" Then
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
    End If
End Sub

Private Sub mnuAsgnedbay_Click()
    With frmCSMSUpdatebayInfo
        .labRO.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
        '.labCust.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 14).Text
        '.labItemNo.Caption = lstJob4Service.SelectedItem
    End With
    frmCSMSUpdatebayInfo.Show 1
End Sub

Private Sub mnuAsgnedTech_Click()
    If grdCounter.ActiveCell.Row = 0 Then
        MsgBox "Please Select RO From the List", vbInformation
        Exit Sub
    End If

    Dim Index                                          As Integer

    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Please Select a Repair Order to be edit.", vbInformation, "CSMS"
        Exit Sub
    End If


    Index = lstJob4Service.SelectedItem.Index
    If Null2String(lstJob4Service.SelectedItem.SubItems(9)) = "Finish Job" Then
        MsgBox "Job Is Already Finish", vbInformation, "CSMS"
        Exit Sub
    End If

    If Null2String(lstJob4Service.SelectedItem.SubItems(4)) = "" Then
        With frmCSMSUpdateCustomerInfo
            .labRO.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
            .lblJobCode.Caption = Null2String(lstJob4Service.SelectedItem.Text)
            .labCust.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 14).Text
            .LABITEMNO.Caption = Null2String(lstJob4Service.SelectedItem)
        End With
        frmCSMSUpdateCustomerInfo.Show 1
    Else
        If CheckIfJobIsFinish(lstJob4Service.ListItems(Index).ListSubItems(8)) = False Then
            If MsgBox("Technician is already assign to this job, do you Want to change", vbQuestion + vbYesNo, "Are you sure") = vbNo Then Exit Sub
            cmdViewRODetails_Click
        Else
            MsgBox "Job Already Finish", vbInformation, "CSMS"
        End If
    End If
End Sub

Private Sub MnuAsscontractor_Click()
    Dim Index                                          As Integer

    Index = lstJob4Service.SelectedItem.Index
    If CheckIfJobIsFinish(lstJob4Service.ListItems(Index).SubItems(8)) = False Then
        With frmCSMSSelectContractor
            .labCust.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 14).Text
            .LABITEMNO.Caption = lstJob4Service.SelectedItem

            .lblRO.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
            .lblCustomer = grdCounter.Cell(grdCounter.ActiveCell.Row, 2).Text
            .lblplate = grdCounter.Cell(grdCounter.ActiveCell.Row, 3).Text
            .lblModel = grdCounter.Cell(grdCounter.ActiveCell.Row, 4).Text

            .lblCONCODE.Caption = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index).ListSubItems(7)
            .lblLineNo.Caption = lstJob4Service.ListItems(lstJob4Service.SelectedItem.Index).ListSubItems(8)
        End With
        frmCSMSSelectContractor.Show 1
        frmCSMSSelectContractor.ZOrder 0
    Else
        MsgBox "Job Already Finish", vbInformation
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
    frmCSMSDataEntry.Show
End Sub

Private Sub mnuCanedJob_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to add Canned job", vbInformation, "CSMS"
        Exit Sub
    End If

    frmMain.MousePointer = 11

    frmCSMSGetCannedLabor.txtCheckMe = "MAIN"
    frmCSMSGetCannedLabor.lblRO.Caption = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
    frmCSMSGetCannedLabor.Show 1

    frmMain.MousePointer = 0
End Sub

Private Sub mnuChangeVehicle_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "You have not select Repair Order", vbInformation, "INFORMATION"
        Exit Sub
    End If
    lblROChange.Caption = theRo
    frmCSMS_ChangeVehicle.lblRONO = theRo
    frmCSMS_ChangeVehicle.Show 1
    lblROChange.Caption = ""
    Call ViewActiveRO
End Sub

Private Sub mnuEstimateAdd_Click()
    frmCSMSNewAppointment.labType(0).Caption = "Estimate"
    frmCSMSNewAppointment.labType(1).Caption = "Estimate"
    frmCSMSNewAppointment.GetDefaultTransactionType
    Timer1.Enabled = False
    frmCSMSNewAppointment.Show 1
    Timer1.Enabled = True
End Sub

Private Sub mnujobdone_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to View", vbInformation, "CSMS"
        Exit Sub
    End If

    With frmCSMS_Jobdone
        Dim Index                                      As Integer
        Index = lstJob4Service.SelectedItem.Index

        .lblJobCode = THEJOBCODE
        .lbljobdesc = THEJOBDEST
        .lblRO = theRo

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
        .txtjobdesc = THEJOBDEST
    End With

    frmCSMS_Jobdone.Show 1, frmCSMSServiceCounter

    'End If
End Sub

Private Sub mnuOption1_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to add General job", vbInformation, "CSMS"
        Exit Sub
    End If

    Dim theAnswer                                      As String
    With frmCSMSReqJobs
        .txtCustomer = grdCounter.Cell(grdCounter.ActiveCell.Row, 2).Text
        .txtActNo = grdCounter.Cell(grdCounter.ActiveCell.Row, 14).Text
        .txtROno = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
        .txtCheckMe = "main"
    End With
    If StrComp(Trim(thestatus), "Finish Job") = 0 Then
        theAnswer = MsgBox("This Job Is Already Finish!,Do You Want To Add New Job?", vbQuestion + vbYesNo, "information")
        If theAnswer = vbYes Then
            frmCSMSReqJobs.Show 1
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    frmCSMSReqJobs.Show 1
End Sub

Private Sub mnuOption13_Click()
    'On Error Resume Next

    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to add PMS job", vbInformation, "CSMS"
        Exit Sub
    End If

    With frmCSMSPMS
        .txtRO.Text = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
        .dtpromise.Value = grdCounter.Cell(grdCounter.ActiveCell.Row, 11).Text
    End With

    If frmCSMSPMS.txtRO.Text = "" Then
        Exit Sub
    Else
        frmCSMSPMS.Show 1
    End If
End Sub

Private Sub mnuOption2_2_Click()
    Dim rsJOBSTATUS                                    As New ADODB.Recordset

    Set rsJOBSTATUS = gconDMIS.Execute("SELECT * FROM CSMS_RO_DET WHERE REP_OR = '" & theRo & "' AND LIVIL = '1' AND LINE_NO = '" & vlineNo & "'")
    If Not (rsJOBSTATUS.BOF And rsJOBSTATUS.EOF) Then
        If Null2String(rsJOBSTATUS!STATUS) = "Y" Then
            PERJOBSTATUS = "Finish Job"
        ElseIf Null2String(rsJOBSTATUS!STATUS) = "I" Then
            PERJOBSTATUS = "Idle Time"
        ElseIf Null2String(rsJOBSTATUS!STATUS) = "B" Then
            PERJOBSTATUS = "Break Time"
        ElseIf Null2String(rsJOBSTATUS!STATUS) = "G" Then
            PERJOBSTATUS = "Going Home"
        ElseIf Null2String(rsJOBSTATUS!STATUS) = "L" Then
            PERJOBSTATUS = "Lunch Break"
        ElseIf Null2String(rsJOBSTATUS!STATUS) = "W" Then
            PERJOBSTATUS = "Working"
        Else
            PERJOBSTATUS = ""
        End If
    End If

    If PERJOBSTATUS = "Finish Job" Then
        MsgBox "Cannot Remove Job. Job Already Finish", vbExclamation, "CSMS"
        Exit Sub
    ElseIf PERJOBSTATUS = "Idle Time" Or PERJOBSTATUS = "Going Home" Or PERJOBSTATUS = "Break Time" Or PERJOBSTATUS = "Lunch Break" Or PERJOBSTATUS = "Working" Then
        MsgBox "Cannot Remove Job. Job Already Started", vbExclamation, "CSMS"
        Exit Sub
    Else

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

    cmdRefresh.Value = True
End Sub

Private Sub mnuOption2_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Please Select a Repair Order to be edit.", vbInformation, "CSMS"
        Exit Sub
    End If

    Load frmCSMSEditRO
    frmCSMSEditRO.txtRep_Or = grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text
    frmCSMSEditRO.lblOLDRO.Caption = grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text

    frmCSMSEditRO.Show 1

    With frmCSMSServiceCounter
        .cmdRefresh = True
    End With
End Sub

Private Sub mnuOtherJobs_Click()
    If theRo = "" Or theRo = "R/O" Or grdCounter.Cell(frmCSMSServiceCounter.grdCounter.ActiveCell.Row, 6).Text = "R/O" Then
        MsgBox "Choose a Repair Order to add Other job", vbInformation, "CSMS"
        Exit Sub
    End If

    Dim theAnswer                                      As String
    With frmCSMSOtherJobs
        .txtCustomer = grdCounter.Cell(grdCounter.ActiveCell.Row, 2).Text
        .txtActNo = grdCounter.Cell(grdCounter.ActiveCell.Row, 14).Text
        .txtROno = grdCounter.Cell(grdCounter.ActiveCell.Row, 6).Text
        .txtCheckMe = "main"
    End With
    If StrComp(Trim(thestatus), "Finish Job") = 0 Then
        theAnswer = MsgBox("This RO Is Already Finish!,Do You Want To Add New Job?", vbQuestion + vbYesNo, "CSMS")
        If theAnswer = vbYes Then
            frmCSMSOtherJobs.Show 1
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    frmCSMSOtherJobs.Show 1
End Sub

Private Sub mnuremovebay_Click()
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim ans                                            As String

    SQL = "SELECT Ro from CSMS_baymonitoring where RO = '" & theRo & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        ans = MsgBox("Are you sure do you want to remove this Repair Order to Bay", vbQuestion + vbYesNo)
        If ans = vbYes Then
            gconDMIS.Execute "Update CSMS_Baymonitoring set ro=null,bay_status='Available' where ro='" & theRo & "'"
            ShowSuccessFullyUpdated
        End If
    Else
        MsgBox "Repair Order not in the bay.", vbInformation, "Infomartion"
    End If
End Sub

Private Sub mnuUpdateEstimate_Click()
    frmCSMSLoadEstimateToRO.Show 1
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
    gconDMIS.Execute ("UPDATE CSMS_JOBCLOCKOPENRO SET HRSWORKED=ROUND(DATEDIFF(MINUTE,CASE WHEN ISDATE(CLOCKIN)=0 THEN GETDATE() ELSE CLOCKIN END,CASE WHEN ISDATE(CLOCKOUT)=0 THEN GETDATE() ELSE CLOCKOUT END)/60.00,2)")
    gconDMIS.Execute ("UPDATE CSMS_REPAIRORDER SET TODAY=GETDATE(), XHRSWORK =T.TOTAL_HRS" & vbCrLf _
                    & " FROM  (SELECT RO_NO, SUM(ISNULL(HRS,0)) AS TOTAL_HRS  FROM CSMS_VW_JOBCLOCKTORO GROUP BY RO_NO) T" & vbCrLf _
                    & " INNER JOIN  CSMS_REPAIRORDER  ON  CSMS_REPAIRORDER.RO_NO=T.RO_NO ")
    Screen.MousePointer = 0
    '''''''''PROCESS UPDATE REVISED




    Exit Sub

    Dim rsProces                                       As ADODB.Recordset
    Set rsProces = gconDMIS.Execute("Select ID,ClockIn,ClockOut,HrsWorked from CSMS_JobClockOpenRO")

    If Not rsProces.EOF And Not rsProces.BOF Then
        Dim xTimeInAM, xTimeOutAM                      As String
        Dim xMam                                       As Double
        Dim hrAM                                       As Double
        Dim tlHrs                                      As Double

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

Sub roSearch()
    'UPDATED BY: JUN CEDRON
    'DATE UPDATED: 12-03-2008
    'DESCRIPTION: THIS IS FOR REPAIR ORDER SEARCHING USER ALLOW ONLY TO INPUT THE
    Dim k                                              As Integer
    Dim xRepairOrder2 As String, xRepairOrder3         As String

    XREPAIRORDER = UCase(txtSearchName)
    If XREPAIRORDER <> "" Then
        If IsNumeric(XREPAIRORDER) = True Then
            XREPAIRORDER = Format(Left(XREPAIRORDER, 1), "R-") & Format(Right(XREPAIRORDER, 6), "00000000")
        Else
            For k = 1 To Len(XREPAIRORDER)
                xRepairOrder2 = Mid(XREPAIRORDER, k, 1)
                If IsNumeric(xRepairOrder2) = True Then xRepairOrder3 = xRepairOrder3 + xRepairOrder2
            Next
            xRepairOrder3 = Format(xRepairOrder3, "00000000"): XREPAIRORDER = Format(Left(xRepairOrder3, 1), "R-") & Format(Right(xRepairOrder3, 6), "00000000")
        End If
    End If
End Sub

Function SetCusVehCondNo(XXX As String, yyy As String) As String
    Dim rsCusVeh                                       As ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select VCOND_NO from CSMS_CUSVEH where CUSCDE = '" & yyy & "' AND PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        SetCusVehCondNo = Null2String(rsCusVeh!VCOND_NO)
    End If
    Set rsCusVeh = Nothing
End Function

Function SetCusVehDesc(XXX As String, yyy As String) As String
    Dim rsCusVeh                                       As ADODB.Recordset
    Set rsCusVeh = gconDMIS.Execute("Select Description from CSMS_CUSVEH where CUSCDE = '" & yyy & "' AND PLATE_NO = '" & XXX & "'")
    If Not rsCusVeh.EOF And Not rsCusVeh.BOF Then
        SetCusVehDesc = Null2String(rsCusVeh!Description)
    End If
    Set rsCusVeh = Nothing
End Function

Function SetCusVehDetail(XXX As String, yyy As String) As String()
    Dim rsCusVeh                                       As ADODB.Recordset
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
'
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

Private Sub txtSearchName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command12.Value = True
    End If
End Sub

Sub ViewActiveRO()
    Dim lng                                            As Long
    Screen.MousePointer = 11
    Dim RSUPLOAD                                       As ADODB.Recordset

    CleanListViewDetails

    lstCounter.Sorted = False
    lstCounter.ListItems.Clear
    'InitGrid: DoEvents
    Dim xx                                             As Integer
    grdCounter.Rows = 1
    xx = 0
    grdCounter.AutoRedraw = False


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


    If Check1.Value = 1 Then
        If CHKSTATUS = "All" Then
            If cboSearch.Text = "Search By Customer Name" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND customer like '%" & txtSearchName.Text & "%' order by AppointmentDate asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND AppointmentDate <= '" & DateValue(MonthView1) & "' order by AppointmentDate asc")    'BTT - 05242007
                End If
            ElseIf cboSearch.Text = "Search By Plate No" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND plate_no = '" & txtSearchName.Text & "' order by AppointmentDate asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND AppointmentDate <= '" & DateValue(MonthView1) & "' order by AppointmentDate asc")    'BTT - 05242007
                End If
            Else
                Call roSearch

                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND RO_NO = '" & XREPAIRORDER & "' order by AppointmentDate asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where (TransType = 'R' and status <> 'Released') AND AppointmentDate <= '" & DateValue(MonthView1) & "' order by AppointmentDate asc")    'BTT - 05242007
                End If
            End If
        Else
            If cboSearch.Text = "Search By Customer Name" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' AND customer like '%" & txtSearchName.Text & "%' Order by RO_No asc")    'BTT - 05242007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' Order by RO_No asc")    'BTT - 05242007
                End If
            ElseIf cboSearch.Text = "Search By Plate No" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' AND plate_no = '" & txtSearchName.Text & "' Order by RO_No asc")    'BTT - 05242007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' Order by RO_No asc")    'BTT - 05242007
                End If
            Else
                Call roSearch

                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' AND RO_NO = '" & XREPAIRORDER & "' Order by RO_No asc")    'BTT - 05242007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,Model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,datefinish from CSMS_vw_RepairOrder where TransType = 'R' and Status = '" & CHKSTATUS & "' Order by RO_No asc")    'BTT - 05242007
                End If
            End If
        End If
        'End If

        'If Check2.Value = 1 Then
    ElseIf Check2.Value = 1 Then
        If CHKSTATUS = "All" Then
            If cboSearch.Text = "Search By Customer Name" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and (TransType = 'R') AND customer like '%" & txtSearchName.Text & "%' order by RO_No asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and (AppointmentDate = '" & DateValue(MonthView1) & "') order by RO_No asc")
                End If
            ElseIf cboSearch.Text = "Search By Plate No" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and (TransType = 'R') AND plate_no = '" & txtSearchName.Text & "' order by RO_No asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and (AppointmentDate = '" & DateValue(MonthView1) & "') order by RO_No asc")
                End If
            Else
                Call roSearch

                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and (TransType = 'R') AND RO_NO = '" & XREPAIRORDER & "' order by RO_No asc")
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and (AppointmentDate = '" & DateValue(MonthView1) & "') order by RO_No asc")
                End If
            End If
        Else
            If cboSearch.Text = "Search By Customer Name" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' and customer like '%" & txtSearchName.Text & "%' order by RO_No asc")    'MJP - 07172007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' order by RO_No asc")    'MJP - 07172007
                End If
            ElseIf cboSearch.Text = "Search By Plate No" Then
                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' and plate_no = '" & txtSearchName.Text & "' order by RO_No asc")    'MJP - 07172007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' order by RO_No asc")    'MJP - 07172007
                End If
            Else
                Call roSearch

                If Trim(txtSearchName.Text) <> "" Then
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' and RO_NO = '" & XREPAIRORDER & "' order by RO_No asc")    'MJP - 07172007
                Else
                    Set RSUPLOAD = gconDMIS.Execute("Select [AppointmentDate],Customer,model,PLATE_NO,RO_No,[Hours],[xHrsWork],[Percentage],PromiseDate,[Today],status,Writer,TECH1,TECH2,TECH3,ACCT_NO,Datefinish from CSMS_vw_RepairOrder where TransType = 'R' and AppointmentDate = '" & DateValue(MonthView1) & "' and Status = '" & CHKSTATUS & "' order by RO_No asc")    'MJP - 07172007
                End If
            End If
        End If
    End If





    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Dim pecnt                                      As Double
        'Dim veh_desc                                   As String
        'Dim veh_Cond_No                                As String
        Dim VEH_DET                                    As Variant
        Dim Grid_Text10                                As String
        Do While Not RSUPLOAD.EOF
            If RSUPLOAD![xHrsWork] > 0 And RSUPLOAD![HOURS] > 0 Then
                pecnt = Format(Round(((RSUPLOAD![xHrsWork] / RSUPLOAD![HOURS]) * 100), 2), "0000.#0")
            Else
                pecnt = 0
            End If

            xx = xx + 1
            'veh_desc = ""
            'veh_Cond_No = ""

            'veh_desc = SetCusVehDesc(Null2String(rsUpload![PLATE_NO]), Null2String(rsUpload![ACCT_NO]))
            'veh_Cond_No = SetCusVehCondNo(Null2String(rsUpload![PLATE_NO]), Null2String(rsUpload![ACCT_NO]))

            VEH_DET = SetCusVehDetail(Null2String(RSUPLOAD![PLATE_NO]), Null2String(RSUPLOAD![ACCT_NO]))

            If VEH_DET(0) = "" Then VEH_DET(0) = Null2String(RSUPLOAD![Model])



            '   If veh_desc = "" Then veh_desc = Null2String(rsUpload![MODEL])
            '  grdCounter.AddItem Format(rsUpload![AppointmentDate], "MM/dd/yyyy") & vbTab & _
               '                               rsUpload![Customer] & vbTab & _
               '                               veh_desc & vbTab & _
               '                               rsUpload![PLATE_NO] & vbTab & _
               '                               veh_Cond_No & vbTab & _
               '                               rsUpload![RO_NO] & vbTab & _
               '                               rsUpload![Hours] & vbTab & _
               '                               rsUpload![xHrsWork] & vbTab & _
               '                               pecnt & vbTab & _
               '                               rsUpload![Status] & vbTab & _
               '                               rsUpload![promisedate] & vbTab & _
               '                               rsUpload![Today] & vbTab & _
               '                               rsUpload![writer] & vbTab & _
               '                               rsUpload![datefinish] & vbTab & _
               '                               Replace(Null2String(rsUpload![tech2]), vbCrLf, "") & vbTab & _
               '                               rsUpload![tech3] & vbTab & _
               '                               rsUpload![ACCT_NO], False

            grdCounter.AddItem Format(RSUPLOAD![AppointmentDate], "MM/DD/YYYY") & vbTab & _
                               RSUPLOAD![Customer] & vbTab & _
                               VEH_DET(0) & vbTab & _
                               RSUPLOAD![PLATE_NO] & vbTab & _
                               VEH_DET(1) & vbTab & _
                               RSUPLOAD![RO_NO] & vbTab & _
                               RSUPLOAD![HOURS] & vbTab & _
                               RSUPLOAD![xHrsWork] & vbTab & _
                               pecnt & vbTab & _
                               RSUPLOAD![STATUS] & vbTab & _
                               RSUPLOAD![PromiseDate] & vbTab & _
                               RSUPLOAD![Today] & vbTab & _
                               RSUPLOAD![writer] & vbTab & _
                               Format(RSUPLOAD![datefinish], "MM/DD/YYYY") & vbTab & _
                               Replace(Null2String(RSUPLOAD![tech2]), vbCrLf, "") & vbTab & _
                               RSUPLOAD![tech3] & vbTab & _
                               RSUPLOAD![ACCT_NO], False



            Grid_Text10 = Trim(grdCounter.Cell(xx, 10).Text)

            If chkWithColors.Value = 1 Then
                If Grid_Text10 <> "" Then
                    If Grid_Text10 = "Park" Then
                        grdCounter.Range(xx, 1, xx, 13).BackColor = &HFBFFFF
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &H0&
                    ElseIf Grid_Text10 = "Working" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &HC0C000
                        If CDate(grdCounter.Cell(xx, 11).Text) < Now Then
                            grdCounter.Range(xx, 1, xx, 13).ForeColor = vbRed
                        End If
                    ElseIf Grid_Text10 = "Over" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = vbRed
                    ElseIf Grid_Text10 = "Finish Job" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = vbBlue
                    ElseIf Grid_Text10 = "Billed" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &H800080
                    ElseIf Grid_Text10 = "Released" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &H8000&
                    ElseIf Grid_Text10 = "Back Job" Then
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &H80FF&
                    Else
                        grdCounter.Range(xx, 1, xx, 13).FontBold = True
                        grdCounter.Range(xx, 1, xx, 13).ForeColor = &H808080
                    End If
                End If
            End If
            RSUPLOAD.MoveNext
        Loop

        grdCounter.AutoRedraw = True
        grdCounter.Refresh
    End If
    Screen.MousePointer = 0
    Exit Sub

ERRORCODE:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Sub ViewJobs()
    Dim RSUPLOAD                                       As ADODB.Recordset
    Dim Item                                           As ListItem

    CleanListViewDetails
    'JOBS

    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETAIL,FLATRATE,det_hrs,TECHNICIAN,HRSWRK,REP_OR,TechCode,LINE_NO,status from CSMS_Ro_Det where LIVIL='1' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstJob4Service.ListItems.Add(, , Null2String(RSUPLOAD!DETCDE))

            Item.SubItems(1) = Replace(Null2String(RSUPLOAD!Detail), vbCrLf, " ")
            Item.SubItems(2) = Format(NumericVal(RSUPLOAD!FLATRATE), MAXIMUM_DIGIT)
            Item.SubItems(3) = Null2String(RSUPLOAD!DET_HRS)
            Item.SubItems(4) = FindTechName(LTrim(RTrim(Null2String(RSUPLOAD!TechCode))))
            'ITEM.SubItems(4) = Null2String(rsUpload!Technician)
            Item.SubItems(5) = Null2String(RSUPLOAD!HRSWRK)
            Item.SubItems(6) = Null2String(RSUPLOAD!REP_OR)
            Item.SubItems(7) = Null2String(RSUPLOAD!TechCode)
            Item.SubItems(8) = Null2String(RSUPLOAD!LINE_NO)

            If Null2String(RSUPLOAD!STATUS) = "W" Then Item.SubItems(9) = "Working": Call CHECK_IN_OUT(RSUPLOAD!REP_OR, RSUPLOAD!DETCDE, RSUPLOAD!TechCode, RSUPLOAD!STATUS) 'UPDATED BY JUN: 03-23-2009
            If Null2String(RSUPLOAD!STATUS) = "I" Then Item.SubItems(9) = "Idle Time"
            If Null2String(RSUPLOAD!STATUS) = "L" Then Item.SubItems(9) = "Lunch Break"
            If Null2String(RSUPLOAD!STATUS) = "G" Then Item.SubItems(9) = "Going Home"
            If Null2String(RSUPLOAD!STATUS) = "B" Then Item.SubItems(9) = "Break Time"
            If Null2String(RSUPLOAD!STATUS) = "Y" Or Null2String(RSUPLOAD!STATUS) = "R" Then Item.SubItems(9) = "Finish Job"

            If Null2String(RSUPLOAD!STATUS) = "J" Then Item.SubItems(9) = "Back Job"
            If Null2String(RSUPLOAD!STATUS) = "Q" Then Item.SubItems(9) = "Waiting for QC"
            RSUPLOAD.MoveNext
        Loop
    End If

    'PMS JOBS
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,PMS_MODEL from CSMS_PMS_Job_det where REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstPMSJobs.ListItems, RSUPLOAD
    End If
    Set RSUPLOAD = Nothing

    'PARTS
    'Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='2' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
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
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='3' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
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
    Set RSUPLOAD = gconDMIS.Execute("Select DETCDE,DETDSC,detprc,DetVol,DetPRC,Det_AMT from CSMS_Ro_Det where LIVIL='4' AND REP_OR = '" & zRONO & "' Order by [LINE_NO] Asc")
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
                                                " STATUS = 'Y', Done = 'Y' where LIVIL = '1' AND RO_NO = '" & xRO_NO & _
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
