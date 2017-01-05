VERSION 5.00
Object = "{1DCD8E4D-ACCD-407E-9486-0B6E4A62CFED}#1.0#0"; "wizXPForm.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSAddEstimate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estimate Data Entry"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
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
   Icon            =   "FrmAddEstimate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   12735
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picqty 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4170
      ScaleHeight     =   1095
      ScaleWidth      =   2055
      TabIndex        =   34
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdClosePicQTY 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   60
         Width           =   465
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   660
         TabIndex        =   35
         Text            =   "1"
         Top             =   510
         Width           =   855
      End
      Begin wizXPForm.frm frm1 
         Height          =   1485
         Left            =   0
         Top             =   -390
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2619
         Caption         =   "Enter Qty."
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty."
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
         Left            =   210
         TabIndex        =   36
         Top             =   600
         Width           =   465
      End
   End
   Begin VB.TextBox txtEstNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   120
      Width           =   1515
   End
   Begin VB.TextBox txtPlateNo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11130
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   930
      Width           =   1515
   End
   Begin VB.TextBox txtVehicle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   930
      Width           =   3555
   End
   Begin VB.TextBox txtCustomer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   510
      Width           =   6345
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estimate Category"
      Height          =   735
      Left            =   90
      TabIndex        =   7
      Top             =   60
      Width           =   4845
      Begin VB.OptionButton Option1 
         Caption         =   "Materials"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1980
         TabIndex        =   41
         Top             =   270
         Width           =   1245
      End
      Begin VB.OptionButton CatAll 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton CatAcesories 
         Caption         =   "Accessories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3300
         TabIndex        =   8
         Top             =   270
         Width           =   1425
      End
      Begin VB.OptionButton CatParts 
         Caption         =   "Parts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   930
         TabIndex        =   9
         Top             =   270
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Details"
      Height          =   6180
      Left            =   5040
      TabIndex        =   11
      Top             =   1350
      Width           =   7665
      Begin TabDlg.SSTab SSTab1 
         Height          =   3735
         Left            =   90
         TabIndex        =   31
         Top             =   1500
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   6588
         _Version        =   393216
         TabHeight       =   529
         TabCaption(0)   =   "PARTS"
         TabPicture(0)   =   "FrmAddEstimate.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label11(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "ListView1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "MATERIALS"
         TabPicture(1)   =   "FrmAddEstimate.frx":05A6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ListView2"
         Tab(1).Control(1)=   "Label12"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "ACCESSORIES"
         TabPicture(2)   =   "FrmAddEstimate.frx":05C2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListView3"
         Tab(2).Control(1)=   "Label11(1)"
         Tab(2).ControlCount=   2
         Begin MSComctlLib.ListView ListView1 
            Height          =   2955
            Left            =   60
            TabIndex        =   32
            Top             =   405
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   5212
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
            MouseIcon       =   "FrmAddEstimate.frx":05DE
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Parts No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Parts Description"
               Object.Width           =   7056
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
         Begin MSComctlLib.ListView ListView2 
            Height          =   2955
            Left            =   -74940
            TabIndex        =   43
            Top             =   405
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   5212
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
            MouseIcon       =   "FrmAddEstimate.frx":0740
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Parts No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Parts Description"
               Object.Width           =   7056
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
         Begin MSComctlLib.ListView ListView3 
            Height          =   2955
            Left            =   -74940
            TabIndex        =   45
            Top             =   405
            Width           =   7305
            _ExtentX        =   12885
            _ExtentY        =   5212
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
            MouseIcon       =   "FrmAddEstimate.frx":08A2
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Type"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Parts No"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Parts Description"
               Object.Width           =   7056
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "* DOUBLE CLICK TO REMOVE ITEM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   270
            Index           =   1
            Left            =   -74910
            TabIndex        =   46
            Top             =   3420
            Width           =   2760
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "* DOUBLE CLICK TO REMOVE ITEM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   -74910
            TabIndex        =   44
            Top             =   3420
            Width           =   2760
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "* DOUBLE CLICK TO REMOVE ITEM"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   3420
            Width           =   2760
         End
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
         Left            =   6780
         MouseIcon       =   "FrmAddEstimate.frx":0A04
         MousePointer    =   99  'Custom
         Picture         =   "FrmAddEstimate.frx":0B56
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Cancel"
         Top             =   5325
         Width           =   705
      End
      Begin VB.TextBox txtIsHariParts 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   990
         Width           =   495
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   990
         Width           =   1845
      End
      Begin VB.TextBox txtsrp 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   1395
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   570
         Width           =   5955
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1875
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
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
         Left            =   6090
         MouseIcon       =   "FrmAddEstimate.frx":0E94
         MousePointer    =   99  'Custom
         Picture         =   "FrmAddEstimate.frx":0FE6
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Save"
         Top             =   5325
         Width           =   705
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   300
         TabIndex        =   39
         Top             =   5610
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label labType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3780
         TabIndex        =   33
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "HARI Parts ?"
         Height          =   225
         Left            =   5940
         TabIndex        =   20
         Top             =   1080
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         Height          =   225
         Left            =   3180
         TabIndex        =   17
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SRP"
         Height          =   225
         Left            =   1140
         TabIndex        =   16
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Part Description"
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   660
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Part No. "
         Height          =   225
         Left            =   840
         TabIndex        =   12
         Top             =   270
         Width           =   705
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estimate Search Materials"
      Height          =   6690
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   4875
      Begin VB.CommandButton cmdSelect 
         Height          =   345
         Left            =   4350
         Picture         =   "FrmAddEstimate.frx":1281
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.OptionButton optDescription 
         Caption         =   "Description "
         Height          =   225
         Left            =   1320
         TabIndex        =   4
         Top             =   810
         Width           =   2055
      End
      Begin VB.ComboBox cboModel 
         Height          =   345
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   3615
      End
      Begin VB.TextBox txtKeyword 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   60
         TabIndex        =   1
         Top             =   1140
         Width           =   4755
      End
      Begin MSComctlLib.ListView lstParts 
         Height          =   4980
         Left            =   60
         TabIndex        =   6
         Top             =   1590
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   8784
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
         MouseIcon       =   "FrmAddEstimate.frx":180B
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parts No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Parts Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SRP"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "On-hand"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "None/Hari Parts"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkModel 
         Caption         =   "Model :"
         Height          =   285
         Left            =   210
         TabIndex        =   3
         Top             =   330
         Width           =   1005
      End
      Begin VB.OptionButton optPartNo 
         Caption         =   "Parts No."
         Height          =   225
         Left            =   60
         TabIndex        =   5
         Top             =   810
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Estimate No."
      Height          =   225
      Left            =   5100
      TabIndex        =   29
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Plate No."
      Height          =   225
      Left            =   10290
      TabIndex        =   27
      Top             =   990
      Width           =   750
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Vehicle "
      Height          =   225
      Left            =   5520
      TabIndex        =   25
      Top             =   1020
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Customer "
      Height          =   225
      Left            =   5295
      TabIndex        =   23
      Top             =   630
      Width           =   870
   End
End
Attribute VB_Name = "frmCSMSAddEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUPLOAD                                           As ADODB.Recordset

Function CgeckIfPartsAlreadyOnTheList(VPARTNO As String, CompareMe As ListView) As Boolean
    Dim X                                              As Integer

    For X = 1 To CompareMe.ListItems.Count
        If VPARTNO = CompareMe.ListItems(X).SubItems(1) Then
            CgeckIfPartsAlreadyOnTheList = True
        Else
            CgeckIfPartsAlreadyOnTheList = False
        End If
    Next
End Function

Public Function LimitChar(ByVal alpha As String, ByVal k As Integer)
    If InStr(alpha, Chr(k)) > 0 Or k = 8 Then
        LimitChar = k
    Else
        LimitChar = 0
    End If
End Function

Private Sub CatAcesories_Click()
    If CatAcesories.Value = True Then
        txtkeyword.SetFocus
    End If
End Sub

Private Sub CatAll_Click()
    If CatAll.Value = True Then
        txtkeyword.SetFocus
    End If
End Sub

Private Sub CatParts_Click()
    If CatParts.Value = True Then
        txtkeyword.SetFocus
    End If
End Sub

Private Sub cmdClosePicQTY_Click()
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True

    picqty.Visible = False
End Sub

Private Sub cmdOK_Click()
    Dim X                                              As Long

    If Not ListView1.ListItems.Count = 0 Then
        For X = 1 To ListView1.ListItems.Count
            With frmCSMSNewAppointment.ListView1
                .Sorted = False
                .ListItems.Add , , ListView1.ListItems(X)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , ListView1.ListItems(X).SubItems(1)
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , ListView1.ListItems(X).SubItems(2)
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , ListView1.ListItems(X).SubItems(3)
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , Format(ListView1.ListItems(X).SubItems(4), MAXIMUM_DIGIT)

            End With
        Next X
    End If

    If Not ListView2.ListItems.Count = 0 Then
        For X = 1 To ListView2.ListItems.Count
            With frmCSMSNewAppointment.ListView1
                .Sorted = False
                .ListItems.Add , , ListView2.ListItems(X)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , ListView2.ListItems(X).SubItems(1)
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , ListView2.ListItems(X).SubItems(2)
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , ListView2.ListItems(X).SubItems(3)
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , Format(ListView2.ListItems(X).SubItems(4), MAXIMUM_DIGIT)

            End With
        Next X
    End If

    If Not ListView3.ListItems.Count = 0 Then
        For X = 1 To ListView3.ListItems.Count
            With frmCSMSNewAppointment.ListView1
                .Sorted = False
                .ListItems.Add , , ListView3.ListItems(X)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , ListView3.ListItems(X).SubItems(1)
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , ListView3.ListItems(X).SubItems(2)
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , ListView3.ListItems(X).SubItems(3)
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , Format(ListView3.ListItems(X).SubItems(4), MAXIMUM_DIGIT)

            End With
        Next X
    End If

    cmdCancel.Value = True
End Sub

Private Sub cmdSelect_Click()
    Dim EXIST                                          As Boolean
    Dim Index                                          As Integer
    Dim VTYPE                                          As String

    If lstParts.ListItems.Count = 0 Then Exit Sub

    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame3.Enabled = False


    Index = lstParts.SelectedItem.Index
    VTYPE = lstParts.ListItems(Index).SubItems(7)

    If CgeckIfPartsAlreadyOnTheList(lstParts.ListItems(Index).Text, frmCSMSNewAppointment.ListView1) = True Then    'FROM NEW APPOINTMENTS
        If VTYPE = "A" Then
            MsgBox "Accessories already in the list", vbInformation, "CSMS"
        ElseIf VTYPE = "P" Then
            MsgBox "Part already in the list", vbInformation, "CSMS"
        Else
            MsgBox "Material already in the list", vbInformation, "CSMS"
        End If
        Frame1.Enabled = True: Frame2.Enabled = True: Frame3.Enabled = True
        Exit Sub
    Else
        If VTYPE = "P" Then
            If CgeckIfPartsAlreadyOnTheList(lstParts.ListItems(Index).Text, ListView1) = True Then
                MsgBox "Part already in the list", vbInformation, "CSMS"
                Frame1.Enabled = True: Frame2.Enabled = True: Frame3.Enabled = True
                Exit Sub
            End If
        ElseIf VTYPE = "M" Then
            If CgeckIfPartsAlreadyOnTheList(lstParts.ListItems(Index).Text, ListView1) = True Then
                MsgBox "Materials already in the list", vbInformation, "CSMS"
                Frame1.Enabled = True: Frame2.Enabled = True: Frame3.Enabled = True
                Exit Sub
            End If
        Else
            If CgeckIfPartsAlreadyOnTheList(lstParts.ListItems(Index).Text, ListView3) = True Then
                MsgBox "Accessories already in the list", vbInformation, "CSMS"
                Frame1.Enabled = True: Frame2.Enabled = True: Frame3.Enabled = True
                Exit Sub
            End If
        End If
    End If

    'If Not EXIST Then
    picqty.Visible = True
    On Error Resume Next
    txtQty.SetFocus
    SendKeys "{HOME}+{END}"
    'Else
    '    MsgBox "Parts Already on the List", vbInformation, "Add Parts"
    '    lstParts.SetFocus
    'End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Set RSUPLOAD = New ADODB.Recordset
    Set RSUPLOAD = gconDMIS.Execute("select MODELCODE from PMIS_PartMas where MODELCODE is not null GROUP BY MODELCODE order by MODELCODE asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        cboModel.Clear
        Do Until RSUPLOAD.EOF
            cboModel.AddItem Null2String(RSUPLOAD![MODELCODE])
            RSUPLOAD.MoveNext
        Loop
    End If
    txtkeyword = "aga": txtkeyword = ""
    SSTab1.Tab = 0
End Sub

Private Sub frm1_Closed()
    picqty.Visible = False
End Sub

Private Sub ListView1_DblClick()
    Dim Index                                          As Integer

    If ListView1.ListItems.Count = 0 Then Exit Sub

    Index = ListView1.SelectedItem.Index

    ListView1.ListItems.Remove (Index)
End Sub

Private Sub ListView2_DblClick()
    Dim Index                                          As Integer

    If ListView2.ListItems.Count = 0 Then Exit Sub

    Index = ListView2.SelectedItem.Index

    ListView2.ListItems.Remove (Index)
End Sub

Private Sub ListView3_DblClick()
    Dim Index                                          As Integer

    If ListView3.ListItems.Count = 0 Then Exit Sub

    Index = ListView3.SelectedItem.Index

    ListView3.ListItems.Remove (Index)
End Sub

Private Sub lstParts_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtPartNo = lstParts.SelectedItem
    txtDesc = lstParts.SelectedItem.SubItems(1)
    txtLOCATION = lstParts.SelectedItem.SubItems(4)
    txtIsHariParts = lstParts.SelectedItem.SubItems(6)
    txtSRP = lstParts.SelectedItem.SubItems(2)

    If lstParts.SelectedItem.SubItems(7) = "P" Then
        labType.Caption = "Parts"
    ElseIf lstParts.SelectedItem.SubItems(7) = "A" Then
        labType.Caption = "Accessories"
    ElseIf lstParts.SelectedItem.SubItems(7) = "M" Then
        labType.Caption = "Materials"
    Else
        labType.Caption = "Please verify this item no identification...."
    End If
End Sub

Private Sub lstParts_DblClick()
    If lstParts.ListItems.Count = 0 Then Exit Sub

    cmdSelect.Value = True
End Sub

Private Sub lstParts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdSelect.Value = True
    End If
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        txtkeyword.SetFocus
    End If
End Sub

Private Sub txtKeyword_Change()
    Set RSUPLOAD = New ADODB.Recordset

    lstParts.Enabled = False

    On Error GoTo ErrorCode

    lstParts.Sorted = False: lstParts.ListItems.Clear
    If chkModel.Value = 1 Then
        If CatAll.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and MODELCODE ='" & cboModel & "'order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and MODELCODE ='" & cboModel & "'order by STOCKDESC asc")
            End If
        ElseIf CatParts.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='P' and MODELCODE ='" & cboModel & "'order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='P' and MODELCODE ='" & cboModel & "'order by STOCKDESC asc")
            End If
        ElseIf CatAcesories.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='A' and MODELCODE ='" & cboModel & "'order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='A' and MODELCODE ='" & cboModel & "'order by STOCKDESC asc")
            End If
        ElseIf Option1.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='M' and MODELCODE ='" & cboModel & "'order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='M' and MODELCODE ='" & cboModel & "'order by STOCKDESC asc")
            End If
        End If
    Else
        If CatAll.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' order by STOCKDESC asc")
            End If
        ElseIf CatParts.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='P' order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='P' order by STOCKDESC asc")
            End If
        ElseIf CatAcesories.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='A' order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='A' order by STOCKDESC asc")
            End If
        ElseIf Option1.Value = True Then
            If optPartNo.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKNO like '" & txtkeyword & "%' and [Type]='M' order by STOCKNO asc")
            ElseIf optDescription.Value = True Then
                Set RSUPLOAD = gconDMIS.Execute("select STOCKNO,STOCKDESC,SRP,MODELCODE,LOCATION,ONHAND,NON_HARI,[type] from PMIS_STOCKMAS where STOCKDESC like '" & txtkeyword & "%' and [Type]='M' order by STOCKDESC asc")
            End If
        End If
    End If
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstParts.ListItems, RSUPLOAD
    End If

    lstParts.Enabled = True

    Exit Sub


ErrorCode:
    ShowVBError
    Exit Sub

End Sub

Private Sub txtKeyword_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstParts.Enabled = True Then
            lstParts.SetFocus
        End If
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorCode:

    If Val(txtQty) <= 0 Then Exit Sub
    If IsNumeric(txtQty) = False Then
        MsgBox "Enter a Valid Qty", vbInformation, "CSMS"
        txtQty.SetFocus
        Exit Sub
    End If

    If KeyAscii = 13 Then
        If labType = "Parts" Then
            ListView1.Enabled = False
            With ListView1
                .Sorted = False
                .ListItems.Add , , Left(labType.Caption, 1)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , txtPartNo
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtDesc
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtQty
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtSRP

            End With

            picqty.Visible = False
            On Error Resume Next
            txtkeyword.SetFocus

            ListView1.Enabled = True
        ElseIf labType = "Materials" Then
            ListView2.Enabled = False
            With ListView2
                .Sorted = False
                .ListItems.Add , , Left(labType.Caption, 1)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , txtPartNo
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtDesc
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtQty
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtSRP

            End With

            picqty.Visible = False
            On Error Resume Next
            txtkeyword.SetFocus

            ListView2.Enabled = True
        ElseIf labType = "Accessories" Then
            ListView3.Enabled = False
            With ListView3
                .Sorted = False
                .ListItems.Add , , Left(labType.Caption, 1)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , txtPartNo
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtDesc
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtQty
                .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtSRP
            End With

            picqty.Visible = False
            On Error Resume Next
            txtkeyword.SetFocus

            ListView3.Enabled = True
        End If

        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame3.Enabled = True
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

