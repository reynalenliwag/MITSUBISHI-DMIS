VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_PurchaseOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sublet Repair Purchase Order Data Entry"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_PurchaseOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   11565
   Begin VB.PictureBox Picture1 
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
      Height          =   870
      Left            =   2160
      ScaleHeight     =   870
      ScaleWidth      =   9405
      TabIndex        =   39
      Top             =   6690
      Width           =   9405
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   8580
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   795
         Left            =   7800
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelPO 
         Caption         =   "Cancel"
         Height          =   795
         Left            =   7020
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost "
         Height          =   795
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":1E7E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":1FD0
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
         Height          =   795
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":2315
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":2467
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   4680
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":278C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":28DE
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3900
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":2C3A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":2D8C
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   795
         Left            =   3120
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":309F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":31F1
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   795
         Left            =   2340
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":3541
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":3693
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1560
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":39F1
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":3B43
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   780
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":3E3D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":3F8F
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   0
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":42E7
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":4439
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox picJobs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   285
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   9345
      TabIndex        =   28
      Top             =   6330
      Width           =   9375
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Jobs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   33
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Jobs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   32
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Jobs"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   31
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4950
         TabIndex        =   30
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7110
         TabIndex        =   29
         Top             =   30
         Width           =   2445
      End
   End
   Begin VB.PictureBox picParts 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   285
      Left            =   2160
      ScaleHeight     =   255
      ScaleWidth      =   9345
      TabIndex        =   14
      Top             =   6330
      Width           =   9375
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   18
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   17
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5070
         TabIndex        =   16
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7110
         TabIndex        =   15
         Top             =   30
         Width           =   2445
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   2160
      TabIndex        =   27
      Top             =   3360
      Width           =   9375
      Begin TabDlg.SSTab SSTab1 
         Height          =   2900
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5106
         _Version        =   393216
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Sublet Labor"
         TabPicture(0)   =   "frmCSMS_PurchaseOrder.frx":4798
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstJobSublet"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Materials"
         TabPicture(1)   =   "frmCSMS_PurchaseOrder.frx":47B4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lstMaterials"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Parts"
         TabPicture(2)   =   "frmCSMS_PurchaseOrder.frx":47D0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lstparts"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSComctlLib.ListView lstJobSublet 
            Height          =   2450
            Left            =   60
            TabIndex        =   80
            Top             =   360
            Width           =   9230
            _ExtentX        =   16272
            _ExtentY        =   4313
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Contractor_Amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstMaterials 
            Height          =   2450
            Left            =   -74940
            TabIndex        =   81
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4313
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Contractor_amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lstparts 
            Height          =   2450
            Left            =   -74940
            TabIndex        =   82
            Top             =   360
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   4313
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LineNo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Job Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Job Description"
               Object.Width           =   8643
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Sublet Cost"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "LIVIL"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Contractor_amt"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "WCODE"
               Object.Width           =   0
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   2160
      TabIndex        =   20
      Top             =   0
      Width           =   9375
      Begin Crystal.CrystalReport rptSubletPo 
         Left            =   180
         Top             =   1410
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.TextBox txtContractorCode 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1950
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1860
         Width           =   1845
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   390
         Top             =   870
      End
      Begin VB.TextBox txtContractorAdd 
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
         ForeColor       =   &H00404040&
         Height          =   645
         Left            =   1950
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2250
         Width           =   7215
      End
      Begin VB.TextBox txtCustAdd 
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
         ForeColor       =   &H00404040&
         Height          =   555
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1260
         Width           =   4095
      End
      Begin VB.TextBox txtVatAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   35
         Top             =   900
         Width           =   2295
      End
      Begin VB.TextBox txtNetAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   34
         Top             =   1290
         Width           =   2295
      End
      Begin VB.TextBox txtContactPerson 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1950
         TabIndex        =   8
         Top             =   2940
         Width           =   3495
      End
      Begin VB.ComboBox cboContractor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3930
         TabIndex        =   6
         Text            =   "cboContractor"
         Top             =   1860
         Width           =   5235
      End
      Begin VB.TextBox txtCustName 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   3
         Top             =   870
         Width           =   4065
      End
      Begin VB.TextBox txtRoNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   4590
         TabIndex        =   22
         Top             =   120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   104136705
         CurrentDate     =   39559
      End
      Begin VB.TextBox txtPoNumber 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Top             =   150
         Width           =   1875
      End
      Begin VB.TextBox txtTotalAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   6870
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   36
         Top             =   510
         Width           =   2295
      End
      Begin VB.Label lbltech 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "lbltech"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   12
         Top             =   2610
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1095
         TabIndex        =   62
         Top             =   2250
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Vat Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5700
         TabIndex        =   60
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label labID 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "labID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   59
         Top             =   2250
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblPOSTED 
         Alignment       =   2  'Center
         Caption         =   "POSTED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   7080
         TabIndex        =   58
         Top             =   150
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label labDET 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "labDET"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   57
         Top             =   2610
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5550
         TabIndex        =   38
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5685
         TabIndex        =   37
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   405
         TabIndex        =   26
         Top             =   2970
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Service Contractor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   1950
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   24
         Top             =   570
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "PO Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   210
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7515
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   2115
      Begin VB.OptionButton optPONo 
         Caption         =   "PO number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   630
         Width           =   1875
      End
      Begin VB.TextBox txtSearch 
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
         TabIndex        =   9
         Top             =   960
         Width           =   1995
      End
      Begin MSComctlLib.ListView lvwTran 
         Height          =   6105
         Left            =   60
         TabIndex        =   78
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10769
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
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":47EC
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tranno"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label22 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   855
      Left            =   9885
      ScaleHeight     =   855
      ScaleWidth      =   1650
      TabIndex        =   52
      Top             =   6690
      Width           =   1650
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   810
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":494E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":4AA0
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   30
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":4DDE
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":4F30
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Height          =   315
      Left            =   30
      ScaleHeight     =   285
      ScaleWidth      =   11505
      TabIndex        =   63
      Top             =   7560
      Width           =   11535
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SI NO"
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
         Height          =   225
         Left            =   30
         TabIndex        =   67
         Top             =   30
         Width           =   615
      End
      Begin VB.Label lblSINO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   660
         TabIndex        =   66
         Top             =   30
         Width           =   1530
      End
      Begin VB.Label lblRRNO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   225
         Left            =   2970
         TabIndex        =   65
         Top             =   30
         Width           =   810
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RR NO"
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
         Height          =   225
         Left            =   2220
         TabIndex        =   64
         Top             =   30
         Width           =   720
      End
   End
   Begin VB.PictureBox picPrintPOExcel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Height          =   4515
      Left            =   4920
      ScaleHeight     =   4485
      ScaleWidth      =   3795
      TabIndex        =   68
      Top             =   1260
      Visible         =   0   'False
      Width           =   3825
      Begin VB.CommandButton Command4 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   2820
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":5280
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":53D2
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Exit Window"
         Top             =   3300
         Width           =   795
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Save"
         Height          =   795
         Left            =   2040
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":5738
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":588A
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Save this Record"
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox txtowner 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   55
         Top             =   2580
         Width           =   3375
      End
      Begin VB.CommandButton cmdPrintPOExcel 
         Caption         =   "&Print"
         Height          =   795
         Left            =   1260
         MouseIcon       =   "frmCSMS_PurchaseOrder.frx":5BDA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_PurchaseOrder.frx":5D2C
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Print this Record"
         Top             =   3300
         Width           =   795
      End
      Begin VB.TextBox txtSIG_Notedby 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   71
         Top             =   1350
         Width           =   3345
      End
      Begin VB.TextBox txtSIG_PreparedBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   70
         Top             =   630
         Width           =   3345
      End
      Begin VB.TextBox txtSIG_NotedbyDesign 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   210
         TabIndex        =   69
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   56
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   61
         Top             =   390
         Width           =   1200
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   77
         Top             =   0
         Width           =   4305
         _Version        =   655364
         _ExtentX        =   7594
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "PO SIGNITORIES"
         ForeColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14606302
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED BY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   76
         Top             =   1770
         Width           =   1230
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   315
         Index           =   0
         Left            =   -60
         TabIndex        =   75
         Top             =   4200
         Width           =   4305
         _Version        =   655364
         _ExtentX        =   7594
         _ExtentY        =   556
         _StockProps     =   14
         ForeColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   14606302
      End
   End
End
Attribute VB_Name = "frmCSMS_PurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim AddorEdit                                          As String
Dim getPOSTATUS                                        As String
Dim rsINFO                                             As ADODB.Recordset
Dim rsFilllstPO_HD                                     As ADODB.Recordset
Dim rsPO_RC_HD                                         As ADODB.Recordset
Dim lstLivil                                           As String
Dim lstLine_No                                         As String
Dim vSublet_TOTAL_AMT                                  As Double
Dim vSublet_TOTAl_VAT                                  As Double
Dim vSublet_NET_AMT                                    As Double
Dim AUDIT_SQL                                          As String
Dim xxCODE                                             As String
Dim PASTPONUMBER                                       As String
Dim str_MSG                                            As String
Dim ERROR_MSG                                          As String


Function SetContractorAdd(XXX As String) As String
    On Error Resume Next
    Dim rsContractorAdd                                As New ADODB.Recordset
    'Set rsContractorAdd = gconDMIS.Execute("Select * from CSMS_Contractor Where CompanyName = '" & Repleys(XXX) & "'")
    Set rsContractorAdd = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE Where NAMEOFVENDOR = '" & LTrim(RTrim(XXX)) & "' AND CODE IS NOT NULL")
    If Not rsContractorAdd.EOF And Not rsContractorAdd.BOF Then
        SetContractorAdd = Null2String(rsContractorAdd!Address)
        txtContractorCode.Text = Null2String(rsContractorAdd!Code)
        If rsContractorAdd!NONVAT = "Y" Then
            txtVatAmount.Locked = True
            txtVatAmount.Locked = ""
        Else
            txtVatAmount.Locked = False
        End If
    End If
    Set rsContractorAdd = Nothing
End Function

Function SetPoStatus(XXX As String) As String
    Dim rsPOSTATUS                                     As New ADODB.Recordset
    Set rsPOSTATUS = gconDMIS.Execute("Select * from CSMS_Po_Hd Where Po_No = '" & XXX & "'")
    If Not rsPOSTATUS.EOF And Not rsPOSTATUS.BOF Then
        SetPoStatus = Null2String(rsPOSTATUS!Status)
    End If
    Set rsPOSTATUS = Nothing
End Function

Function CheckIfPosted(XXX As Variant) As Variant
    Dim rsPosted                                       As New ADODB.Recordset

    Set rsPosted = gconDMIS.Execute("Select * from CSMS_Po_hd where STATUS = 'P' and PO_NO ='" & XXX & "'")
    If Not rsPosted.EOF And Not rsPosted.BOF Then
        CheckIfPosted = True
        lblPOSTED.Visible = True
    Else
        CheckIfPosted = False
        lblPOSTED.Visible = False
    End If
    Set rsPosted = Nothing
End Function

Function passID(XXX As Variant) As Variant
    Call rsRefresh
    rsINFO.Find ("id=" & labid)
    Call StoreMemVars
End Function

Sub editJobs()
    Dim rsEditJobs                                     As New ADODB.Recordset
    Dim vTypeOfJob                                     As String

    On Error Resume Next

    'check Purchase Order if Already  posted
    getPOSTATUS = SetPoStatus(txtPoNumber.Text)

    Set rsEditJobs = gconDMIS.Execute("Select TRANSTATUS,JOBTYPE,DETCDE,DETDSC,DET_AMT,CONTRACTAMOUNT,COMPAMOUNT,WCODE,DETAIL,TECHCODE,SUBLET_TYPE from CSMS_Po_Dt where ID ='" & labDET & "'")
    If Not (rsEditJobs.BOF And rsEditJobs.EOF) Then
        If Trim(rsEditJobs!DETCDE) = "SRLABOR" Then
            vTypeOfJob = "SUBLET LABOR"
        ElseIf Trim(rsEditJobs!DETCDE) = "SRPARTS" Then
            vTypeOfJob = "SUBLET PARTS"
        Else
            vTypeOfJob = "SUBLET MATERIALS"
        End If

        If Not rsEditJobs.EOF And Not rsEditJobs.BOF Then
            With frmCSMS_SubletJob
                .txtCustomer = Null2String(txtCustName.Text)
                .txtROno = Null2String(txtROno)
                .txtOPCODE = Null2String(rsEditJobs!DETCDE)
                .txtJobDesc = Null2String(rsEditJobs!DETDSC)
                .txtSubletAmount = Format(NumericVal(rsEditJobs!DET_AMT), MAXIMUM_DIGIT)
                .txtContracAmount = Format(NumericVal(rsEditJobs!CONTRACTAMOUNT), MAXIMUM_DIGIT)
                .txtCompAmount = Format(NumericVal(rsEditJobs!COMPAMOUNT), MAXIMUM_DIGIT)
                .cboJobChargeTo = Null2String(rsEditJobs!wCode)
                .txtNote = Null2String(rsEditJobs!Detail)
                .cboSubletCategory = vTypeOfJob
                .cboBPorGJ = Null2String(rsEditJobs!JOBTYPE)
                .lbltechcode.Caption = Null2String(rsEditJobs!TechCode)
                
                If COMPANY_CODE = "CMC" Or COMPANY_CODE = "DSSC" Then
                     If Null2String(rsEditJobs!SUBLET_TYPE) = "T" Then
                           .cbosublettype.Text = "Tinsmith"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "P" Then
                           .cbosublettype.Text = "Painting"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "A" Then
                           .cbosublettype.Text = "Aircon"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "U" Then
                           .cbosublettype.Text = "Undercoating"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "D" Then
                           .cbosublettype.Text = "Detailing"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "S" Then
                           .cbosublettype.Text = "Sublet"
                       ElseIf Null2String(rsEditJobs!SUBLET_TYPE) = "O" Then
                           .cbosublettype.Text = "Towing"
                       End If
                       
                 End If
                
                'UPDATE BY : MJP 07222008
                If UCase(.cboSubletCategory.Text) = "SUBLET LABOR" Then
                    If Null2String(rsEditJobs!JOBTYPE) = "BP" Then
                        .cboBP_TYPE.Visible = True
                        .Label4.Visible = True
                        If Null2String(rsEditJobs!transtatus) = "M" Then
                            .cboBP_TYPE.Text = "Major"
                        Else
                            .cboBP_TYPE.Text = "Minor"
                        End If
                    Else
                        .cboBP_TYPE.Visible = False
                        .Label4.Visible = False
                    End If
                End If

                .labDET = labDET.Caption
            End With
        End If
    End If
    Set rsEditJobs = Nothing
End Sub

Sub initMemvars()
    Dim rsPO_Counter                                   As New ADODB.Recordset

    rsPO_Counter.Open "Select Max(Right(Po_No,6)) as PO_NO from CSMS_Po_hd", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'JJE Prefix
'    If COMPANY_CODE = "DJM" Then
'        If Not rsPO_Counter.EOF And Not rsPO_Counter.BOF Then
'            txtPoNumber.Text = Format(Right(rsPO_Counter!PO_NO, 6) + 1, "000000")
'            txtPoNumber.Text = "SP" + txtPoNumber.Text
'        Else
'            txtPoNumber.Text = "SP" + "000001"
'        End If
'    Else
    If COMPANY_CODE = "DJM" Then
        If Not rsPO_Counter.EOF And Not rsPO_Counter.BOF Then
            txtPoNumber.Text = Format(NumericVal(Mid$(Null2String(rsPO_Counter!PO_NO), 1, 6)) + 1, "000000")
        Else
            txtPoNumber.Text = "000001"
        End If
        txtPoNumber.Locked = True
    Else
        If Not rsPO_Counter.EOF And Not rsPO_Counter.BOF Then
            txtPoNumber.Text = Format(NumericVal(Mid$(Null2String(rsPO_Counter!PO_NO), 1, 6)) + 1, "000000")
        Else
            txtPoNumber.Text = "000001"
        End If
        txtPoNumber.Locked = True
    End If
'    End If
    'JJE
    
    lblSINO.Caption = ""
    lblRRNO.Caption = ""
    txtROno.Text = ""
    txtCustName.Text = ""
    txtCustAdd.Text = ""
    txtContractorAdd.Text = ""
    txtContactPerson.Text = ""
    txtVatAmount.Text = ""
    txtNetAmount.Text = ""
    txtTotalAmount.Text = ""
    DTPicker1.Value = Date
    cboContractor.ListIndex = -1
    lblPOSTED.Caption = ""
    txtContractorCode.Text = ""
    labDET = 0
    labid = 0
End Sub

Sub initCboContractor()
    Dim rsContractor                                   As New ADODB.Recordset
    
    'Set rsContractor = gconDMIS.Execute("Select * from CSMS_Contractor Order by code asc")
    'UPDATED BY: JUN
    'DATE UPDATED: 0904/2008
    'DESCRIPTION: RETRIEVE THE NAME OF VENDOR REGISTERED IN AMIS MODULE
    Set rsContractor = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE WHERE CODE IS NOT NULL Order by code asc")
    If Not rsContractor.EOF And Not rsContractor.BOF Then
        rsContractor.MoveFirst: cboContractor.Clear
        Do While Not rsContractor.EOF
            cboContractor.AddItem Null2String(rsContractor!nameofvendor)
            'cboContractor.AddItem Null2String(rsContractor!CompanyName)
            rsContractor.MoveNext
        Loop
    End If
    Set rsContractor = Nothing
End Sub

Function FindRRno(vPONO As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT RC_NO FROM CSMS_PO_RC_HD WHERE PO_NO = '" & vPONO & "' and STATUS <> 'C'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindRRno = Null2String(rstmp!RC_NO)
    Else
        FindRRno = ""
    End If
    Set rstmp = Nothing
End Function

Function findSINO(vRONO As String) As String
    Dim rstmp                                          As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE REP_OR = '" & vRONO & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        findSINO = Null2String(rstmp!INVOICE)
    Else
        findSINO = ""
    End If
    Set rstmp = Nothing
End Function

Sub StoreMemVars()
    If Not (rsINFO.EOF Or rsINFO.BOF) Then
        labid = rsINFO!ID
        txtPoNumber.Text = Null2String(rsINFO!PO_NO)
        txtROno.Text = Null2String(rsINFO!RO_NO)

        '-----------------------------------------------------------------------------------------
        'UPDATE BY : MJP 09022008 04:57 PM
        lblRRNO.Caption = FindRRno(txtPoNumber)
        lblSINO.Caption = findSINO(txtROno)
        '-----------------------------------------------------------------------------------------

        DTPicker1.Value = rsINFO!Po_Date
        txtCustName.Text = Null2String(rsINFO!Cust_name)
        txtCustAdd.Text = Null2String(rsINFO!Customer_Add)
        txtContractorCode.Text = Null2String(rsINFO!Contractor_Code)
        cboContractor.Text = Null2String(rsINFO!Contractor_Name)
        txtContractorAdd.Text = Null2String(rsINFO!Contractor_Address)
        txtContactPerson = Null2String(rsINFO!Contact_Person)
        lbltech.Caption = Null2String(rsINFO!Contractor_Code)
        txtTotalAmount.Text = Format((NumericVal(rsINFO!Sublet_TOTAL_AMT)), MAXIMUM_DIGIT)
        txtVatAmount.Text = Format((NumericVal(rsINFO!Sublet_TOTAl_VAT)), MAXIMUM_DIGIT)
        txtNetAmount.Text = Format((NumericVal(rsINFO!SUBLET_TOTAL_NET_AMT)), MAXIMUM_DIGIT)

        Call FillListview(txtPoNumber)

        If Null2String(rsINFO!Status) = "P" Then
            cmdPrint.Enabled = True
            cmdPost.Enabled = False

            Dim rsUnpostDisabled                       As New ADODB.Recordset
            Set rsUnpostDisabled = gconDMIS.Execute("Select DTE_COMP from CSMS_repor where rep_or ='" & txtROno & "' and dte_comp is null")
            If Not rsUnpostDisabled.EOF And Not rsUnpostDisabled.BOF Then
                cmdUnPost.Enabled = True
            Else
                cmdUnPost.Enabled = False
            End If
            'cmdUnPost.Enabled = True
            cmdEdit.Enabled = False
            cmdCancelPO.Enabled = False
            lblPOSTED = "**POSTED**"
        ElseIf Null2String(rsINFO!Status) = "C" Then
            cmdEdit.Enabled = False
            cmdPrint.Enabled = False
            cmdPost.Enabled = False
            cmdCancelPO.Enabled = False
            cmdUnPost.Enabled = False
            lblPOSTED = "**CANCELLED**"
        Else
            cmdEdit.Enabled = True
            cmdPrint.Enabled = False
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelPO.Enabled = True
            lblPOSTED = ""
        End If
    Else
        Call ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsINFO = New ADODB.Recordset
    rsINFO.Open "select * from CSMS_Po_Hd order by Po_No desc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub passINFO()
    frmCSMS_SubletJob.txtCustomer = txtCustName.Text
    frmCSMS_SubletJob.txtROno = txtROno.Text
'Updated by:    IEBV_08262010_0200pm
    With frmCSMS_SubletJob
         .lbltechcode.Caption = lbltech.Caption
    End With
End Sub

Sub deleteJobs()
    Dim rsDelJob                                       As New ADODB.Recordset
    Dim ans                                            As String

    ans = MsgBox("Are you sure do you want to DELETE this Job?", vbQuestion + vbYesNo)
    If ans = vbYes Then
        SQL_STATEMENT = "Delete from CSMS_Po_Dt where ID = '" & labDET & "' and LIVIL = '" & lstLivil & "'"
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "CODE: " & xxCODE, "", labDET)
        'NEW LOG AUDIT-----------------------------------------------------


        Dim rsComputeTotalCost                         As New ADODB.Recordset
        Set rsComputeTotalCost = gconDMIS.Execute("Select DETAMT,TAXVAL,DET_AMT from CSMS_PO_DT where PO_NO ='" & txtPoNumber & "'")

        vSublet_TOTAL_AMT = 0
        vSublet_TOTAl_VAT = 0
        vSublet_NET_AMT = 0


        If Not rsComputeTotalCost.EOF And Not rsComputeTotalCost.BOF Then
            Do While Not rsComputeTotalCost.EOF
                vSublet_TOTAL_AMT = N2Str2Zero(rsComputeTotalCost!DETAMT) + N2Str2Zero(vSublet_TOTAL_AMT)
                vSublet_TOTAl_VAT = N2Str2Zero(rsComputeTotalCost!TAXVAL) + N2Str2Zero(vSublet_TOTAl_VAT)
                vSublet_NET_AMT = N2Str2Zero(rsComputeTotalCost!DET_AMT) + N2Str2Zero(vSublet_NET_AMT)
                rsComputeTotalCost.MoveNext
            Loop
        End If

        SQL_STATEMENT = "Update CSMS_PO_HD set " & _
            "SUBLET_TOTAL_AMT = " & vSublet_TOTAL_AMT & "," & _
            "SUBLET_TOTAL_VAT = " & vSublet_TOTAl_VAT & "," & _
            "SUBLET_TOTAL_NET_AMT = " & vSublet_NET_AMT & " " & _
            "where PO_NO =" & txtPoNumber
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber & " - DELETE DETAILS", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call ShowDeletedMsg
    Else
        labDET.Caption = ""
        Exit Sub
    End If

    Call rsRefresh
    rsINFO.Find ("ID =" & labid)
    Call StoreMemVars
End Sub

Sub PostPurchaseOrder()
    Dim rsstatus                                       As New ADODB.Recordset
    Dim ans                                            As String
    Dim vSTAT                                          As String

    vSTAT = "'P'"
    getPOSTATUS = SetPoStatus(txtPoNumber.Text)       'check if PO is already posted

    If getPOSTATUS = "P" Then
        MsgBox ("Purchase order already posted..."), vbOKOnly + vbInformation, "INFORMATION"
        Exit Sub
    Else
        ans = MsgBox("Are you sure you want to POST this" & vbCrLf & "Purchase order?", vbYesNo + vbInformation)
        If ans = vbYes Then
            gconDMIS.Execute "update CSMS_Po_Hd set " & _
                "Status = " & vSTAT & "" & _
                "where Po_No = " & txtPoNumber
        Else
            Exit Sub
        End If
    End If
End Sub

Sub UnPostPurchaseOrder()
    Dim rsstatus                                       As New ADODB.Recordset
    Dim ans                                            As String

    getPOSTATUS = SetPoStatus(txtPoNumber.Text)       'check if PO is already posted

    If getPOSTATUS = "P" Then
        ans = MsgBox("Are you sure you want to UNPOST this" & vbCrLf & "Purchase order", vbYesNo + vbInformation)
        If ans = vbYes Then
            gconDMIS.Execute "update CSMS_Po_Hd set " & _
                "Status = NULL " & _
                "where Po_No =" & txtPoNumber
            Exit Sub
        End If
    Else
        MsgBox "This transaction is not yet Posted"
        Exit Sub
    End If
End Sub

Sub FilllstPO_HD()
    Dim i                                              As Integer
    Listview_Loadval lvwTran.ListItems, gconDMIS.Execute("Select top 30 po_no,id  from CSMS_Po_Hd  order by Po_no desc")
End Sub

Sub refreshPO_RC_HD()
    Set rsPO_RC_HD = New ADODB.Recordset
    rsPO_RC_HD.Open "select * from CSMS_PO_RC_DT order by RC_NO asc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub FillListview(XXX As String)
    Dim Item                                           As ListItem
    Dim rsPO_dt                                        As New ADODB.Recordset

    'LABOR
    Me.lstJobSublet.Sorted = True: Me.lstJobSublet.ListItems.Clear: Me.lstJobSublet.Enabled = False
    
    'JJE
'    If COMPANY_CODE = "DJM" Then   ** FOR APPROVAL **
'        Set rsPO_dt = gconDMIS.Execute("Select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = '" & XXX & "' and livil = '1' order by LINE_NO asc")
'    Else
        Set rsPO_dt = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = " & XXX & " and livil = '1' order by LINE_NO asc")
'    End If
    'JJE
    
    If Not rsPO_dt.EOF And Not rsPO_dt.BOF Then
        Do While Not rsPO_dt.EOF
            Set Item = lstJobSublet.ListItems.Add(, , Null2String(rsPO_dt!LINE_NO))
            Item.SubItems(1) = Null2String(rsPO_dt!DETCDE)
            Item.SubItems(2) = Null2String(rsPO_dt!Detail)
            Item.SubItems(3) = Format(NumericVal(rsPO_dt!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsPO_dt!ID)
            Item.SubItems(5) = Null2String(rsPO_dt!LIVIL)
            Item.SubItems(6) = Null2String(rsPO_dt!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsPO_dt!wCode)
            rsPO_dt.MoveNext
        Loop
        Me.lstJobSublet.Enabled = True: Me.lstJobSublet.Sorted = False: Me.lstJobSublet.Refresh
    End If

    Set rsPO_dt = Nothing

    'MATERIALS
    Dim rsMaterials                                    As New ADODB.Recordset

    Me.lstMaterials.Sorted = True: Me.lstMaterials.ListItems.Clear: Me.lstMaterials.Enabled = False
    Set rsMaterials = New ADODB.Recordset
    'JJE
'    If COMPANY_CODE = "DJM" Then   ** FOR APPROVAL **
'        Set rsMaterials = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = '" & XXX & "' and livil = '3' order by LINE_NO asc")
'    Else
        Set rsMaterials = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = " & XXX & " and livil = '3' order by LINE_NO asc")
'    End If
    'JJE
    
    If Not rsMaterials.EOF And Not rsMaterials.BOF Then
        Do While Not rsMaterials.EOF
            Set Item = lstMaterials.ListItems.Add(, , Null2String(rsMaterials!LINE_NO))
            Item.SubItems(1) = Null2String(rsMaterials!DETCDE)
            Item.SubItems(2) = Null2String(rsMaterials!Detail)
            Item.SubItems(3) = Format(NumericVal(rsMaterials!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsMaterials!ID)
            Item.SubItems(5) = Null2String(rsMaterials!LIVIL)
            Item.SubItems(6) = Null2String(rsMaterials!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsMaterials!wCode)
            rsMaterials.MoveNext
        Loop
        Me.lstMaterials.Enabled = True: Me.lstMaterials.Sorted = False: Me.lstMaterials.Refresh
    End If
    Set rsMaterials = Nothing

    'PARTS
    Dim rsParts                                        As New ADODB.Recordset

    Me.lstparts.Sorted = True: Me.lstparts.ListItems.Clear: Me.lstparts.Enabled = False
    Set rsParts = New ADODB.Recordset
    'JJE
'    If COMPANY_CODE = "DJM" Then   ** FOR APPROVAL **
'        Set rsParts = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = '" & XXX & "' and livil = '2' order by LINE_NO asc")
'    Else
        Set rsParts = gconDMIS.Execute("select Line_No,DETCDE,DETDSC,DET_AMT,ID,LIVIL,CONTRACTAMOUNT,WCODE,DETAIL from CSMS_Po_Dt where Po_no = " & XXX & " and livil = '2' order by LINE_NO asc")
'    End If
    'JJE

    If Not rsParts.EOF And Not rsParts.BOF Then
        Do While Not rsParts.EOF
            Set Item = lstparts.ListItems.Add(, , Null2String(rsParts!LINE_NO))
            Item.SubItems(1) = Null2String(rsParts!DETCDE)
            Item.SubItems(2) = Null2String(rsParts!Detail)
            Item.SubItems(3) = Format(NumericVal(rsParts!DET_AMT), MAXIMUM_DIGIT)
            Item.SubItems(4) = Null2String(rsParts!ID)
            Item.SubItems(5) = Null2String(rsParts!LIVIL)
            Item.SubItems(6) = Null2String(rsParts!CONTRACTAMOUNT)
            Item.SubItems(7) = Null2String(rsParts!wCode)
            rsParts.MoveNext
        Loop
        Me.lstparts.Enabled = True: Me.lstparts.Sorted = False: Me.lstparts.Refresh
    End If

    Set rsParts = Nothing
End Sub

Sub UnpostDelete()
    'check if already received
    Dim rscheckIfReceived                              As New ADODB.Recordset

    Set rscheckIfReceived = gconDMIS.Execute("Select Status from CSMS_PO_RC_HD where PO_NO ='" & txtPoNumber & "' and (status ='R' or status ='P')")
    If Not rscheckIfReceived.EOF And Not rscheckIfReceived.BOF Then
        MsgBox "You cannot Unpost this Purchase Order It's Already Received."
        Exit Sub
        '     Else
        '        cmdUnPost.Value = True
    End If


    Dim rsUnpostDelete                                 As New ADODB.Recordset
    Dim rsChekIfInvoice                                As New ADODB.Recordset
    Dim alreadyInvoice                                 As Boolean

    Set rsChekIfInvoice = gconDMIS.Execute("Select DTE_COMP from CSMS_repor where rep_or ='" & txtROno & "' and dte_comp is null")
    If Not rsChekIfInvoice.EOF And Not rsChekIfInvoice.BOF Then
        alreadyInvoice = True
    Else
        alreadyInvoice = False
    End If

    If alreadyInvoice = True Then
        If MsgBox("Do You Want to Un Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub

        SQL_STATEMENT = "update CSMS_PO_HD set STATUS = NULL WHERE ID = " & labid
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("U", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        SQL_STATEMENT = "update CSMS_PO_dt set STATUS = NULL WHERE PO_No = " & N2Str2Null(txtPoNumber)
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("UU", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        SQL_STATEMENT = "Delete from CSMS_ro_Det where rep_or = '" & txtROno & "' and ROTYPE = 'SR' and SUBPOCODE = '" & txtPoNumber & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("XX", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtROno), "REP_OR", "CSMS_REPOR"), "", "SUBLET DETAILS REMOVE", "", "")
        'NEW LOG AUDIT-----------------------------------------------------

        Call rsRefresh
        rsINFO.Find ("id=" & labid)
        Call StoreMemVars
    Else
        MsgBox ("You cannot Unpost this Purchase Order." & vbCrLf & "Already Invoice."), vbOKOnly + vbInformation, "INFORMATION"
        Exit Sub
    End If
    Call CheckRO_Status(RTrim(LTrim(txtROno)))
End Sub

Private Sub cboContractor_Change()
    txtContractorAdd.Text = SetContractorAdd(cboContractor)
End Sub

Private Sub cboContractor_Click()
    txtContractorAdd.Text = SetContractorAdd(cboContractor)
End Sub

Private Sub cboContractor_LostFocus()
    cboContractor.Text = SetContractorName(txtContractorCode)
End Sub

Function SetContractorName(XXX As String) As String
    Dim rsContractorName                               As New ADODB.Recordset
    'Set rsContractorAdd = gconDMIS.Execute("Select * from CSMS_Contractor Where CompanyName = '" & Repleys(XXX) & "'")
    Set rsContractorName = gconDMIS.Execute("Select * from ALL_VENDOR_TABLE Where Code = '" & Repleys(XXX) & "' AND CODE IS NOT NULL")
    If Not rsContractorName.EOF And Not rsContractorName.BOF Then
        SetContractorName = Null2String(rsContractorName!nameofvendor)
    End If
    Set rsContractorName = Nothing
End Function

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "ACESS_ADD", "SUBLET PURCHASE") = False Then Exit Sub
    AddorEdit = "ADD"
    Picture1.Visible = False
    Picture2.Visible = True
    Frame1.Enabled = True
    txtSearch.Enabled = False
    lblPOSTED.Enabled = False
    Frame2.Enabled = False
    Call initMemvars

    lstJobSublet.ListItems.Clear
    lstMaterials.ListItems.Clear
    lstparts.ListItems.Clear
    On Error Resume Next
    txtROno.SetFocus
    initCboContractor
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    Frame2.Enabled = True
    lblPOSTED.Enabled = True
    txtSearch.Enabled = True
    lblPOSTED.Enabled = True
    txtSearch.Enabled = True
    'initMemvars
    Call StoreMemVars
End Sub

Private Sub cmdCancelPO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "SUBLET PURCHASE") = False Then Exit Sub
    If MsgBox("Do You Want to Cancel this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub

    SQL_STATEMENT = "update CSMS_PO_HD set STATUS = 'C' WHERE ID = " & labid
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("C", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Update CSMS_PO_dt set STATUS = 'C' WHERE PO_NO = '" & txtPoNumber & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("CC", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Call rsRefresh
    rsINFO.Find ("id = " & labid)
    Call StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "SUBLET PURCHASE") = False Then Exit Sub
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    'txtSearch.Enabled = False
    'JJE Prefix  ** FOR APPROVAL **
'    If COMPANY_CODE <> "DJM" Then
'        txtPoNumber.Locked = False
'    End If
    'JJE
    lblPOSTED.Enabled = False
    Frame2.Enabled = False
    PASTPONUMBER = GETPASTPONO(LTrim(RTrim((txtPoNumber.Text))))
End Sub

Function GETPASTPONO(PONUMBER) As String
    Dim SQLTXT                                  As String
    Dim rstmp                                   As New ADODB.Recordset
    
    SQLTXT = "SELECT PO_NO,RO_NO FROM CSMS_PO_HD WHERE PO_NO = '" & PONUMBER & "'"
    Set rstmp = gconDMIS.Execute(SQLTXT)
    
    If Not (rstmp.BOF And rstmp.EOF) Then
        GETPASTPONO = Trim(rstmp!PO_NO)
    End If
    
    rstmp.Close
    Set rstmp = Nothing
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsINFO.MoveFirst
    Call ShowFirstRecordMsg
    Call StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsINFO.MoveLast
    Call ShowLastRecordMsg
    Call StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsINFO.MoveNext
    If rsINFO.EOF Then
        rsINFO.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "ACESS_POST", "SUBLET PURCHASE") = False Then Exit Sub
    If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
        MsgBox ("You cannot post this Transaction" & vbCrLf & "There is no Job Selected."), vbOKOnly + vbInformation, "Information"
        Exit Sub
    End If

    'UPDATE BY   : MJP 010509 1112AM
    'DESCRIPTION : TO NOT ALLOW USER TO POST A TRANSACTION WHEN THE RO IS ALREADY INVOICE
        If CheckIfRoIsAlreadyInvoice(txtROno) = True Then
            MsgBox "This Transaction cannot be post, RO no. " & txtROno & " is already invoiced", vbExclamation, "DMIS"
            Exit Sub
        End If
        
    If MsgBox("Do You Want to Post this Transaction", vbInformation + vbYesNo) = vbNo Then Exit Sub
     
    
    'UPDATE BY  : MJP 010509 1112AM
    
    gconDMIS.BeginTrans
    If post = False Then
    
        str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
        str_MSG = str_MSG & "Description: "
        str_MSG = str_MSG & " " & ERROR_MSG
        str_MSG = str_MSG & " " & vbCrLf
        str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
        str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
        str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
        
        str_MSG = Replace(str_MSG, "@UTX83912839123", "Posting of Transaction")
        MsgBox str_MSG, vbCritical, "Posting Error"
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.CommitTrans
   
    

    Call rsRefresh
    rsINFO.Find ("id=" & labid)
    Call StoreMemVars
End Sub
Function getline(xRO As String, xLIVIL As Integer) As String
Dim rsline                          As ADODB.Recordset
Dim rschk                           As ADODB.Recordset
Dim ictr                            As Integer
Set rsline = gconDMIS.Execute("Select max (line_no) + 1  as line from csms_ro_det where rep_or = '" & xRO & "' and livil = '" & xLIVIL & "'")
If Not (rsline.EOF And rsline.BOF) Then
    If N2Str2IntZero(rsline!Line) = "0" Then
        getline = "01"
    ElseIf N2Str2IntZero(rsline!Line) > 99 Then
        For ictr = 1 To N2Str2IntZero(rsline!Line)
            Set rschk = gconDMIS.Execute("select line_no from csms_ro_det where line_no = '" & Format(ictr, "00") & "' AND [REP_OR] = '" & xRO & "' and LIVIL ='" & xLIVIL & "'")
            If Not (rschk.EOF And rschk.BOF) Then
            Else
                getline = Format(ictr, "00")
                Exit Function
            End If
        Next
    Else
        getline = Format(N2Str2IntZero(rsline!Line), "00")
    End If
Else
    getline = "01"
End If
Set rsline = Nothing
End Function
Function post() As Boolean
On Error GoTo errordaa
    SQL_STATEMENT = "update CSMS_PO_HD set STATUS = 'P' WHERE ID = " & labid
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("P", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    SQL_STATEMENT = "Update CSMS_PO_Dt set STATUS = 'P' WHERE PO_NO = '" & txtPoNumber & "'"
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("PP", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: '" & txtPoNumber, "'", "")
    'NEW LOG AUDIT-----------------------------------------------------

    'UPDATE THE TABLE CSMS_ro_DET------------------------------------------------------

    Dim rsPutJobToService                              As New ADODB.Recordset

    Dim pRep_or                                        As String
    Dim pJOBTYPE                                       As String
    Dim pLIVIL                                         As String
    Dim pLINE_NO                                       As String
    Dim pDETCDE                                        As String
    Dim pDETDSC                                        As String
    Dim pTECHNICIAN                                    As String
    Dim pDETAMT                                        As Double
    Dim pwCode                                         As String
    Dim pTAXRATE                                       As Double
    Dim pTAXVAL                                        As Double
    Dim pDETAIL                                        As String
    Dim pDET_AMT                                       As Double
    Dim pUSERCODE                                      As String
    Dim pSAVEDATE                                      As String
    Dim pTECHCODE                                      As String
    Dim pSTATUS                                        As String
    Dim pDONE                                          As String
    Dim pROTYPE                                        As String
    Dim vBP_TYPE                                       As String
    Dim vContractorAmount                              As Double
    Dim vDetvol                                        As Integer
    Dim vSUBLET_TYPE                                   As String
    
    'UPDATE BY: JUN
    'DATE UPDATE: 07/19/2008
    'DESCRIPTION: UPDATE DETCOST,DETPRICE AND DISCRATE

    'Dim pDETCOST As Double
    Dim pDETPRICE                                      As Double
    Dim pDISRATE                                       As Double
    '-------------------------------------------------
    'pDETCDE = 0:
    pDETPRICE = 0: pDISRATE = 0
    vDetvol = 0

    Set rsPutJobToService = gconDMIS.Execute("Select TRANSTATUS,Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,TECHNICIAN,DETAMT,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,USERCODE,SAVEDATE,TECHCODE,CONTRACTAMOUNT,SUBLET_TYPE from CSMS_PO_DT where PO_NO = '" & txtPoNumber.Text & "'")

    If Not rsPutJobToService.EOF And Not rsPutJobToService.BOF Then
        Do While Not rsPutJobToService.EOF
            vBP_TYPE = N2Str2Null(rsPutJobToService!transtatus)
            pRep_or = N2Str2Null(rsPutJobToService!REP_OR)
            pROTYPE = LTrim(RTrim(N2Str2Null(rsPutJobToService!ROTYPE)))
            pJOBTYPE = N2Str2Null(rsPutJobToService!JOBTYPE)
            pLIVIL = N2Str2Null(rsPutJobToService!LIVIL)
            pLINE_NO = N2Str2Null(getline(Null2String(rsPutJobToService!REP_OR), Null2String(rsPutJobToService!LIVIL)))
            pDETCDE = N2Str2Null(rsPutJobToService!DETCDE)
            pTECHNICIAN = N2Str2Null(rsPutJobToService!Technician)
            pDETAMT = NumericVal(rsPutJobToService!DETAMT)
            pwCode = N2Str2Null(rsPutJobToService!wCode)
            pTAXRATE = NumericVal((rsPutJobToService!taxrate) * 100)
            pTAXVAL = NumericVal(rsPutJobToService!TAXVAL)
            vSUBLET_TYPE = N2Str2Null(rsPutJobToService!SUBLET_TYPE)
            
            If pLIVIL = "'2'" Or pLIVIL = "'3'" Then
                pDETAIL = "NULL"
            Else
                pDETAIL = N2Str2Null(rsPutJobToService!Detail)
            End If

            If pLIVIL = "'2'" Or pLIVIL = "'3'" Then
                pDETDSC = N2Str2Null(rsPutJobToService!Detail)
                vDetvol = 1
            Else
                pDETDSC = N2Str2Null(rsPutJobToService!DETDSC)
            End If



            pDET_AMT = NumericVal(rsPutJobToService!DET_AMT)
            pUSERCODE = N2Str2Null(rsPutJobToService!USERCODE)
            pSAVEDATE = N2Str2Null(rsPutJobToService!savedate)
            pTECHCODE = N2Str2Null(rsPutJobToService!TechCode)
            pDONE = "'Y'"
            pSTATUS = "'Y'"
            vContractorAmount = ToDoubleNumber((rsPutJobToService!CONTRACTAMOUNT))

            'UPDATE BY: JUN
            'DATE UPDATE: 07/19/2008
            'DESCRIPTION: UPDATE DETCOST,DETPRICE AND DISCRATE

            pDETPRICE = NumericVal(rsPutJobToService!DET_AMT)

            SQL_STATEMENT = "Insert into CSMS_RO_DET" & _
                "(transtatus, Rep_or,ROTYPE,JOBTYPE,LIVIL,LINE_NO,DETCDE,DETDSC,DETPRC,DETVOL,DETCOST,TECHNICIAN,DETAMT,DISCRATE,WCODE,TAXRATE,TAXVAL,DETAIL,DET_AMT,STATUS,USERCDE,SAVEDATE,DONE,SUBPOCODE,TECHCODE,SUBLET_TYPE)" & _
                "values(" & vBP_TYPE & "," & pRep_or & _
                "," & pROTYPE & _
                "," & pJOBTYPE & _
                "," & pLIVIL & _
                "," & pLINE_NO & _
                "," & pDETCDE & _
                "," & pDETDSC & _
                "," & pDETPRICE & "," & vDetvol & _
                "," & vContractorAmount & _
                "," & pTECHNICIAN & _
                "," & pDETAMT & _
                "," & pDISRATE & _
                "," & pwCode & _
                "," & pTAXRATE & _
                "," & pTAXVAL & _
                "," & pDETAIL & _
                "," & pDET_AMT & _
                "," & pSTATUS & _
                "," & pUSERCODE & _
                "," & pSAVEDATE & _
                "," & pDONE & _
                ",'" & txtPoNumber & _
                "'," & pTECHCODE & ", " & vSUBLET_TYPE & ")"
            gconDMIS.Execute SQL_STATEMENT
            
            'NEW LOG AUDIT-----------------------------------------------------
                Call NEW_LogAudit("AA", "BILLING SYSTEM", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtROno), "REP_OR", "CSMS_REPOR"), "", "CODE : " & Null2String(pDETCDE), "", "")
            'NEW LOG AUDIT-----------------------------------------------------

            rsPutJobToService.MoveNext
        Loop
    End If

    Call CheckRO_Status(RTrim(LTrim(txtROno)))
    post = True
    Exit Function
errordaa:
    post = False
    ERROR_MSG = Error
End Function


Private Sub cmdPrevious_Click()
    rsINFO.MovePrevious
    If rsINFO.BOF Then
        rsINFO.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    rptSubletPo.Reset
    If Function_Access(LOGID, "ACESS_PRINT", "SUBLET PURCHASE") = False Then Exit Sub
        rptSubletPo.Formulas(0) = "Company Name = '" & COMPANY_NAME & "'"
        rptSubletPo.Formulas(1) = "Company Address = '" & COMPANY_ADDRESS & "'"
    If COMPANY_CODE = "HPI" Then
        rptSubletPo.Formulas(3) = "PREPAREDBY = '" & GetSetting("CSMS", "SIGNATORIES", "PO-PREPBY", "") & "'"
        rptSubletPo.Formulas(4) = "NOTEDBY= '" & GetSetting("CSMS", "SIGNATORIES", "PO-NOTEDBY", "") & "'"
        rptSubletPo.Formulas(5) = "APPROVEDBY= '" & GetSetting("CSMS", "SIGNATORIES", "PO-APPROVEDBY", "") & "'"
        rptSubletPo.Formulas(2) = "OWNER= '" & GetSetting("CSMS", "SIGNATORIES", "PO-OWNER", "") & "'"
    Else
        rptSubletPo.Formulas(2) = "G_M = '" & GENERAL_MANAGER & "'"
    End If
    rptSubletPo.ReportTitle = "Purchase Order "
    rptSubletPo.WindowTitle = "Sublet Purchase Order"
    PrintSQLReport rptSubletPo, CSMS_REPORT_PATH & "SubletPO.rpt", "{CSMS_PO_HD.PO_NO} = '" & txtPoNumber & "' and {CSMS_PO_HD.STATUS} <> 'C'", CSMS_REPORT_CONNECTION, 1

    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "SUBLET PURCHASE", "", labid, "", "PO NO: " & txtPoNumber, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
End Sub

Private Sub cmdSave_Click()
    Dim vtxtPoNumber                                    As String
    Dim VTXTRONO                                        As String
    Dim VtxtCustName                                    As String
    Dim vtxtContractorCode                              As String
    Dim vtxtCustAdd                                     As String
    Dim vcboContractor                                  As String
    Dim vtxtContractorAdd                               As String
    Dim vtxtVat                                         As String
    Dim vtxtContactPerson                               As String
    Dim vtxtVatVal                                      As Double
    Dim vtxtVatAmount                                   As Double
    Dim vtxtNetAmount                                   As Double
    Dim vtxtTotalAmount                                 As Double
    Dim vDTPicker1                                      As String
    Dim rsPoDup                                         As New ADODB.Recordset

    If txtPoNumber.Text = "" Then
        MsgBox ("PO number is required"), vbOKOnly + vbCritical, "Invalid PO"
        Exit Sub
    End If

    If txtROno.Text = "" Then
        MsgBox ("RO number is required"), vbOKOnly + vbCritical, "Invalid RO"
        Exit Sub
    End If

    If cboContractor.Text = "" Then
        MsgBox ("Contractor name is required"), vbOKOnly + vbCritical, "Invalid Contractor"
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Set rsPoDup = New ADODB.Recordset
        rsPoDup.Open "Select Po_No from CSMS_Po_Hd where Po_No = '" & txtPoNumber.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPoDup.EOF And Not rsPoDup.BOF Then
            MsgSpeechBox "Purchase Order Number already exist!"
            On Error Resume Next
            txtPoNumber.SetFocus
            Exit Sub
        End If
    Else
        Set rsPoDup = New ADODB.Recordset
        'COMMENT BY  : MJP 11202009 0558PM
        'DESCRIPTION : TCN 12943 STATUS <> 'P' IS NOT CHECK
            'rsPoDup.Open "select Po_No from CSMS_Po_Hd where Po_No = '" & txtPoNumber.Text & "' and Status = 'P' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
'            If Not (rsPoDup.EOF And rsPoDup.BOF) Then
'                MsgSpeechBox "Purchase Order Number already exist!"
'                On Error Resume Next
'                txtPoNumber.SetFocus
'                Exit Sub
'            End If
        'COMMENT BY  : MJP 11202009 0558PM
        
        'UPDATE BY   : MJP 11202009 0558PM
            rsPoDup.Open "select ID, Po_No from CSMS_Po_Hd where " & _
                " Po_No = '" & txtPoNumber.Text & _
                "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not (rsPoDup.EOF And rsPoDup.BOF) Then
                If labid.Caption <> rsPoDup!ID Then
                    MsgBox "Purchase Order Number already exist!", vbInformation, "Info."
                    On Error Resume Next
                    txtPoNumber.SetFocus
                    Exit Sub
                End If
            End If
        'UPDATE BY   : MJP 11202009 0558PM
    End If
    
    
    vtxtPoNumber = N2Str2Null(txtPoNumber.Text)
    VTXTRONO = N2Str2Null(txtROno.Text)
    VtxtCustName = N2Str2Null(txtCustName.Text)
    vtxtCustAdd = N2Str2Null(txtCustAdd.Text)
    vtxtContractorCode = N2Str2Null(txtContractorCode.Text)
    vcboContractor = N2Str2Null(cboContractor.Text)
    vtxtContractorAdd = N2Str2Null(txtContractorAdd.Text)
    vtxtContactPerson = N2Str2Null(txtContactPerson.Text)
    vtxtVatAmount = NumericVal(txtVatAmount.Text)
    vtxtNetAmount = NumericVal(txtNetAmount.Text)
    vtxtTotalAmount = NumericVal(txtTotalAmount.Text)
    vDTPicker1 = N2Date2Null(DTPicker1.Value)

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into CSMS_Po_Hd" & _
            "(Po_No, Ro_No, Po_Date, Cust_Name, Customer_Add, Contractor_Code, Contractor_Name, Contractor_Address, Contact_Person)" & _
            " values(" & vtxtPoNumber & _
            ", " & VTXTRONO & _
            ", " & vDTPicker1 & _
            ", " & VtxtCustName & _
            ", " & vtxtCustAdd & _
            ", " & vtxtContractorCode & _
            ", " & vcboContractor & _
            ", " & vtxtContractorAdd & _
            ", " & vtxtContactPerson & ")"
        gconACCESS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT------------------------------------------------------------------------
            Call NEW_LogAudit("A", "SUBLET PURCHASE", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtPoNumber), "PO_NO", "CSMS_PO_HD"), "", "PO NO: " & txtPoNumber, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------
        Call ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_Po_Hd set " & _
            "Po_No = " & vtxtPoNumber & "," & _
            "Ro_No = " & VTXTRONO & "," & _
            "Po_Date = " & vDTPicker1 & "," & _
            "Cust_Name = " & VtxtCustName & "," & _
            "Customer_Add = " & vtxtCustAdd & "," & _
            "Contractor_code =" & vtxtContractorCode & "," & _
            "Contractor_Name = " & vcboContractor & "," & _
            "Contractor_Address = " & vtxtContractorAdd & "," & _
            "Contact_Person = " & vtxtContactPerson & " " & _
            "where ID = " & labid
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT------------------------------------------------------------------------
            Call NEW_LogAudit("E", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------

        SQL_STATEMENT = "update CSMS_PO_DT set po_no = " & vtxtPoNumber & ",Rep_or = " & VTXTRONO & " where PO_NO = '" & PASTPONUMBER & "'"
        gconDMIS.Execute SQL_STATEMENT
        
        PASTPONUMBER = "0"
        'NEW LOG AUDIT------------------------------------------------------------------------
            Call NEW_LogAudit("EE", "SUBLET PURCHASE", SQL_STATEMENT, labid, "", "PO NO: " & txtPoNumber, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------
        Call ShowSuccessFullyAdded
    End If
    
    Call rsRefresh
    Call FilllstPO_HD
    If AddorEdit = "EDIT" Then
        rsINFO.Find ("ID =" & labid)
    Else
        rsINFO.Find ("PO_NO = '" & txtPoNumber & "'")
    End If
    txtPoNumber.Locked = True
    cmdCancel.Value = True
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "ACESS_UNPOST", "SUBLET PURCHASE") = False Then Exit Sub
    Call UnpostDelete
End Sub

Function EnabledFrame(COND As Boolean)
    Picture1.Enabled = COND
    Frame3.Enabled = COND
    Frame2.Enabled = COND
End Function

Private Sub Command2_Click()
    Call SaveSetting("CSMS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
    Call SaveSetting("CSMS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
    Call SaveSetting("CSMS", "SIGNATORIES", "PO-APPROVEDBY", txtSIG_NotedbyDesign)
    Call SaveSetting("CSMS", "SIGNATORIES", "PO-OWNER", txtowner)
    
    picPrintPOExcel.Visible = False
End Sub

Private Sub Command4_Click()
     picPrintPOExcel.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SUBLET PURCHASE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "SUBLET PURCHASE", "")
        Case vbKeyF8 And Shift = 1:
            cmdPost.Value = True
        
        Case vbKeyF2
            If COMPANY_CODE = "HPI" Then
                txtSIG_PreparedBy = GetSetting("CSMS", "SIGNATORIES", "PO-PREPBY", "")
                txtSIG_Notedby = GetSetting("CSMS", "SIGNATORIES", "PO-NOTEDBY", "")
                txtSIG_NotedbyDesign = GetSetting("CSMS", "SIGNATORIES", "PO-APPROVEDBY", "")
                txtowner = GetSetting("CSMS", "SIGNATORIES", "PO-OWNER", "")
                txtowner.Visible = True
                cmdPrintPOExcel.Enabled = False
                picPrintPOExcel.Visible = True
                picPrintPOExcel.ZOrder 0
            End If
        
        Case vbKeyEscape
            'COMMENT BY  : MJP 08292008 01:06 AM
            'REASON      : IF ACCIDENTALLY CLICK WHILE IN THE MIDDLE OF TRANSACTION, SO IVE COMMENT IT
                'Unload Me
            'COMMENT BY  : MJP 08292008 01:06 AM

        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsINFO!Status) = "C" Then
                    MsgBox "Purchase order already Cancelled. Cannot Add Further Job.", vbInformation, "INFORMATION"
                    Exit Sub

                ElseIf Null2String(rsINFO!Status) = "P" Then
                    MsgBox "Purchase order already posted. Cannot Add Further Job.", vbInformation, "INFORMATION"
                    Exit Sub
                Else
                    frmCSMS_SubletJob.lblAddorEdit.Caption = "ADD"
                    Call EnabledFrame(False)
                    passINFO
                    frmCSMS_SubletJob.Show
                End If
            End If

        Case vbKeyF4

            MsgBox "Please double click the Item to Edit..", vbOKOnly, "INFORMATION"
            '            If lstJobSublet.ListItems.Count = 0 And lstMaterials.ListItems.Count = 0 And lstparts.ListItems.Count = 0 Then
            '                MsgBox "No Job to be edit", vbInformation, "INFORMATION"
            '                Exit Sub
            '            End If
            '            If Picture1.Visible = True Then
            '                If Null2String(rsInfo!Status) = "C" Then
            '                    MsgBox "Purchase order already Cancelled. Cannot EDIT this Job.", vbInformation, "INFORMATION"
            '                     Exit Sub
            '
            '                ElseIf Null2String(rsInfo!Status) = "P" Then
            '                     MsgBox "Purchase order already posted. Cannot EDIT this Job", vbInformation, "INFORMATION"
            '                     Exit Sub
            '                Else
            '                 Call editJobs
            '                 frmCSMS_SubletJob.lblAddorEdit.Caption = "EDIT"
            '                 End If
            '            End If
        Case vbKeyF5
            If Picture1.Visible = True Then
                If Null2String(rsINFO!Status) = "C" Then
                    MsgBox "Purchase order already Cancelled. Cannot DELETE This Job.", vbInformation, "INFORMATION"
                    Exit Sub
                ElseIf Null2String(rsINFO!Status) = "P" Then
                    MsgBox "Purchase order already posted. Cannot DELETE this Job", vbInformation, "INFORMATION"
                    Exit Sub
                Else
                    Call deleteJobs
                End If
            End If

        Case vbKeyF8
            If cmdPost.Enabled = True And Picture1.Visible = True Then
                cmdPost.Value = True
            End If
            

        Case vbKeyF12
            'If cmdUnPost.Enabled = True And Picture1.Visible = True Then
            cmdUnPost.Value = True
            ' End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
    'Me.Show
    
    'JJE Prefix  ** FOR APPROVAL **
'    If COMPANY_CODE = "DJM" Then
'        txtPoNumber.MaxLength = 8
'        txtPoNumber.Enabled = False
'    End If
    'JJE
    
    Call refreshPO_RC_HD
    Call rsRefresh
    Frame1.Enabled = False
    Call FilllstPO_HD
    Call initCboContractor
    Call initMemvars
    Call StoreMemVars
    Call optPONo_Click
    Screen.MousePointer = 0
End Sub

Private Sub lstJobSublet_DblClick()
    'enable to edit job
    If lstJobSublet.ListItems.Count = 0 Then
        MsgBox "No Job to be edit", vbInformation, "INFORMATION"
        Exit Sub
    End If
    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot EDIT this Job.", vbInformation, "INFORMATION"
            Exit Sub
        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot EDIT this Job", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "EDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lstJobSublet_Click()
    If lstJobSublet.ListItems.Count = 0 Then Exit Sub

    xxCODE = (lstJobSublet.SelectedItem.Text)
    labDET.Caption = (lstJobSublet.SelectedItem.SubItems(4))
    lstLivil = lstJobSublet.SelectedItem.SubItems(5)
    lstLine_No = lstJobSublet.SelectedItem.Text
End Sub

Private Sub lstMaterials_DblClick()
    If lstMaterials.ListItems.Count = 0 Then
        MsgBox "No Job to be edit", vbInformation, "INFORMATION"
        Exit Sub
    End If
    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot EDIT this Job.", vbInformation, "INFORMATION"
            Exit Sub
        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot EDIT this Job", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "EDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lstMaterials_Click()
    If lstMaterials.ListItems.Count = 0 Then Exit Sub

    'xxCODE = (lstJobSublet.SelectedItem.Text)
    xxCODE = (lstMaterials.SelectedItem.Text)
    labDET.Caption = lstMaterials.SelectedItem.SubItems(4)
    lstLivil = lstMaterials.SelectedItem.SubItems(5)
    lstLine_No = lstMaterials.SelectedItem.Text
End Sub

Private Sub lstparts_Click()
    If lstparts.ListItems.Count = 0 Then Exit Sub

    'xxCODE = (lstJobSublet.SelectedItem.Text)
    xxCODE = (lstparts.SelectedItem.Text)
    labDET.Caption = lstparts.SelectedItem.SubItems(4)
    lstLivil = lstparts.SelectedItem.SubItems(5)
    lstLine_No = lstparts.SelectedItem.Text
End Sub

Private Sub lstParts_DblClick()
    'enable to edit job
    If lstparts.ListItems.Count = 0 Then
        MsgBox "No Job to be edit", vbInformation, "INFORMATION"
        Exit Sub
    End If
    If Picture1.Visible = True Then
        If Null2String(rsINFO!Status) = "C" Then
            MsgBox "Purchase order already Cancelled. Cannot EDIT this Job.", vbInformation, "INFORMATION"
            Exit Sub
        ElseIf Null2String(rsINFO!Status) = "P" Then
            MsgBox "Purchase order already posted. Cannot EDIT this Job", vbInformation, "INFORMATION"
            Exit Sub
        Else
            frmCSMS_SubletJob.LINE_NO.Caption = lstLine_No
            Call editJobs
            frmCSMS_SubletJob.lblAddorEdit.Caption = "EDIT"
            Call EnabledFrame(False)
            frmCSMS_SubletJob.Show
        End If
    End If
End Sub

Private Sub lvwTran_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsINFO.MoveFirst
    rsINFO.Find ("ID=" & Item.ListSubItems(1).Text)
    Call StoreMemVars
End Sub

Private Sub optPONo_Click()
    Dim rsoptPONo                                      As New ADODB.Recordset
    lvwTran.Enabled = False
    lvwTran.Sorted = False: lvwTran.ListItems.Clear
    Set rsoptPONo = gconDMIS.Execute("Select top 30 Po_No, ID from CSMS_Po_hd order by Po_no desc")
    If Not (rsoptPONo.EOF And rsoptPONo.BOF) Then
        Listview_Loadval Me.lvwTran.ListItems, rsoptPONo
        lvwTran.Refresh
    End If
    lvwTran.Enabled = True
    Set rsoptPONo = Nothing
    
    On Error Resume Next
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub optRONo_Click()
    Dim rsoptRONo                                      As New ADODB.Recordset
    lvwTran.Enabled = False
    lvwTran.Sorted = False: lvwTran.ListItems.Clear
    Set rsoptRONo = gconDMIS.Execute("Select top 30 RO_NO,ID from CSMS_Po_hd order by RO_NO asc")
    If Not (rsoptRONo.EOF And rsoptRONo.BOF) Then
        Listview_Loadval Me.lvwTran.ListItems, rsoptRONo
        lvwTran.Refresh
    End If
    lvwTran.Enabled = True
    Set rsoptRONo = Nothing
    
    On Error Resume Next
    txtSearch.Text = ""
    txtSearch.SetFocus
End Sub

Private Sub Timer1_Timer()
    If lblPOSTED.Caption <> "" Then
        If lblPOSTED.Visible = True Then
            lblPOSTED.Visible = False
        Else
            lblPOSTED.Visible = True
        End If
    End If
End Sub

Private Sub txtROno_Change()
    txtROno.Text = UCase(txtROno.Text)
End Sub

Private Sub txtROno_LostFocus()
    Dim CUSTCODE                                       As String
    Dim rsGetCustomer                                  As ADODB.Recordset
    Dim rsGetName                                      As ADODB.Recordset
    Dim rsCheckRoExist                                 As ADODB.Recordset

    Dim RepairOrder, RepairOrder2, RepairOrder3        As String
    Dim k                                              As Integer
    
    
    If COMPANY_CODE = "CMC" Then
        If Len(txtROno) >= 3 And IsNumeric(Left(txtROno.Text, 2)) = False Then
            txtROno.Text = Mid(txtROno.Text, 3, Len(txtROno.Text) - 2)
        ElseIf IsNumeric(txtROno.Text) = True Then
            txtROno.Text = Format(Right(txtROno.Text, 8), "00000000")
        End If
        txtROno.Text = "R-" + Format(Right(NumericVal(Mid(txtROno.Text, 1, Len(txtROno.Text))), 8), "00000000")
    End If
            
            
    RepairOrder = UCase(txtROno.Text)
    If IsNumeric(RepairOrder) = True Then
        If VALID_COMPANY_CODE_FORHAI = True Then
        Else
'            If FOR_J = True Then
'                RepairOrder = Format(Left(RepairOrder, 1), "J-") & Format(Right(RepairOrder, 8), "00000000")
'            Else
'                RepairOrder = Format(Left(RepairOrder, 1), "R-") & Format(Right(RepairOrder, 8), "00000000")
'            End If
        End If
        txtROno.Text = UCase(RepairOrder)
    Else
        For k = 1 To Len(RepairOrder)
            RepairOrder2 = Mid(RepairOrder, k, 1)
            If IsNumeric(RepairOrder2) = True Then RepairOrder3 = RepairOrder3 + RepairOrder2
        Next
        If VALID_COMPANY_CODE_FORHAI = True Then
        Else
'            If FOR_J = True Then
'                RepairOrder = Format(Left(RepairOrder, 1), "J-") & Format(Right(RepairOrder, 8), "00000000")
'            Else
'                RepairOrder = Format(Left(RepairOrder, 1), "R-") & Format(Right(RepairOrder, 8), "00000000")
'            End If
        End If
        txtROno.Text = UCase(RepairOrder)
    End If

    If checkifro_void(txtROno.Text) = True Then MsgBox "Repair already voided", vbInformation, "Invalid RO": txtROno.SetFocus: Exit Sub
    
    Set rsCheckRoExist = gconDMIS.Execute("Select * from CSMS_RepairOrder where Ro_No ='" & RepairOrder & "'")
    If Not rsCheckRoExist.EOF And Not rsCheckRoExist.BOF Then
        Set rsGetCustomer = New ADODB.Recordset
        Set rsGetCustomer = gconACCESS.Execute("Select ACCT_NO from CSMS_RepairOrder where RO_no ='" & RepairOrder & "'")

        If Not rsGetCustomer.EOF And Not rsGetCustomer.BOF Then
            CUSTCODE = Null2String(rsGetCustomer!ACCT_NO)
            Set rsGetName = New ADODB.Recordset
            Set rsGetName = gconACCESS.Execute(" Select ACCTNAME,CUSTOMERADD from All_Customer_Table where cuscde ='" & CUSTCODE & "'")
            If Not rsGetName.EOF And Not rsGetName.BOF Then
                txtCustName.Text = Null2String(rsGetName!AcctName)
                txtCustAdd.Text = Null2String(rsGetName!CUSTOMERADD)
            End If
        End If
        
        'UPDATE BY   : MJP 011609 1122AM
        'DESCRIPTION : TO NOT ALLOW USER TO ENTER TRANSACTION WHEN THE RO IS ALREADY INVOICE
            Dim rstmp                                          As New ADODB.Recordset
            Set rstmp = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE REP_OR = " & N2Str2Null(txtROno) & "")
            If Not (rstmp.BOF And rstmp.EOF) Then
                If Not Null2String(rstmp!INVOICE) = "" Then
                    MsgBox "RO no. " & txtROno & " is already invoiced, this transaction cannot proceed", vbExclamation, "DMIS"
                    txtROno.SetFocus
                    Exit Sub
                End If
            End If
            Set rstmp = Nothing
        'UPDATE BY  : MJP 011609 1122AM
    Else
        If txtROno.Text = "" Then
            Exit Sub
        Else
            txtROno.Text = ""
            MsgBox ("Repair order does not exist"), vbOKOnly + vbCritical, "Invalid RO"
            txtROno.SetFocus
            Exit Sub
        End If
    End If
    txtROno.Text = UCase(txtROno.Text)
    
    Set rsGetCustomer = Nothing
    Set rsGetName = Nothing
    Set rsCheckRoExist = Nothing
End Sub

Private Sub txtSearch_Change()
    Dim rsSearch                                        As New ADODB.Recordset
    Dim PONUMBER                                        As String
    Dim poNUMBER2                                       As String
    Dim poNUMBER3                                       As String
    Dim RepairOrder                                     As String
    Dim RepairOrder2                                    As String
    Dim RepairOrder3                                    As String
    Dim k                                               As Integer


    If optPONo.Value = True Then
        PONUMBER = UCase(txtSearch.Text)
        If txtSearch = "" Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 PO_NO ,ID from CSMS_Po_hd order by PO_NO desc ")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        ElseIf PONUMBER <> "" Then
'            If IsNumeric(PONUMBER) = True Then
'                PONUMBER = Format(Right(PONUMBER, 6), "000000")
'            Else
'                For k = 1 To Len(PONUMBER)
'                    poNUMBER2 = Mid(PONUMBER, k, 1)
'                    If IsNumeric(poNUMBER2) = True Then poNUMBER3 = poNUMBER3 + poNUMBER2
'                Next
'                PONUMBER = Format(poNUMBER3, "000000")
'            End If
        End If
'        If IsNumeric(PONUMBER) = True Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select top 30 PO_NO ,ID from CSMS_Po_hd where Po_No like '" & PONUMBER & "%'")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
'        End If
    Else
        If txtSearch = "" Then
            lvwTran.Enabled = False
            lvwTran.Sorted = False: lvwTran.ListItems.Clear
            Set rsSearch = gconDMIS.Execute("Select TOP 30 RO_NO ,ID from CSMS_Po_hd order by RO_NO asc ")
            If Not (rsSearch.EOF And rsSearch.BOF) Then
                Listview_Loadval Me.lvwTran.ListItems, rsSearch
                lvwTran.Refresh
            End If
            lvwTran.Enabled = True
        Else
            RepairOrder = UCase(txtSearch.Text)
'            If RepairOrder <> "" Then
'                If IsNumeric(RepairOrder) = True Then
'                    RepairOrder = Format(Left(RepairOrder, 1), "R-") & Format(Right(RepairOrder, 8), "00000000")
'                Else
'                    For k = 1 To Len(RepairOrder)
'                        RepairOrder2 = Mid(RepairOrder, k, 1)
'                        If IsNumeric(RepairOrder2) = True Then RepairOrder3 = RepairOrder3 + RepairOrder2
'                    Next
'                    RepairOrder3 = Format(RepairOrder3, "00000000"): RepairOrder = Format(Left(RepairOrder3, 1), "R-") & Format(Right(RepairOrder3, 8), "00000000")
'                End If
'            End If
'            If Left(RepairOrder, 2) = "R-" Then
                lvwTran.Enabled = False
                lvwTran.Sorted = False: lvwTran.ListItems.Clear
                Set rsSearch = gconDMIS.Execute("Select TOP 30 RO_No , ID from CSMS_Po_hd where RO_NO like '" & RepairOrder & "%'")
                If Not (rsSearch.EOF And rsSearch.BOF) Then
                    Listview_Loadval Me.lvwTran.ListItems, rsSearch
                    lvwTran.Refresh
                End If
'            End If
        End If
        lvwTran.Enabled = True
    End If
    Set rsSearch = Nothing
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub

Function CheckRO_Status(XXX As String)
    Dim RS                                             As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT DONE  FROM CSMS_RO_DET WHERE " & _
        " LIVIL = '1' " & _
        " AND (DONE = 'N' OR DONE = 'W' OR DONE IS NULL) " & _
        " and REP_OR = '" & XXX & "'")
    If RS.EOF And RS.BOF Then
        gconDMIS.Execute "Update CSMS_RepairOrder set dateFinish = '" & LOGDATE & "', JSTATUS ='F', STATUS ='Finish Job' where RO_NO ='" & XXX & "'"
    Else
        'Retain Status
    End If
    Set RS = Nothing
End Function



