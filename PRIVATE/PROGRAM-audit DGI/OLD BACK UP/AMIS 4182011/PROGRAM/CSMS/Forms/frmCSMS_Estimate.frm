VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMSEstimateEntry_NEO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Estimate"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_Estimate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   10320
   Begin VB.PictureBox picDETAILS 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   30
      ScaleHeight     =   3405
      ScaleWidth      =   10185
      TabIndex        =   46
      Top             =   3210
      Width           =   10215
      Begin TabDlg.SSTab SSTab1 
         Height          =   3030
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   10185
         _ExtentX        =   17965
         _ExtentY        =   5345
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   1058
         BackColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Details"
         TabPicture(0)   =   "frmCSMS_Estimate.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lsvDET"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "F3-Jobs"
         TabPicture(1)   =   "frmCSMS_Estimate.frx":08AC
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lsvJOBS"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "F4-Parts"
         TabPicture(2)   =   "frmCSMS_Estimate.frx":0BC6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lsvPARTS"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "F5-Materials"
         TabPicture(3)   =   "frmCSMS_Estimate.frx":0EE0
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lsvMAT"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "F6-Accessories"
         TabPicture(4)   =   "frmCSMS_Estimate.frx":11FA
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "lsvACC"
         Tab(4).ControlCount=   1
         Begin MSComctlLib.ListView lsvDET 
            Height          =   2235
            Left            =   60
            TabIndex        =   49
            Top             =   690
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.ListView lsvJOBS 
            Height          =   2235
            Left            =   -74940
            TabIndex        =   50
            Top             =   690
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LINE NO."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "JOB CODE"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "JOB DESCRIPTION"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "AMOUNT"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "WCS"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "DISCOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lsvPARTS 
            Height          =   2235
            Left            =   -74940
            TabIndex        =   51
            Top             =   690
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LINE NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PART NO."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPTION"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "QTY"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "AMOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "WSC"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "DISCOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lsvMAT 
            Height          =   2235
            Left            =   -74940
            TabIndex        =   57
            Top             =   690
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LINE NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PART NO."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPTION"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "QTY"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "AMOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "WSC"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "DISCOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lsvACC 
            Height          =   2235
            Left            =   -74940
            TabIndex        =   58
            Top             =   690
            Width           =   10035
            _ExtentX        =   17701
            _ExtentY        =   3942
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "LINE NO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PART NO."
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "DESCRIPTION"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "QTY"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "AMOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "WSC"
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "DISCOUNT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ID"
               Object.Width           =   0
            EndProperty
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   10245
         _Version        =   655364
         _ExtentX        =   18071
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "ESTIMATE DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   30
      ScaleHeight     =   3045
      ScaleWidth      =   10185
      TabIndex        =   13
      Top             =   60
      Width           =   10215
      Begin VB.TextBox txtDesc 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2610
         Width           =   5865
      End
      Begin VB.TextBox txtParticipat 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   29
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox txtKm_rdg 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   9
         TabIndex        =   28
         Top             =   1470
         Width           =   1815
      End
      Begin VB.TextBox txtAcct_No 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4770
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   27
         Top             =   390
         Width           =   1035
      End
      Begin VB.TextBox txtPlate_No 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         MaxLength       =   8
         TabIndex        =   26
         Top             =   1830
         Width           =   1815
      End
      Begin VB.TextBox txtNiym 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   390
         Width           =   3945
      End
      Begin VB.TextBox txtestNo 
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
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1230
         TabIndex        =   24
         Top             =   450
         Width           =   1815
      End
      Begin VB.ComboBox cboRecd_by 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   4800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1500
         Width           =   2295
      End
      Begin VB.ComboBox cboModel 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1230
         Locked          =   -1  'True
         Sorted          =   -1  'True
         TabIndex        =   22
         Text            =   "cboModel"
         Top             =   2220
         Width           =   1845
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   8730
         TabIndex        =   21
         Top             =   1080
         Width           =   1425
      End
      Begin VB.CheckBox chkParticipat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   4470
         TabIndex        =   20
         Top             =   750
         Width           =   225
      End
      Begin VB.TextBox txtDte_recd 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1860
         Width           =   2295
      End
      Begin VB.TextBox txtDte_comp 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2220
         Width           =   2295
      End
      Begin VB.TextBox txtVIN 
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
         ForeColor       =   &H00A00000&
         Height          =   360
         Left            =   7200
         MaxLength       =   35
         TabIndex        =   17
         Top             =   1890
         Width           =   2925
      End
      Begin VB.TextBox txtParticipation 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   750
         Width           =   3945
      End
      Begin VB.CommandButton cmdCust 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   15
         Top             =   390
         Width           =   285
      End
      Begin VB.CommandButton cmdPart 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5820
         TabIndex        =   14
         Top             =   750
         Width           =   285
      End
      Begin VB.TextBox txtAddress 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1230
         TabIndex        =   31
         Top             =   1110
         Width           =   8895
      End
      Begin VB.Label labID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
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
         Height          =   255
         Left            =   3270
         TabIndex        =   56
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   8205
         TabIndex        =   55
         Top             =   30
         Width           =   1845
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   30
         Width           =   10245
         _Version        =   655364
         _ExtentX        =   18071
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "ESTIMATE INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   44
         Top             =   2670
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   660
         TabIndex        =   43
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Vin No"
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
         Height          =   240
         Left            =   7230
         TabIndex        =   42
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Participation"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   3330
         TabIndex        =   41
         Top             =   750
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Uploaded"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3450
         TabIndex        =   40
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Estimated"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3450
         TabIndex        =   39
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Advisor"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3450
         TabIndex        =   38
         Top             =   1560
         Width           =   1245
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "KM Reading"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   37
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   450
         TabIndex        =   36
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   450
         TabIndex        =   35
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Estimate No."
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   9
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   3330
         TabIndex        =   33
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label labPrevID 
         Caption         =   "Label59"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2670
         TabIndex        =   32
         Top             =   420
         Width           =   345
      End
   End
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
      Height          =   855
      Left            =   90
      ScaleHeight     =   855
      ScaleWidth      =   10215
      TabIndex        =   0
      Top             =   6690
      Width           =   10215
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         CausesValidation=   0   'False
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
         Left            =   9480
         MouseIcon       =   "frmCSMS_Estimate.frx":39AC
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":3AFE
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   8790
         MouseIcon       =   "frmCSMS_Estimate.frx":3E64
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":3FB6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print this Record"
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
         Left            =   8100
         MouseIcon       =   "frmCSMS_Estimate.frx":431C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":446E
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
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
         Left            =   7410
         MouseIcon       =   "frmCSMS_Estimate.frx":47CA
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":491C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Left            =   6720
         MouseIcon       =   "frmCSMS_Estimate.frx":4C2F
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":4D81
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   705
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
         Left            =   6030
         MouseIcon       =   "frmCSMS_Estimate.frx":50D1
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":5223
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move to First Record"
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
         Left            =   5340
         MouseIcon       =   "frmCSMS_Estimate.frx":5581
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":56D3
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   4650
         MouseIcon       =   "frmCSMS_Estimate.frx":59CD
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":5B1F
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "Pre&v"
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
         Left            =   3960
         MouseIcon       =   "frmCSMS_Estimate.frx":5E77
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":5FC9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
      Begin Crystal.CrystalReport rptRepairOrder 
         Left            =   2910
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Repair Order Print Out"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - Participation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   30
         TabIndex        =   54
         Top             =   30
         Width           =   1425
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F7 - Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   30
         TabIndex        =   53
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Customer Vehicle Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   30
         TabIndex        =   52
         Top             =   570
         Width           =   2325
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
      Height          =   885
      Left            =   8685
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   10
      Top             =   6750
      Visible         =   0   'False
      Width           =   1590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         Left            =   870
         MouseIcon       =   "frmCSMS_Estimate.frx":6328
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":647A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   0
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
         Left            =   180
         MouseIcon       =   "frmCSMS_Estimate.frx":67B8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_Estimate.frx":690A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMSEstimateEntry_NEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEST                                              As ADODB.Recordset

Function FindCustomerAddress(vCUSCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT CUSTOMERADD FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & vCUSCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindCustomerAddress = Null2String(RSTMP!CUSTOMERADD)
    Else
        FindCustomerAddress = ""
    End If

    Set RSTMP = Nothing
End Function

Function FindSA(SCODE As String) As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT NAYM FROM CSMS_VW_EMPNO WHERE CODE = '" & SCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindSA = Null2String(RSTMP!NAYM)
    Else
        FindSA = "SA NOT FOUND"
    End If
    Set RSTMP = Nothing
End Function

Sub FillSA()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO ORDER BY NAYM")
    cboRecd_by.Clear
    cboRecd_by.AddItem "SA NOT FOUND"
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboRecd_by.AddItem LTrim(RTrim(Null2String(RSTMP!NAYM)))
            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FindVehicleInformation(vPLATENO As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = '" & vPLATENO & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtDesc.Text = Null2String(RSTMP!Description)
        txtVIN.Text = UCase(Null2String(RSTMP!Vin))
    Else
        txtDesc.Text = ""
        txtVIN.Text = ""
    End If
    Set RSTMP = Nothing
End Sub

Sub StoreMemVars()
    If Not (rsEST.BOF And rsEST.EOF) Then
        txtestNo.Text = Null2String(rsEST!EstimateNo)
        txtAcct_No.Text = Null2String(rsEST!ACCT_NO)
        txtNiym.Text = Null2String(rsEST!NIYM)
        txtAddress.Text = FindCustomerAddress(Null2String(rsEST!ACCT_NO))

        If rsEST!PARTICIPATION = 0 Then chkParticipat.Value = 0
        If rsEST!PARTICIPATION = 1 Then chkParticipat.Value = 1
        txtParticipat.Text = Null2String(rsEST!INS_CODE)
        txtParticipation.Text = Null2String(rsEST!INS_NAME)

        txtKm_rdg.Text = Null2String(rsEST!KM_READ)
        txtPlate_No.Text = Null2String(rsEST!PLATE_NO)
        Call FindVehicleInformation(Null2String(rsEST!PLATE_NO))

        cboRecd_by.Text = FindSA(Null2String(rsEST!SA))
        txtDte_recd.Text = Null2String(rsEST!DATE_EST)
        txtDte_comp.Text = Null2String(rsEST!DATE_UPL)

        'Call FillJobs
    Else
        ShowNoRecord
    End If
End Sub

Sub rsRefresh()
    Set rsEST = New ADODB.Recordset
    Set rsEST = gconDMIS.Execute("select * from CSMS_EstHd order by id asc")
End Sub

Sub FillDetails()

End Sub

Sub FillJobs(xESTNO As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_ESTDETAILS WHERE ESTIMATENO = '" & xESTNO & "'")
    lsvDET.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = Nothing
End Sub

Sub FillDet()

End Sub

Sub FillModel()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT MODEL FROM CSMS_MODELs ORDER BY MODEL")
    cboModel.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboModel.AddItem LTrim(RTrim(Null2String(RSTMP!MODEL)))

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdAdd_Click()
    '
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    rsEST.MoveFirst
    Call StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsEST.MoveLast
    Call StoreMemVars
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsEST.MoveNext
    If rsEST.EOF Then
        rsEST.MoveLast
        ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsEST.MovePrevious
    If rsEST.BOF Then
        rsEST.MoveFirst
        ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call FillModel
    Call FillSA

    Call rsRefresh
    Call StoreMemVars
End Sub

