VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCSMS_INQUIRY_JobEstimate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listing By Job Estimate"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_INQUIRY_JobEstimate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   8520
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   30
      ScaleHeight     =   5265
      ScaleWidth      =   8415
      TabIndex        =   7
      Top             =   1410
      Width           =   8445
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   30
         ScaleHeight     =   2865
         ScaleWidth      =   8325
         TabIndex        =   8
         Top             =   30
         Width           =   8355
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   8325
            _Version        =   655364
            _ExtentX        =   14684
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "PERSONAL INFORMATION"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   1410
            Width           =   8325
            _Version        =   655364
            _ExtentX        =   14684
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "OTHER INFORMATION"
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
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "km Rdg. :"
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
            Index           =   10
            Left            =   5430
            TabIndex        =   31
            Top             =   2190
            Width           =   660
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Record By :"
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
            Index           =   9
            Left            =   5220
            TabIndex        =   30
            Top             =   2550
            Width           =   855
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   6270
            TabIndex        =   29
            Top             =   2100
            Width           =   1905
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   6270
            TabIndex        =   28
            Top             =   2460
            Width           =   1905
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Recorded :"
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
            Index           =   8
            Left            =   645
            TabIndex        =   27
            Top             =   2190
            Width           =   1170
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plate no. :"
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
            Index           =   7
            Left            =   1110
            TabIndex        =   26
            Top             =   2550
            Width           =   705
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Model :"
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
            Index           =   6
            Left            =   720
            TabIndex        =   25
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   1890
            TabIndex        =   24
            Top             =   2100
            Width           =   1905
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   1890
            TabIndex        =   23
            Top             =   2460
            Width           =   1905
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   1920
            TabIndex        =   15
            Top             =   1020
            Width           =   4185
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   14
            Top             =   660
            Width           =   5655
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1920
            TabIndex        =   13
            Top             =   300
            Width           =   5655
         End
         Begin VB.Label lblCap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No. :"
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
            Index           =   4
            Left            =   885
            TabIndex        =   11
            Top             =   1080
            Width           =   930
         End
         Begin VB.Label lblCap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CustomerAddress :"
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
            Index           =   3
            Left            =   405
            TabIndex        =   10
            Top             =   750
            Width           =   1410
         End
         Begin VB.Label lblCap 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name :"
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
            Index           =   2
            Left            =   585
            TabIndex        =   9
            Top             =   390
            Width           =   1230
         End
         Begin VB.Label lblInfo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1890
            TabIndex        =   22
            Top             =   1710
            Width           =   6285
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2220
         Left            =   30
         TabIndex        =   12
         Top             =   3000
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   3916
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   1058
         BackColor       =   14606302
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "JOBS"
         TabPicture(0)   =   "frmCSMS_INQUIRY_JobEstimate.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lsvJobs"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "PARTS"
         TabPicture(1)   =   "frmCSMS_INQUIRY_JobEstimate.frx":05A6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lsvParts"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "MATERIALS"
         TabPicture(2)   =   "frmCSMS_INQUIRY_JobEstimate.frx":05C2
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lsvMat"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "ACCESSORIES"
         TabPicture(3)   =   "frmCSMS_INQUIRY_JobEstimate.frx":05DE
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lsvAcc"
         Tab(3).ControlCount=   1
         Begin MSComctlLib.ListView lsvJobs 
            Height          =   1425
            Left            =   90
            TabIndex        =   16
            Top             =   690
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2514
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "no."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lsvParts 
            Height          =   1425
            Left            =   -74910
            TabIndex        =   17
            Top             =   690
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2514
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lsvMat 
            Height          =   1425
            Left            =   -74910
            TabIndex        =   18
            Top             =   690
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2514
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2646
            EndProperty
         End
         Begin MSComctlLib.ListView lsvAcc 
            Height          =   1425
            Left            =   -74910
            TabIndex        =   19
            Top             =   690
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   2514
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No."
               Object.Width           =   882
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Code"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Description"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Amount"
               Object.Width           =   2646
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   30
      ScaleHeight     =   1275
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   30
      Width           =   8445
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7170
         TabIndex        =   20
         Top             =   750
         Width           =   1185
      End
      Begin VB.ComboBox cboENO 
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
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   750
         Width           =   2295
      End
      Begin VB.ComboBox cboPlateNo 
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
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   2745
      End
      Begin VB.ComboBox cboEstNo 
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
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2745
      End
      Begin VB.OptionButton optPlateNo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "By Plate no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   870
         Width           =   1605
      End
      Begin VB.OptionButton optEstNo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "By Estimate no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   2
         Top             =   420
         Width           =   1875
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   8415
         _Version        =   655364
         _ExtentX        =   14843
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "SELECT SEARCH OPTION"
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
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SELECT SEARCH OPTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   30
         Width           =   2115
      End
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
      Height          =   795
      Left            =   7740
      MouseIcon       =   "frmCSMS_INQUIRY_JobEstimate.frx":05FA
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_INQUIRY_JobEstimate.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exit Window"
      Top             =   6510
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMS_INQUIRY_JobEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function FindSAName(SACODE As String)
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("SELECT * FROM CSMS_VW_EMPNO WHERE CODE = '" & SACODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        FindSAName = Null2String(RSTMP!NAYM)
    Else
        FindSAName = ""
    End If
    Set RSTMP = Nothing
End Function

Sub FindCustomerInfo(CUSTCODE As String)
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * From All_Customer_Table Where CusCde = '" & CUSTCODE & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        lblinfo(0).Caption = Null2String(RSTMP!lastname) & ", " & Null2String(RSTMP!Firstname)
        lblinfo(1).Caption = Null2String(RSTMP!CUSTOMERADD)
        lblinfo(2).Caption = Null2String(RSTMP!HomePhone) & "/ " & Null2String(RSTMP!TelephoneNo)
    End If

    Set RSTMP = Nothing
End Sub

Sub DisplayPLateEstimate()
    Dim RSTMP                                          As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select EstimateNo from CSMS_EstHD Where Plate_no = '" & cboPlateNo.Text & "'")
    cboENO.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            cboENO.AddItem Null2String(RSTMP!EstimateNo)
            RSTMP.MoveNext
        Loop
        cboENO.ListIndex = 0
    End If

    Set RSTMP = Nothing
End Sub

Sub DisplayInformation(vESTNO As String, vLIVIL As String, vLSV As ListView, vMAINTABLE As String)
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ITEM                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * from " & vMAINTABLE & " Where EstimateNo = '" & vESTNO & "' And Livil = '" & vLIVIL & "' Order By Line_No asc")
    vLSV.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = vLSV.ListItems.Add(, , Null2String(RSTMP!LINE_NO))
            ITEM.SubItems(1) = Null2String(RSTMP!DETCDE)
            ITEM.SubItems(2) = Null2String(RSTMP!DETDSC)
            ITEM.SubItems(3) = Format(Null2String(RSTMP!DETAMT), "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cboPlateNo_Change()
    DisplayPLateEstimate
End Sub

Private Sub cboPlateNo_Click()
    DisplayPLateEstimate
End Sub

Private Sub cboPlateNo_LostFocus()
    DisplayPLateEstimate
End Sub

Private Sub cmdDisplay_Click()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim vESTNO                                         As String
    Dim RSTMP1                                         As New ADODB.Recordset

    If optEstNo.Value = True Then
        If cboEstNo.Text = "" Then Exit Sub

        vESTNO = cboEstNo.Text
        Set RSTMP = gconDMIS.Execute("Select * From CSMS_ESTHD Where EstimateNo = '" & cboEstNo.Text & "'")
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Call FindCustomerInfo(RSTMP!ACCT_NO)
        End If

        GoTo VIEW_VEHICLE_RECORD
    Else
        If cboPlateNo.Text = "" Then Exit Sub
        lblinfo(4).Caption = cboPlateNo.Text

VIEW_VEHICLE_RECORD:

        If optEstNo.Value = True Then
            Set RSTMP = gconDMIS.Execute("Select * From CSMS_REPOR Where EstimateNo = '" & cboEstNo.Text & "'")
            vESTNO = cboEstNo.Text
        Else
            Set RSTMP = gconDMIS.Execute("Select * From CSMS_REPOR Where EstimateNo = '" & cboENO.Text & "'")
            vESTNO = cboENO.Text
        End If
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Call FindCustomerInfo(RSTMP!ACCT_NO)
            lblinfo(3).Caption = Null2String(RSTMP!MODEL)
            lblinfo(5).Caption = Null2String(RSTMP!DTE_RECD)
            lblinfo(4).Caption = Null2String(RSTMP!PLATE_NO)
            lblinfo(7).Caption = Null2String(RSTMP!km_rdg)
            lblinfo(6).Caption = FindSAName(RSTMP!RECD_BY)
        End If
    End If

    Call DisplayInformation(vESTNO, "1", lsvJobs, "CSMS_RO_DET")
    Call DisplayInformation(vESTNO, "2", lsvParts, "CSMS_ESTDetails")
    Call DisplayInformation(vESTNO, "3", lsvMat, "CSMS_ESTDetails")
    Call DisplayInformation(vESTNO, "4", lsvAcc, "CSMS_ESTDetails")

    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("I", "JOB ESTIMATE LISTING", "", "", "", "EST NO: " & vESTNO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Set RSTMP = Nothing
End Sub

Private Sub cmdPARTSINQUIRYExit_Click()
    Unload Me
End Sub

Private Sub cmdExit_Click()
    '    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub

            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (JOB ESTIMATE LISTING)"
            Call frmALL_AuditInquiry.DisplayHistory("", "JOB ESTIMATE LISTING", "PRINTING")
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    optEstNo.Value = True
End Sub

Private Sub optEstNo_Click()
    On Error Resume Next
    Dim RSTMP                                          As New ADODB.Recordset

    If optEstNo.Value = True Then
        cboEstNo.Visible = True
        cboPlateNo.Visible = False
        cboENO.Visible = False

        Set RSTMP = gconDMIS.Execute("Select EstimateNo From CSMS_ESthd Order By EstimateNO ASC ")
        cboEstNo.Clear
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                cboEstNo.AddItem Null2String(RSTMP!EstimateNo)

                RSTMP.MoveNext
            Loop
            cboEstNo.ListIndex = 0
        End If
        cboEstNo.SetFocus
    End If

    Set RSTMP = Nothing
End Sub

Private Sub optPlateNo_Click()
    Dim RSTMP                                          As New ADODB.Recordset

    If optPlateNo.Value = True Then
        cboEstNo.Visible = False
        cboPlateNo.Visible = True
        cboENO.Visible = True

        Set RSTMP = gconDMIS.Execute("Select Distinct Plate_no From CSMS_EStHD Order By Plate_no ASC")
        cboPlateNo.Clear
        If Not (RSTMP.BOF And RSTMP.EOF) Then
            Do While Not RSTMP.EOF
                cboPlateNo.AddItem Null2String(RSTMP!PLATE_NO)
                RSTMP.MoveNext
            Loop
            cboPlateNo.ListIndex = 0
        End If
        cboPlateNo.SetFocus
    End If

    Set RSTMP = Nothing
End Sub

