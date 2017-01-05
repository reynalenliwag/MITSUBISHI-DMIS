VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAMISFILESChartOfAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chart of Account"
   ClientHeight    =   7170
   ClientLeft      =   4230
   ClientTop       =   3165
   ClientWidth     =   7425
   ForeColor       =   &H00F5F5F5&
   Icon            =   "ChartOfAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   7425
   Tag             =   "CHART OF ACCOUNTS"
   Begin VB.CheckBox chkSchedule 
      Caption         =   "Is Schedule Accounts ?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   3600
      TabIndex        =   47
      Top             =   930
      Width           =   3645
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   7305
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   930
         Width           =   1935
      End
      Begin VB.TextBox txtDescription 
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
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   540
         Width           =   5685
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   1560
         MaxLength       =   16
         TabIndex        =   2
         Text            =   "XX-XXXXX-XX"
         Top             =   150
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   6
         Top             =   600
         Width           =   465
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3690
         TabIndex        =   5
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         Height          =   255
         Left            =   30
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport rptChartOfAccounts 
      Left            =   6870
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Chart of Accounts"
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
   Begin VB.PictureBox picWizard 
      Height          =   4305
      Left            =   135
      ScaleHeight     =   4245
      ScaleWidth      =   7125
      TabIndex        =   13
      Top             =   1890
      Width           =   7185
      Begin VB.CommandButton cmdOkDepartment 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3645
         TabIndex        =   30
         Top             =   2325
         Width           =   615
      End
      Begin VB.CommandButton cmdOkTitleCode 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3675
         TabIndex        =   19
         Top             =   975
         Width           =   615
      End
      Begin VB.CommandButton cmdOkSubHeader 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3645
         TabIndex        =   17
         Top             =   510
         Width           =   615
      End
      Begin VB.CommandButton cmdOkHeader 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3645
         TabIndex        =   15
         Top             =   75
         Width           =   615
      End
      Begin VB.ComboBox cboAccountType 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2760
         Width           =   2715
      End
      Begin VB.TextBox txtAccountName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Left            =   60
         MaxLength       =   100
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   3570
         Width           =   7005
      End
      Begin VB.PictureBox picAccount 
         BackColor       =   &H00EBFAFA&
         Height          =   855
         Left            =   30
         ScaleHeight     =   795
         ScaleWidth      =   4215
         TabIndex        =   20
         Top             =   1380
         Width           =   4275
         Begin VB.TextBox txtCode6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   3480
            MaxLength       =   2
            TabIndex        =   27
            Text            =   "XX"
            Top             =   90
            Width           =   675
         End
         Begin VB.TextBox txtCode4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   2020
            MaxLength       =   1
            TabIndex        =   25
            Text            =   "X"
            Top             =   90
            Width           =   450
         End
         Begin VB.TextBox txtCode5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   2505
            MaxLength       =   2
            TabIndex        =   26
            Text            =   "XX"
            Top             =   90
            Width           =   675
         End
         Begin VB.TextBox txtCode2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   600
            Left            =   540
            MaxLength       =   1
            TabIndex        =   22
            Text            =   "X"
            Top             =   90
            Width           =   450
         End
         Begin VB.TextBox txtCode1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   600
            Left            =   60
            MaxLength       =   1
            TabIndex        =   21
            Text            =   "X"
            Top             =   90
            Width           =   450
         End
         Begin VB.TextBox txtCode3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   630
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   24
            Text            =   "XX"
            Top             =   90
            Width           =   675
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   3240
            TabIndex        =   28
            Top             =   150
            Width           =   255
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1050
            TabIndex        =   23
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.ComboBox cboTitleCode 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   960
         Width           =   3585
      End
      Begin VB.ComboBox cboSubHeader 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   510
         Width           =   3585
      End
      Begin VB.ComboBox cboHeader 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   60
         Width           =   3585
      End
      Begin VB.ComboBox cboDepartment 
         Appearance      =   0  'Flat
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
         Height          =   360
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2310
         Width           =   3585
      End
      Begin MSComctlLib.ListView lstRelatedAccounts 
         Height          =   3195
         Left            =   4350
         TabIndex        =   33
         ToolTipText     =   "Related Accounts"
         Top             =   30
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ChartOfAccount.frx":08CA
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TYPE"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         Height          =   255
         Left            =   -30
         TabIndex        =   31
         Top             =   2790
         Width           =   1485
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name"
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
         Height          =   255
         Left            =   45
         TabIndex        =   34
         Top             =   3240
         Width           =   1995
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5025
      Left            =   30
      TabIndex        =   9
      Top             =   1260
      Width           =   7335
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4425
         Left            =   60
         TabIndex        =   11
         ToolTipText     =   "Press F1 - View Scheduled Accounts Only, F2 - View Non Scheduled Accounts, F3 - All Accounts"
         Top             =   540
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   7805
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
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
         MouseIcon       =   "ChartOfAccount.frx":0A2C
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   7585
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TYPE"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Schedule"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1650
         TabIndex        =   49
         Top             =   210
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   210
         Value           =   -1  'True
         Width           =   1425
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
         Left            =   3240
         MaxLength       =   35
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   150
         Width           =   3945
      End
   End
   Begin wizButton.cmd cmdWizard 
      Height          =   4410
      Left            =   30
      TabIndex        =   12
      Top             =   1800
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   7779
      TX              =   "Ok"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "ChartOfAccount.frx":0B8E
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -75
      ScaleHeight     =   900
      ScaleWidth      =   8415
      TabIndex        =   36
      Top             =   6300
      Width           =   8415
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
         Left            =   6570
         MouseIcon       =   "ChartOfAccount.frx":0BAA
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":0CFC
         Style           =   1  'Graphical
         TabIndex        =   44
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
         Left            =   5880
         MouseIcon       =   "ChartOfAccount.frx":1062
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   43
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
         Left            =   5190
         MouseIcon       =   "ChartOfAccount.frx":151A
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":166C
         Style           =   1  'Graphical
         TabIndex        =   42
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
         Left            =   4500
         MouseIcon       =   "ChartOfAccount.frx":1997
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":1AE9
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Left            =   3810
         MouseIcon       =   "ChartOfAccount.frx":1E45
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":1F97
         Style           =   1  'Graphical
         TabIndex        =   40
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
         Left            =   3090
         MouseIcon       =   "ChartOfAccount.frx":22AA
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":23FC
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   735
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
         Left            =   2370
         MouseIcon       =   "ChartOfAccount.frx":274C
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":289E
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Move to First Record"
         Top             =   30
         Width           =   735
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
         Left            =   1680
         MouseIcon       =   "ChartOfAccount.frx":2BFC
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":2D4E
         Style           =   1  'Graphical
         TabIndex        =   39
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
         Left            =   990
         MouseIcon       =   "ChartOfAccount.frx":3048
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":319A
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Left            =   300
         MouseIcon       =   "ChartOfAccount.frx":34F2
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":3644
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5700
      ScaleHeight     =   885
      ScaleWidth      =   2340
      TabIndex        =   50
      Top             =   6270
      Width           =   2340
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
         Left            =   780
         MouseIcon       =   "ChartOfAccount.frx":39A3
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":3AF5
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Cancel"
         Top             =   60
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
         Left            =   90
         MouseIcon       =   "ChartOfAccount.frx":3E33
         MousePointer    =   99  'Custom
         Picture         =   "ChartOfAccount.frx":3F85
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Save Entry"
         Top             =   60
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmAMISFILESChartOfAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsChartAccount                                As ADODB.Recordset
Dim rsHeader                                      As ADODB.Recordset
Dim rsSubHeader                                   As ADODB.Recordset
Dim rsTitleCode                                   As ADODB.Recordset
Dim rsDepartment                                  As ADODB.Recordset
Dim rsAccType                                     As ADODB.Recordset
Dim AddorEdit                                     As String
Dim PREV_ACCT_CODE                                As String

Function SetAccType(Acc As String) As String
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select * from AMIS_Acctype where code = " & N2Str2Null(Acc))
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        SetAccType = Null2String(rsAccType!Description)
    Else
        SetAccType = "Not Defined"
    End If
End Function

Function SetAccCode(Acc As String) As String
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select * from AMIS_Acctype where description = " & N2Str2Null(Acc))
    If Not rsAccType.EOF And Not rsAccType.BOF Then
        SetAccCode = Null2String(rsAccType!Code)
    Else
        SetAccCode = ""
    End If
End Function

Function SetHeaderCode(Acc As String) As String
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("select * from AMIS_Header where description = " & N2Str2Null(Acc))
    If Not rsHeader.EOF And Not rsHeader.BOF Then
        SetHeaderCode = Null2String(rsHeader!Code)
    End If
End Function

Function SetHeaderDesc(Acc As String) As String
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("select * from AMIS_Header where code = " & N2Str2Null(Acc))
    If Not rsHeader.EOF And Not rsHeader.BOF Then
        SetHeaderDesc = Null2String(rsHeader!Description)
    End If
End Function

Function SetSubHeaderCode(Acc As String) As String
    Set rsSubHeader = New ADODB.Recordset
    Set rsSubHeader = gconDMIS.Execute("select * from AMIS_SubHeader where description = " & N2Str2Null(Acc))
    If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then
        SetSubHeaderCode = Null2String(rsSubHeader!SubHeaderCode)
    End If
End Function

Function SetSubHeaderDesc(Acc As String) As String
    Set rsSubHeader = New ADODB.Recordset
    Set rsSubHeader = gconDMIS.Execute("select * from AMIS_SubHeader where code = " & N2Str2Null(Acc))
    If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then
        SetSubHeaderDesc = Null2String(rsSubHeader!Description)
    End If
End Function

Function SetTitleCode(Acc As String) As String
    Set rsTitleCode = New ADODB.Recordset
    Set rsTitleCode = gconDMIS.Execute("select * from AMIS_TitleCode where description = " & N2Str2Null(Acc))
    If Not rsTitleCode.EOF And Not rsTitleCode.BOF Then
        SetTitleCode = Null2String(rsTitleCode!SubTitleCode)
    End If
End Function

Function SetTitleCodeDesc(Acc As String) As String
    Set rsTitleCode = New ADODB.Recordset
    Set rsTitleCode = gconDMIS.Execute("select * from AMIS_TitleCode where code = " & N2Str2Null(Acc))
    If Not rsTitleCode.EOF And Not rsTitleCode.BOF Then
        SetTitleCodeDesc = Null2String(rsTitleCode!Description)
    End If
End Function

Function SetDeptCode(Acc As String) As String
    Set rsDepartment = New ADODB.Recordset
    Set rsDepartment = gconDMIS.Execute("select * from AMIS_Department where DeptName = " & N2Str2Null(Acc))
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SetDeptCode = Null2String(rsDepartment!DeptCode)
    End If
End Function

Function SetDeptName(Acc As String) As String
    Set rsDepartment = New ADODB.Recordset
    Set rsDepartment = gconDMIS.Execute("select * from AMIS_Department where DeptCode = " & N2Str2Null(Acc))
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then
        SetDeptName = Null2String(rsDepartment!DeptName)
    End If
End Function

Sub rsRefresh()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select * from AMIS_ChartAccount order by AcctCode asc")
End Sub

Sub initMemvars()
    txtCode.Text = "XX-XXXXX-XX": txtCode1.Text = "XX": txtCode2.Text = "XXXXX": txtCode3.Text = "XX"
    txtDescription.Text = "": txtAccountName.Text = ""
    picWizard.Visible = True: picWizard.Visible = True
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select Description from AMIS_Acctype order by code asc")
    If Not rsAccType.EOF And Not rsAccType.BOF Then Combo_Loadval cboType, rsAccType
    Set rsAccType = New ADODB.Recordset
    Set rsAccType = gconDMIS.Execute("select Description from AMIS_Acctype order by code asc")
    If Not rsAccType.EOF And Not rsAccType.BOF Then Combo_Loadval cboAccountType, rsAccType
    Set rsHeader = New ADODB.Recordset
    Set rsHeader = gconDMIS.Execute("select Description from AMIS_Header order by code asc")
    If Not rsHeader.EOF And Not rsHeader.BOF Then Combo_Loadval cboHeader, rsHeader
    Set rsSubHeader = New ADODB.Recordset
    Set rsSubHeader = gconDMIS.Execute("select Description from AMIS_SubHeader order by code asc")
    If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then Combo_Loadval cboSubHeader, rsSubHeader
    Set rsTitleCode = New ADODB.Recordset
    Set rsTitleCode = gconDMIS.Execute("select Description from AMIS_TitleCode order by Code asc")
    If Not rsTitleCode.EOF And Not rsTitleCode.BOF Then Combo_Loadval cboTitleCode, rsTitleCode
    Set rsDepartment = New ADODB.Recordset
    Set rsDepartment = gconDMIS.Execute("select DeptName from AMIS_Department order by DeptCode asc")
    If Not rsDepartment.EOF And Not rsDepartment.BOF Then Combo_Loadval cboDepartment, rsDepartment
    cboHeader.Enabled = True: cboSubHeader.Enabled = False: cmdOkHeader.Enabled = True
    cboTitleCode.Enabled = False: cboDepartment.Enabled = False
    cmdOkSubHeader.Enabled = False: cmdOkTitleCode.Enabled = False: cmdOkDepartment.Enabled = False
    cboAccountType.Enabled = False
    txtCode1.Text = "0": txtCode2.Text = "0": txtCode3.Text = "00"
    txtCode4.Text = "0": txtCode5.Text = "00": txtCode6.Text = "00"
    lstRelatedAccounts.ListItems.Clear
End Sub

Sub StoreMemVars()
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Frame1.Enabled = False
        picWizard.Visible = False
        cmdWizard.Visible = False
        labID.Caption = rsChartAccount!ID
        txtCode.Text = Null2String(rsChartAccount!ACCTCODE)
        txtDescription.Text = Null2String(rsChartAccount!Description)
        If Null2Bool(rsChartAccount!Is_Schedule_Accnt) = True Then
            chkSchedule.Value = 1
        Else
            chkSchedule.Value = 0
        End If

        If Null2String(rsChartAccount!ACCTTYPE) <> "" Then
            cboType.Text = SetAccType(Null2String(rsChartAccount!ACCTTYPE))
        Else
            cboType.ListIndex = -1
        End If
'        If Null2String(rsChartAccount!Trantype1) <> "" Or Null2String(rsChartAccount!Trantype2) <> "" Or Null2String(rsChartAccount!Trantype3) <> "" Or Null2String(rsChartAccount!Trantype4) <> "" Then
'            cmdDelete.Enabled = False
'        Else
            cmdDelete.Enabled = True
'        End If
    Else
        'MsgBox "No Such Record!"
        MessagePop RecNotFound, "Not Found", "No Such Record!"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreEntry(XXX As Variant)
    Dim rsChartOfAccount                          As ADODB.Recordset
    Set rsChartOfAccount = New ADODB.Recordset
    Set rsChartOfAccount = gconDMIS.Execute("select * from AMIS_ChartAccount where ID = " & XXX)
    If Not rsChartOfAccount.EOF And Not rsChartOfAccount.BOF Then
        On Error Resume Next
        fraDetails.Enabled = False
        initMemvars
        labID.Caption = rsChartOfAccount!ID
        picWizard.Visible = True
        cmdWizard.Visible = True
        txtCode.Text = Null2String(rsChartOfAccount!ACCTCODE)
        txtCode1.Text = Null2String(rsChartOfAccount!HeaderCode)
        txtCode2.Text = Null2String(rsChartOfAccount!SubHeaderCode)
        txtCode3.Text = Null2String(rsChartOfAccount!TitleCode)
        txtCode4.Text = Null2String(rsChartOfAccount!SubTitleCode)
        txtCode5.Text = Null2String(rsChartOfAccount!DetailCode)
        txtCode6.Text = Null2String(rsChartOfAccount!DepartmentCode)
        txtDescription.Text = Null2String(rsChartOfAccount!Description)
        cboAccountType.Text = SetAccType(Null2String(rsChartOfAccount!ACCTTYPE))
        cboHeader.Text = SetHeaderDesc(Null2String(rsChartOfAccount!HeaderCode))
        cboSubHeader.Text = SetSubHeaderDesc(Null2String(rsChartOfAccount!HeaderCode) & Null2String(rsChartOfAccount!SubHeaderCode))
        cboTitleCode.Text = SetTitleCodeDesc(Null2String(rsChartOfAccount!HeaderCode) & Null2String(rsChartOfAccount!SubHeaderCode) & Null2String(rsChartOfAccount!TitleCode))
        cboDepartment.Text = SetDeptName(Null2String(rsChartOfAccount!DepartmentCode))
        txtAccountName.Text = Null2String(rsChartOfAccount!Description)
        FillRelatedGrid
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartOfAccount                          As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Dim xChart                                    As ListItem
    Dim lvCount                                   As Integer
    Set rsChartOfAccount = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If Option1.Value = True Then
        Set rsChartOfAccount = gconDMIS.Execute("select acctcode,Description,Accttype,ID,Is_Schedule_Accnt from AMIS_ChartAccount where description like'" & Repleys(XXX) & "%'")
    Else
        Set rsChartOfAccount = gconDMIS.Execute("select acctcode,Description,Accttype,ID,Is_Schedule_Accnt from AMIS_ChartAccount where acctcode like'" & Repleys(XXX) & "%'")
    End If



    If Not (rsChartOfAccount.EOF And rsChartOfAccount.BOF) Then
        'Listview_Loadval Me.lstAccounts.ListItems, rsChartOfAccount
        Do While Not rsChartOfAccount.EOF
            Set xChart = lstAccounts.ListItems.Add(, , Null2String(rsChartOfAccount!ACCTCODE))
            xChart.SubItems(1) = Null2String(rsChartOfAccount!Description)
            xChart.SubItems(2) = Null2String(rsChartOfAccount!ACCTTYPE)
            xChart.SubItems(3) = Null2String(rsChartOfAccount!ID)
            xChart.SubItems(4) = Null2String(rsChartOfAccount!Is_Schedule_Accnt)
            rsChartOfAccount.MoveNext
        Loop

        For lvCount = 1 To lstAccounts.ListItems.Count
            If lstAccounts.ListItems(lvCount).SubItems(4) = "True" Then
                lstAccounts.ListItems(lvCount).ForeColor = &HC00000
                lstAccounts.ListItems(lvCount).ListSubItems.Item(1).ForeColor = &HC00000
                lstAccounts.ListItems(lvCount).ListSubItems.Item(2).ForeColor = &HC00000
            End If
        Next lvCount
        lstAccounts.Refresh
        lstAccounts.Enabled = True
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Sub FillRelatedGrid()
    Dim rsChartOfAccount                          As ADODB.Recordset
    lstRelatedAccounts.Enabled = False
    lstRelatedAccounts.Sorted = False: lstRelatedAccounts.ListItems.Clear
    Set rsChartOfAccount = New ADODB.Recordset
    Set rsChartOfAccount = gconDMIS.Execute("select acctcode,Description,Accttype,ID from AMIS_ChartAccount where Titles = '" & txtCode1.Text & txtCode2.Text & txtCode3.Text & "' order by AcctCode asc")
    If Not (rsChartOfAccount.EOF And rsChartOfAccount.BOF) Then
        Listview_Loadval Me.lstRelatedAccounts.ListItems, rsChartOfAccount
        lstRelatedAccounts.Refresh
        lstRelatedAccounts.Enabled = True
        lstRelatedAccounts.Enabled = True
    Else
        lstRelatedAccounts.Enabled = False
    End If

End Sub

Sub SetAccountCode()
    txtCode.Text = txtCode1.Text & txtCode2.Text & "-" & txtCode3.Text & txtCode4.Text & txtCode5.Text & "-" & txtCode6.Text
End Sub

Private Sub cboAccountType_Click()
    If AddorEdit <> "" Then cboType.Text = cboAccountType.Text
End Sub

Private Sub cboDepartment_Click()
    txtCode6.Text = SetDeptCode(cboDepartment.Text)
End Sub

Private Sub cboHeader_Click()
    txtCode1.Text = SetHeaderCode(cboHeader.Text)
End Sub

Private Sub cboSubHeader_Click()
    txtCode2.Text = SetSubHeaderCode(cboSubHeader.Text)
End Sub

Private Sub cboTitleCode_Click()
    txtCode3.Text = SetTitleCode(cboTitleCode.Text)
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "Chart of Accounts") = False Then Exit Sub

    AddorEdit = "ADD"
    initMemvars
    Picture1.Visible = False
    Picture2.Visible = True
    chkSchedule.Enabled = True
    lstAccounts.Enabled = False
    On Error Resume Next
    '    txtCode1.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    chkSchedule.Enabled = False
    lstAccounts.Enabled = True
    AddorEdit = ""
    StoreMemVars
    fraDetails.Enabled = True
    txtSearch_Change
    picWizard.Visible = False
    cmdWizard.Visible = False
    On Error Resume Next
    lstAccounts.FindItem(txtCode.Text).EnsureVisible
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "CHART OF ACCOUNTS") = False Then Exit Sub
    If MsgBox("Delete Current Record", vbQuestion + vbYesNo, "Delete") = vbYes Then
        Dim rsCheckUsedAcctCode                   As ADODB.Recordset
        Set rsCheckUsedAcctCode = New ADODB.Recordset
        Set rsCheckUsedAcctCode = gconDMIS.Execute("Select * from AMIS_Journal_Det Where Acct_Code = " & N2Str2Null(txtCode.Text))
        If Not rsCheckUsedAcctCode.EOF And Not rsCheckUsedAcctCode.BOF Then

            MessagePop RecLocekd, "Required Record", "Account is in used and cannot be deleted. Pls. check this account in General Ledger"
            'MsgBox "Account is in used and cannot be deleted." & vbCrLf & _
             '       "Pls. check this account in General Ledger", vbCritical, "Warning"
            Exit Sub
        End If
        SQL_STATEMENT = "delete from AMIS_ChartAccount where id = " & Trim(Me.lstAccounts.SelectedItem.SubItems(3))
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "X", "CHART OF ACCOUNTS", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", ""
    End If
    rsRefresh
    StoreMemVars
    FillGrid
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "CHART OF ACCOUNTS") = False Then Exit Sub
    AddorEdit = "EDIT"
    StoreEntry labID.Caption
    lstAccounts.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    chkSchedule.Enabled = True
    PREV_ACCT_CODE = txtCode.Text

    'UPDATED BY: JUN --- DESCRIPTION: USER IS NOT ALLOWED TO EDIT THE ACCT CODE
    If Function_Access(LOGID, "Acess_Add", "Chart of Accounts") = False Then
        cboHeader.Enabled = False
        cmdOkHeader.Enabled = False
        txtCode4.Enabled = False
        txtCode5.Enabled = False
        Exit Sub
    Else
        cboHeader.Enabled = True
        cmdOkHeader.Enabled = True
        txtCode4.Enabled = True
        txtCode5.Enabled = True
    End If
    'UPDATED BY: JUN

    On Error Resume Next
    'txtDescription.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdFind_Click()
    On Error GoTo ErrorCode:

    txtSearch.Enabled = True
    On Error Resume Next
    txtSearch.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

    rsChartAccount.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:03
Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:

    rsChartAccount.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsChartAccount.MoveNext
    If rsChartAccount.EOF Then
        rsChartAccount.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdOkDepartment_Click()
    cboDepartment.Enabled = False
    cboAccountType.Enabled = True
    cmdOkDepartment.Enabled = False
    picAccount.Enabled = False
End Sub

Private Sub cmdOkHeader_Click()
    cboHeader.Enabled = False
    cboSubHeader.Enabled = True
    cmdOkSubHeader.Enabled = True
    cmdOkHeader.Enabled = False
    Set rsSubHeader = New ADODB.Recordset
    Set rsSubHeader = gconDMIS.Execute("select Description from AMIS_SubHeader where HeaderCode = '" & txtCode1.Text & "' order by code asc")
    If Not rsSubHeader.EOF And Not rsSubHeader.BOF Then
        Combo_Loadval cboSubHeader, rsSubHeader
    End If
End Sub

Private Sub cmdOkSubHeader_Click()
    cboSubHeader.Enabled = False
    cboTitleCode.Enabled = True
    cmdOkTitleCode.Enabled = True
    cmdOkSubHeader.Enabled = False
    Set rsTitleCode = New ADODB.Recordset
    Set rsTitleCode = gconDMIS.Execute("select Description from AMIS_TitleCode where SubHeaderCode = '" & txtCode1.Text & txtCode2.Text & "' order by Code asc")
    If Not rsTitleCode.EOF And Not rsTitleCode.BOF Then
        Combo_Loadval cboTitleCode, rsTitleCode
    End If
End Sub

Private Sub cmdOkTitleCode_Click()
    cboTitleCode.Enabled = False
    cboDepartment.Enabled = True
    cmdOkDepartment.Enabled = True
    cmdOkTitleCode.Enabled = False
    picAccount.Enabled = True
    FillRelatedGrid
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsChartAccount.MovePrevious
    If rsChartAccount.BOF Then
        rsChartAccount.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "CHART OF ACCOUNTS") = False Then Exit Sub
    Screen.MousePointer = 11
    ShowReport "ChartofAccounts", "AccountFiles", "", "Chart of Accounts", "AS OF: " & LOGDATE, True
    Screen.MousePointer = 0
    LogAudit "V", "CHART OF ACCOUNTS", cboType & " - " & txtDescription
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0707200713:04
Private Sub cmdSave_Click()
    Dim rsChartAccountDup                         As ADODB.Recordset
    Dim VtxtCode, vtxtDescription, VcboType       As String
    Dim VHeaderCode, VSubHeaderCode, VHeaders, VTitleCode, VSubTitleCode, VTitles, VDetailCode, VDepartmentCode As String
    Dim VtxtShedAccount                           As Integer
    On Error GoTo ErrorCode:

    If chkSchedule.Value = 1 Then
        VtxtShedAccount = 1
    Else
        VtxtShedAccount = 0
    End If

    VtxtCode = N2Str2Null(txtCode.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    VcboType = N2Str2Null(SetAccCode(cboType.Text))
    VHeaderCode = N2Str2Null(txtCode1.Text)
    VSubHeaderCode = N2Str2Null(txtCode2.Text)
    VHeaders = N2Str2Null(txtCode1.Text & txtCode2.Text)
    VTitleCode = N2Str2Null(txtCode3.Text)
    VSubTitleCode = N2Str2Null(txtCode4.Text)
    VTitles = N2Str2Null(txtCode1.Text & txtCode2.Text & txtCode3.Text)
    VDetailCode = N2Str2Null(txtCode5.Text)
    VDepartmentCode = N2Str2Null(txtCode6.Text)
    If cboHeader.Text = "" Then
        'MsgBox "Invalid Account Header!", vbOKOnly + vbCritical, "Error!"
        MessagePop InfoVoid, "Invalid Entry:", "Invalid Account Header!"
        Exit Sub
    End If
    If cboDepartment.Text = "" Then
        MessagePop InfoVoid, "Invalid Entry:", "Invalid Department Name!"
        'MsgBox "Invalid Department Name!", vbOKOnly + vbCritical, "Error!"
        Exit Sub
    End If
    If AddorEdit = "EDIT" Then
        If txtCode.Text <> PREV_ACCT_CODE Then
            Set rsChartAccountDup = New ADODB.Recordset
            rsChartAccountDup.Open "select Acctcode from AMIS_ChartAccount where Acctcode = " & VtxtCode, gconDMIS
            If Not rsChartAccountDup.EOF And Not rsChartAccountDup.BOF Then
                'MsgBox "ChartAccount Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
                MessagePop RecSaveError, "Duplicate Entry", "ChartAccount Code Already Exist!"
                Exit Sub
            End If
        End If
    End If
    If AddorEdit = "ADD" Then
        Set rsChartAccountDup = New ADODB.Recordset
        rsChartAccountDup.Open "select Acctcode from AMIS_ChartAccount where Acctcode = " & VtxtCode, gconDMIS
        If Not rsChartAccountDup.EOF And Not rsChartAccountDup.BOF Then
            'MsgBox "ChartAccount Code Already Exist!", vbCritical, "Duplicate Code Not Allowed"
            MessagePop RecSaveError, "Duplicate Entry", "ChartAccount Code Already Exist!"
            Exit Sub
        End If
        SQL_STATEMENT = "Insert into AMIS_ChartAccount " & _
                        "(AcctCode,Description,HeaderCode,SubHeaderCode,Headers,TitleCode,SubTitleCode,Titles,DetailCode,DepartmentCode, IS_SCHEDULE_ACCNT ,AcctType) " & _
                        " values (" & VtxtCode & _
                        ", " & vtxtDescription & _
                        ", " & VHeaderCode & _
                        ", " & VSubHeaderCode & _
                        ", " & VHeaders & _
                        ", " & VTitleCode & _
                        ", " & VSubTitleCode & _
                        ", " & VTitles & _
                        ", " & VDetailCode & _
                        ", " & VDepartmentCode & _
                        ", " & VtxtShedAccount & _
                        ", " & VcboType & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "CHART OF ACCOUNTS", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", N2Str2Null(vtxtDescription)
    Else
        SQL_STATEMENT = "update AMIS_ChartAccount set" & _
                        " AcctCode = " & VtxtCode & "," & _
                        " AcctType = " & VcboType & "," & _
                        " HeaderCode = " & VHeaderCode & "," & _
                        " SubHeaderCode = " & VSubHeaderCode & "," & _
                        " Headers = " & VHeaders & "," & _
                        " TitleCode = " & VTitleCode & "," & _
                        " SubtitleCode = " & VSubTitleCode & "," & _
                        " Titles = " & VTitles & "," & _
                        " DetailCode = " & VDetailCode & "," & _
                        " DepartmentCode = " & VDepartmentCode & "," & _
                        " Description = " & vtxtDescription & "," & _
                        " IS_SCHEDULE_ACCNT = " & VtxtShedAccount & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "CHART OF ACCOUNTS", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", N2Str2Null(vtxtDescription)
        SQL_STATEMENT = "update AMIS_Journal_Det Set Acct_Code = " & VtxtCode & ", Acct_Name = " & vtxtDescription & " Where Acct_Code = " & N2Str2Null(PREV_ACCT_CODE)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "CHART OF ACCOUNTS", SQL_STATEMENT, labID.Caption, "", txtCode.Text, "", N2Str2Null(vtxtDescription)
    End If
    rsRefresh
    FillGrid
    rsChartAccount.Find "AcctCode = " & VtxtCode
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

'Private Sub Command1_Click()
'Dim rsJournal_Det As ADODB.Recordset
'Dim rsChartAccount As ADODB.Recordset
'Set rsJournal_Det = New ADODB.Recordset
'Set rsJournal_Det = gconDMIS.Execute("select * from AMIS_Journal_Det Order by ID ASC")
'If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
'   Do While Not rsJournal_Det.EOF
'      Set rsChartAccount = New ADODB.Recordset
'      Set rsChartAccount = gconDMIS.Execute("SELECT * from AMIS_ChartAccount WHERE ACCTCODE = " & N2Str2Null(rsJournal_Det!acct_code))
'      'If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
'      '   Label4.Caption = Null2String(rsChartAccount!AcctCode)
'      '   gconDMIS.Execute "update AMIS_Journal_Det SET " & _
       '      '                    " ACCT_CODE = " & N2Str2Null(rsChartAccount!AcctCode) & "," & _
       '      '                    " ACCT_NAME = " & N2Str2Null(rsChartAccount!Description) & _
       '      '                    " WHERE ID  = " & rsJournal_Det!ID
'      '   DoEvents
'      'End If
'      If rsChartAccount.EOF And rsChartAccount.BOF Then
'         MsgBox rsJournal_Det!acct_code
'      End If
'      rsJournal_Det.MoveNext
'   Loop
'   MsgBox "Update Completed"
'End If
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next

    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = "CHART OF ACCOUNTS"
        Call frmALL_AuditInquiry.DisplayHistory(labID, "CHART OF ACCOUNTS")

    Case vbKeyF1
        Call FillGrid(1)
    Case vbKeyF2
        Call FillGrid(0)
    Case vbKeyF3
        FillGrid
    Case vbKeyEscape
        txtSearch.Text = ""
        txtSearch.Enabled = False
    Case vbKeyDelete
        If Mid(Me.ActiveControl.Name, 1, 3) = "cbo" Then Me.ActiveControl.ListIndex = -1
    Case Else
        If Left(Me.ActiveControl.Name, 3) <> "lst" Then MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    initMemvars
    AddorEdit = ""
    txtSearch.Text = ""
    StoreMemVars
    Call FillGrid
    cmdWizard.ZOrder 0
    picWizard.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub FillGrid(Optional XXX As String)
    Dim rsChartOfAccount                          As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartOfAccount = New ADODB.Recordset
    Dim xChart                                    As ListItem
    Dim lvCount                                   As Integer
    If XXX = "" Then
        Set rsChartOfAccount = gconDMIS.Execute("select acctcode,Description,Accttype,ID,Is_Schedule_Accnt from AMIS_ChartAccount Order by AcctCode asc")
    Else
        Set rsChartOfAccount = gconDMIS.Execute("select acctcode,Description,Accttype,ID,Is_Schedule_Accnt from AMIS_ChartAccount WHERE ISNULL(IS_SCHEDULE_ACCNT,0) = '" & XXX & "' Order by AcctCode asc")
    End If
    If Not (rsChartOfAccount.EOF And rsChartOfAccount.BOF) Then
        'Listview_Loadval Me.lstAccounts.ListItems, rsChartOfAccount
        Do While Not rsChartOfAccount.EOF
            Set xChart = lstAccounts.ListItems.Add(, , Null2String(rsChartOfAccount!ACCTCODE))
            xChart.SubItems(1) = Null2String(rsChartOfAccount!Description)
            xChart.SubItems(2) = Null2String(rsChartOfAccount!ACCTTYPE)
            xChart.SubItems(3) = Null2String(rsChartOfAccount!ID)
            xChart.SubItems(4) = Null2String(rsChartOfAccount!Is_Schedule_Accnt)
            rsChartOfAccount.MoveNext
        Loop
        For lvCount = 1 To lstAccounts.ListItems.Count
            If lstAccounts.ListItems(lvCount).SubItems(4) = "True" Then
                lstAccounts.ListItems(lvCount).ForeColor = &HC00000
                lstAccounts.ListItems(lvCount).ListSubItems.Item(1).ForeColor = &HC00000
                lstAccounts.ListItems(lvCount).ListSubItems.Item(2).ForeColor = &HC00000
            End If
        Next lvCount
        lstAccounts.Refresh
        lstAccounts.Enabled = True
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REFRESH_ACCOUNT = True Then
        Unload Me
        'frmAMISJournalEntry.txtSearch.Text = txtCode.Text
    End If
End Sub

Private Sub lstAccounts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstAccounts
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstAccounts_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsChartAccount.Bookmark = rsFind(rsChartAccount.Clone, "acctcode", lstAccounts.SelectedItem).Bookmark
    StoreMemVars
End Sub

Private Sub lstAccounts_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown Then
        If lstAccounts.ListItems.Count > 0 And lstAccounts.Enabled = True Then
            lstAccounts.SetFocus
        End If
    End If
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdEdit.Value = True
End Sub

Private Sub Option1_Click()
    txtSearch.SetFocus
    Call txtSearch_Change
End Sub

Private Sub Option2_Click()
    txtSearch.SetFocus
    Call txtSearch_Change
End Sub

Private Sub txtAccountName_Change()
    If AddorEdit <> "" Then txtDescription.Text = txtAccountName.Text
End Sub

Private Sub txtAccountName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCode1_Change()
    SetAccountCode
End Sub

Private Sub txtCode2_Change()
    SetAccountCode
End Sub

Private Sub txtCode3_Change()
    SetAccountCode
End Sub

Private Sub txtCode4_Change()
    SetAccountCode
End Sub

Private Sub txtCode4_GotFocus()
    If txtCode4.Text = "0" Then txtCode4.Text = ""
End Sub

Private Sub txtCode4_LostFocus()
    If Trim(txtCode4.Text) = "" Then txtCode4.Text = "0"
End Sub

Private Sub txtCode5_Change()
    SetAccountCode
End Sub

Private Sub txtCode5_GotFocus()
    If txtCode5.Text = "00" Then txtCode5.Text = ""
End Sub

Private Sub txtCode5_LostFocus()
    If Trim(txtCode5.Text) = "" Then txtCode5.Text = "00"
End Sub

Private Sub txtCode6_Change()
    SetAccountCode
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstAccounts.ListItems.Count > 0 And lstAccounts.Enabled = True Then
            lstAccounts.SetFocus
        End If
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

