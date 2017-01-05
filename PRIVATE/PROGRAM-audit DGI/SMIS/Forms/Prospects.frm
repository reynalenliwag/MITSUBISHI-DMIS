VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Files_Prospects 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prospect Information"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9765
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Prospects.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   9765
   Begin VB.PictureBox picSearch 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8160
      Left            =   0
      ScaleHeight     =   8160
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox cboStatus 
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
         ItemData        =   "Prospects.frx":08CA
         Left            =   90
         List            =   "Prospects.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   300
         Width           =   2475
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   30
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   11245
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
         MouseIcon       =   "Prospects.frx":08CE
         NumItems        =   0
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   60
         TabIndex        =   1
         Top             =   900
         Width           =   2490
      End
      Begin VB.Label Label17 
         Caption         =   "Filter by Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Search Account Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   64
         Top             =   660
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   9795
      TabIndex        =   42
      Top             =   7200
      Width           =   9795
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   2100
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2670
         Top             =   90
      End
      Begin VB.PictureBox picSaves 
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
         Left            =   8235
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   45
         Top             =   15
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   795
            Left            =   720
            MouseIcon       =   "Prospects.frx":0A30
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":0B82
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Cancel"
            Top             =   90
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   30
            MouseIcon       =   "Prospects.frx":0EC0
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":1012
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Save Prospect"
            Top             =   90
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
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
         Height          =   945
         Left            =   3840
         ScaleHeight     =   945
         ScaleWidth      =   6075
         TabIndex        =   48
         Top             =   45
         Width           =   6075
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   5130
            MouseIcon       =   "Prospects.frx":1362
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":14B4
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   5130
            MouseIcon       =   "Prospects.frx":181A
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":196C
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Delete Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   4440
            MouseIcon       =   "Prospects.frx":1C97
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":1DE9
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   3750
            MouseIcon       =   "Prospects.frx":2145
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":2297
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   3060
            MouseIcon       =   "Prospects.frx":25AA
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":26FC
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Find a Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   2370
            MouseIcon       =   "Prospects.frx":29F6
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":2B48
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   1680
            MouseIcon       =   "Prospects.frx":2EA0
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":2FF2
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdMM_ProspectLog 
            Caption         =   "Option"
            Height          =   795
            Left            =   990
            MouseIcon       =   "Prospects.frx":3351
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":34A3
            Style           =   1  'Graphical
            TabIndex        =   62
            Tag             =   "1102"
            ToolTipText     =   "Prospect Logs"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton Command1 
            Caption         =   "From Customer"
            Height          =   795
            Left            =   210
            MouseIcon       =   "Prospects.frx":3B1E
            MousePointer    =   99  'Custom
            Picture         =   "Prospects.frx":3C70
            Style           =   1  'Graphical
            TabIndex        =   72
            Tag             =   "1102"
            ToolTipText     =   "New Prospect Information From Existing Customer"
            Top             =   60
            Width           =   795
         End
      End
      Begin VB.Label labConflictCUSCODE 
         Height          =   465
         Left            =   1350
         TabIndex        =   44
         Top             =   225
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label labCusCode 
         Height          =   375
         Left            =   180
         TabIndex        =   43
         Top             =   30
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox picAdvanceProspecting 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   4170
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2715
      ScaleWidth      =   4350
      TabIndex        =   56
      Top             =   2250
      Visible         =   0   'False
      Width           =   4380
      Begin VB.CommandButton cmdOption_View 
         BackColor       =   &H8000000D&
         Caption         =   "View Transaction Info"
         Height          =   435
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Prospects.frx":3DDA
         MousePointer    =   99  'Custom
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Remove Customer Information"
         Top             =   1650
         Width           =   2325
      End
      Begin VB.CommandButton cmdOption_Delete 
         BackColor       =   &H8000000D&
         Caption         =   "Delete Prospect Information"
         Height          =   405
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Prospects.frx":3F2C
         MousePointer    =   99  'Custom
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "Remove Customer Information"
         Top             =   1260
         Width           =   2325
      End
      Begin VB.CommandButton cmdOption_New 
         BackColor       =   &H8000000D&
         Caption         =   "Convert Into New Customer"
         Height          =   375
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Prospects.frx":407E
         MousePointer    =   99  'Custom
         TabIndex        =   60
         TabStop         =   0   'False
         ToolTipText     =   "Remove Customer Information"
         Top             =   900
         Width           =   2325
      End
      Begin VB.CommandButton cmdOption_Add 
         BackColor       =   &H8000000D&
         Caption         =   "Add To Existing Customer"
         Height          =   465
         Left            =   1890
         MaskColor       =   &H00FFFFFF&
         MouseIcon       =   "Prospects.frx":41D0
         MousePointer    =   99  'Custom
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Add/Change Customer Information"
         Top             =   450
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancelStatus 
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
         Height          =   255
         Index           =   1
         Left            =   3990
         TabIndex        =   57
         Top             =   30
         Width           =   285
      End
      Begin VB.Label labid 
         Caption         =   "0"
         Height          =   465
         Left            =   1890
         TabIndex        =   71
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label labINFO 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2325
         Left            =   30
         TabIndex        =   61
         Top             =   360
         Width           =   1815
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Width           =   4515
         _Version        =   655364
         _ExtentX        =   7964
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Prospect Option"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picProspectEntry 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   2655
      ScaleHeight     =   8160
      ScaleWidth      =   8385
      TabIndex        =   3
      Top             =   0
      Width           =   8385
      Begin VB.TextBox txtProsName 
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
         Left            =   1530
         MaxLength       =   90
         TabIndex        =   11
         Tag             =   "@R"
         Top             =   1380
         Width           =   5445
      End
      Begin VB.TextBox txtTelNo 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1530
         MaxLength       =   15
         TabIndex        =   18
         Tag             =   "0"
         Top             =   2550
         Width           =   1905
      End
      Begin VB.TextBox txtAddres 
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
         ForeColor       =   &H00000000&
         Height          =   1170
         Left            =   3810
         MaxLength       =   200
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   2460
         Width           =   3165
      End
      Begin VB.ComboBox cboSAE 
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
         Left            =   1530
         TabIndex        =   9
         Tag             =   "@R"
         Top             =   960
         Width           =   3285
      End
      Begin VB.TextBox txtEmailAdd 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   21
         Tag             =   "1"
         Top             =   2970
         Width           =   1905
      End
      Begin VB.TextBox txtContactPerson 
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
         Left            =   1530
         MaxLength       =   90
         TabIndex        =   13
         Top             =   1770
         Width           =   5445
      End
      Begin VB.TextBox txtCellPhone 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   1530
         MaxLength       =   30
         TabIndex        =   25
         Tag             =   "2"
         Top             =   3330
         Width           =   1905
      End
      Begin VB.CommandButton cmdSee 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   3480
         MouseIcon       =   "Prospects.frx":4322
         MousePointer    =   99  'Custom
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "EMAIL"
         Top             =   2970
         Width           =   285
      End
      Begin VB.CommandButton cmdSee 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   3480
         MouseIcon       =   "Prospects.frx":4474
         MousePointer    =   99  'Custom
         TabIndex        =   19
         TabStop         =   0   'False
         Tag             =   "TELEPHONE"
         Top             =   2550
         Width           =   285
      End
      Begin VB.CommandButton cmdSee 
         Caption         =   "?"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   3480
         MouseIcon       =   "Prospects.frx":45C6
         MousePointer    =   99  'Custom
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "MOBILE"
         Top             =   3330
         Width           =   285
      End
      Begin VB.TextBox txtProsCode 
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
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   4830
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "@R"
         Top             =   960
         Width           =   2115
      End
      Begin VB.TextBox txtComments 
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
         ForeColor       =   &H00000000&
         Height          =   1140
         Left            =   135
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   5985
         Width           =   6915
      End
      Begin VB.ComboBox cboProspectType 
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
         Left            =   1530
         TabIndex        =   15
         Text            =   "cboProspectType"
         Top             =   2160
         Width           =   2235
      End
      Begin MSMask.MaskEdBox txtDtInquiry 
         Height          =   345
         Left            =   1530
         TabIndex        =   6
         Tag             =   "@R"
         Top             =   120
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4194304
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "mm-dd-yyyy"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCusCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   1530
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   510
         Width           =   1275
      End
      Begin VB.Frame Frame2 
         Height          =   2115
         Left            =   135
         TabIndex        =   27
         Top             =   3615
         Width           =   6885
         Begin VB.ComboBox cboClassification 
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
            Left            =   3600
            TabIndex        =   31
            Top             =   420
            Width           =   3165
         End
         Begin VB.ComboBox cboLeadSource 
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
            Left            =   3600
            TabIndex        =   35
            Tag             =   "@R"
            Top             =   1020
            Width           =   3165
         End
         Begin VB.ComboBox cboColor 
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
            ForeColor       =   &H00400000&
            Height          =   345
            Left            =   240
            TabIndex        =   37
            Tag             =   "@R"
            Top             =   1605
            Width           =   3150
         End
         Begin VB.ComboBox cboVariant 
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
            Left            =   240
            TabIndex        =   34
            Top             =   1020
            Width           =   3135
         End
         Begin VB.CheckBox chkFollowUps 
            Caption         =   "For Followups"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3615
            TabIndex        =   38
            Top             =   1620
            Width           =   1530
         End
         Begin VB.ComboBox cboModel 
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
            Left            =   240
            TabIndex        =   30
            Tag             =   "@R"
            Text            =   "cboModel"
            Top             =   420
            Width           =   3135
         End
         Begin MSMask.MaskEdBox txtFollowupDate 
            Height          =   345
            Left            =   5160
            TabIndex        =   39
            Top             =   1590
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   4194304
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   20
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "mm-dd-yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classification "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   3600
            TabIndex        =   28
            Top             =   120
            Width           =   1200
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lead Source"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   3615
            TabIndex        =   33
            Top             =   810
            Width           =   1080
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   240
            TabIndex        =   36
            Top             =   1395
            Width           =   450
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Variant "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   240
            TabIndex        =   32
            Top             =   780
            Width           =   660
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model "
            DataField       =   "s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   240
            TabIndex        =   29
            Top             =   180
            Width           =   555
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   5730
         Width           =   540
      End
      Begin VB.Label labStatus 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   5220
         TabIndex        =   4
         Top             =   60
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prospect Name "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Inquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   450
         TabIndex        =   5
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. No. "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   795
         TabIndex        =   17
         Top             =   2580
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   3840
         TabIndex        =   16
         Top             =   2190
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Executive"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   990
         TabIndex        =   20
         Top             =   2925
         Width           =   465
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Phone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   555
         TabIndex        =   23
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   135
         TabIndex        =   12
         Top             =   1800
         Width           =   1320
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prospect Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   225
         TabIndex        =   14
         Top             =   2190
         Width           =   1230
      End
      Begin VB.Label labCustomerName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2820
         TabIndex        =   69
         Top             =   510
         Width           =   4110
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   270
         TabIndex        =   68
         Top             =   570
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmSMIS_Files_Prospects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMain                                                            As ADODB.Recordset
Dim ctl                                                               As Control
Dim AddorEdit                                                         As String
Dim StatusProspect                                                    As String
Dim PROSPECTID                                                        As Long

Private WithEvents FormCustomer                                       As frmSMIS_Mis_SearchMaster
Attribute FormCustomer.VB_VarHelpID = -1
Private WithEvents FormSearchCustomer                                 As frmSMIS_Mis_SearchMaster
Attribute FormSearchCustomer.VB_VarHelpID = -1
Private WithEvents FormConflictingCustomer                            As frmSMIS_Trans_Confilct
Attribute FormConflictingCustomer.VB_VarHelpID = -1
Private WithEvents FormAllCustomer                                    As frmAllCustomer
Attribute FormAllCustomer.VB_VarHelpID = -1
Private SpecialProspect                                               As String
Dim WithEvents SEARCHFORM                                             As frmSMIS_Mis_SearchMaster
Attribute SEARCHFORM.VB_VarHelpID = -1

Function GetProspectType() As String
    If cboProspectType.ListIndex = 0 Then
        GetProspectType = "P"
    ElseIf cboProspectType.ListIndex = 1 Then
        GetProspectType = "C"
    ElseIf cboProspectType.ListIndex = 2 Then
        GetProspectType = "G"
    ElseIf cboProspectType.ListIndex = 3 Then
        GetProspectType = "F"
    Else
        GetProspectType = vbNullString
    End If
End Function
''end get set prospect type

Private Function SOEXISTS() As Boolean
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT COUNT(*) as Total FROM SMIS_SALESORDER WHERE PROSPECTID=" & labid)
    If TEMPRS(0).Value <> 0 Then
        SOEXISTS = True
    Else
        SOEXISTS = False
    End If
    Set TEMPRS = Nothing
End Function

Sub HanapID(XXX As Long)
    rsMain.MoveFirst
    rsMain.Find "ProsCode =" & txtProsCode
    XXX = rsMain.Fields("PROSPECTID").Value
End Sub

Sub EditProspect(xProsID As Long)

    PROSPECTID = xProsID
    '    cmdEdit.Value = True

End Sub

Sub FillSearchGrid(XXX As String)
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim LIMITKEY                                                      As String
    ListView1.Enabled = False

    Select Case cbostatus.ListIndex
        Case 0                                              'OPEN
            LIMITKEY = "('O' , NULL)"
        Case 1                                              'CLOSED
            LIMITKEY = "('C')"
        Case 2                                              ' INACTIVE
            LIMITKEY = "('I')"
        Case 3                                              ' INACTIVE
            LIMITKEY = "('L')"
        Case Else                                           'ALL
            LIMITKEY = "('O','C','I',NULL,'L')"
    End Select

    '    If optAcctName.Value = True Then
    ListView1.ColumnHeaders(1).Text = " PROSPECT NAME"
    If LOGSAE <> "" Then
        Set TEMPRS = gconDMIS.Execute("select  AcctName , SAE,MODEL  , PROSPECTID from CRIS_PROSPECTS WHERE USERCODE='" & LOGSAE & " ' AND  ACCTNAME like'" & ReplaceQuote(XXX) & "%'  AND status IN " & LIMITKEY & " order by 1  asc")
    Else
        Set TEMPRS = gconDMIS.Execute("select  AcctName , SAE,MODEL  , PROSPECTID  from CRIS_PROSPECTS where ACCTNAME like'" & ReplaceQuote(XXX) & "%' AND status IN " & LIMITKEY & " order by 1  asc")
    End If
    '    Else
    '        ListView1.ColumnHeaders(1).Text = " MODEL"
    '        If LOGSAE <> "" Then
    '            Set temprs = gconDMIS.Execute("select  AcctName , SAE,MODEL  + ISNULL('-' + VARIANT , ''), PROSPECTID  from CRIS_PROSPECTS WHERE USERCODE='" & LOGSAE & " ' AND MODEL  + ISNULL(VARIANT , '') like'" & ReplaceQuote(xxx) & "%' AND status IN " & LimitKey & " order by 1 asc")
    '        Else
    '            Set temprs = gconDMIS.Execute("select  AcctName , SAE,MODEL  + ISNULL('-' + VARIANT , ''), PROSPECTID  from CRIS_PROSPECTS WHERE LOGINITIALINQUIRY LIKE '" & ReplaceQuote(xxx) & "%' AND status IN " & LimitKey & " order by 1 asc")
    '        End If
    '    End If

    If Not TEMPRS.EOF And Not TEMPRS.BOF Then
        ListView1.Enabled = True
    End If


    flex_FillListView TEMPRS, ListView1


End Sub

'''''''''''''Conflicing Sub
Sub FindConfilict(ColumnKeyFind As String, MessageKey As String, tbox As TextBox)

    Dim TotalCount                                                    As Long

    If Trim(CUSCODE) = vbNullString Then
        'TotalCount = gconDMIS.Execute("Select COUNT(*) FROM CRIS_vW_AllProfile Where " & ColumnKeyFind & "=" & N2Str2Null(ReplaceQuote(tbox.Text))).Fields(0).Value
        TotalCount = gconDMIS.Execute("Select COUNT(*) FROM CRIS_Prospects Where " & ColumnKeyFind & "=" & N2Str2Null(ReplaceQuote(tbox.Text))).Fields(0).Value

        'Else
        '   TotalCount = gconDMIS.Execute("Select COUNT(*) FROM CRIS_vW_AllProfile Where CUSCDE<>" & N2Str2Null(CusCode) & " AND " & ColumnKeyFind & "=" & N2Str2Null(ReplaceQuote(tbox.Text))).Fields(0).Value

    End If

    If labid = 0 And TotalCount > 0 Then
        cmdSee(tbox.Tag).Enabled = True
        MessagePop InfoWarning, "Duplicate Name Detected", " There is Duplicated Entry Matching to This " & MessageKey & " Click On The Button [?] At Side to See Details"
    Else
        cmdSee(tbox.Tag).Enabled = False
    End If
End Sub

Sub InitData()
    txtDtInquiry.Text = Format(Now, "MM/dd/yyyy")
    FillCombo "Select ID, Name from SMIS_vw_Srep Order By 2 ", 0, 1, cboSAE
    FillCombo "Select DISTINCT MODEL FROM ALL_MODEL Order By 1 ", -1, 0, cboModel
    cboModel.AddItem "ALL UNIT", 0
    FillCombo "Select DISTINCT Color_Desc from ALl_COLOR Order By 1 ", -1, 0, cboColor
    FillCombo "SELECT Notes,  DataDesc  From CRIS_vW_MasterPullDown WHERE  MasterDesc ='Customer Classification' Order by 2", 0, 1, cboClassification
    FillCombo "SELECT DataDesc  From CRIS_vW_MasterPullDown WHERE  MasterDesc ='Lead Source' Order by 1", -1, 0, cboLeadSource
    AddColumnHeader "ACCOUNT NAME,SAE, MODEL", ListView1
    ResizeColumnHeader ListView1, "90,40,40"

    SetComboWidth cboVariant, 250
    SetComboWidth cboColor, 250
    SetComboWidth cboModel, 200
    With cboProspectType
        .AddItem ("Personal")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet")
        .ListIndex = 0
    End With
    With cbostatus
        .AddItem ("OPEN")                                   '0
        .AddItem ("CLOSED")                                 '1
        .AddItem ("INACTIVE")                               '2
        .AddItem ("LOST SALES")                             '3
        .AddItem ("(ANY)")                                  '4
        .ListIndex = 0
    End With
End Sub

Sub InitVar()
    txtProsCode = vbNullString
    txtDtInquiry.Text = FormatDateTime(Now, vbShortDate)
    cboSAE.Text = vbNullString
    txtProsName = vbNullString
    txtContactPerson = vbNullString
    txtTelNo = vbNullString
    txtEmailAdd = vbNullString
    txtCellPhone = vbNullString
    txtAddres = vbNullString
    cboModel = vbNullString
    cboVariant = vbNullString
    cboColor = vbNullString
    cboLeadSource = vbNullString
    cboClassification = vbNullString
    txtComments = vbNullString
    labCusCode = vbNullString
    chkFollowUps.Value = 0
    labStatus = ""
    labCusCode = ""
    labCustomerName = ""
    labConflictCUSCODE = ""

    txtCusCode = ""
    ' cmdOption_Add.Caption = "Existing Customer"
    CUSCODE = ""




End Sub

Sub rsRefresh()

    Set rsMain = New ADODB.Recordset

    If LOGSAE <> "" Then
        rsMain.Open "select * from CRIS_Prospects  WHERE USERCODE='" & LOGSAE & " ' order by ProspectID DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsMain.Open "select * from CRIS_Prospects  order by ProspectID DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

'''''''''''END COMBOS
''get set prospect type
Sub SetProspectType(XXX As String)
    If XXX = "P" Then
        cboProspectType.ListIndex = 0
    ElseIf XXX = "C" Then
        cboProspectType.ListIndex = 1
    ElseIf XXX = "G" Then
        cboProspectType.ListIndex = 2
    ElseIf XXX = "F" Then
        cboProspectType.ListIndex = 3
    Else
        cboProspectType.ListIndex = -1
    End If
End Sub

Sub SpecialProspectOnly(xCustCode)
    SpecialProspect = xCustCode
End Sub

Sub StoreMemVars()

    If Not rsMain.EOF And Not rsMain.BOF Then
        Dim TEMPRS                                                    As ADODB.Recordset
        labid = rsMain!PROSPECTID
        txtProsCode = Null2String(rsMain!ProsCode)
        txtDtInquiry.Text = Null2String(rsMain!loginitialinquiry)
        cboSAE.Text = Null2String(rsMain!SAE)
        txtProsName = Null2String(rsMain!AcctName)
        txtContactPerson = Null2String(rsMain!ContactPerson)
        txtTelNo = Null2String(rsMain!Telephone)
        txtEmailAdd = Null2String(rsMain!EMAIL)
        txtCellPhone = Null2String(rsMain!Mobile)
        txtAddres = Null2String(rsMain!Address)
        cboModel = Null2String(rsMain!Model)
        cboVariant = Null2String(rsMain!Variant)
        cboColor = Null2String(rsMain!Color)
        cboLeadSource = Null2String(rsMain!LeadSource)
        cboClassification = Null2String(rsMain!Classification)
        txtComments = Null2String(rsMain!Notes)
        labCusCode = Null2String(rsMain!CUSCDE)

        txtCusCode = Null2String(rsMain!CUSCDE)
        ''CUSTOMER NAME
        labCustomerName = ""
        If Not txtCusCode = "" Then
            Set TEMPRS = gconDMIS.Execute("SELECT customername from CRIS_vw_allprofile where cuscde='" & txtCusCode & "'")
            If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
                labCustomerName = " " & Null2String(TEMPRS!CUSTOMERNAME)
            End If
            Set TEMPRS = Nothing
        End If
        ''END CUSTOMER NAME
        CUSCODE = labCusCode
        SetProspectType Null2String(rsMain!ProspectType)

        If IsNull(rsMain!LogFollowupDate) = False Then
            chkFollowUps.Value = 1
            txtFollowupDate = rsMain!LogFollowupDate
        Else
            chkFollowUps.Value = 0
            txtFollowupDate = vbNullString
        End If

        StatusProspect = UCase(Null2String(rsMain!STATUS))

        '**************TEMPORARY REMOVAL FOR EDITING OF PROSPECT ONLY NEEDED IN ACTUAL RUN*****************************
        If StatusProspect = "C" Then
            cboProspectType.Enabled = False
            cboSAE.Enabled = False
            cboModel.Enabled = False
            cboVariant.Enabled = False
            cboColor.Enabled = False
            cboLeadSource.Enabled = False
            cmdDelete.Enabled = False
        Else
            cboProspectType.Enabled = True
            If LOGSAE = "" Then
                cboSAE.Enabled = True
            Else
                cboSAE.Enabled = False
            End If

            cboModel.Enabled = True
            cboVariant.Enabled = True
            cboColor.Enabled = True
            cboLeadSource.Enabled = True
            cmdDelete.Enabled = True
            '  cmdOption_Remove.Enabled = True
        End If

        '**************TEMPORARY REMOVAL FOR EDITING OF PROSPECT ONLY ***************************************


        If StatusProspect = "O" Then
            labStatus = "**OPEN**"
        ElseIf StatusProspect = "C" Then
            labStatus = "**CLOSED**"
        ElseIf StatusProspect = "I" Then
            labStatus = "**INACTIVE**"
        ElseIf StatusProspect = "L" Then
            labStatus = "**LOST SALES**"
        Else
            labStatus = "OPEN"
        End If


        'cboProspectType.Enabled = False
        cboProspectType.Enabled = Not IsDate(rsMain!LogApplication)

        If Null2String(rsMain!CUSCDE) <> "" Then
            '      cmdOption_Add.Caption = "Change Customer"

        Else
            '   cmdOption_Add.Caption = "Exisiting Customer"
            ' cmdOption_Remove.Enabled = False

        End If



    Else
        ShowNoRecord
        cmdAdd.Value = True

    End If
End Sub

Private Sub cboColor_Validate(Cancel As Boolean)
    cboColor.ListIndex = SelectCombo(cboColor, cboColor)
End Sub

Private Sub cboModel_CLick()
    If cboModel.ListIndex = -1 Then: Exit Sub
    If cboVariant.Tag = cboModel.Text Then: Exit Sub
    FillCombo "Select DISTINCT Descript FROM ALL_MODEL WHERE MODEL=" & N2Str2Null(cboModel.Text), -1, 0, cboVariant
    cboVariant.Tag = cboModel.Text
    'If cboVariant.ListCount > 0 Then
    'cboVariant.ListIndex = 0
    'End If
End Sub

Private Sub cboModel_GotFocus()
    'Set cComb.AttachCombo = cboModel
End Sub

Private Sub cboSAE_GotFocus()
    'Set cComb.AttachCombo = cboSAE
End Sub

Private Sub cboSAE_LostFocus()
    If cboSAE <> "" Then
        If SelectSAE(cboSAE, cboSAE) = False Then
            On Error Resume Next
            cboSAE = ""
        End If
    End If
End Sub

Private Sub cboVariant_GotFocus()
    'Set cComb.AttachCombo = cboVariant
End Sub

Private Sub chkFollowUps_Click()
    If chkFollowUps.Value = 1 Then
        txtFollowupDate.Enabled = True
        txtFollowupDate.Text = FormatDateTime(Now, vbShortDate)
        txtFollowupDate.BackColor = vbWhite
    Else
        txtFollowupDate.Enabled = False
        txtFollowupDate.BackColor = vbButtonFace
        txtFollowupDate.Text = vbNullString
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "PROSPECT") = False Then Exit Sub
    On Error GoTo ErrorCode:
    labid = 0
    InitVar
    picProspectEntry.Enabled = True
    PICSEARCH.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True
    txtProsCode = GenerateCode("CRIS_PROSPECTS", "PROSCODE", "0000000000")
    'cmdOption_Add.Enabled = True
    cboProspectType.Enabled = True
    cboModel.Enabled = True
    If LOGSAE = "" Then
        cboSAE.Enabled = True
    Else
        cboSAE.Text = SAENAME
        cboSAE.Enabled = False
    End If

    cboVariant.Enabled = True
    cboColor.Enabled = True
    cboLeadSource.Enabled = True
    AddorEdit = "ADD"
    On Error Resume Next
    txtDtInquiry.SetFocus
    Exit Sub


ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    picProspectEntry.Enabled = False
    PICSEARCH.Enabled = True
    labConflictCUSCODE = ""
    StoreMemVars

End Sub

Private Sub cmdCancelStatus_Click(Index As Integer)
    ShowHidePictureBox2 picAdvanceProspecting, False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "PROSPECT") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Delete that name...  " & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton1, "Message Box") = vbNo Then
        Exit Sub
    End If
    gconDMIS.Execute "Delete from CRIS_Prospects where PROSPECTID= " & labid
    rsRefresh
    StoreMemVars
    FillSearchGrid ""
    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "PROSPECT") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If StatusProspect = "C" Then
        MessagePop InfoWarning, "Prospect Closed", "Prospect Information Closed.. Editing Is Limited"
    ElseIf StatusProspect = "I" Then
        MessagePop InfoWarning, "Inactive Prospect", "Prospect Information Inactive.. Editing Is Limited"
    End If
    If LOGSAE <> "" Then
        txtDtInquiry.Enabled = False
    Else
        txtDtInquiry.Enabled = True
    End If
    picAdds.Visible = False
    picSaves.Visible = True
    AddorEdit = "EDIT"

    PICSEARCH.Enabled = False
    picProspectEntry.Enabled = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub cmdMM_ProspectLog_Click()
    'If LOGSAE = "" Then
    If Module_Access(LOGID, "PROSPECT CONVERSION", "SYSTEM") = False Then Exit Sub
    ShowHidePictureBox2 picAdvanceProspecting, True
    'Else
    '        MessagePop InfoVoid, "Access denied!", "Access denied! Contact Sys Ad!"
    'End If
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsMain.MoveNext
    If rsMain.EOF Then
        rsMain.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdOption_Add_Click()
    If SOEXISTS Then
        If MsgBox("Current Prospect has Sales Order." & vbCrLf & "Are you sure you want to Append it to Change Customer Information ?", vbYesNo Or vbExclamation Or vbDefaultButton1, App.TITLE) = vbNo Then: Exit Sub
    End If
    Set FormSearchCustomer = New frmSMIS_Mis_SearchMaster
    FormSearchCustomer.SearchForCustomers
    FormSearchCustomer.Show 1
End Sub

Private Sub cmdOption_Add_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labINFO.Caption = "Add Prospect Info Into Your Existing Customer Information " & vbCrLf & " USE THIS OPTION " & vbCrLf & " When Prospect Is Your Customer " & vbCrLf & " Prospect Is Returing Customer."
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdOption_Delete_Click
' DateTime  : 10/31/2007 23:30
' Author    : Ashish
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cmdOption_Delete_Click()
    On Error GoTo ErrorCode
    If SOEXISTS Then
        MsgSpeechBox (" Prospect Information already exist in Sales Order. Cannot delete the record.")
        Exit Sub
    End If
    '''''''''
    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT  COUNT(LOGNAME) as Total FROM CRIS_VIEWLOG WHERE  PROSPECTID=" & labid)
    If TEMPRS(0).Value > 1 Then
        On Error GoTo lnine
        oVoice.Speak "Prospect has " & TEMPRS(0).Value & " Log Entries." & vbCrLf & "Are You Sure You Want to Delete this Information ?", SVSFlagsAsync
lnine:
        If MsgBox("Prospect has " & TEMPRS(0).Value & " Log Entries." & vbCrLf & "Are You Sure You Want to Delete this Information ?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    '''''''''
    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("Delete  FROM CRIS_PROSPECTS WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM DBO.CRIS_PROSPECT_CALLS WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM DBO.CRIS_PROSPECT_EMAIL WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM   DBO.CRIS_PROSPECT_LETTER WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM   DBO.CRIS_PROSPECT_VISITS WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM   DBO.CRIS_QUOTATION WHERE PROSPECTID=" & labid)
        'gconDMIS.Execute ("Delete FROM   DBO.CRIS_REMINDERS WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM   DBO.CRIS_SALESAPPOINTMENTS WHERE PROSPECTID=" & labid)
        gconDMIS.Execute ("Delete FROM   dbo.CRIS_TestDriveSchedules WHERE PROSPECTID=" & labid)
        LogAudit "X", "PROSPECT DELETION", txtProsName
        MessagePop RecSave, "Prospect deleted", "Prospect Information sucessfully deleted"
        rsRefresh
        StoreMemVars
        ShowHidePictureBox2 picAdvanceProspecting, False
        
        FillSearchGrid ""
        
        Exit Sub
    End If


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdOption_Delete_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labINFO.Caption = "Check Current Prospect Information Against Your Database." & vbCrLf & " Searches Customer Database Againsts Your Current Prospect Information"
End Sub

Private Sub cmdOption_New_Click()

    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & labid)
    If Not TEMPRS.BOF Or Not TEMPRS.EOF Then
        Set FormAllCustomer = New frmAllCustomer
        FormAllCustomer.cmdAdd.Value = True
        Call FormAllCustomer.AddCustomerFromProspect(TEMPRS, "")
        FormAllCustomer.Show
    End If



End Sub

Private Sub cmdOption_New_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labINFO.Caption = "To Convert Into New Customer." & vbCrLf & " Use this Option to Convert Prospect Into New Customer"
End Sub

'
'Private Sub cmdOption_Remove_Click()
'
'    On Error GoTo ErrorCode:
'
'    If MsgBox("Are You Sure You Want to Remove Customer Information From This Prospect", vbQuestion + vbYesNo) = vbYes Then
'        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=NULL WHERE prospectID=" & labid)
'
'        cmdOption_Remove.Enabled = False
'        rsRefresh
'        rsMain.Find "id =" & labid
'        StoreMemVars
'
'
'    End If
'    Exit Sub
'ErrorCode:
'    ShowVBError
'
'End Sub

Private Sub cmdOption_Remove_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    labINFO.Caption = "To Remove Customer Information From Prospect Info "
End Sub

Private Sub cmdOption_View_Click()
    Call frmSMIS_Inquiry_ViewLog.SHOWPROSPECTLOG(labid, "")
    frmSMIS_Inquiry_ViewLog.Show
    LogAudit "V", "PROSPECT TRANSACTION DETAIL", txtProsName
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsMain.MovePrevious
    If rsMain.BOF Then
        rsMain.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdSave_Click()



    ' If Validateform("@R", Me, Timer1) = False Then: Exit Sub

    On Error GoTo ErrorCode:

    If LTrim(RTrim(cboSAE)) = "" Then
        ShowIsRequiredMsg " SAE NAME"
        On Error Resume Next
        cboSAE.SetFocus
        Exit Sub
    End If

    If SelectSAE(cboSAE, cboSAE) = False Then
        On Error Resume Next
        MsgBox (" INVALID SAE NAME ")
        cboSAE.SetFocus
        Exit Sub
    End If

    If LTrim(RTrim(txtProsName)) = "" Then
        ShowIsRequiredMsg " Prospect Name"
        On Error Resume Next
        txtProsName.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(cboProspectType)) = "" Then
        ShowIsRequiredMsg " Prospect Type"
        On Error Resume Next
        cboProspectType.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(cboModel)) = "" Then
        ShowIsRequiredMsg " Vehicle Model"
        On Error Resume Next
        cboModel.SetFocus
        Exit Sub
    End If


    Dim vtxtProsName                                                  As String
    Dim vtxtmodel                                                     As String
    Dim vtxtTelNo                                                     As String
    Dim vtxtAddres                                                    As String
    Dim vtxtComments                                                  As String
    Dim vtxtSAE                                                       As String
    Dim vtxtEmailAdd                                                  As String
    Dim vtxtVariant                                                   As String
    Dim vTxtColor                                                     As String
    Dim vtxtLeadSource                                                As String
    Dim vtxtClassification                                            As String
    Dim vtxtContactPerson                                             As String
    Dim vtxtCellPhone                                                 As String
    Dim vtxtProsCode                                                  As String
    Dim vtxtDtInquiry                                                 As String
    Dim vtxtFollowupDate                                              As String
    Dim vtxtCusCode                                                   As String
    Dim vtxtProspectType                                              As String
    Dim VTXTSACODE
    Dim prosID                                                        As Long
    Dim SQL                                                           As String

    If chkFollowUps.Value = 1 Then
        vtxtFollowupDate = N2Str2Null(txtFollowupDate.Text)
    Else
        vtxtFollowupDate = N2Str2Null("")
    End If

    vtxtProsCode = N2Str2Null(txtProsCode.Text)
    vtxtDtInquiry = N2Str2Null(txtDtInquiry.Text)
    vtxtSAE = N2Str2Null(cboSAE.Text)
    VTXTSACODE = N2Str2Null(GetSAECode(cboSAE))
    vtxtProsName = N2Str2Null(txtProsName.Text)
    vtxtContactPerson = N2Str2Null(txtContactPerson.Text)

    vtxtTelNo = N2Str2Null(txtTelNo.Text)
    vtxtCellPhone = N2Str2Null(txtCellPhone.Text)
    vtxtEmailAdd = N2Str2Null(txtEmailAdd.Text)
    vtxtAddres = N2Str2Null(txtAddres.Text)

    vtxtmodel = N2Str2Null(cboModel.Text)
    vtxtVariant = N2Str2Null(cboVariant.Text)
    vTxtColor = N2Str2Null(cboColor.Text)

    vtxtLeadSource = N2Str2Null(cboLeadSource.Text)
    vtxtClassification = N2Str2Null(cboClassification.Text)
    vtxtComments = N2Str2Null(txtComments.Text)
    vtxtCusCode = N2Str2Null(txtCusCode)


    vtxtProspectType = N2Str2Null(GetProspectType)

    Dim TEMPRS                                                        As ADODB.Recordset
    Dim rsHanap                                                       As ADODB.Recordset

    Set rsHanap = New ADODB.Recordset

    If labid = 0 Then

        SQL = " Insert INTO CRIS_Prospects  (  ProsCode, LogInitialInquiry, LogFollowUpDate ,SAE, AcctName, ContactPerson, Telephone, Mobile,Email, Address, Model, Variant, Color, LeadSource, Classification ,Notes,CUSCDE, ProspectType, Status,USERCODE,LASTUPDATED) "
        SQL = SQL & " values (  " & vbCrLf
        SQL = SQL & vtxtProsCode & ", " & vtxtDtInquiry & ", " & vtxtFollowupDate & ", " & vtxtSAE & ", " & vtxtProsName & ", " & vtxtContactPerson & "," & vbCrLf
        SQL = SQL & vtxtTelNo & ", " & vtxtCellPhone & ", " & vtxtEmailAdd & ", " & vtxtAddres & ", " & vbCrLf
        SQL = SQL & vtxtmodel & ", " & vtxtVariant & ", " & vTxtColor & ", " & vtxtLeadSource & ", " & vtxtClassification & ", " & vtxtComments & ", " & vtxtCusCode & ", " & vtxtProspectType & ", 'O'," & VTXTSACODE & "," & N2Str2Null(LOGDATE) & vbCrLf
        SQL = SQL & ")" & vbCrLf & "SELECT @@IDENTITY "
        Set TEMPRS = gconDMIS.Execute(SQL)
        SQL_STATEMENT = SQL
        Set rsHanap = gconDMIS.Execute("Select * From CRIS_Prospects WHERE ProsCode=" & Null2String(txtProsCode) & "")
        If Not (rsHanap.BOF Or rsHanap.EOF) Then
            prosID = Null2String(rsHanap!PROSPECTID)
        End If
        NEW_LogAudit "A", "PROSPECT", SQL_STATEMENT, Null2String(prosID), "", "Customer ID: " & Null2String(prosID), "", ""
    Else
        SQL = " UPDATE CRIS_Prospects  SET " & vbCrLf
        SQL = SQL & " LogInitialInquiry=" & vtxtDtInquiry
        SQL = SQL & " , LogFollowUpDate=" & vtxtFollowupDate
        SQL = SQL & " , SAE=" & vtxtSAE
        SQL = SQL & " , AcctName=" & vtxtProsName
        SQL = SQL & " , ContactPerson=" & vtxtContactPerson
        SQL = SQL & " , Telephone=" & vtxtTelNo
        SQL = SQL & " , Mobile=" & vtxtCellPhone
        SQL = SQL & " , Email=" & vtxtEmailAdd
        SQL = SQL & " , Address=" & vtxtAddres
        SQL = SQL & " , Model=" & vtxtmodel
        SQL = SQL & " , Variant=" & vtxtVariant
        SQL = SQL & " , Color=" & vTxtColor
        SQL = SQL & " , LeadSource=" & vtxtLeadSource
        SQL = SQL & " , Classification=" & vtxtClassification
        SQL = SQL & " , Notes=" & vtxtComments
        SQL = SQL & " , ProspectType=" & vtxtProspectType
        SQL = SQL & " , CUSCDE=" & vtxtCusCode
        SQL = SQL & " , USERCODE=" & N2Str2Null(GetSAECode(cboSAE))
        SQL = SQL & " , LASTUPDATED=" & N2Str2Null(LOGDATE)
        SQL = SQL & " WHERE PROSPECTID=" & labid
        Set TEMPRS = gconDMIS.Execute(SQL)
        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "PROSPECT", SQL_STATEMENT, Null2String(labid), "", "Customer ID:" & Null2String(labid), "", ""

    End If



    If labid <= 0 Then
        MessagePop RecSaveOk, "Record Added ", "New Prospect Sucessfully Added"
    Else
        MessagePop RecSaveOk, "RecordSaved", "Prospect  Information Updated "
    End If


    Set TEMPRS = TEMPRS.NextRecordset
    If Not TEMPRS Is Nothing Then
        labid = TEMPRS.Collect(0)
    End If

    If AddorEdit = "ADD" And chkFollowUps.Value = 1 Then
        SQL = "INSERT INTO CRIS_Reminders "
        SQL = SQL & " (USERID, CSCDE, ReminderType, DateTimeRemind, ReminderNotes, Subject,Snoozed,NextTime, USERCODE,ENTITYTYPE,Priority) "
        SQL = SQL & " VALUES("
        SQL = SQL & N2Str2Null(LOGID) & ","
        SQL = SQL & N2Str2Null(labid) & ",'FOLLOW UP',"
        SQL = SQL & N2Str2Null(txtFollowupDate) & ","
        SQL = SQL & N2Str2Null(txtComments) & ",'PROSPECT FOLLOW UPS'" & ", 0, "
        SQL = SQL & N2Str2Null(txtFollowupDate) & "," & VTXTSACODE & ", 'P' ,'H' )"
        gconDMIS.Execute SQL
    End If

    Set TEMPRS = Nothing
    rsRefresh
    If AddorEdit = "EDIT" Then
        rsMain.Find "PROSPECTID =" & labid
    End If



    cmdCancel.Value = True
    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    If FormExist("MainSAE") Then
        MainSAE.ShowData
    End If

    FillSearchGrid txtSEARCH
    LogAudit "A", "PROSPECT INFORMATION", "Prospect No:" & txtProsCode & "-" & txtProsName    '''''*RYAN DC CULAWAY MAY 24,08
    Exit Sub

ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSee_Click(Index As Integer)

    Select Case Index
        Case 0
            Call FormConflictingCustomer.Conflict("TELEPHONE", txtTelNo.Text, CUSCODE)
            FormConflictingCustomer.lblError = " There is Already Prospect(s) With Such Phone Number.. Below Are the Details of Customer With Conflicting Phone Number"
        Case 2
            Call FormConflictingCustomer.Conflict("Mobile", txtCellPhone.Text, CUSCODE)
            FormConflictingCustomer.lblError = " There is Already Prospect(s) With Such Cell Phone Number .. Below Are the Details of Prospect With Conflicting Cell Phone Number"
        Case 1
            Call FormConflictingCustomer.Conflict("Email", txtEmailAdd.Text, CUSCODE)
            FormConflictingCustomer.lblError = " There is Already Prospect(s) With Such Email Address.. Below Are the Details of Prospect With Conflicting Email Address"
    End Select

    FormConflictingCustomer.Show vbModal
End Sub

Private Sub cboStatus_Click()
    FillSearchGrid (txtSEARCH.Text)
End Sub

Private Sub Command1_Click()
    InitVar
    Set SEARCHFORM = New frmSMIS_Mis_SearchMaster
    SEARCHFORM.SearchForCustomers
    SEARCHFORM.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If picAdvanceProspecting.Visible = True Then
            ShowHidePictureBox2 picAdvanceProspecting, False
            Exit Sub
        End If
        'If picAdds.Visible = True Then
        '    Unload Me
        'Else
        '   cmdCancel.Value = True
        'End If
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PROSPECTING)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "PROSPECT")
            'End If
    End Select

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitVar
    InitData
    rsRefresh
    picAdds.Visible = True
    picSaves.Visible = False
    picProspectEntry.Enabled = False
    If PROSPECTID > 0 Then
        rsMain.Find ("ProspectID=" & PROSPECTID)
    End If
    StoreMemVars
    Set FormConflictingCustomer = New frmSMIS_Trans_Confilct
    FillSearchGrid txtSEARCH

End Sub

Private Sub Form_Unload(Cancel As Integer)
    labid = 0
    SpecialProspect = vbNullString
    Set rsMain = Nothing
    Set ctl = Nothing
    PROSPECTID = 0
    Set FormSearchCustomer = Nothing
    Set FormConflictingCustomer = Nothing
End Sub

Private Sub FormAllCustomer_ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)


    gconDMIS.Execute ("Update CRIS_PROSPECTS SET  CUSCDE='" & CustomerCode & "' WHERE PROSPECTID=" & labid)
    Call MsgBox(" Prospect Sucessfully Converted Into Customer ", vbInformation)
    rsRefresh
    LogAudit "A", "PROSPECT CONVERSION", "FROM " & txtProsName & " TO " & CustomerCode

    rsMain.Find ("PROSPECTID=" & labid)
    StoreMemVars
    Unload FormAllCustomer
    Set FormAllCustomer = Nothing
    Exit Sub





End Sub

Private Sub FormConflictingCustomer_SelectionMade(oRs As ADODB.Recordset)


    On Error GoTo ErrorCode

    'labConflictCUSCODE = Null2String(oRs!CUSCDE)
    AddorEdit = ""
    picAdds.Visible = True
    picSaves.Visible = False
    rsMain.Find ("ProspectID=" & oRs!PROSPECTID)
    StoreMemVars
    'labCusCode = Null2String(oRs!CUSCDE)
    'CusCode = Null2String(oRs!CUSCDE)
    'txtContactPerson.Text = Null2String(oRs!ContactPerson)
    'txtTelNo.Text = Null2String(oRs!Telephone)
    'txtAddres.Text = Null2String(oRs!Address)
    ' txtCellPhone.Text = Null2String(oRs!Mobile)
    ' txtEmailAdd.Text = Null2String(oRs!EMAIL)

    Dim i                                                             As Integer
    For i = 0 To cmdSee.Count - 1
        cmdSee(i).Enabled = False
    Next
    Unload FormConflictingCustomer


    Exit Sub
ErrorCode:
    ShowVBError
    Unload FormConflictingCustomer

End Sub

Private Sub FormCustomer_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtCusCode = Null2String(oCusRs!CUSCDE)
    Unload FormCustomer
End Sub

Private Sub FormSearchCustomer_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    If labid > 0 Then
        If XSelection = "CUSTOMER" Then
            If MsgBox(" Do You Want To Add Current Prospect To " & Null2String(oCusRs!AcctName), vbYesNo + vbInformation) = vbNo Then Exit Sub
            gconDMIS.Execute ("Update CRIS_PROSPECTS SET PROSPECTTYPE='" & Null2String(oCusRs("CUSTYPE")) & "' , CUSCDE=" & N2Str2Null(oCusRs("CUSCDE")) & " WHERE PROSPECTID=" & labid)
            MessagePop RecSave, "Prospect Converted", "Prospect Information has been converted Sucessfully"
            rsRefresh
            rsMain.Find ("PROSPECTID=" & labid)
            StoreMemVars
            Unload FormSearchCustomer
            Set FormSearchCustomer = Nothing
            Exit Sub
        End If

    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With ListView1
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

Private Sub LISTVIEW1_DblClick()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    cmdEdit.Value = True
End Sub

Private Sub ListView1_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsMain.MoveFirst
    rsMain.Find ("PROSPECTID=" & ITEM.ListSubItems(3).Text)
    StoreMemVars

End Sub

Private Sub SEARCHFORM_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtCusCode = Null2String(oCusRs!CUSCDE)
    labCustomerName = Null2String(oCusRs!lastname) & ", " & Null2String(oCusRs!Firstname) & "." & Left(Null2String(oCusRs!Firstname), 1)
    txtProsCode = GenerateCode("CRIS_PROSPECTS", "PROSCODE", "0000000000")
    txtProsName = Null2String(oCusRs!AcctName)
    txtContactPerson = Null2String(oCusRs!AcctName)
    txtTelNo = Null2String(oCusRs!TelephoneNo)
    txtEmailAdd = Null2String(oCusRs!EMAIL)
    txtCellPhone = Null2String(oCusRs!Mobile)
    txtAddres = Null2String(oCusRs!CUSTOMERADD)
    picAdds.Visible = False
    picSaves.Visible = True
    picProspectEntry.Enabled = True
    PICSEARCH.Enabled = False
    labid = 0
    AddorEdit = "ADD"
    Unload SEARCHFORM
End Sub

Private Sub Timer1_Timer()
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If
        End If
    Next
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    If labStatus.Caption <> "" Then
        If labStatus.Visible = True Then
            labStatus.Visible = False
        Else
            labStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtAddres_GotFocus()
    txtAddres.SelStart = Len(txtAddres)
    txtAddres.SelLength = 0
End Sub

Private Sub txtCellPhone_Change()
    cmdSee(2).Enabled = False
End Sub

Private Sub txtCellPhone_LostFocus()
    If Len(Trim(txtCellPhone)) = 0 Or cmdSee(2).Enabled = True Then: Exit Sub
    Call FindConfilict(" Mobile ", " CellPhome Number", txtCellPhone)
End Sub

Private Sub txtComments_LostFocus()
    txtComments = Trim(txtComments)
End Sub

Private Sub txtContactPerson_Gotfocus()
    'Set cComb.AttachTextBox = txtContactPerson
End Sub

Private Sub txtContactPerson_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCusCode_Change()
    If txtCusCode <> "" Then
        cmdOption_New.Enabled = False
    Else
        cmdOption_New.Enabled = True
    End If

End Sub

Private Sub txtDtInquiry_Validate(Cancel As Boolean)
    If IsDate(txtDtInquiry.Text) = False Then
        txtDtInquiry.Text = FormatDateTime(Now, vbShortDate)
    End If
End Sub

Private Sub txtEmailAdd_Change()
    cmdSee(1).Enabled = False
End Sub

Private Sub txtEmailAdd_LostFocus()
    If Len(Trim(txtEmailAdd)) = 0 Or cmdSee(1).Enabled = True Or Len(labCusCode) <> 0 Then: Exit Sub
    Call FindConfilict("Email", " Email Address", txtEmailAdd)
End Sub

Private Sub txtFollowupDate_Validate(Cancel As Boolean)
    If IsDate(txtFollowupDate.Text) = False Then
        txtFollowupDate.Text = FormatDateTime(Now, vbShortDate)
    End If
End Sub

Private Sub txtProsCode_LostFocus()
    txtProsCode = Format(txtProsCode, "0000000000")
End Sub

Private Sub txtProsName_Change()
    If labid = 0 Then
        txtContactPerson.Text = txtProsName.Text
    End If
    'cmdOption_Add.Enabled = IIf(Len(Trim(txtProsName.Text)) > 0, True, False)
End Sub

Private Sub txtProsName_GotFocus()
    'Set cComb.AttachTextBox = txtProsName
End Sub

Private Sub txtProsName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSEARCH.Text)
End Sub

Private Sub txtTelNo_Change()
    cmdSee(0).Enabled = False
End Sub

Private Sub txtTelNo_LostFocus()
    If Len(Trim(txtTelNo)) = 0 Or cmdSee(0).Enabled = True Then: Exit Sub
    Call FindConfilict("TELEPHONE", " Telephone Number", txtTelNo)

End Sub

