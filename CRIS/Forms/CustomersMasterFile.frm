VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAllCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   8910
   ClientLeft      =   525
   ClientTop       =   735
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CustomersMasterFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11895
   Begin VB.PictureBox picContactAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFCCCF&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   5280
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4305
      ScaleWidth      =   4350
      TabIndex        =   105
      Top             =   1380
      Visible         =   0   'False
      Width           =   4380
      Begin VB.TextBox txtContactName 
         Height          =   345
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   109
         Top             =   390
         Width           =   3045
      End
      Begin VB.CommandButton cmdCloseContactsAE 
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
         Height          =   675
         Index           =   1
         Left            =   3600
         MouseIcon       =   "CustomersMasterFile.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Cancel Entry"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdSaveContact 
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
         Height          =   675
         Left            =   2970
         MouseIcon       =   "CustomersMasterFile.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Save Details"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdDeleteContact 
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
         Height          =   675
         Left            =   2340
         MouseIcon       =   "CustomersMasterFile.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Delect Details"
         Top             =   3480
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseContactsAE 
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
         Height          =   315
         Index           =   0
         Left            =   3990
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cboContactRelation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "CustomersMasterFile.frx":1679
         Left            =   1140
         List            =   "CustomersMasterFile.frx":167B
         TabIndex        =   111
         Top             =   790
         Width           =   3045
      End
      Begin VB.TextBox txtContactPosition 
         Height          =   345
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   113
         Top             =   1190
         Width           =   3045
      End
      Begin VB.TextBox txtContactDepartment 
         Height          =   345
         Left            =   1140
         MaxLength       =   40
         TabIndex        =   114
         Top             =   1590
         Width           =   3045
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   345
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   116
         Top             =   1990
         Width           =   3045
      End
      Begin VB.TextBox txtContactMobile 
         Height          =   345
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   118
         Top             =   2390
         Width           =   3045
      End
      Begin VB.TextBox txtContactAddress 
         Height          =   645
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   121
         Top             =   2790
         Width           =   3045
      End
      Begin VB.Label labIDContacts 
         Height          =   555
         Left            =   1350
         TabIndex        =   122
         Top             =   3570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Relation:"
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
         Left            =   375
         TabIndex        =   110
         Top             =   870
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   330
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "CONTACTS INFORMATION"
         ForeColor       =   -2147483630
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
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   570
         TabIndex        =   108
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
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
         Left            =   375
         TabIndex        =   112
         Top             =   1290
         Width           =   735
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
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
         Left            =   60
         TabIndex        =   115
         Top             =   1710
         Width           =   1050
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone:"
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
         Left            =   525
         TabIndex        =   117
         Top             =   2130
         Width           =   585
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile:"
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
         Left            =   510
         TabIndex        =   119
         Top             =   2550
         Width           =   600
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         Left            =   345
         TabIndex        =   120
         Top             =   2970
         Width           =   765
      End
   End
   Begin VB.TextBox labOLDCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   12450
      TabIndex        =   126
      Top             =   2430
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   12450
      TabIndex        =   98
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin Crystal.CrystalReport rptCustomer 
      Left            =   1230
      Top             =   8430
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picContactList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   4650
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   99
      Top             =   1350
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdCancelContactList 
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
         Height          =   645
         Left            =   5010
         MouseIcon       =   "CustomersMasterFile.frx":167D
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":17CF
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Cancel"
         Top             =   4110
         Width           =   705
      End
      Begin VB.CommandButton cmdEditContact 
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
         Height          =   645
         Left            =   4320
         MouseIcon       =   "CustomersMasterFile.frx":1B0D
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":1C5F
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Edit Contact"
         Top             =   4110
         Width           =   705
      End
      Begin MSComctlLib.ListView lvContactList 
         Height          =   3735
         Left            =   60
         TabIndex        =   101
         Top             =   330
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6588
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomersMasterFile.frx":1FBB
         NumItems        =   0
      End
      Begin VB.CommandButton Command4 
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
         Height          =   645
         Left            =   3630
         MouseIcon       =   "CustomersMasterFile.frx":211D
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":226F
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Add Contact"
         Top             =   4110
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   1
         Left            =   -30
         TabIndex        =   100
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   ":: LIST OF CONTACTS::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.01
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
   Begin VB.PictureBox picCredit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00A9B8C2&
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   4650
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2295
      ScaleWidth      =   3390
      TabIndex        =   140
      Top             =   2760
      Visible         =   0   'False
      Width           =   3420
      Begin VB.TextBox txtCreditDays 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   146
         Text            =   "Text1"
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   420
         Width           =   1875
      End
      Begin VB.CommandButton cmdCloseTerm 
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
         Height          =   675
         Index           =   1
         Left            =   2460
         MouseIcon       =   "CustomersMasterFile.frx":2582
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":26D4
         Style           =   1  'Graphical
         TabIndex        =   149
         ToolTipText     =   "Cancel Entry"
         Top             =   1410
         Width           =   645
      End
      Begin VB.CommandButton Command12 
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
         Height          =   675
         Left            =   1830
         MouseIcon       =   "CustomersMasterFile.frx":2A12
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":2B64
         Style           =   1  'Graphical
         TabIndex        =   148
         ToolTipText     =   "Save Entry"
         Top             =   1410
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseTerm 
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
         Height          =   315
         Index           =   0
         Left            =   3090
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Days:"
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
         Left            =   180
         TabIndex        =   145
         Top             =   930
         Width           =   1020
      End
      Begin VB.Label labTermID 
         Height          =   555
         Left            =   360
         TabIndex        =   147
         Top             =   1320
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit:"
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
         Left            =   690
         TabIndex        =   143
         Top             =   480
         Width           =   465
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   330
         Left            =   0
         TabIndex        =   142
         Top             =   0
         Width           =   3405
         _Version        =   655364
         _ExtentX        =   6006
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::CREDITS AND TERMS::"
         ForeColor       =   -2147483630
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
         Alignment       =   1
         ForeColor       =   -2147483630
      End
   End
   Begin VB.PictureBox picChildAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFFDFD&
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   3930
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2475
      ScaleWidth      =   4350
      TabIndex        =   127
      Top             =   3180
      Visible         =   0   'False
      Width           =   4380
      Begin VB.CommandButton cmdCloseChildAE 
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
         Height          =   675
         Index           =   3
         Left            =   3480
         MouseIcon       =   "CustomersMasterFile.frx":2EB4
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":3006
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Cancel Entry"
         Top             =   1650
         Width           =   645
      End
      Begin VB.CommandButton cmdSaveChild 
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
         Height          =   675
         Left            =   2850
         MouseIcon       =   "CustomersMasterFile.frx":3344
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":3496
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Save Children Information"
         Top             =   1650
         Width           =   645
      End
      Begin VB.CommandButton cmdCloseChildAE 
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
         Height          =   315
         Index           =   2
         Left            =   3990
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdDeleteChild 
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
         Height          =   675
         Left            =   2220
         MouseIcon       =   "CustomersMasterFile.frx":37E6
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":3938
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Add Children Information"
         Top             =   1650
         Width           =   645
      End
      Begin VB.TextBox txtChildName 
         Height          =   345
         Left            =   1200
         TabIndex        =   131
         Top             =   390
         Width           =   3015
      End
      Begin VB.ComboBox cboChildSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "CustomersMasterFile.frx":3C63
         Left            =   1200
         List            =   "CustomersMasterFile.frx":3C70
         TabIndex        =   135
         Top             =   1170
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtChildDate 
         Height          =   345
         Left            =   1200
         TabIndex        =   133
         Top             =   780
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   "mm/dd/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
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
         Left            =   570
         TabIndex        =   130
         Top             =   390
         Width           =   540
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::CHILDREN INFORMATION::"
         ForeColor       =   -2147483630
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
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Brith:"
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
         Left            =   45
         TabIndex        =   132
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label labIdCHILD 
         Height          =   555
         Left            =   1290
         TabIndex        =   139
         Top             =   1800
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SEX:"
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
         Left            =   720
         TabIndex        =   134
         Top             =   1200
         Width           =   390
      End
   End
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   9315
      Left            =   0
      ScaleHeight     =   9315
      ScaleWidth      =   12345
      TabIndex        =   0
      Top             =   0
      Width           =   12345
      Begin VB.PictureBox picToolFrame 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2460
         ScaleHeight     =   855
         ScaleWidth      =   9855
         TabIndex        =   10
         Top             =   0
         Width           =   9855
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   8550
            Top             =   570
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmdCUSTINFO_Contact 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   3600
            MouseIcon       =   "CustomersMasterFile.frx":3C7D
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":3DCF
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Contact Information"
            Top             =   300
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Child 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   2100
            MouseIcon       =   "CustomersMasterFile.frx":44C1
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":4613
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "View Customers Number of Children"
            Top             =   300
            Width           =   585
         End
         Begin VB.CommandButton cmdCustInfo_Credit 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            MouseIcon       =   "CustomersMasterFile.frx":4C34
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":4D86
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Update Credit and Terms of Customers"
            Top             =   300
            Width           =   585
         End
         Begin VB.Label labCustInfo_Contact 
            Caption         =   "Contact Information"
            Height          =   195
            Left            =   4260
            MouseIcon       =   "CustomersMasterFile.frx":53E9
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   480
            Width           =   1995
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   270
            Index           =   2
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Width           =   9780
            _Version        =   655364
            _ExtentX        =   17251
            _ExtentY        =   476
            _StockProps     =   14
            Caption         =   "Customers Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.74
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
         Begin VB.Label labCustInfo_Child 
            Caption         =   "Children"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2700
            MouseIcon       =   "CustomersMasterFile.frx":56F3
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   420
            Width           =   975
         End
         Begin VB.Label labCustInfo_Credit 
            Caption         =   "Credit && Terms"
            Height          =   225
            Left            =   780
            MouseIcon       =   "CustomersMasterFile.frx":59FD
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   465
            Width           =   1335
         End
      End
      Begin VB.Frame fraSearch 
         Height          =   8865
         Left            =   0
         TabIndex        =   1
         Top             =   -90
         Width           =   2475
         Begin VB.TextBox txtSearch 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   60
            MaxLength       =   35
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   1860
            Width           =   2295
         End
         Begin VB.OptionButton optSearchKeyLast 
            Caption         =   "Search By Last Name"
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
            Height          =   255
            Left            =   60
            MouseIcon       =   "CustomersMasterFile.frx":5D07
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   180
            Width           =   2295
         End
         Begin VB.OptionButton optSearchKeyCompany 
            Caption         =   "Search By Company"
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
            Height          =   255
            Left            =   60
            MouseIcon       =   "CustomersMasterFile.frx":5E59
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   435
            Width           =   2295
         End
         Begin VB.OptionButton optSearchKeyAcctName 
            Caption         =   "Search By A/C Name"
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
            Height          =   255
            Left            =   60
            MouseIcon       =   "CustomersMasterFile.frx":5FAB
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   690
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton optSearchKeyAddress 
            Caption         =   "Search By Address"
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
            Height          =   255
            Left            =   60
            MouseIcon       =   "CustomersMasterFile.frx":60FD
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   945
            Width           =   2295
         End
         Begin VB.OptionButton optSearchKeyEmail 
            Caption         =   "Search By Email"
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
            Height          =   255
            Left            =   60
            MouseIcon       =   "CustomersMasterFile.frx":624F
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1200
            Width           =   2295
         End
         Begin VB.ComboBox cboSearchCustype 
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
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1500
            Width           =   2355
         End
         Begin MSComctlLib.ListView lstCustomer 
            Height          =   6525
            Left            =   30
            TabIndex        =   9
            Top             =   2280
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   11509
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
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "CustomersMasterFile.frx":63A1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "A/C Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   10440
         ScaleHeight     =   885
         ScaleWidth      =   1590
         TabIndex        =   74
         Top             =   7920
         Width           =   1590
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
            Left            =   690
            MouseIcon       =   "CustomersMasterFile.frx":6503
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":6655
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Cancel"
            Top             =   75
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
            Left            =   0
            MouseIcon       =   "CustomersMasterFile.frx":6993
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":6AE5
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Save this Record"
            Top             =   75
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7125
         Left            =   2520
         TabIndex        =   18
         Top             =   780
         Width           =   9315
         Begin VB.Frame Frame4 
            Caption         =   "Delivery Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1335
            Left            =   4290
            TabIndex        =   69
            Top             =   2610
            Width           =   4995
            Begin VB.TextBox txtDeliveryAddress 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   825
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   420
               Width           =   4845
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Same As above"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3480
               TabIndex        =   70
               Top             =   150
               Width           =   1395
            End
         End
         Begin VB.TextBox txtAcctName 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   5190
            MaxLength       =   100
            TabIndex        =   22
            Top             =   195
            Width           =   3945
         End
         Begin VB.ComboBox cboCustType 
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
            ItemData        =   "CustomersMasterFile.frx":6E35
            Left            =   1320
            List            =   "CustomersMasterFile.frx":6E37
            TabIndex        =   19
            Text            =   "cboCustType"
            Top             =   180
            Width           =   2835
         End
         Begin VB.Frame Frame3 
            Caption         =   "Notes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3165
            Left            =   4290
            TabIndex        =   72
            Top             =   3900
            Width           =   4965
            Begin VB.TextBox txtNotes 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   2835
               Left            =   60
               MaxLength       =   300
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   73
               Top             =   240
               Width           =   4845
            End
         End
         Begin VB.Frame fraEntity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   2100
            Left            =   30
            TabIndex        =   23
            Top             =   450
            Width           =   9225
            Begin VB.ComboBox cboPersonalCity 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   4200
               TabIndex        =   37
               Text            =   "cboApod"
               Top             =   1020
               Width           =   1995
            End
            Begin VB.TextBox txtPersonalStreet 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   375
               Left            =   120
               ScrollBars      =   2  'Vertical
               TabIndex        =   36
               Top             =   1020
               Width           =   4035
            End
            Begin VB.TextBox txtPersonalState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   6240
               MaxLength       =   30
               TabIndex        =   38
               Top             =   1020
               Width           =   1695
            End
            Begin VB.TextBox txtPersonalZIP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   7980
               MaxLength       =   6
               TabIndex        =   39
               Top             =   1020
               Width           =   1155
            End
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1020
               TabIndex        =   29
               ToolTipText     =   "LAST NAME OR COMPANY NAME"
               Top             =   420
               Width           =   2715
            End
            Begin VB.ComboBox cboApod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   28
               Text            =   "cboApod"
               Top             =   420
               Width           =   855
            End
            Begin VB.TextBox txtMiddleName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   6435
               MaxLength       =   50
               TabIndex        =   31
               Top             =   420
               Width           =   2700
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3780
               TabIndex        =   30
               Top             =   420
               Width           =   2625
            End
            Begin VB.ComboBox cboSex 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2040
               TabIndex        =   44
               Text            =   "cboSex"
               Top             =   1650
               Width           =   855
            End
            Begin VB.TextBox txtBirthDate 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   120
               MaxLength       =   10
               TabIndex        =   43
               Top             =   1650
               Width           =   1875
            End
            Begin VB.TextBox txtSpouse 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   2940
               MaxLength       =   100
               TabIndex        =   45
               Top             =   1650
               Width           =   6195
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "City"
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
               Height          =   225
               Left            =   4200
               TabIndex        =   34
               Top             =   810
               Width           =   315
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Street"
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
               Height          =   225
               Left            =   120
               TabIndex        =   33
               Top             =   810
               Width           =   525
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "State/Province"
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
               Height          =   225
               Left            =   6240
               TabIndex        =   35
               Top             =   810
               Width           =   1245
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Zip Code"
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
               Height          =   225
               Left            =   7980
               TabIndex        =   32
               Top             =   780
               Width           =   735
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Salutation"
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
               Height          =   225
               Left            =   120
               TabIndex        =   24
               Top             =   150
               Width           =   855
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
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
               Height          =   225
               Left            =   6420
               TabIndex        =   27
               Top             =   210
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
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
               Height          =   225
               Left            =   3780
               TabIndex        =   26
               Top             =   210
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
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
               Height          =   225
               Left            =   1050
               TabIndex        =   25
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
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
               Height          =   225
               Left            =   2010
               TabIndex        =   42
               Top             =   1440
               Width           =   330
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
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
               Height          =   225
               Left            =   120
               TabIndex        =   40
               Top             =   1410
               Width           =   840
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Name"
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
               Height          =   225
               Left            =   2910
               TabIndex        =   41
               Top             =   1410
               Width           =   1185
            End
         End
         Begin VB.Frame fraMiscellenous 
            Height          =   4455
            Left            =   60
            TabIndex        =   46
            Top             =   2610
            Width           =   4185
            Begin VB.TextBox txtTin 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               MaxLength       =   15
               TabIndex        =   48
               Top             =   210
               Width           =   2775
            End
            Begin VB.TextBox txtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   64
               Top             =   3300
               Width           =   2775
            End
            Begin VB.TextBox txtHomePhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   62
               Top             =   2925
               Width           =   2775
            End
            Begin VB.TextBox txtMobile 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   60
               Top             =   2550
               Width           =   2775
            End
            Begin VB.TextBox txtCusphon1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   58
               Top             =   2175
               Width           =   2775
            End
            Begin VB.TextBox txtAsstPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   68
               Top             =   3975
               Width           =   2775
            End
            Begin VB.TextBox txtAssistant 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   66
               Top             =   3675
               Width           =   2775
            End
            Begin VB.TextBox txtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   56
               Top             =   1785
               Width           =   2775
            End
            Begin VB.TextBox txtDepartment 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   1245
               TabIndex        =   54
               Top             =   1380
               Width           =   2775
            End
            Begin VB.TextBox txtTitle 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1245
               TabIndex        =   52
               Top             =   1005
               Width           =   2775
            End
            Begin VB.ComboBox cboLeadSource 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1245
               TabIndex        =   50
               Text            =   "cboLeadSource"
               Top             =   615
               Width           =   2775
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Tin Number"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   47
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   870
               TabIndex        =   63
               Top             =   3165
               Width           =   285
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Home Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   61
               Top             =   2805
               Width           =   1095
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   615
               TabIndex        =   59
               Top             =   2445
               Width           =   540
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Office Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   57
               Top             =   2085
               Width           =   1050
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Asst. Phone"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   165
               TabIndex        =   67
               Top             =   3885
               Width           =   990
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Assistant"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   390
               TabIndex        =   65
               Top             =   3525
               Width           =   765
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   675
               TabIndex        =   55
               Top             =   1725
               Width           =   480
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   180
               TabIndex        =   53
               Top             =   1365
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   480
               TabIndex        =   51
               Top             =   1005
               Width           =   675
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Lead Source"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   105
               TabIndex        =   49
               Top             =   645
               Width           =   1050
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Acct. Name"
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
            Left            =   4200
            TabIndex        =   21
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Type"
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
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   960
         Left            =   -90
         ScaleHeight     =   960
         ScaleWidth      =   12315
         TabIndex        =   77
         Top             =   7920
         Width           =   12315
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
            Left            =   11220
            MouseIcon       =   "CustomersMasterFile.frx":6E39
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":6F8B
            Style           =   1  'Graphical
            TabIndex        =   87
            ToolTipText     =   "Exit Window"
            Top             =   75
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
            Left            =   10530
            MouseIcon       =   "CustomersMasterFile.frx":72F1
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":7443
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Print this Record"
            Top             =   75
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
            Left            =   9845
            MouseIcon       =   "CustomersMasterFile.frx":77A9
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":78FB
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Delete Selected Record"
            Top             =   75
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
            Left            =   9150
            MouseIcon       =   "CustomersMasterFile.frx":7C26
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":7D78
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Edit Selected Record"
            Top             =   75
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
            Left            =   8465
            MouseIcon       =   "CustomersMasterFile.frx":80D4
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":8226
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Add Record"
            Top             =   75
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
            Left            =   7740
            MouseIcon       =   "CustomersMasterFile.frx":8539
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":868B
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Move to Last Record"
            Top             =   75
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
            Left            =   7025
            MouseIcon       =   "CustomersMasterFile.frx":89DB
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":8B2D
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Move to First Record"
            Top             =   75
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
            Left            =   6330
            MouseIcon       =   "CustomersMasterFile.frx":8E8B
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":8FDD
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Find a Record"
            Top             =   75
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
            Left            =   5645
            MouseIcon       =   "CustomersMasterFile.frx":92D7
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":9429
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Move to Next Record"
            Top             =   75
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
            Left            =   4950
            MouseIcon       =   "CustomersMasterFile.frx":9781
            MousePointer    =   99  'Custom
            Picture         =   "CustomersMasterFile.frx":98D3
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Move to Previous Record"
            Top             =   75
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picChildList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   3810
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   91
      Top             =   1530
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton cmdCancelChildList 
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
         Height          =   645
         Left            =   5040
         MouseIcon       =   "CustomersMasterFile.frx":9C32
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":9D84
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Cancel"
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton cmdSelectChild 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4350
         MouseIcon       =   "CustomersMasterFile.frx":A0C2
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":A214
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Select"
         Top             =   4080
         Width           =   705
      End
      Begin MSComctlLib.ListView lvChildList 
         Height          =   3735
         Left            =   60
         TabIndex        =   93
         Top             =   330
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6588
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomersMasterFile.frx":A550
         NumItems        =   0
      End
      Begin VB.CommandButton cmdAddChildInfo 
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
         Height          =   645
         Left            =   3660
         MouseIcon       =   "CustomersMasterFile.frx":A6B2
         MousePointer    =   99  'Custom
         Picture         =   "CustomersMasterFile.frx":A804
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Add Children/Dependent"
         Top             =   4080
         Width           =   705
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   92
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   ":: LIST OF CHILDRENS/DEPENDENTS::"
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
   Begin VB.Label labid 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   12450
      TabIndex        =   88
      Top             =   210
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQ 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   12450
      TabIndex        =   90
      Top             =   1050
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labSEQMAX 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
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
      Left            =   12450
      TabIndex        =   97
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label labPrev 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12450
      TabIndex        =   89
      Top             =   660
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmAllCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs                                                                As ADODB.Recordset
Dim rsCusCtl                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim AccountCode                                                       As String
Dim CustType                                                          As String
Dim EntryPoint                                                        As String

Dim TempProspectID                                                    As Long
Event ChangedData(xCUSCODE As String)
Event ProspectConverted(CustomerCode As String, xGoingWhere As String, ProspectID As Long)
Dim GoingWhere                                                        As String

Public Sub AddCustomerFromProspect(oRs As Recordset, xGoingWhere As String)
    Dim ar                                                            As Variant
    GoingWhere = xGoingWhere
    If Not (oRs.EOF Or oRs.BOF) Then
        EntryPoint = "PROSPECT"
        AddorEdit = "ADD"
        picAdds.Visible = False: picSaves.Visible = True: Frame1.Enabled = True: fraSearch.Enabled = False
        txtAcctName.Text = Null2String(oRs!AcctName)
        TempProspectID = oRs!ProspectID
        CustType = Null2String(oRs!ProspectType)
        If CustType = "P" Then
            ar = Split(Null2String(oRs!AcctName))
            If UBound(ar) = 0 Then
                txtLastName.Text = ar(0)
            ElseIf UBound(ar) = 1 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
            ElseIf UBound(ar) >= 2 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(2)
                txtMiddleName.Text = ar(1)
            End If
        Else
            ar = Split(Null2String(oRs!ContactPerson))

            If UBound(ar) = 0 Then
                txtLastName.Text = ar(0)
            ElseIf UBound(ar) = 1 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
            ElseIf UBound(ar) >= 2 Then
                txtFirstName.Text = ar(0)
                txtLastName.Text = ar(1)
                txtMiddleName.Text = ar(2)
            End If
        End If
        txtCusphon1 = Null2String(oRs!Telephone)
        txtMobile = Null2String(oRs!Mobile)
        txtEmail = Null2String(oRs!EMAIL)
        txtPersonalStreet = Null2String(oRs!Address)
        txtNotes = Null2String(oRs!Notes)

        If CustType = "P" Then
            cboCustType.ListIndex = 0
        ElseIf CustType = "C" Then
            cboCustType.ListIndex = 1
        ElseIf CustType = "F" Then
            cboCustType.ListIndex = 2
        ElseIf CustType = "G" Then
            cboCustType.ListIndex = 3
        End If
        cboLeadSource.Text = Null2String(oRs!LeadSource)
    End If

End Sub

Friend Sub AddEditCustomer(xAcCode As String)
    AccountCode = xAcCode
End Sub

Private Sub cboApod_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub



Private Sub cboCustType_Click()
    Select Case cboCustType.Text
        Case "Personal"
            Label1.Caption = "Last Name"
            Label2.Visible = True: Label3.Visible = True
            txtLastName.Width = 2625: txtFirstName.Visible = True: txtMiddleName.Visible = True
            Label7.Caption = "Birth Date"
            Label24.Caption = "Spouse Name"
            CustType = "P"

            cmdCustInfo_Child.Enabled = True
            labCustInfo_Child.Enabled = True

        Case "Company/Agency"
            Label7.Caption = "Est Date"
            Label2.Visible = False: Label3.Visible = False
            txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
            Label1.Caption = "Company Name"
            Label24.Caption = "Contact Person"
            CustType = "C"
            cmdCustInfo_Child.Enabled = False
            labCustInfo_Child.Enabled = False
        Case "Fleet Account"
            CustType = "F"
            Label7.Caption = "Est Date"
            Label2.Visible = False: Label3.Visible = False
            txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
            Label1.Caption = "Company Name"
            Label24.Caption = "Contact Person"
            cmdCustInfo_Child.Enabled = False
            labCustInfo_Child.Enabled = False

        Case "Government"
            Label7.Caption = "Est Date"
            Label2.Visible = False: Label3.Visible = False
            txtLastName.Width = 8115: txtFirstName.Visible = False: txtMiddleName.Visible = False
            Label1.Caption = "Establisment Name"
            Label24.Caption = "Contact Person"
            CustType = "G"
            cmdCustInfo_Child.Enabled = False
            labCustInfo_Child.Enabled = False

    End Select


End Sub

Private Sub cboSearchCustype_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CUSTOMER") = False Then Exit Sub

    AddorEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picToolFrame.Enabled = True
    lstCustomer.Enabled = False
    txtSearch.Enabled = False
    initMemvars
    On Error Resume Next
    cboCustType.SetFocus
End Sub

Private Sub cmdAddChildInfo_Click()
    cmdDeleteChild.Enabled = False
    txtChildDate = ""
    txtChildName = ""
    cboChildSex = ""
    labIdCHILD = ""
    ShowPictureBox picChildAE, True, picMain
    On Error Resume Next
    txtChildName.SetFocus

End Sub



Private Sub cmdCancel_Click()
    ShowPictureBox picChildList, False
    ShowPictureBox picChildAE, False, picMain
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    picToolFrame.Enabled = True
    lstCustomer.Enabled = True
    fraSearch.Enabled = True
    AddorEdit = ""
    txtSearch.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCancelChildList_Click()
    ShowPictureBox picChildList, False, picMain

End Sub

Private Sub cmdCancelContactList_Click()
    ShowPictureBox picContactList, False, picMain
End Sub

Private Sub cmdCloseChildAE_Click(Index As Integer)
    ShowPictureBox picChildAE, False, picMain
End Sub

Private Sub cmdCloseContactsAE_Click(Index As Integer)
    ShowPictureBox picContactAE, False, picMain
End Sub



Private Sub cmdCloseTerm_Click(Index As Integer)
    ShowPictureBox picCredit, False, picMain
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "CUSTOMER") = False Then Exit Sub


    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Dim lng                                                           As Integer
    Load frmSplash
    
    Screen.MousePointer = 11
    frmSplash.labCon.Caption = "Checking Customer Record(s)... Please wait..."
    frmSplash.Show
        'IS PROSPECT
        lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) from CRIS_PROSPECTS WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Prospect Information Exists": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT RECEIPTS
        lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) from cmis_off_hd WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Official Receipt Exists For this Customer..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT VEHICLES
        lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) from csms_cusveh WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Record For Service..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT Appointment
        lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) from csms_appointment WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Appointment Information..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT Appointment
        lng = gconDMIS.Execute("SELECT COUNT(CUSCDE) from cmis_nonvat WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Official Receipt Exists For this Customer..": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT PARTS TRANS
        lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) from pmis_ord_hist WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Record For Parts Transactions.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'GOT PARTS TRANS
        lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) from pmis_ord_hd WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Is A Parts Customer and has Parts Transactions.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'ACCOUNTING
        lng = gconDMIS.Execute("SELECT COUNT(CUSTCODE) from amis_openinvoice WHERE CUSTCODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Record Finance and Accounting.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
        'SALES
        lng = gconDMIS.Execute("SELECT COUNT(CODE) from smis_salesorder WHERE CODE=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Sales Record.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
 
        'SERVICE
        lng = gconDMIS.Execute("SELECT COUNT(ACCT_NO) from CSMS_repairorder WHERE ACCT_NO=" & N2Str2Null(txtCuscde)).Fields(0).Value
        If lng > 0 Then: MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Customer Has Service Record.": Screen.MousePointer = 0: Unload frmSplash: Exit Sub
 
 
  Unload frmSplash
Screen.MousePointer = 0

    
    
 

    
    
    
        
        
        If ShowConfirmDelete = True Then
            Screen.MousePointer = 11
            gconDMIS.Execute "Delete from ALL_CUSTOMER  where ID=" & labid
            gconDMIS.Execute "Delete from ALL_CUSTOMER_CONTACTS  where CUSCDE=" & N2Str2Null(txtCuscde)
            gconDMIS.Execute "Delete from ALL_CUSTOMER_CHILD  where CUSCDE=" & N2Str2Null(txtCuscde)
            gconDMIS.Execute "Delete from ALL_CusCtl"

            Dim rsCustomer                                            As ADODB.Recordset
            Dim k                                                     As Integer
            Dim NewCtlCde                                             As String
            For k = 65 To 90
                Set rsCustomer = New ADODB.Recordset
                rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
                If Not rsCustomer.EOF And Not rsCustomer.BOF Then
                    NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!CODE, 2, 5)) + 1, "00000")
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
                Else
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
                End If
            Next
            Screen.MousePointer = 0
            FillSearchGrid ""
            rsRefresh
            StoreMemVars
            MessagePop Delete, "Record Deleted", "Customer Information Deleted. "
        End If

    
    
    rsRefresh
    rs.Bookmark = rsFind(rs.Clone, "ID", labid).Bookmark
    initMemvars
    StoreMemVars

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdDeleteChild_Click()

'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub



    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CHILD WHERE id=" & labIdCHILD
    ShowPictureBox picChildAE, False, picMain

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdDeleteContact_Click()

'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub
    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CONTACTS WHERE id=" & labIDContacts
    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If
    
    ShowPictureBox picContactAE, False, picMain

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "CUSTOMER") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picToolFrame.Enabled = False
    lstCustomer.Enabled = False
    fraSearch.Enabled = False
    On Error Resume Next
    txtLastName.SetFocus
End Sub

Private Sub cmdEditContact_Click()
    lvContactList_KeyPress 13
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    rs.MoveFirst
    ShowLastRecordMsg
End Sub
Private Sub cmdLast_Click()
    On Error Resume Next
    rs.MoveLast
    ShowLastRecordMsg
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CUSTOMER") = False Then Exit Sub
    CrystalReport1.Formulas(0) = "CompanyName = '" & Company_name & "'"
    CrystalReport1.Formulas(1) = "CompanyAddress = '" & Company_Address & "'"


    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Customers.rpt", "", DMIS_REPORT_Connection, 1
    '    'updating code:    JAA - 07112007
    '    On Error GoTo ErrorCode:
    '
    '    '    frmSMIS_ReportChoice.REPORTNAME = "CUSTOMERLISTING"
    '    Select Case MsgBox("Do you Want to Print This Customer or Print All Customer Listing" _
         '                     & vbCrLf & "Click Yes To Print All Customer Listing" _
         '                     & vbCrLf & "Click No To Print This Customer " _
         '         , vbYesNo Or vbExclamation Or vbDefaultButton2, App.TITLE)
    '
    '        Case vbYes
    '
    '        Case vbNo
    '
    '    End Select
    '    '    frmSMIS_ReportChoice.Show 1

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSave_Click()

'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If txtAcctName = "" Then
        ShowIsRequiredMsg "Account Name "
        On Error Resume Next
        txtAcctName.SetFocus
        Exit Sub
    End If
    If CustType = "P" And txtLastName = "" Then
        ShowIsRequiredMsg "Last Name"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If

    If CustType = "C" And txtLastName = "" Then
        ShowIsRequiredMsg "Company Name"
        On Error Resume Next
        txtLastName.SetFocus
        Exit Sub
    End If




    If AddorEdit = "ADD" Then
        Dim rsfindDup                                                 As ADODB.Recordset
        Set rsfindDup = New ADODB.Recordset
        rsfindDup.Open "select * from ALL_CustMaster_Smis where Code = '" & txtCuscde & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsfindDup.EOF And Not rsfindDup.BOF Then
            MsgSpeechBox "Code already exist!"
            Exit Sub
        End If
        txtCuscde = GetCustomerCode(txtLastName)
    End If






    Dim vtxtCusCde                                                    As String
    Dim VTXTLASTNAME                                                  As String
    Dim VTXTFIRSTNAME                                                 As String
    Dim vtxtMiddleInitial                                             As String
    Dim vtxtCUSCOMP                                                   As String

    Dim vcboSex                                                       As String
    Dim vtxtCusadd1                                                   As String
    Dim vtxtCusadd2                                                   As String
    Dim vtxtCuszipc                                                   As String
    Dim vtxtCusphon1                                                  As String
    Dim vtxtAcctName                                                  As String
    Dim vcboApod                                                      As String
    Dim vcboLeadSource                                                As String
    Dim vtxtTitle                                                     As String
    Dim vtxtDepartment                                                As String
    Dim vtxtEmail                                                     As String
    Dim vtxtMobile                                                    As String
    Dim vtxtHomePhone                                                 As String
    Dim VtxtFax                                                       As String
    Dim vtxtAssistant                                                 As String
    Dim vtxtAsstPhone                                                 As String
    Dim vtxtCity                                                      As String
    Dim vTxtBirthDate                                                 As String
    Dim vTxtSpouse                                                    As String
    Dim vtxtDescription                                               As String
    Dim vtxtCustType                                                  As String
    Dim vtxtCompanyAdd                                                As String
    Dim TEMPSQL                                                       As String
    Dim vtxtDeliveryAddress                                           As String
    Dim vtxtTIN                                                       As String
    vtxtTIN = N2Str2Null(txtTIN)
    vtxtCompanyAdd = N2Str2Null(UCase(txtPersonalStreet))
    vtxtCustType = N2Str2Null(CustType)
    vcboApod = N2Str2Null(UCase(cboApod))
    vtxtCusCde = N2Str2Null(txtCuscde)
    VTXTLASTNAME = N2Str2Null(UCase(txtLastName))
    VTXTFIRSTNAME = N2Str2Null(UCase(txtFirstName))
    vtxtMiddleInitial = N2Str2Null(txtMiddleName)
    vtxtAcctName = N2Str2Null(txtAcctName)
    vtxtCUSCOMP = N2Str2Null(UCase(txtLastName))
    vcboSex = N2Str2Null(cboSex)
    vtxtCusadd1 = N2Str2Null(Trim(UCase(txtPersonalStreet)))
    vtxtCusadd2 = N2Str2Null(UCase(txtPersonalState))
    vtxtCuszipc = N2Str2Null(txtPersonalZIP)
    vtxtCusphon1 = N2Str2Null(txtCusphon1)

    vcboLeadSource = N2Str2Null(cboLeadSource)
    vtxtTitle = N2Str2Null(txtTitle)
    vtxtDepartment = N2Str2Null(txtDepartment)
    vtxtEmail = N2Str2Null(txtEmail)
    vtxtMobile = N2Str2Null(txtMobile)
    vtxtHomePhone = N2Str2Null(txtHomePhone)
    VtxtFax = N2Str2Null(txtFax)
    vtxtAssistant = N2Str2Null(txtAssistant)
    vtxtAsstPhone = N2Str2Null(txtAsstPhone)

    vtxtCity = N2Str2Null(UCase(cboPersonalCity))
    vTxtBirthDate = N2Str2Null(txtBirthDate)
    vTxtSpouse = N2Str2Null(txtSpouse)
    vtxtDescription = N2Str2Null(txtNotes)

    vtxtDeliveryAddress = N2Str2Null(txtDeliveryAddress)

    If AddorEdit = "ADD" Then
        TEMPSQL = "INSERT INTO ALL_CUSTOMER(" & vbCrLf
        TEMPSQL = TEMPSQL & " TIN, CUSCOMP, APOD , CUSCDE , LASTNAME, FIRSTNAME, MIDDLEINITIAL,ACCTNAME,SEX,CUSTOMERADD,PROVINCIALADD,ZIPCODE,TELEPHONENO,LEADSOURCE,TITLE,DEPARTMENT,EMAIL,MOBILE,HOMEPHONE,FAX,ASSISTANT,ASSTPHONE,CITY,BIRTHDATE,SPOUSE,DESCRIPTION, CUSTYPE, COMPANYADD , DELIVERYADDRESS " & vbCrLf
        TEMPSQL = TEMPSQL & " ) VALUES ( " & vbCrLf
        TEMPSQL = TEMPSQL & vtxtTIN & ", "
        TEMPSQL = TEMPSQL & vtxtCUSCOMP & ", "
        TEMPSQL = TEMPSQL & vcboApod & ","
        TEMPSQL = TEMPSQL & vtxtCusCde & ", "
        TEMPSQL = TEMPSQL & VTXTLASTNAME & ", "
        TEMPSQL = TEMPSQL & VTXTFIRSTNAME & ", "
        TEMPSQL = TEMPSQL & vtxtMiddleInitial & ", "
        TEMPSQL = TEMPSQL & vtxtAcctName & ","
        TEMPSQL = TEMPSQL & vcboSex & "," & vbCrLf
        TEMPSQL = TEMPSQL & vtxtCusadd1 & ", "
        TEMPSQL = TEMPSQL & vtxtCusadd2 & ", "
        TEMPSQL = TEMPSQL & vtxtCuszipc & ", "
        TEMPSQL = TEMPSQL & vtxtCusphon1 & ","
        TEMPSQL = TEMPSQL & vcboLeadSource & ","
        TEMPSQL = TEMPSQL & vtxtTitle & ","
        TEMPSQL = TEMPSQL & vtxtDepartment & ","
        TEMPSQL = TEMPSQL & vtxtEmail & ","
        TEMPSQL = TEMPSQL & vtxtMobile & ","
        TEMPSQL = TEMPSQL & vtxtHomePhone & ","
        TEMPSQL = TEMPSQL & VtxtFax & ","
        TEMPSQL = TEMPSQL & vtxtAssistant & ","
        TEMPSQL = TEMPSQL & vtxtAsstPhone & ","
        TEMPSQL = TEMPSQL & vtxtCity & ","
        TEMPSQL = TEMPSQL & vTxtBirthDate & ","
        TEMPSQL = TEMPSQL & vTxtSpouse & ","
        TEMPSQL = TEMPSQL & vtxtDescription & ","
        TEMPSQL = TEMPSQL & vtxtCustType & ","
        TEMPSQL = TEMPSQL & vtxtCompanyAdd & ","
        TEMPSQL = TEMPSQL & vtxtDeliveryAddress
        TEMPSQL = TEMPSQL & ")"

        gconDMIS.Execute TEMPSQL



    Else
        TEMPSQL = "UPDATE ALL_CUSTOMER SET" & vbCrLf
        TEMPSQL = TEMPSQL & " CUSCOMP = " & vtxtCUSCOMP & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TIN = " & vtxtTIN & "," & vbCrLf
        TEMPSQL = TEMPSQL & " COMPANYADD = " & vtxtCompanyAdd & "," & vbCrLf
        TEMPSQL = TEMPSQL & " APOD = " & vcboApod & "," & vbCrLf
        TEMPSQL = TEMPSQL & " LASTNAME = " & VTXTLASTNAME & "," & vbCrLf
        TEMPSQL = TEMPSQL & " FIRSTNAME = " & VTXTFIRSTNAME & "," & vbCrLf
        TEMPSQL = TEMPSQL & " MIDDLEINITIAL = " & vtxtMiddleInitial & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ACCTNAME = " & vtxtAcctName & "," & vbCrLf
        TEMPSQL = TEMPSQL & " SEX = " & vcboSex & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CUSTOMERADD = " & vtxtCusadd1 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " PROVINCIALADD = " & vtxtCusadd2 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ZIPCODE = " & vtxtCuszipc & "," & vbCrLf
        TEMPSQL = TEMPSQL & " LEADSOURCE = " & vcboLeadSource & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TITLE = " & vtxtTitle & "," & vbCrLf
        TEMPSQL = TEMPSQL & " DEPARTMENT = " & vtxtDepartment & "," & vbCrLf
        TEMPSQL = TEMPSQL & " EMAIL = " & vtxtEmail & "," & vbCrLf
        TEMPSQL = TEMPSQL & " MOBILE = " & vtxtMobile & "," & vbCrLf
        TEMPSQL = TEMPSQL & " TELEPHONENO  = " & vtxtCusphon1 & "," & vbCrLf
        TEMPSQL = TEMPSQL & " HOMEPHONE = " & vtxtHomePhone & "," & vbCrLf
        TEMPSQL = TEMPSQL & " FAX = " & VtxtFax & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ASSISTANT = " & vtxtAssistant & "," & vbCrLf
        TEMPSQL = TEMPSQL & " ASSTPHONE = " & vtxtAsstPhone & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CITY = " & vtxtCity & "," & vbCrLf
        TEMPSQL = TEMPSQL & " BIRTHDATE = " & vTxtBirthDate & "," & vbCrLf
        TEMPSQL = TEMPSQL & " SPOUSE = " & vTxtSpouse & "," & vbCrLf
        TEMPSQL = TEMPSQL & " CUSTYPE = " & vtxtCustType & "," & vbCrLf
        TEMPSQL = TEMPSQL & " DESCRIPTION = " & vtxtDescription & "," & vbCrLf

        TEMPSQL = TEMPSQL & " DeliveryAddress = " & vtxtDeliveryAddress & vbCrLf
        TEMPSQL = TEMPSQL & " WHERE CUSCDE = '" & txtCuscde & "'" & vbCrLf
        gconDMIS.Execute TEMPSQL
    End If



    '''CUSTOMER CONTROL
    Screen.MousePointer = 11
    gconDMIS.Execute "delete from ALL_CusCtl"
    Dim rsCustomer                                                    As ADODB.Recordset
    Dim k                                                             As Integer
    Dim NewCtlCde                                                     As String
    For k = 65 To 90
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!CODE, 2, 5)) + 1, "00000")
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
    Next
    Screen.MousePointer = 0
    '''END CONTROL


    If EntryPoint = "PROSPECT" Then
        gconDMIS.Execute " UPDATE CRIS_PROSPECTS SET PROSPECTTYPE=" & vtxtCustType & " WHERE  PROSPECTID= " & TempProspectID
        RaiseEvent ProspectConverted(Replace(vtxtCusCde, "'", ""), GoingWhere, TempProspectID)
        Screen.MousePointer = 0
        'Unload Me: Exit Sub
    Else
        RaiseEvent ChangedData(Replace(vtxtCusCde, "'", ""))
        Screen.MousePointer = 0
        'Unload Me: Exit Sub

        Screen.MousePointer = 0
        MessagePop RecSave, "Record Saved", " Customer Information Saved"

        rs.Requery
        If AddorEdit = "EDIT" Then
            rs.Find "id =" & labid
        End If
        FillSearchGrid txtSearch
    End If


    cmdCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub



Private Sub cmdSaveContact_Click()

'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtContactName)) = "" Then
        ShowIsRequiredMsg "CONTACT NAME"
        On Error Resume Next
        txtContactName.SetFocus
        Exit Sub
    End If
    Dim vtxtCusCde, ContactName, Relation, ContactPosition, Department, Phone, Mobile, Address, SQL



    vtxtCusCde = N2Str2Null(txtCuscde)
    ContactName = N2Str2Null(txtContactName)
    Relation = N2Str2Null(cboContactRelation)
    ContactPosition = N2Str2Null(txtContactPosition)
    Department = N2Str2Null(txtContactDepartment)
    Phone = N2Str2Null(txtContactPhone)
    Mobile = N2Str2Null(txtContactMobile)
    Address = N2Str2Null(txtContactAddress)



    If NumericVal(labIDContacts) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CONTACTS "
        SQL = SQL & "(ContactName , CUSCDE, Relation, ContactPosition, Department, Phone, Mobile, Address) VALUES ("
        SQL = SQL & ContactName & " ,"
        SQL = SQL & vtxtCusCde & " ,"
        SQL = SQL & Relation & " ,"
        SQL = SQL & ContactPosition & " ,"
        SQL = SQL & Department & " ,"
        SQL = SQL & Phone & " ,"
        SQL = SQL & Mobile & " ,"
        SQL = SQL & Address & " )"
    Else
        SQL = "UPDATE ALL_CUSTOMER_CONTACTS SET "
        SQL = SQL & " ContactName =" & ContactName & ", "
        SQL = SQL & " Relation =" & Relation & ", "
        SQL = SQL & " ContactPosition =" & ContactPosition & ", "
        SQL = SQL & " Department =" & Department & ", "
        SQL = SQL & " Phone =" & Phone & ", "
        SQL = SQL & " Address =" & Address & ", "
        SQL = SQL & " Mobile =" & Mobile
        SQL = SQL & "  where id=" & labIDContacts



    End If
    gconDMIS.Execute SQL
    If picContactList.Visible = True Then
        cmdCUSTINFO_Contact_Click
    End If
   ShowPictureBox picContactAE, False, picMain


    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelectChild_Click()
    lvChildList_KeyPress 13
End Sub

Private Sub cmdCUSTINFO_Child_Click()
'AddColumnHeader "ChildName,Sex,DateofBirth,Age", lvChildList
'        lvChildList.ListItems , )
    Dim TempRs                                                        As ADODB.Recordset
    Set TempRs = gconDMIS.Execute("SELECT CHILDNAME,SEX,DOB, DATEDIFF(YEAR, DOB, GETDATE()) ,ID FROM ALL_CUSTOMER_CHILD WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    FillCustomerListView TempRs, lvChildList
    ShowPictureBox picChildList, True, picMain
    If lvChildList.ListItems.Count = 0 Then
        cmdSelectChild.Enabled = False
    Else
        cmdSelectChild.Enabled = True
    End If
End Sub



Private Sub Command1_Click()
    Dim rstable1                                                      As ADODB.Recordset
    Dim rsSTOCKS                                                      As ADODB.Recordset
    Dim rsDNPP                                                        As ADODB.Recordset

    Dim HARI_SRP                                                      As Double

    Set rstable1 = New ADODB.Recordset
    Set rstable1 = gconDMIS.Execute("Select * from table1 order by partno asc")
    If Not rstable1.EOF And Not rstable1.BOF Then
        rstable1.MoveFirst
        Screen.MousePointer = 11
        Do While Not rstable1.EOF
            Set rsDNPP = New ADODB.Recordset
            Set rsDNPP = gconDMIS.Execute("Select * from PMIS_DNPP where (partnumber = " & N2Str2Null(rstable1!partno) & ") or partnumber = '" & Left(Null2String(rstable1!partno), 5) & " " & Right(Null2String(rstable1!partno), Len(Null2String(rstable1!partno)) - 5) & "'")
            If Not rsDNPP.EOF And Not rsDNPP.BOF Then
                HARI_SRP = N2Str2Zero(rsDNPP!srp)
            Else
                HARI_SRP = N2Str2Zero(rstable1!srp)
            End If
            Set rsSTOCKS = New ADODB.Recordset
            Set rsSTOCKS = gconDMIS.Execute("Select * from pmis_stockmas where (stockno = '" & Left(Null2String(rstable1!partno), 5) & " " & Right(Null2String(rstable1!partno), Len(Null2String(rstable1!partno)) - 5) & "')")
            If Not rsSTOCKS.EOF And Not rsSTOCKS.BOF Then
                'MsgBox rsstocks!stockno & vbCrLf & rstable1!partno
                gconDMIS.Execute ("Update pmis_stockmas set onhand = " & N2Str2Zero(rstable1!ONHAND) & ", mac = " & N2Str2Zero(rstable1!Mac) & ", dnp = " & N2Str2Zero(rstable1!DNP) & ", srp = " & N2Str2Zero(rstable1!srp) & ", location = " & N2Str2Null(rstable1!Location) & " where stockno = " & N2Str2Null(rsSTOCKS!StockNo))
            Else
                gconDMIS.Execute ("Insert into pmis_stockmas (stockno,stockdesc,onhand,mac,dnp,srp,location) values (" & N2Str2Null(rstable1!partno) & "," & N2Str2Null(rstable1!PartDesc) & "," & N2Str2Zero(rstable1!ONHAND) & "," & N2Str2Zero(rstable1!Mac) & "," & N2Str2Zero(rstable1!DNP) & "," & N2Str2Zero(rstable1!srp) & "," & N2Str2Null(rstable1!Location) & ")")
            End If
            rstable1.MoveNext
        Loop
        'Me.Caption = rstable1!partno & " --> " & rstable1!ONHAND
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Command12_Click()
    Dim cTerm, cLimit, cDays

    cLimit = NumericVal(txtCreditLimit)
    cDays = NumericVal(txtCreditDays)



    gconDMIS.Execute "update all_Customer set CreditLimit=" & cLimit & ", CREDITTERM='C', CREDITDAYS=" & cDays & " WHERE ID=" & labid


    MessagePop RecSaveOk, "Credit Info", "Credit Information Updated"
    ShowPictureBox picCredit, False, picMain
End Sub



Private Sub Command3_Click()
    txtDeliveryAddress = txtPersonalStreet & "," & txtPersonalState & "," & cboPersonalCity & "," & txtPersonalZIP
End Sub

Private Sub Command4_Click()
    labCUSTINFO_Contact_Click
End Sub



Private Sub cmdSaveChild_Click()

    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtChildName)) = "" Then
        ShowIsRequiredMsg "Children Name "
        On Error Resume Next
        txtChildName.SetFocus
        Exit Sub
    End If

    Dim vtxtCHILDNAME, vtxtSEX, vtxtDOB, SQL, vtxtCusCde
    vtxtCHILDNAME = N2Str2Null(txtChildName)
    vtxtCusCde = N2Str2Null(txtCuscde)
    If cboChildSex = "M" Then
        vtxtSEX = "'M'"
    ElseIf cboChildSex = "F" Then

        vtxtSEX = "'F'"
    Else
        vtxtSEX = "'U'"
    End If

    vtxtDOB = N2Date2Null(txtChildDate)

    If NumericVal(labIdCHILD) = 0 Then
        SQL = "INSERT INTO ALL_CUSTOMER_CHILD (CUSCDE,CHILDNAME,SEX,DOB)VALUES("
        SQL = SQL & vtxtCusCde
        SQL = SQL & "," & vtxtCHILDNAME & " ,"
        SQL = SQL & vtxtSEX & " ,"
        SQL = SQL & vtxtDOB & " )"
    Else
        SQL = "UPDATE ALL_CUSTOMER_CHILD SET "
        SQL = SQL & " CHILDNAME =" & vtxtCHILDNAME & " , "
        SQL = SQL & " SEX=" & vtxtSEX & " , "
        SQL = SQL & " DOB=" & vtxtDOB
        SQL = SQL & " where id=" & labIdCHILD
    End If

    gconDMIS.Execute SQL

    If picChildList.Visible = True Then
        cmdCUSTINFO_Child_Click
    End If
    ShowPictureBox picChildAE, False, picMain

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCUSTINFO_Contact_Click()

'SELECT  ContactName,    Relation,Phone, Mobile,  ContactPosition, Department, Address, ID
    Dim TempRs                                                        As ADODB.Recordset
    Set TempRs = gconDMIS.Execute("SELECT  CONTACTNAME, RELATION,PHONE, MOBILE, CONTACTPOSITION, DEPARTMENT, ADDRESS, ID FROM ALL_CUSTOMER_CONTACTS WHERE CUSCDE=" & N2Str2Null(txtCuscde))
    FillCustomerListView TempRs, lvContactList
    ShowPictureBox picContactList, True, picMain
    If lvContactList.ListItems.Count = 0 Then
        cmdEditContact.Enabled = False
    Else
        cmdEditContact.Enabled = True
    End If
End Sub

Private Sub cmdCUSTINFO_CREDIT_Click()
    If Module_Access(LOGID, "CUSTOMER CREDIT LIMIT", "DATA ENTRY") = False Then Exit Sub
    ShowPictureBox picCredit, True, picMain
    On Error Resume Next
    txtCreditLimit.SetFocus
End Sub

Sub FillSearchGrid(XXX As String)

    Dim rsCustomer2                                                   As ADODB.Recordset
    Dim Key                                                           As String
    Dim LIMITKEY                                                      As String

    lstCustomer.Enabled = False: lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    If optSearchKeyAcctName.Value = True Then
        Key = "AcctName"
    ElseIf optSearchKeyAddress.Value = True Then
        Key = "CustomerAdd"
    ElseIf optSearchKeyCompany.Value = True Then
        Key = "CUSCOMP"
    ElseIf optSearchKeyLast.Value = True Then
        Key = "LastName"
    ElseIf optSearchKeyEmail.Value = True Then
        Key = "Email"
    End If

    Select Case cboSearchCustype.ListIndex
        Case 0                                              'Search All
            LIMITKEY = "'P','C','F','G', NULL"
        Case 1                                              'Only Personal Customers
            LIMITKEY = "'P'"
        Case 2                                              ' Only Company/Agency Customers
            LIMITKEY = "'C'"
        Case 3                                              'Only Government Customer
            LIMITKEY = "'G'"
        Case 4                                              'Only Fleet Account Customer
            LIMITKEY = "'F'"
    End Select


    Set rsCustomer2 = gconDMIS.Execute("select TOP 100  " & Key & "  as CustomerName, id  from ALL_CUSTOMER where CUSCDE <> '999999' and " & Key & " like'" & XXX & "%' AND CUSTYPE IN (" & LIMITKEY & " )  order by 1 asc")

    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
        lstCustomer.Enabled = True
        lstCustomer.Refresh
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    If picChildAE.Visible = True And KeyCode = 27 Then
        cmdCloseChildAE_Click 0
    ElseIf picChildList.Visible = True And KeyCode = 27 Then
        cmdCancelChildList_Click

    ElseIf picContactAE.Visible = True And KeyCode = 27 Then
        cmdCloseContactsAE_Click 0
    ElseIf picContactList.Visible = True And KeyCode = 27 Then
        cmdCancelContactList_Click

    ElseIf picCredit.Visible = True And KeyCode = 27 Then
        cmdCloseTerm_Click 0
        'ElseIf picAdds.Visible = True And KeyCode = vbKeyEscape Then
        '    Unload Me
    Else
        MoveKeyPress KeyCode
    End If

End Sub

Private Sub Form_Load()

    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Frame1.Enabled = False
    picAdds.Visible = True
    picToolFrame.Enabled = True
    picSaves.Visible = False
    initMemvars
    InitData

    With lvChildList.ColumnHeaders
        .Add 1, , "ChildName", 0.5 * lvChildList.Width
        .Add 2, , "SEX", 0.15 * lvChildList.Width
        .Add 3, , "DATEOFBIRTH", 0.15 * lvChildList.Width
        .Add 4, , "AGE", 0.15 * lvChildList.Width
    End With
    With lvContactList.ColumnHeaders
        .Add 1, , "CONTACTNAME", 0.4 * lvChildList.Width
        .Add 2, , "RELATION", 0.2 * lvChildList.Width
        .Add 3, , "PHONE", 0.17 * lvChildList.Width
        .Add 4, , "MOBILE", 0.17 * lvChildList.Width
    End With

    rsRefresh




    If AccountCode <> "" Then
        rs.Find ("CUSCDE=" & N2Str2Null(AccountCode))
        StoreMemVars
        cmdEdit.Value = True
    End If
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsCusCtl = Nothing
    AddorEdit = vbNullString
    AccountCode = vbNullString
    CustType = vbNullString
    EntryPoint = vbNullString
    TempProspectID = 0

End Sub

Sub InitData()
    Combo_Loadval cboPersonalCity, gconDMIS.Execute("Select Distinct CITY FROM ALL_CUSTOMER WHERE CITY IS NOT NULL")
    With cboCustType
        .AddItem ("Personal")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
         .AddItem ("Fleet Account")
        .ListIndex = 0
    End With
    With cboSearchCustype
        .AddItem ("Search All")
        .AddItem ("Individual")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet")
        .ListIndex = 0
    End With




End Sub

Sub initMemvars()
    Dim TempRs                                                        As ADODB.Recordset
    txtSearch.Text = vbNullString
    labSEQ.Caption = gconDMIS.Execute("SELECT isnull(MAX(ID),0) FROM ALL_CUSTOMER").Collect(0)
    txtCuscde.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleName.Text = ""
    txtAcctName.Text = ""
    cboLeadSource.Text = ""
    cboSex.Text = ""
    txtTitle.Text = ""
    txtDepartment.Text = ""
    txtEmail.Text = ""
    txtCusphon1.Text = ""
    txtMobile.Text = ""
    txtHomePhone.Text = ""
    txtFax.Text = ""
    txtAssistant.Text = ""
    txtAsstPhone.Text = ""
    txtPersonalStreet.Text = ""
    cboPersonalCity.Text = ""
    txtPersonalState.Text = ""
    txtPersonalZIP.Text = ""
    txtBirthDate.Text = ""
    txtSpouse.Text = ""
    txtNotes.Text = ""
    cboApod.Clear
    
    
    txtDeliveryAddress = ""
    txtCreditDays = "0"
    txtCreditLimit = "0.00"

    Dim rsAPOD                                                        As ADODB.Recordset
    Set rsAPOD = New ADODB.Recordset
    Set rsAPOD = gconDMIS.Execute("Select distinct apod from ALL_CustMaster_Smis Where APOD is Not Null")

    If Not rsAPOD.EOF And Not rsAPOD.BOF Then
        rsAPOD.MoveFirst
        Do While Not rsAPOD.EOF
            cboApod.AddItem Null2String(rsAPOD!APOD)
            rsAPOD.MoveNext
        Loop
    End If
    Set rsAPOD = Nothing
    Set TempRs = gconDMIS.Execute("Select DataDesc from CRIS_vW_MasterPullDown where  Masterdesc='Lead Source'")
    cboLeadSource.Clear

    While Not TempRs.EOF
        cboLeadSource.AddItem TempRs.Collect(0)
        TempRs.MoveNext
    Wend
    cboSex.Clear
    cboSex.AddItem "NA"
    cboSex.AddItem "M"
    cboSex.AddItem "F"

End Sub








Private Sub fraCompany_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub labCustInfo_Child_Click()
    cmdAddChildInfo_Click
End Sub

Private Sub labCustInfo_Child_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = &H400000
    labCustInfo_Child.FontBold = True
End Sub

Private Sub labCUSTINFO_CREDIT_Click()
    cmdCUSTINFO_CREDIT_Click
End Sub

Private Sub labCUSTINFO_CREDIT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Credit.ForeColor = &H400000
    labCustInfo_Credit.FontBold = True
End Sub





Private Sub labCUSTINFO_Contact_Click()
labIDContacts = 0
txtContactName = ""
cboContactRelation = ""
txtContactPosition = ""
txtContactDepartment = ""
txtContactPhone = ""
txtContactMobile = ""
txtContactAddress = ""
cmdDeleteContact.Enabled = False
    ShowHidePictureBox2 picContactAE, True, picMain
End Sub

Private Sub labCUSTINFO_Contact_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Contact.ForeColor = &H400000
    labCustInfo_Contact.FontBold = True
    
End Sub




Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    With lstCustomer
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

Private Sub lstCustomer_DblClick()
    If lstCustomer.Enabled = True Then
        cmdEdit.Value = True
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rs.MoveFirst
    rs.Find ("ID=" & Item.ListSubItems(1).Text)

    StoreMemVars
End Sub

Private Sub lvChildList_DblClick()
    lvChildList_KeyPress 13
End Sub

Private Sub lvChildList_KeyPress(KeyAscii As Integer)
    If lvChildList.SelectedItem Is Nothing Then Exit Sub
    On Error GoTo ADDER:

    If KeyAscii <> 13 Then Exit Sub
    'AddColumnHeader "ID, ChildName,Sex,DateofBirth,Age", lvChildList
    txtChildName = lvChildList.SelectedItem
    cboChildSex = lvChildList.SelectedItem.ListSubItems(1).Text
    txtChildDate = lvChildList.SelectedItem.ListSubItems(2).Text
    labIdCHILD = lvChildList.SelectedItem.ListSubItems(3).Text
    cmdDeleteChild.Enabled = True

    ShowPictureBox picChildList, False
    ShowPictureBox picChildAE, True, picMain
    On Error Resume Next
    txtChildName.SetFocus
    Exit Sub
ADDER:
    ShowVBError
End Sub

Private Sub lvContactList_DblClick()
    lvContactList_KeyPress 13
End Sub

Private Sub lvContactList_KeyPress(KeyAscii As Integer)
    If lvContactList.SelectedItem Is Nothing Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    'CONTACTNAME, RELATION,PHONE, MOBILE, CONTACTPOSITION, DEPARTMENT, ADDRESS, ID
    ShowPictureBox picContactAE, True, picMain
    With lvContactList.SelectedItem
        txtContactName = .Text
        cboContactRelation = .ListSubItems(1).Text
        txtContactPhone = .ListSubItems(2).Text
        txtContactMobile = .ListSubItems(3).Text
        txtContactPosition = .ListSubItems(4).Text
        txtContactDepartment = .ListSubItems(5).Text
        txtContactAddress = .ListSubItems(6).Text
        labIDContacts = .ListSubItems(7).Text

    End With

    cmdDeleteContact.Enabled = True
   ' ShowPictureBox picContactList, False, picMain
    On Error Resume Next

    txtContactName.SetFocus
End Sub

Sub rsRefresh()
    Set rs = New ADODB.Recordset
    rs.Open "Select * from ALL_Customer order by id DESC", gconDMIS, adOpenKeyset, adLockReadOnly


End Sub

Sub SetCustomerAccountName()
    If EntryPoint = "PROSPECT" Or AddorEdit = "EDIT" Or AddorEdit = "" Then: Exit Sub

    'If CustType = "P" Then
    txtAcctName = UCase(txtLastName & IIf(txtFirstName = "", "", ",") & txtFirstName & IIf(txtMiddleName = "", "", ".") & Left(txtMiddleName, 1))
    'Else
    '   txtAcctName.Text = UCase(txtCusComp)
    'End If


End Sub

Sub StoreMemVars()

    If Not rs.EOF And Not rs.BOF Then
        labid.Caption = rs!ID
        cboApod.Text = Null2String(rs!APOD)
        txtCuscde.Text = Null2String(rs!CUSCDE)
        txtLastName.Text = Null2String(rs!lastname)
        txtFirstName.Text = Null2String(rs!firstname)
        txtMiddleName.Text = Null2String(rs!MiddleInitial)

        
        txtTIN = Null2String(rs!Tin)
        
        cboSex.Text = Null2String(rs!Sex)
        txtAcctName = Null2String(rs!AcctName)

        txtPersonalStreet.Text = Null2String(rs!customeradd)
        txtPersonalState.Text = Null2String(rs!provincialadd)
        txtPersonalZIP.Text = Null2String(rs!ZIPCODE)
        txtCusphon1.Text = Null2String(rs!TelephoneNo)

        cboLeadSource.Text = Null2String(rs!LeadSource)
        txtTitle.Text = Null2String(rs!TITLE)
        txtDepartment.Text = Null2String(rs!Department)
        txtEmail.Text = Null2String(rs!EMAIL)
        txtMobile.Text = Null2String(rs!Mobile)
        txtHomePhone.Text = Null2String(rs!HomePhone)
        txtFax.Text = Null2String(rs!Fax)
        txtAssistant.Text = Null2String(rs!Assistant)
        txtAsstPhone.Text = Null2String(rs!AsstPhone)
        cboPersonalCity.Text = Null2String(rs!City)
        txtBirthDate.Text = Null2String(rs!BirthDate)
        txtSpouse.Text = Null2String(rs!Spouse)
        txtDeliveryAddress = Null2String(rs!DELIVERYADDRESS)
        txtNotes.Text = Null2String(rs!Description)

        txtCreditDays = NumericVal(rs!CREDITDAYS)
        txtCreditLimit = FormatNumber(NumericVal(rs!CreditLimit))



        If Null2String(rs!CUSTYPE) = "P" Then
            cboCustType.ListIndex = 0
        ElseIf Null2String(rs!CUSTYPE) = "C" Then
            cboCustType.ListIndex = 1
        ElseIf Null2String(rs!CUSTYPE) = "G" Then
            cboCustType.ListIndex = 2
        ElseIf Null2String(rs!CUSTYPE) = "F" Then
            cboCustType.ListIndex = 3
        Else
            cboCustType.ListIndex = 0
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub optSearchKeyAcctName_Click()
    lstCustomer.ColumnHeaders(1).Text = "A/C Name"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyAddress_Click()
    lstCustomer.ColumnHeaders(1).Text = "Address"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyCompany_Click()
    lstCustomer.ColumnHeaders(1).Text = "Company"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyEmail_Click()
    lstCustomer.ColumnHeaders(1).Text = "Email"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub optSearchKeyLast_Click()
    lstCustomer.ColumnHeaders(1).Text = "LastName"
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub picToolFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    labCustInfo_Child.ForeColor = vbBlack
    labCustInfo_Child.FontBold = False
    labCustInfo_Credit.ForeColor = vbBlack
    labCustInfo_Credit.FontBold = False
    labCustInfo_Contact.ForeColor = vbBlack
    labCustInfo_Contact.FontBold = False


End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCreditDays_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditDays_GotFocus()

    If NumericVal(txtCreditDays.Text) <= 0 Then txtCreditDays = ""
End Sub

Private Sub txtCreditDays_LostFocus()
    If NumericVal(txtCreditDays) <= 0 Then txtCreditDays = "0"
    txtCreditDays = NumericVal(txtCreditDays)
End Sub



Private Sub txtCreditLimit_GotFocus()

    If NumericVal(txtCreditLimit.Text) <= 0 Then txtCreditLimit = ""
End Sub

Private Sub txtCreditLimit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtCreditLimit_LostFocus()
    If NumericVal(txtCreditLimit) <= 0 Then txtCreditLimit = "0.00"
    txtCreditLimit = FormatNumber(NumericVal(txtCreditLimit))
End Sub

Private Sub txtCusComp_Change()

    SetCustomerAccountName
End Sub

Private Sub txtFirstName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtLastName_Change()
    
    If AddorEdit = "ADD" And LTrim(RTrim(txtLastName)) <> "" Then
        txtCuscde = GetCustomerCode(txtLastName)
        SetCustomerAccountName
    End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtMiddleName_Change()
    SetCustomerAccountName
End Sub

Private Sub txtMiddleName_KeyPress(KeyAscii As Integer)
    UpperAscii (KeyAscii)
End Sub

Private Sub txtPersonalStreet_KeyPress(KeyAscii As Integer)
    UpperAscii KeyAscii
End Sub

Private Sub txtPersonalZIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtsearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub


Sub ShowPictureBox(cntl As Object, State As Boolean, Optional ByVal MasterObject As Object)
    cntl.Visible = State

    If Not (MasterObject Is Nothing) Then
        MasterObject.Enabled = Not State
    End If
    If State = True Then
        cntl.ZOrder 0
    Else
        cntl.ZOrder 1
    End If
End Sub



Function GetCustomerCode(lastname As String) As String
    Dim TempRs                                                        As ADODB.Recordset
    If Len(lastname) = 0 Then
        Exit Function
    End If
    Dim lAlpha                                                        As String
    lAlpha = Left(Trim(lastname), 1)
    Set TempRs = gconDMIS.Execute("Select CTLCDE From ALL_CUSCTL Where LEFT(CTLCDE,1)='" & lAlpha & "'")
    If Not (TempRs.EOF Or TempRs.BOF) Then
        GetCustomerCode = Left(lastname, 1) & Format(Mid(TempRs.Collect(0), 2, 5), "00000")
    Else
        GetCustomerCode = Left(lastname, 1) & "00001"
    End If
End Function

Public Sub FillCustomerListView(rs As Recordset, grd As ListView, Optional WithSN As Boolean = False, Optional WITHCOLUMNHEADER As Boolean = False)
    Dim fld                                                           As Field
    Dim j                                                             As Long
    Dim ijx                                                           As Integer
    Dim lst                                                           As ListItem
    Dim I                                                             As Integer

    grd.Enabled = False

    grd.ListItems.Clear

    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For I = 0 To rs.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , rs.Fields(I).Name)
        Next
        While Not rs.EOF
            j = j + 1
            Set lst = grd.ListItems.Add(, , j)
            For Each fld In rs.Fields
                If IsNull(fld.Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , fld.Value
                End If
            Next
            rs.MoveNext
        Wend

    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then

        While Not rs.EOF
            j = j + 1
            Set lst = grd.ListItems.Add(, , j)
            For Each fld In rs.Fields
                If IsNull(fld.Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , fld.Value
                End If
            Next
            rs.MoveNext
        Wend

    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For I = 0 To rs.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , rs.Fields(I).Name)
        Next
        j = rs.Fields.Count
        While Not rs.EOF
            Set lst = grd.ListItems.Add(, , rs.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(rs.Fields(ijx).Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , rs.Fields(ijx).Value
                End If
            Next
            rs.MoveNext
        Wend
    Else
        j = rs.Fields.Count
        While Not rs.EOF
            Set lst = grd.ListItems.Add(, , Null2String(rs.Fields(0).Value))
            For ijx = 1 To j - 1
                If IsNull(rs.Fields(ijx).Value) Then
                    lst.ListSubItems.Add , , vbNullString
                Else
                    lst.ListSubItems.Add , , rs.Fields(ijx).Value
                End If
            Next
            rs.MoveNext
        Wend
    End If

    grd.Enabled = True

    Set lst = Nothing
    'Set rs = Nothing
End Sub
