VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCRIS_Customer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers"
   ClientHeight    =   8295
   ClientLeft      =   525
   ClientTop       =   735
   ClientWidth     =   12270
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
   Icon            =   "TCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12270
   Begin VB.TextBox labOLDCuscde 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00400000&
      Height          =   450
      Left            =   12450
      TabIndex        =   160
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
      TabIndex        =   136
      Top             =   1920
      Visible         =   0   'False
      Width           =   1500
   End
   Begin Crystal.CrystalReport rptCustomer 
      Left            =   3270
      Top             =   7770
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox picCredit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00A9B8C2&
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   4230
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2295
      ScaleWidth      =   3390
      TabIndex        =   150
      Top             =   2910
      Visible         =   0   'False
      Width           =   3420
      Begin VB.TextBox txtCreditDays 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1230
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   840
         Width           =   1875
      End
      Begin VB.TextBox txtCreditLimit 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1230
         TabIndex        =   154
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
         Left            =   2430
         MouseIcon       =   "TCustomers.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   1350
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
         Left            =   1740
         MouseIcon       =   "TCustomers.frx":0D5A
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":0EAC
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   1350
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
         TabIndex        =   151
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
         TabIndex        =   155
         Top             =   930
         Width           =   1020
      End
      Begin VB.Label labTermID 
         Height          =   555
         Left            =   360
         TabIndex        =   159
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
         TabIndex        =   153
         Top             =   480
         Width           =   465
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   330
         Left            =   0
         TabIndex        =   152
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
   Begin VB.PictureBox picContactAE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DFCCCF&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3930
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4305
      ScaleWidth      =   4350
      TabIndex        =   106
      Top             =   1980
      Visible         =   0   'False
      Width           =   4380
      Begin VB.TextBox txtContactName 
         Height          =   345
         Left            =   1140
         TabIndex        =   110
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
         MouseIcon       =   "TCustomers.frx":11FC
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   126
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
         Left            =   2910
         MouseIcon       =   "TCustomers.frx":168C
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   125
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
         Left            =   2220
         MouseIcon       =   "TCustomers.frx":1B2E
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":1C80
         Style           =   1  'Graphical
         TabIndex        =   124
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
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cboContactRelation 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "TCustomers.frx":1FAB
         Left            =   1140
         List            =   "TCustomers.frx":1FAD
         TabIndex        =   112
         Top             =   790
         Width           =   3045
      End
      Begin VB.TextBox txtContactPosition 
         Height          =   345
         Left            =   1140
         TabIndex        =   114
         Top             =   1190
         Width           =   3045
      End
      Begin VB.TextBox txtContactDepartment 
         Height          =   345
         Left            =   1140
         TabIndex        =   115
         Top             =   1590
         Width           =   3045
      End
      Begin VB.TextBox txtContactPhone 
         Height          =   345
         Left            =   1140
         TabIndex        =   117
         Top             =   1990
         Width           =   3045
      End
      Begin VB.TextBox txtContactMobile 
         Height          =   345
         Left            =   1140
         TabIndex        =   119
         Top             =   2390
         Width           =   3045
      End
      Begin VB.TextBox txtContactAddress 
         Height          =   645
         Left            =   1140
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   122
         Top             =   2790
         Width           =   3045
      End
      Begin VB.Label labIDContacts 
         Height          =   555
         Left            =   1470
         TabIndex        =   123
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
         TabIndex        =   111
         Top             =   870
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   330
         Left            =   0
         TabIndex        =   107
         Top             =   0
         Width           =   4425
         _Version        =   655364
         _ExtentX        =   7805
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "MULTIPLE CONTACTS"
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
         TabIndex        =   109
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
         TabIndex        =   113
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
         TabIndex        =   116
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
         TabIndex        =   118
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
         TabIndex        =   120
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
         TabIndex        =   121
         Top             =   2970
         Width           =   765
      End
   End
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00DD624D&
      ForeColor       =   &H80000008&
      Height          =   3285
      Left            =   3870
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   3255
      ScaleWidth      =   4500
      TabIndex        =   128
      Top             =   2670
      Visible         =   0   'False
      Width           =   4530
      Begin VB.CommandButton Command2 
         Caption         =   "All History"
         Height          =   405
         Left            =   1095
         TabIndex        =   161
         Top             =   2685
         Width           =   2775
      End
      Begin VB.CommandButton cmdHist_Visit 
         Caption         =   "Visit History"
         Height          =   405
         Left            =   1080
         TabIndex        =   135
         Top             =   2250
         Width           =   2775
      End
      Begin VB.CommandButton cmdHist_Vehicles 
         Caption         =   "Customer Vehicles Information"
         Height          =   405
         Left            =   1080
         MouseIcon       =   "TCustomers.frx":1FAF
         MousePointer    =   99  'Custom
         TabIndex        =   133
         Top             =   1365
         Width           =   2775
      End
      Begin VB.CommandButton cmdHist_Calls 
         Caption         =   "Call History"
         Height          =   405
         Left            =   1080
         TabIndex        =   134
         Top             =   1800
         Width           =   2775
      End
      Begin VB.CommandButton cmdHist_Service 
         Caption         =   "Vehicle Service History"
         Height          =   405
         Left            =   1080
         TabIndex        =   132
         Top             =   915
         Width           =   2775
      End
      Begin VB.CommandButton cmdHist_Sales 
         Caption         =   "Vehicle Sales History"
         Height          =   405
         Left            =   1080
         TabIndex        =   131
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton cmdCloseHistory 
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
         Left            =   4140
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   0
         Width           =   315
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   330
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   4515
         _Version        =   655364
         _ExtentX        =   7964
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::Customer History/Information::"
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
      TabIndex        =   137
      Top             =   2880
      Visible         =   0   'False
      Width           =   4380
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
         TabIndex        =   138
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
         Left            =   2100
         MouseIcon       =   "TCustomers.frx":2101
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":2253
         Style           =   1  'Graphical
         TabIndex        =   146
         Top             =   1650
         Width           =   645
      End
      Begin VB.CommandButton Command6 
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
         Left            =   2790
         MouseIcon       =   "TCustomers.frx":257E
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":26D0
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   1650
         Width           =   645
      End
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
         MouseIcon       =   "TCustomers.frx":2A20
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":2B72
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   1650
         Width           =   645
      End
      Begin VB.TextBox txtChildName 
         Height          =   345
         Left            =   1200
         TabIndex        =   141
         Top             =   390
         Width           =   3015
      End
      Begin VB.ComboBox cboChildSex 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00400000&
         Height          =   345
         ItemData        =   "TCustomers.frx":2EB0
         Left            =   1200
         List            =   "TCustomers.frx":2EBD
         TabIndex        =   145
         Top             =   1170
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtChildDate 
         Height          =   345
         Left            =   1200
         TabIndex        =   143
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
         TabIndex        =   140
         Top             =   390
         Width           =   540
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   139
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
         TabIndex        =   142
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label labIdCHILD 
         Height          =   555
         Left            =   1290
         TabIndex        =   149
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
         TabIndex        =   144
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
      Begin VB.Frame fraSearch 
         Height          =   6675
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
            Height          =   4305
            Left            =   30
            TabIndex        =   9
            Top             =   2280
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   7594
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
            MouseIcon       =   "TCustomers.frx":2ECA
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CUSTOMER NAME"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Height          =   7455
         Left            =   2490
         TabIndex        =   10
         Top             =   -90
         Width           =   9735
         Begin VB.Frame fraCompany 
            Caption         =   "Company Information"
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
            Height          =   810
            Left            =   90
            TabIndex        =   31
            Top             =   2220
            Width           =   9555
            Begin VB.TextBox txtCompanyAdd 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3360
               TabIndex        =   35
               Top             =   420
               Width           =   6045
            End
            Begin VB.TextBox txtCusComp 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   120
               TabIndex        =   34
               Top             =   420
               Width           =   3165
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Company Address"
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
               Left            =   3345
               TabIndex        =   33
               Top             =   195
               Width           =   1545
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Company Name"
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
               Left            =   165
               TabIndex        =   32
               Top             =   195
               Width           =   1290
            End
         End
         Begin VB.TextBox txtAcctName 
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00400000&
            Height          =   330
            Left            =   5640
            TabIndex        =   15
            Top             =   495
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   3135
         End
         Begin VB.Frame fraEntity 
            Caption         =   "Personal Contact Name"
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
            Height          =   1410
            Left            =   60
            TabIndex        =   16
            Top             =   810
            Width           =   9555
            Begin VB.ComboBox cboApod 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   21
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
               TabIndex        =   24
               Top             =   420
               Width           =   3030
            End
            Begin VB.TextBox txtFirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3780
               TabIndex        =   23
               Top             =   420
               Width           =   2625
            End
            Begin VB.TextBox txtLastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1020
               TabIndex        =   22
               ToolTipText     =   "LAST NAME OR COMPANY NAME"
               Top             =   420
               Width           =   2625
            End
            Begin VB.ComboBox cboSex 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   120
               TabIndex        =   28
               Text            =   "cboSex"
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtBirthDate 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   1020
               TabIndex        =   29
               Top             =   990
               Width           =   2655
            End
            Begin VB.TextBox txtSpouse 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3780
               TabIndex        =   30
               Top             =   990
               Width           =   5685
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Salutation"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   120
               TabIndex        =   17
               Top             =   210
               Width           =   825
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   6420
               TabIndex        =   20
               Top             =   210
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   3780
               TabIndex        =   19
               Top             =   210
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1020
               TabIndex        =   18
               Top             =   210
               Width           =   915
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Sex"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   120
               TabIndex        =   27
               Top             =   780
               Width           =   300
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Birth Date"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1020
               TabIndex        =   25
               Top             =   750
               Width           =   810
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse Name"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   3780
               TabIndex        =   26
               Top             =   750
               Width           =   1200
            End
         End
         Begin VB.Frame fraMiscellenous 
            Height          =   4425
            Left            =   90
            TabIndex        =   36
            Top             =   2970
            Width           =   4515
            Begin VB.TextBox txtFax 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   52
               Top             =   2745
               Width           =   3165
            End
            Begin VB.TextBox txtHomePhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   50
               Top             =   2385
               Width           =   3165
            End
            Begin VB.TextBox txtMobile 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   48
               Top             =   2025
               Width           =   3165
            End
            Begin VB.TextBox txtCusphon1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   46
               Top             =   1665
               Width           =   3165
            End
            Begin VB.TextBox txtAsstPhone 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   56
               Top             =   3465
               Width           =   3165
            End
            Begin VB.TextBox txtAssistant 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   54
               Top             =   3105
               Width           =   3165
            End
            Begin VB.TextBox txtEmail 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   44
               Top             =   1305
               Width           =   3165
            End
            Begin VB.TextBox txtDepartment 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   42
               Top             =   945
               Width           =   3165
            End
            Begin VB.TextBox txtTitle 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1290
               TabIndex        =   40
               Top             =   585
               Width           =   3165
            End
            Begin VB.ComboBox cboLeadSource 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1290
               TabIndex        =   37
               Text            =   "cboLeadSource"
               Top             =   210
               Width           =   3150
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   870
               TabIndex        =   51
               Top             =   2775
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
               TabIndex        =   49
               Top             =   2415
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
               TabIndex        =   47
               Top             =   2055
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
               TabIndex        =   45
               Top             =   1695
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
               TabIndex        =   55
               Top             =   3495
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
               TabIndex        =   53
               Top             =   3135
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
               TabIndex        =   43
               Top             =   1335
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
               TabIndex        =   41
               Top             =   975
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Title"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   810
               TabIndex        =   39
               Top             =   615
               Width           =   345
            End
            Begin VB.Label Label12 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Lead Source"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   150
               TabIndex        =   38
               Top             =   255
               Width           =   1215
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Personal Address Information"
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
            Height          =   1725
            Left            =   4620
            TabIndex        =   57
            Top             =   3000
            Width           =   4995
            Begin VB.TextBox txtPersonalZIP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1380
               MaxLength       =   5
               TabIndex        =   65
               Top             =   1350
               Width           =   3495
            End
            Begin VB.TextBox txtPersonalState 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   315
               Left            =   1380
               MaxLength       =   30
               TabIndex        =   63
               Top             =   990
               Width           =   3495
            End
            Begin VB.TextBox txtPersonalStreet 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1380
               ScrollBars      =   2  'Vertical
               TabIndex        =   59
               Top             =   240
               Width           =   3495
            End
            Begin VB.ComboBox cboPersonalCity 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   1380
               TabIndex        =   61
               Text            =   "cboApod"
               Top             =   615
               Width           =   3495
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Zip Code"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   540
               TabIndex        =   64
               Top             =   1380
               Width           =   750
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "State/Province"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   120
               TabIndex        =   62
               Top             =   1020
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Street"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   810
               TabIndex        =   58
               Top             =   300
               Width           =   480
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "City"
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   990
               TabIndex        =   60
               Top             =   660
               Width           =   300
            End
         End
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
            Left            =   4620
            TabIndex        =   66
            Top             =   4680
            Width           =   4995
            Begin VB.TextBox txtDeliveryAddress 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   825
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   68
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
               TabIndex        =   67
               Top             =   150
               Width           =   1395
            End
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
            Height          =   1365
            Left            =   4650
            TabIndex        =   69
            Top             =   5970
            Width           =   4965
            Begin VB.TextBox txtNotes 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   1065
               Left            =   60
               MaxLength       =   300
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   70
               Top             =   210
               Width           =   4845
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
            Left            =   4680
            TabIndex        =   14
            Top             =   540
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
            TabIndex        =   13
            Top             =   540
            Width           =   1335
         End
         Begin XtremeShortcutBar.ShortcutCaption CapInfo 
            Height          =   270
            Index           =   2
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   135
            Width           =   9660
            _Version        =   655364
            _ExtentX        =   17039
            _ExtentY        =   476
            _StockProps     =   14
            Caption         =   "Customers Information"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            ForeColor       =   64
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   10650
         ScaleHeight     =   885
         ScaleWidth      =   1590
         TabIndex        =   88
         Top             =   7380
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
            Left            =   765
            MouseIcon       =   "TCustomers.frx":302C
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":317E
            Style           =   1  'Graphical
            TabIndex        =   90
            Top             =   30
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
            Left            =   30
            MouseIcon       =   "TCustomers.frx":34BC
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":360E
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   1860
         Left            =   0
         ScaleHeight     =   1860
         ScaleWidth      =   12315
         TabIndex        =   71
         Top             =   6570
         Width           =   12315
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
            Left            =   6945
            MouseIcon       =   "TCustomers.frx":395E
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":3AB0
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   840
            Width           =   735
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
            Left            =   7695
            MouseIcon       =   "TCustomers.frx":3E0E
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":3F60
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   840
            Width           =   735
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
            Left            =   11385
            MouseIcon       =   "TCustomers.frx":42B0
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":4402
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   840
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
            Left            =   10635
            MouseIcon       =   "TCustomers.frx":4768
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":48BA
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   840
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
            Left            =   9900
            MouseIcon       =   "TCustomers.frx":4C20
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":4D72
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   840
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
            Left            =   9180
            MouseIcon       =   "TCustomers.frx":509D
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":51EF
            Style           =   1  'Graphical
            TabIndex        =   84
            Top             =   840
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
            Left            =   8460
            MouseIcon       =   "TCustomers.frx":554B
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":569D
            Style           =   1  'Graphical
            TabIndex        =   83
            Top             =   840
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
            Left            =   6225
            MouseIcon       =   "TCustomers.frx":59B0
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":5B02
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   840
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
            Left            =   5490
            MouseIcon       =   "TCustomers.frx":5DFC
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":5F4E
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   840
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
            Left            =   4770
            MouseIcon       =   "TCustomers.frx":62A6
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":63F8
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   840
            Width           =   705
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Child Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            MouseIcon       =   "TCustomers.frx":6757
            MousePointer    =   99  'Custom
            TabIndex        =   72
            Top             =   30
            Width           =   1965
         End
         Begin VB.CommandButton cmdAddChildInfo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            MouseIcon       =   "TCustomers.frx":68A9
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":69FB
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   30
            Width           =   375
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Contact Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            MouseIcon       =   "TCustomers.frx":6BC5
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   420
            Width           =   1965
         End
         Begin VB.CommandButton CmdaddContactInfo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2040
            MouseIcon       =   "TCustomers.frx":6D17
            MousePointer    =   99  'Custom
            Picture         =   "TCustomers.frx":6E69
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   420
            Width           =   375
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Credit && Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            MouseIcon       =   "TCustomers.frx":7033
            MousePointer    =   99  'Custom
            TabIndex        =   76
            Top             =   810
            Width           =   1965
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Customer History"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   30
            MouseIcon       =   "TCustomers.frx":7185
            MousePointer    =   99  'Custom
            TabIndex        =   77
            Top             =   1200
            Width           =   1965
         End
      End
   End
   Begin VB.PictureBox picContactList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   3188
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   94
      Top             =   1725
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
         MouseIcon       =   "TCustomers.frx":72D7
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":7429
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   4110
         Width           =   705
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
         Left            =   3570
         MouseIcon       =   "TCustomers.frx":7767
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":78B9
         Style           =   1  'Graphical
         TabIndex        =   97
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
         Left            =   4290
         MouseIcon       =   "TCustomers.frx":7BCC
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":7D1E
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   4110
         Width           =   705
      End
      Begin MSComctlLib.ListView lvContactList 
         Height          =   3735
         Left            =   60
         TabIndex        =   96
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
         MouseIcon       =   "TCustomers.frx":807A
         NumItems        =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   95
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
   Begin VB.PictureBox picChildList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   3188
      ScaleHeight     =   4815
      ScaleWidth      =   5835
      TabIndex        =   100
      Top             =   1725
      Visible         =   0   'False
      Width           =   5865
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
         Left            =   4290
         MouseIcon       =   "TCustomers.frx":81DC
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":832E
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   4080
         Width           =   705
      End
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
         Left            =   5010
         MouseIcon       =   "TCustomers.frx":866A
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":87BC
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   4080
         Width           =   705
      End
      Begin VB.CommandButton Command5 
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
         Left            =   3540
         MouseIcon       =   "TCustomers.frx":8AFA
         MousePointer    =   99  'Custom
         Picture         =   "TCustomers.frx":8C4C
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   4080
         Width           =   705
      End
      Begin MSComctlLib.ListView lvChildList 
         Height          =   3735
         Left            =   60
         TabIndex        =   102
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
         MouseIcon       =   "TCustomers.frx":8F5F
         NumItems        =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   101
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
      TabIndex        =   91
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
      TabIndex        =   93
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
      TabIndex        =   127
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
      TabIndex        =   92
      Top             =   660
      Visible         =   0   'False
      Width           =   1545
   End
End
Attribute VB_Name = "frmCRIS_Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS                                 As ADODB.Recordset
Dim rsCusCtl                           As ADODB.Recordset
Dim ADDOrEdit                          As String
Dim AccountCode                        As String
Dim CustType                           As String
Dim EntryPoint                         As String

Dim TempProspectID                     As Long
Event ChangedData(xCUSCODE As String)
Event ProspectConverted(CustomerCode As String, xGoingWhere As String, ProspectID As Long)
Dim GoingWhere                         As String

Public Sub AddCustomerFromProspect(oRs As Recordset, xGoingWhere As String)
    Dim ar                             As Variant
    GoingWhere = xGoingWhere
    If Not (oRs.EOF Or oRs.BOF) Then
        EntryPoint = "PROSPECT"
        ADDOrEdit = "ADD"
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
        If CustType = "P" Then

            txtPersonalStreet = Null2String(oRs!Address)
        Else
            txtCompanyAdd = Null2String(oRs!Address)
        End If
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
            fraEntity.Caption = "Personal Contact"
            fraCompany.Caption = "Company Information"
            CustType = "P"
        Case "Company/Agency"
            fraEntity.Caption = "Contact Person"
            fraCompany.Caption = "Company Information"
            CustType = "C"
        Case "Fleet Account"
            fraEntity.Caption = "Contact Person"
            fraCompany.Caption = "Company Information"

            CustType = "F"

        Case "Government"
            fraEntity.Caption = "Contact Person"
            fraCompany.Caption = "Government/Agency Information"
            CustType = "G"
    End Select


End Sub

Private Sub cboSearchCustype_Click()
    FillSearchGrid txtSearch
End Sub

Private Sub cmdAdd_Click()
    ADDOrEdit = "ADD"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    lstCustomer.Enabled = False


    InitMemVars

End Sub

Private Sub cmdAddChildInfo_Click()
    cmdDeleteChild.Enabled = False
    txtChildDate = ""
    txtChildName = ""
    cboChildSex = ""
    labIdCHILD = ""
    ShowHidePictureBox2 picChildAE, True, picMain
    txtChildName.SetFocus
End Sub

Private Sub CmdaddContactInfo_Click()

    cmdDeleteContact.Enabled = False
    txtContactAddress = ""
    cboContactRelation = ""
    txtContactDepartment = ""
    txtContactMobile = ""
    txtContactName = ""
    txtContactPhone = ""
    txtContactPosition = ""
    labIDContacts = "0"

    ShowHidePictureBox2 picContactAE, True, picMain
    txtContactName.SetFocus
End Sub

Private Sub cmdCancel_Click()
    ShowHidePictureBox2 picChildList, False
    ShowHidePictureBox2 picChildAE, False, picMain
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    lstCustomer.Enabled = True
    fraSearch.Enabled = True
    ADDOrEdit = ""
    StoreMemvars
End Sub

Private Sub cmdCancelChildList_Click()
    ShowHidePictureBox2 picChildList, False, picMain

End Sub

Private Sub cmdCancelContactList_Click()
    ShowHidePictureBox2 picContactList, False, picMain
End Sub

Private Sub cmdCloseChildAE_Click(Index As Integer)
    ShowHidePictureBox2 picChildAE, False, picMain
End Sub

Private Sub cmdCloseContactsAE_Click(Index As Integer)
    ShowHidePictureBox2 picContactAE, False, picMain
End Sub

Private Sub cmdCloseHistory_Click()
    ShowHidePictureBox2 picHistory, False, picMain
End Sub

Private Sub cmdCloseTerm_Click(Index As Integer)
    ShowHidePictureBox2 picCredit, False, picMain
End Sub

Private Sub cmdDelete_Click()
    Dim lng                            As Integer

    lng = gconDMIS.Execute("SELECT COUNT(*) from CRIS_PROSPECTS WHERE CUSCDE=" & N2Str2Null(txtCuscde)).Fields(0).Value


    If lng = 0 Then
        If ShowConfirmDelete = True Then
            Screen.MousePointer = 11
            gconDMIS.Execute "Delete from ALL_CUSTOMER  where ID=" & labid
            gconDMIS.Execute "Delete from ALL_CUSTOMER_CONTACTS  where CUSCDE=" & N2Str2Null(txtCuscde)
            gconDMIS.Execute "Delete from ALL_CUSTOMER_CHILD  where CUSCDE=" & N2Str2Null(txtCuscde)
            gconDMIS.Execute "Delete from ALL_CusCtl"

            Dim rsCUSTOMER             As ADODB.Recordset
            Dim k                      As Integer
            Dim NewCtlCde              As String
            For k = 65 To 90
                Set rsCUSTOMER = New ADODB.Recordset
                rsCUSTOMER.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
                If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
                    NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCUSTOMER!CODE, 2, 5)) + 1, "00000")
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
                Else
                    gconDMIS.Execute "insert into ALL_CusCtl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
                End If
            Next
            Screen.MousePointer = 0
            FillSearchGrid ""
            rsRefresh
            StoreMemvars
            MessagePop Delete, "Record Deleted", "Customer Information Deleted. "
        End If

    Else
        MessagePop RecLocekd, "Record Cannot Be Deleted", "Customer Information Cannot be deleted. Prospect Information Exists"
    End If
    rsRefresh
    RS.Bookmark = rsFind(RS.Clone, "ID", labid).Bookmark
    InitMemVars
    StoreMemvars
End Sub

Private Sub cmdDeleteChild_Click()
    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub



    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CHILD WHERE id=" & labIdCHILD
    ShowHidePictureBox2 picChildAE, False, picMain


End Sub

Private Sub cmdDeleteContact_Click()
    If MsgBox("Msgbox ""Are You Sure You Want to Delete this Information""", vbQuestion + vbOKCancel, "Delete?") = vbCancel Then: Exit Sub
    gconDMIS.Execute "DELETE FROM ALL_CUSTOMER_CONTACTS WHERE id=" & labIDContacts
    ShowHidePictureBox2 picContactAE, False, picMain

End Sub

Private Sub cmdEdit_Click()
    ADDOrEdit = "EDIT"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
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
    txtSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RS.MoveFirst
    ShowLastRecordMsg
End Sub

Private Sub cmdHist_Calls_Click()
    frmSMIS_Inquiry_CallVisit_History.Otp(2).Value = True
    frmSMIS_Inquiry_CallVisit_History.txtSearchKey_All = txtCuscde
    frmSMIS_Inquiry_CallVisit_History.TabControl.SelectedItem = 1
    frmSMIS_Inquiry_CallVisit_History.Show
End Sub

Private Sub cmdHist_Sales_Click()
    frmSMIS_Inquiry_CustomerSalesHistory.optSearchActiveCode.Value = True
    frmSMIS_Inquiry_CustomerSalesHistory.txtSearchKey_All = txtCuscde
    frmSMIS_Inquiry_CustomerSalesHistory.Show
End Sub

Private Sub cmdHist_Service_Click()
    frmCSMSCustomerHistory.cmdAll.Value = True
    frmCSMSCustomerHistory.Otp(2).Value = True
    frmCSMSCustomerHistory.txtSearchKey_All = txtCuscde
    frmCSMSCustomerHistory.Show
End Sub

Private Sub cmdHist_Vehicles_Click()

    With FrmCSMSAddVehicle
        .CustomerCode = txtCuscde
        .labCustomer.Caption = UCase(txtLastName & IIf(txtFirstName = "", "", ",") & txtFirstName & IIf(txtMiddleName = "", "", ".") & Left(txtMiddleName, 1))
    End With
    FrmCSMSAddVehicle.Show 1
End Sub

Private Sub cmdHist_Visit_Click()
    frmSMIS_Inquiry_CallVisit_History.Otp(2).Value = True
    frmSMIS_Inquiry_CallVisit_History.txtSearchKey_All = txtCuscde
    frmSMIS_Inquiry_CallVisit_History.TabControl.SelectedItem = 0
    frmSMIS_Inquiry_CallVisit_History.Show
End Sub

Private Sub cmdLast_Click()
    RS.MoveLast
    ShowLastRecordMsg
End Sub

Private Sub cmdNext_Click()
    RS.MoveNext
    If RS.EOF Then
        RS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdPrevious_Click()
    RS.MovePrevious
    If RS.BOF Then
        RS.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemvars

End Sub

Private Sub cmdPrint_Click()
'    frmSMIS_ReportChoice.REPORTNAME = "CUSTOMERLISTING"
'    frmSMIS_ReportChoice.Show 1

End Sub

Private Sub cmdSave_Click()

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

    If CustType <> "P" And txtCusComp = "" Then
        ShowIsRequiredMsg "Company Name"
        On Error Resume Next
        txtCusComp.SetFocus
        Exit Sub
    End If




    If ADDOrEdit = "ADD" Then


        Dim rsfindDup                  As ADODB.Recordset
        Set rsfindDup = New ADODB.Recordset
        rsfindDup.Open "select * from ALL_CustMaster_Smis where Code = '" & txtCuscde & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsfindDup.EOF And Not rsfindDup.BOF Then
            MsgSpeechBox "Code already exist!"
            Exit Sub
        End If
        If CustType = "P" Then
            txtCuscde = GetCustomerCode(txtLastName)
        Else
            txtCuscde = GetCustomerCode(txtCusComp)
        End If
    End If






    Dim vtxtCusCde                     As String
    Dim VTXTLASTNAME                   As String
    Dim VTXTFIRSTNAME                  As String
    Dim vtxtMiddleInitial              As String
    Dim vtxtCUSCOMP                    As String

    Dim vcboSex                        As String
    Dim vtxtCusadd1                    As String
    Dim vtxtCusadd2                    As String
    Dim vtxtCuszipc                    As String
    Dim vtxtCusphon1                   As String
    Dim vtxtAcctName                   As String
    Dim vcboApod                       As String
    Dim vcboLeadSource                 As String
    Dim vtxtTitle                      As String
    Dim vtxtDepartment                 As String
    Dim vtxtEmail                      As String
    Dim vtxtMobile                     As String
    Dim vtxtHomePhone                  As String
    Dim VtxtFax                        As String
    Dim vtxtAssistant                  As String
    Dim vtxtAsstPhone                  As String
    Dim vtxtCity                       As String
    Dim vTxtBirthDate                  As String
    Dim vTxtSpouse                     As String
    Dim vtxtDescription                As String
    Dim vtxtCustType                   As String
    Dim vtxtCompanyAdd                 As String
    Dim TEMPSQL                        As String
    Dim vtxtDeliveryAddress            As String
    vtxtCompanyAdd = N2Str2Null(UCase(txtCompanyAdd))
    vtxtCustType = N2Str2Null(CustType)
    vcboApod = N2Str2Null(UCase(cboApod))
    vtxtCusCde = N2Str2Null(txtCuscde)
    VTXTLASTNAME = N2Str2Null(UCase(txtLastName))
    VTXTFIRSTNAME = N2Str2Null(UCase(txtFirstName))
    vtxtMiddleInitial = N2Str2Null(txtMiddleName)
    vtxtAcctName = N2Str2Null(txtAcctName)
    vtxtCUSCOMP = N2Str2Null(UCase(txtCusComp))
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

    If ADDOrEdit = "ADD" Then
        TEMPSQL = "INSERT INTO ALL_CUSTOMER(" & vbCrLf
        TEMPSQL = TEMPSQL & " CUSCOMP, APOD , CUSCDE , LASTNAME, FIRSTNAME, MIDDLEINITIAL,ACCTNAME,SEX,CUSTOMERADD,PROVINCIALADD,ZIPCODE,TELEPHONENO,LEADSOURCE,TITLE,DEPARTMENT,EMAIL,MOBILE,HOMEPHONE,FAX,ASSISTANT,ASSTPHONE,CITY,BIRTHDATE,SPOUSE,DESCRIPTION, CUSTYPE, COMPANYADD " & vbCrLf
        TEMPSQL = TEMPSQL & " ) VALUES ( " & vbCrLf
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
    Dim rsCUSTOMER                     As ADODB.Recordset
    Dim k                              As Integer
    Dim NewCtlCde                      As String
    For k = 65 To 90
        Set rsCUSTOMER = New ADODB.Recordset
        rsCUSTOMER.Open "select Code from ALL_CustMaster_Smis where left(Code,1) = '" & Chr(k) & "' order by Code desc", gconDMIS
        If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCUSTOMER!CODE, 2, 5)) + 1, "00000")
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

        RS.Requery
        If ADDOrEdit = "EDIT" Then
            RS.Find "id =" & labid
        End If
        FillSearchGrid txtSearch
    End If

    UPDATELOGTABLE "ALL_CUSTOMER", labid
    cmdCancel.Value = True

End Sub

Private Sub cmdSaveChild_Click()

End Sub

Private Sub cmdSaveContact_Click()
    If RTrim(LTrim(txtContactName)) = "" Then
        ShowIsRequiredMsg "CONTACT NAME"
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
        Command7_Click
    End If
    ShowHidePictureBox2 picContactAE, False, picMain

End Sub

Private Sub cmdSelectChild_Click()
    lvChildList_KeyPress 13
End Sub

Private Sub Command1_Click()
'AddColumnHeader "ChildName,Sex,DateofBirth,Age", lvChildList

    flex_FillListView gconDMIS.Execute("SELECT CHILDNAME,SEX,DOB, DATEDIFF(YEAR, DOB, GETDATE()) ,ID FROM ALL_CUSTOMER_CHILD WHERE CUSCDE=" & N2Str2Null(txtCuscde)), lvChildList
    ShowHidePictureBox2 picChildList, True, picMain

    If lvChildList.ListItems.Count = 0 Then
        cmdSelectChild.Enabled = False
    Else
        cmdSelectChild.Enabled = True
    End If
End Sub

Private Sub Command10_Click()
    ShowHidePictureBox2 picHistory, True, picMain

End Sub

Private Sub Command12_Click()
    Dim cTerm, cLimit, cDays

    cLimit = NumericVal(txtCreditLimit)
    cDays = NumericVal(txtCreditDays)

    
    
    gconDMIS.Execute "update all_Customer set CreditLimit=" & cLimit & ", CREDITTERM='C', CREDITDAYS=" & cDays & " WHERE ID=" & labid
        
    
    MessagePop RecSaveOk, "Credit Info", "Credit Information Updated"
    ShowHidePictureBox2 picCredit, False, picMain
End Sub

Private Sub Command2_Click()
    frmSMIS_Inquiry_ViewLog.SHOWCUSTOMERLOG txtCuscde, txtAcctName
    frmSMIS_Inquiry_ViewLog.Show
End Sub

Private Sub Command3_Click()
    txtDeliveryAddress = txtPersonalStreet & "," & txtPersonalState & "," & cboPersonalCity & "," & txtPersonalZIP
End Sub

Private Sub Command4_Click()
    CmdaddContactInfo_Click
End Sub

Private Sub Command5_Click()

    cmdAddChildInfo_Click
End Sub

Private Sub Command6_Click()
    If RTrim(LTrim(txtChildName)) = "" Then
        ShowIsRequiredMsg "CHILD NAME"
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
        Command1_Click
    End If
    ShowHidePictureBox2 picChildAE, False, picMain
End Sub

Private Sub Command7_Click()

'SELECT  ContactName,    Relation,Phone, Mobile,  ContactPosition, Department, Address, ID



    flex_FillListView gconDMIS.Execute("SELECT  CONTACTNAME, RELATION,PHONE, MOBILE, CONTACTPOSITION, DEPARTMENT, ADDRESS, ID FROM ALL_CUSTOMER_CONTACTS WHERE CUSCDE=" & N2Str2Null(txtCuscde)), lvContactList
    ShowHidePictureBox2 picContactList, True, picMain

    If lvContactList.ListItems.Count = 0 Then
        cmdEditContact.Enabled = False
    Else
        cmdEditContact.Enabled = True
    End If
End Sub

Private Sub Command8_Click()
    ShowHidePictureBox2 picCredit, True, picMain
    txtCreditLimit.SetFocus
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsCustomer2                    As ADODB.Recordset
    Dim Key                            As String
    Dim LIMITKEY                       As String
    lstCustomer.Sorted = False
    lstCustomer.ListItems.Clear

    Set rsCustomer2 = New ADODB.Recordset
    xxx = Replace(Trim(xxx), "'", "")

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
        Case 0                                        'Search All
            LIMITKEY = "'P','C','F','G', NULL"
        Case 1                                        'Only Personal Customers
            LIMITKEY = "'P'"
        Case 2                                        ' Only Company/Agency Customers
            LIMITKEY = "'C'"
        Case 3                                        'Only Government Customer
            LIMITKEY = "'G'"
        Case 4                                        'Only Fleet Account Customer
            LIMITKEY = "'F'"
    End Select


    Set rsCustomer2 = gconDMIS.Execute("select TOP 100 AcctName as CustomerName, id  from ALL_CUSTOMER where CUSCDE <> '999999' and " & Key & " like'" & xxx & "%' AND CUSTYPE IN (" & LIMITKEY & " )  order by 1 asc")

    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
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
    ElseIf picHistory.Visible = True And KeyCode = 27 Then
        cmdCloseHistory_Click
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
    CenterMe frmMain, Me, 0
    Frame1.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    InitMemVars
    InitData

    AddColumnHeader "CHILDNAME,SEX,DATEOFBIRTH,AGE", lvChildList
    ResizeColumnHeader lvChildList, "50,15,15,15"

    AddColumnHeader "CONTACTNAME,RELATION,PHONE,MOBILE", lvContactList
    ResizeColumnHeader lvContactList, "40,20,17,17"
    rsRefresh




    If AccountCode <> "" Then
        RS.Find ("CUSCDE=" & N2Str2Null(AccountCode))
        StoreMemvars
        cmdEdit.Value = True
    End If
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsCusCtl = Nothing
    ADDOrEdit = vbNullString
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
        .AddItem ("Search Only Personal Customers")
        .AddItem ("Search Only Company/Agency Customers")
        .AddItem ("Search Only Government Customer")
        .AddItem ("Search Only Fleet Account Customer")
        .ListIndex = 0
    End With


    SetComboWidth cboSearchCustype, 300

End Sub

Sub InitMemVars()
    Dim temprs                         As ADODB.Recordset
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
    txtCusComp = ""
    txtCompanyAdd = ""
    txtCreditDays = "0"
    txtCreditLimit = "0.00"
    
    Dim rsAPOD                         As ADODB.Recordset
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
    Set temprs = gconDMIS.Execute("Select DataDesc from CRIS_vW_MasterPullDown where  Masterdesc='Lead Source'")
    cboLeadSource.Clear

    While Not temprs.EOF
        cboLeadSource.AddItem temprs.Collect(0)
        temprs.MoveNext
    Wend
    cboSex.Clear
    cboSex.AddItem "NA"
    cboSex.AddItem "M"
    cboSex.AddItem "F"

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
    cmdEdit.Value = True
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    RS.MoveFirst
    RS.Find ("ID=" & Item.ListSubItems(1).Text)

    StoreMemvars
End Sub

Private Sub lvChildList_DblClick()
    lvChildList_KeyPress 13
End Sub

Private Sub lvChildList_KeyPress(KeyAscii As Integer)
    If lvChildList.SelectedItem Is Nothing Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    'AddColumnHeader "ChildName,Sex,DateofBirth,Age", lvChildList

    txtChildName = lvChildList.SelectedItem.Text
    cboChildSex = lvChildList.SelectedItem.ListSubItems(1).Text
    txtChildDate = lvChildList.SelectedItem.ListSubItems(2).Text
    labIdCHILD = lvChildList.SelectedItem.ListSubItems(4).Text
    cmdDeleteChild.Enabled = True

    ShowHidePictureBox2 picChildList, False
    ShowHidePictureBox2 picChildAE, True, picMain
    txtChildName.SetFocus
End Sub

Private Sub lvContactList_DblClick()
    lvContactList_KeyPress 13
End Sub

Private Sub lvContactList_KeyPress(KeyAscii As Integer)
    If lvContactList.SelectedItem Is Nothing Then Exit Sub
    If KeyAscii <> 13 Then Exit Sub
    'CONTACTNAME, RELATION,PHONE, MOBILE, CONTACTPOSITION, DEPARTMENT, ADDRESS, ID
    ShowHidePictureBox2 picContactAE, True, picMain
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
    ShowHidePictureBox2 picContactList, False, picMain

    txtContactName.SetFocus
End Sub

Sub rsRefresh()
    Set RS = New ADODB.Recordset
    RS.Open "Select * from ALL_CUSTOMER order by id DESC", gconDMIS, adOpenKeyset, adLockReadOnly


End Sub

Sub SetCustomerAccountName()
    If EntryPoint = "PROSPECT" Or ADDOrEdit = "EDIT" Or ADDOrEdit = "" Then: Exit Sub

    'If CustType = "P" Then
    txtAcctName = UCase(txtLastName & IIf(txtFirstName = "", "", ",") & txtFirstName & IIf(txtMiddleName = "", "", ".") & Left(txtMiddleName, 1))
    'Else
    '   txtAcctName.Text = UCase(txtCusComp)
    'End If


End Sub

Sub StoreMemvars()

    If Not RS.EOF And Not RS.BOF Then
        labid.Caption = RS!Id
        cboApod.Text = Null2String(RS!APOD)
        txtCuscde.Text = Null2String(RS!CUSCDE)
        txtLastName.Text = Null2String(RS!lastname)
        txtFirstName.Text = Null2String(RS!FirstName)
        txtMiddleName.Text = Null2String(RS!MiddleInitial)

        txtCusComp.Text = Null2String(RS!CUSCOMP)

        txtCompanyAdd = Null2String(RS!CompanyAdd)
        cboSex.Text = Null2String(RS!Sex)
        txtAcctName = Null2String(RS!AcctName)

        txtPersonalStreet.Text = Null2String(RS!CustomerAdd)
        txtPersonalState.Text = Null2String(RS!provincialadd)
        txtPersonalZIP.Text = Null2String(RS!ZIPCODE)
        txtCusphon1.Text = Null2String(RS!TelephoneNo)

        cboLeadSource.Text = Null2String(RS!LeadSource)
        txtTitle.Text = Null2String(RS!TITLE)
        txtDepartment.Text = Null2String(RS!Department)
        txtEmail.Text = Null2String(RS!EMAIL)
        txtMobile.Text = Null2String(RS!Mobile)
        txtHomePhone.Text = Null2String(RS!HomePhone)
        txtFax.Text = Null2String(RS!Fax)
        txtAssistant.Text = Null2String(RS!Assistant)
        txtAsstPhone.Text = Null2String(RS!AsstPhone)
        cboPersonalCity.Text = Null2String(RS!City)
        txtBirthDate.Text = Null2String(RS!BirthDate)
        txtSpouse.Text = Null2String(RS!Spouse)
        txtDeliveryAddress = Null2String(RS!DELIVERYADDRESS)
        txtNotes.Text = Null2String(RS!Description)

        txtCreditDays = NumericVal(RS!CREDITDAYS)
        txtCreditLimit = FormatNumber(NumericVal(RS!CreditLimit))

        

        If Null2String(RS!CUSTYPE) = "P" Then
            cboCustType.ListIndex = 0
        ElseIf Null2String(RS!CUSTYPE) = "C" Then
            cboCustType.ListIndex = 1
        ElseIf Null2String(RS!CUSTYPE) = "G" Then
            cboCustType.ListIndex = 2
        ElseIf Null2String(RS!CUSTYPE) = "F" Then
            cboCustType.ListIndex = 3
        Else
            cboCustType.ListIndex = 0
        End If
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
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
    If ADDOrEdit = "ADD" And LTrim(RTrim(txtLastName)) <> "" Then
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

