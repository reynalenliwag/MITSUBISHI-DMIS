VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Trans_ApplicationCorporate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Application Data Entry for Corporate"
   ClientHeight    =   16965
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "CorporateApplication.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   16965
   ScaleWidth      =   11505
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   11505
      TabIndex        =   173
      Top             =   15975
      Width           =   11505
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   1140
         Left            =   450
         ScaleHeight     =   1140
         ScaleWidth      =   10875
         TabIndex        =   174
         Top             =   30
         Width           =   10875
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
            Left            =   9930
            MouseIcon       =   "CorporateApplication.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   183
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
            Left            =   9240
            MouseIcon       =   "CorporateApplication.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   182
            ToolTipText     =   "Print this Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdDocumentCheckList 
            Caption         =   "Documents"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   8550
            MouseIcon       =   "CorporateApplication.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   197
            ToolTipText     =   "Add/Remove Require Document for Loan Application"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdUpdateStatus 
            Caption         =   "&Status"
            Height          =   795
            Left            =   7860
            MouseIcon       =   "CorporateApplication.frx":19FF
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":1B51
            Style           =   1  'Graphical
            TabIndex        =   198
            ToolTipText     =   "Update Loan Status"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   7170
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CorporateApplication.frx":2183
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":22D5
            Style           =   1  'Graphical
            TabIndex        =   189
            ToolTipText     =   "Cancel this Transaction"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdUnPost 
            Caption         =   "Unpost"
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
            Left            =   6480
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CorporateApplication.frx":260F
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":2761
            Style           =   1  'Graphical
            TabIndex        =   191
            ToolTipText     =   "Unpost this Transaction"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post"
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
            Left            =   5790
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CorporateApplication.frx":2AA6
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":2BF8
            Style           =   1  'Graphical
            TabIndex        =   190
            ToolTipText     =   "Post this Transaction"
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
            Left            =   5100
            MouseIcon       =   "CorporateApplication.frx":2F1D
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":306F
            Style           =   1  'Graphical
            TabIndex        =   181
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
            Left            =   4410
            MouseIcon       =   "CorporateApplication.frx":33CB
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":351D
            Style           =   1  'Graphical
            TabIndex        =   180
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
            Left            =   3720
            MouseIcon       =   "CorporateApplication.frx":3830
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":3982
            Style           =   1  'Graphical
            TabIndex        =   179
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
            Left            =   3030
            MouseIcon       =   "CorporateApplication.frx":3CD2
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":3E24
            Style           =   1  'Graphical
            TabIndex        =   178
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
            Left            =   2340
            MouseIcon       =   "CorporateApplication.frx":4182
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":42D4
            Style           =   1  'Graphical
            TabIndex        =   177
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
            Left            =   1650
            MouseIcon       =   "CorporateApplication.frx":45CE
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":4720
            Style           =   1  'Graphical
            TabIndex        =   176
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
            Left            =   960
            MouseIcon       =   "CorporateApplication.frx":4A78
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":4BCA
            Style           =   1  'Graphical
            TabIndex        =   175
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   7980
         ScaleHeight     =   885
         ScaleWidth      =   3210
         TabIndex        =   184
         Top             =   0
         Visible         =   0   'False
         Width           =   3210
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
            Left            =   2370
            MouseIcon       =   "CorporateApplication.frx":4F29
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":507B
            Style           =   1  'Graphical
            TabIndex        =   186
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
            Left            =   1680
            MouseIcon       =   "CorporateApplication.frx":53B9
            MousePointer    =   99  'Custom
            Picture         =   "CorporateApplication.frx":550B
            Style           =   1  'Graphical
            TabIndex        =   185
            ToolTipText     =   "Save this Record"
            Top             =   60
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "::"
         Height          =   345
         Left            =   11160
         TabIndex        =   199
         Top             =   60
         Width           =   345
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   3180
         Top             =   30
      End
      Begin VB.TextBox txtDateApplied 
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
         Left            =   8940
         TabIndex        =   6
         Tag             =   "@D"
         Top             =   53
         Width           =   2190
      End
      Begin VB.TextBox labID 
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
         Height          =   360
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtApl_No 
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   60
         Width           =   2010
      End
      Begin VB.Label labTStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6360
         TabIndex        =   192
         Top             =   60
         Width           =   1890
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8340
         TabIndex        =   5
         Top             =   90
         Width           =   495
      End
      Begin VB.Label labLStatus 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3210
         TabIndex        =   4
         Top             =   90
         Width           =   3000
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APL No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Top             =   90
         Width           =   840
      End
   End
   Begin VB.PictureBox picDocumentList 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3000
      ScaleHeight     =   4755
      ScaleWidth      =   5835
      TabIndex        =   152
      Top             =   900
      Visible         =   0   'False
      Width           =   5865
      Begin VB.CommandButton Command3 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "CorporateApplication.frx":585B
         MousePointer    =   99  'Custom
         Picture         =   "CorporateApplication.frx":59AD
         Style           =   1  'Graphical
         TabIndex        =   156
         ToolTipText     =   "Cancel"
         Top             =   3870
         Width           =   705
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Save"
         Height          =   795
         Left            =   4320
         MouseIcon       =   "CorporateApplication.frx":5CEB
         MousePointer    =   99  'Custom
         Picture         =   "CorporateApplication.frx":5E3D
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Save"
         Top             =   3870
         Width           =   705
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3525
         Left            =   60
         TabIndex        =   154
         Top             =   300
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   6218
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
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
         MouseIcon       =   "CorporateApplication.frx":618D
         NumItems        =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   153
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Update Document Check List For Individual Application:::"
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
   Begin VB.PictureBox picFindLoan 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   2280
      ScaleHeight     =   4305
      ScaleWidth      =   7005
      TabIndex        =   157
      Top             =   960
      Visible         =   0   'False
      Width           =   7035
      Begin VB.CommandButton cmdCancelSO 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6660
         TabIndex        =   158
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdCancelSelect 
         Caption         =   "Cancel"
         Height          =   645
         Index           =   0
         Left            =   6150
         Picture         =   "CorporateApplication.frx":62EF
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Cancel"
         Top             =   3615
         Width           =   765
      End
      Begin VB.TextBox txtFindAPL 
         Height          =   330
         Left            =   2220
         TabIndex        =   161
         Top             =   690
         Width           =   4695
      End
      Begin MSComctlLib.ListView lstLoan 
         Height          =   2535
         Left            =   90
         TabIndex        =   163
         Top             =   1050
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4471
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
         MouseIcon       =   "CorporateApplication.frx":662D
         NumItems        =   0
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Height          =   645
         Left            =   5400
         Picture         =   "CorporateApplication.frx":678F
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Select"
         Top             =   3615
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Search For Corporate Application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Index           =   1
         Left            =   135
         TabIndex        =   160
         Top             =   315
         Width           =   5025
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Name"
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   162
         Top             =   720
         Width           =   1425
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   159
         Top             =   0
         Width           =   7020
         _Version        =   655364
         _ExtentX        =   12382
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "::: Edit Loan Application:::"
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
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4545
      Left            =   3877
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   4515
      ScaleWidth      =   3720
      TabIndex        =   166
      Top             =   1680
      Visible         =   0   'False
      Width           =   3750
      Begin VB.ComboBox cboFinCom 
         Height          =   345
         ItemData        =   "CorporateApplication.frx":6ACB
         Left            =   210
         List            =   "CorporateApplication.frx":6ACD
         TabIndex        =   195
         Text            =   "Combo1"
         Top             =   1860
         Width           =   3270
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   210
         TabIndex        =   194
         Top             =   615
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   609
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleBackColor=   4194304
         CalendarTitleForeColor=   16777215
         Format          =   20643841
         CurrentDate     =   39378
      End
      Begin VB.CommandButton cmdCancelStatus 
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
         Index           =   0
         Left            =   2820
         MouseIcon       =   "CorporateApplication.frx":6ACF
         MousePointer    =   99  'Custom
         Picture         =   "CorporateApplication.frx":6C21
         Style           =   1  'Graphical
         TabIndex        =   187
         ToolTipText     =   "Cancel"
         Top             =   3660
         Width           =   675
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
         Left            =   3420
         TabIndex        =   167
         Top             =   0
         Width           =   285
      End
      Begin VB.TextBox txtReasonNote 
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
         Height          =   1095
         Left            =   210
         MaxLength       =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   172
         Top             =   2520
         Width           =   3270
      End
      Begin VB.ComboBox cboLoanStatus 
         Height          =   345
         ItemData        =   "CorporateApplication.frx":6F5F
         Left            =   210
         List            =   "CorporateApplication.frx":6F61
         TabIndex        =   170
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3270
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Update"
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
         Left            =   2130
         MouseIcon       =   "CorporateApplication.frx":6F63
         MousePointer    =   99  'Custom
         Picture         =   "CorporateApplication.frx":70B5
         Style           =   1  'Graphical
         TabIndex        =   188
         ToolTipText     =   "Save Changes"
         Top             =   3660
         Width           =   705
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Financing Company:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   4
         Left            =   210
         TabIndex        =   196
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   193
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   169
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label68 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   171
         Top             =   2265
         Width           =   435
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   330
         Left            =   0
         TabIndex        =   168
         Top             =   0
         Width           =   3735
         _Version        =   655364
         _ExtentX        =   6588
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   ":: Update Status ::"
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
   Begin VB.PictureBox picMiddles 
      Align           =   1  'Align Top
      Height          =   15510
      Left            =   0
      ScaleHeight     =   15450
      ScaleWidth      =   11445
      TabIndex        =   7
      Top             =   465
      Width           =   11505
      Begin VB.PictureBox picLoan 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   15420
         Left            =   0
         ScaleHeight     =   15420
         ScaleWidth      =   11145
         TabIndex        =   8
         Top             =   0
         Width           =   11145
         Begin VB.Frame fraAOR 
            Caption         =   "Loan Applied For"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3045
            Left            =   60
            TabIndex        =   39
            Top             =   3090
            Width           =   10965
            Begin VB.PictureBox Picture2 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   885
               Left            =   6330
               ScaleHeight     =   885
               ScaleWidth      =   4425
               TabIndex        =   44
               Top             =   210
               Width           =   4425
               Begin VB.TextBox txtLoan_MonthlyAmortization 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   390
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   46
                  TabStop         =   0   'False
                  Tag             =   "0.00"
                  Top             =   0
                  Width           =   2700
               End
               Begin VB.TextBox txtLoan_FinBalAmount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00400000&
                  Height          =   390
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Tag             =   "0.00"
                  Top             =   450
                  Width           =   2700
               End
               Begin VB.Label Label53 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Amortization :"
                  ForeColor       =   &H00400000&
                  Height          =   285
                  Index           =   3
                  Left            =   -510
                  TabIndex        =   45
                  Top             =   0
                  Width           =   1905
               End
               Begin VB.Label Label46 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Balance Financed : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   1
                  Left            =   60
                  TabIndex        =   47
                  Top             =   510
                  Width           =   1620
               End
            End
            Begin VB.ComboBox cboLoan_PlaceofUse 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2085
               TabIndex        =   50
               Top             =   1040
               Width           =   3825
            End
            Begin VB.ComboBox cboLoan_SAENAME 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2085
               TabIndex        =   42
               Tag             =   "@R"
               ToolTipText     =   "Sales Account Executives"
               Top             =   640
               Width           =   3825
            End
            Begin VB.ComboBox cboLoan_UnitModel 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2085
               TabIndex        =   41
               ToolTipText     =   "Unit Model"
               Top             =   240
               Width           =   3825
            End
            Begin VB.TextBox txtLoan_BankTerms 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   4650
               MaxLength       =   4
               TabIndex        =   62
               Tag             =   "0"
               Top             =   2490
               Width           =   1290
            End
            Begin VB.TextBox txtLoan_UnitCost 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2085
               TabIndex        =   56
               Tag             =   "0.00"
               Top             =   1720
               Width           =   3825
            End
            Begin VB.TextBox txtLoan_Downpayment 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2085
               TabIndex        =   58
               Tag             =   "0.00"
               Top             =   2105
               Width           =   3825
            End
            Begin VB.Frame Frame7 
               Caption         =   "Surety Information"
               Height          =   1905
               Index           =   1
               Left            =   5970
               TabIndex        =   63
               Top             =   1050
               Width           =   4815
               Begin VB.TextBox txtSuretyIncome 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1950
                  TabIndex        =   71
                  Tag             =   "0.00"
                  Top             =   1470
                  Width           =   2715
               End
               Begin VB.TextBox txtSuretyCompanyIncome 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1950
                  TabIndex        =   69
                  Tag             =   "0.00"
                  Top             =   1065
                  Width           =   2715
               End
               Begin VB.TextBox txtSuretyAddress 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1950
                  TabIndex        =   67
                  Top             =   645
                  Width           =   2715
               End
               Begin VB.TextBox txtSuretyName 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1950
                  TabIndex        =   65
                  Top             =   240
                  Width           =   2715
               End
               Begin VB.Label Label53 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Surety Income : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   2
                  Left            =   630
                  TabIndex        =   70
                  Top             =   1560
                  Width           =   1305
               End
               Begin VB.Label Label53 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Company Income : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   1
                  Left            =   345
                  TabIndex        =   68
                  Top             =   1080
                  Width           =   1590
               End
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Surety Address : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   1
                  Left            =   555
                  TabIndex        =   66
                  Top             =   645
                  Width           =   1380
               End
               Begin VB.Label Label43 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Surety Name : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   1
                  Left            =   735
                  TabIndex        =   64
                  Top             =   270
                  Width           =   1200
               End
            End
            Begin VB.TextBox txtLoan_AORPercentage 
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
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2085
               MaxLength       =   4
               TabIndex        =   60
               Tag             =   "0.00"
               Top             =   2490
               Width           =   1440
            End
            Begin VB.OptionButton optPurposePrivate 
               Caption         =   "Private"
               Height          =   225
               Left            =   2085
               TabIndex        =   52
               Top             =   1440
               Width           =   1035
            End
            Begin VB.OptionButton optPurposeBusiness 
               Caption         =   "Business"
               Height          =   225
               Left            =   3352
               TabIndex        =   53
               Top             =   1440
               Width           =   1155
            End
            Begin VB.OptionButton optPurposePublic 
               Caption         =   "Public"
               Height          =   225
               Left            =   4740
               TabIndex        =   54
               Top             =   1440
               Width           =   1035
            End
            Begin VB.Label Label56 
               BackStyle       =   0  'Transparent
               Caption         =   "Place of Use : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   870
               TabIndex        =   49
               Top             =   1080
               Width           =   1185
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purpose : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   1215
               TabIndex        =   51
               Top             =   1410
               Width           =   840
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Unit/Model : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   1050
               TabIndex        =   40
               Top             =   300
               Width           =   1005
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   630
               TabIndex        =   43
               Top             =   780
               Width           =   1425
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Downpayment:"
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   795
               TabIndex        =   57
               Top             =   2100
               Width           =   1230
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Net Unit Price: "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   840
               TabIndex        =   55
               Top             =   1755
               Width           =   1215
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AOR : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   1530
               TabIndex        =   59
               Top             =   2520
               Width           =   510
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank Term : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   3585
               TabIndex        =   61
               Top             =   2550
               Width           =   1035
            End
         End
         Begin VB.Frame fraCompanyProfile 
            Caption         =   "Company Profile"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3105
            Left            =   90
            TabIndex        =   9
            Top             =   0
            Width           =   10965
            Begin VB.TextBox txtComp_MajorProduct 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2100
               TabIndex        =   23
               Top             =   1851
               Width           =   3750
            End
            Begin VB.OptionButton optCorporationTypePartnership 
               Caption         =   "Partnership"
               Height          =   285
               Left            =   3195
               TabIndex        =   18
               Top             =   1124
               Width           =   1275
            End
            Begin VB.OptionButton optCorporationTypeCorporation 
               Caption         =   "Corporation"
               Height          =   285
               Left            =   4590
               TabIndex        =   19
               Top             =   1124
               Width           =   1305
            End
            Begin VB.OptionButton optCorporationTypeSingle 
               Caption         =   "Single"
               Height          =   285
               Left            =   2100
               TabIndex        =   17
               Top             =   1124
               Width           =   825
            End
            Begin VB.TextBox txtComp_PaidUpCapital 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2100
               TabIndex        =   29
               Tag             =   "0.00"
               Top             =   2655
               Width           =   3765
            End
            Begin VB.TextBox txtComp_TelNo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   7800
               TabIndex        =   13
               Top             =   330
               Width           =   2940
            End
            Begin VB.TextBox txtComp_OfficeAdd 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2100
               MaxLength       =   100
               TabIndex        =   15
               Top             =   723
               Width           =   8670
            End
            Begin VB.TextBox txtComp_Busname 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2100
               TabIndex        =   11
               Tag             =   "@R"
               Top             =   322
               Width           =   4695
            End
            Begin VB.ComboBox cboComp_NatureOfBusiness 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   345
               Left            =   2100
               TabIndex        =   21
               Top             =   1450
               Width           =   3750
            End
            Begin VB.Frame Frame1 
               Caption         =   "Tin Info"
               Height          =   1860
               Left            =   6450
               TabIndex        =   30
               Top             =   1140
               Width           =   4425
               Begin VB.ComboBox cboComp_IssuedAt 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   345
                  Left            =   1590
                  TabIndex        =   38
                  Top             =   1440
                  Width           =   2715
               End
               Begin VB.TextBox txtComp_TinNO 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1590
                  TabIndex        =   32
                  Top             =   240
                  Width           =   2715
               End
               Begin VB.TextBox txtComp_CCINo 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1590
                  TabIndex        =   34
                  Top             =   640
                  Width           =   2715
               End
               Begin VB.TextBox txtComp_IssuedOn 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00400000&
                  Height          =   330
                  Left            =   1590
                  TabIndex        =   36
                  Top             =   1040
                  Width           =   2715
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "TIN : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   4
                  Left            =   1140
                  TabIndex        =   31
                  Top             =   300
                  Width           =   420
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "C && CI No : "
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   5
                  Left            =   615
                  TabIndex        =   33
                  Top             =   675
                  Width           =   945
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Issued On:"
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   6
                  Left            =   630
                  TabIndex        =   35
                  Top             =   1080
                  Width           =   900
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Issued At:"
                  ForeColor       =   &H00400000&
                  Height          =   225
                  Index           =   7
                  Left            =   750
                  TabIndex        =   37
                  Top             =   1500
                  Width           =   810
               End
            End
            Begin VB.TextBox txtComp_DateEstablised 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2100
               TabIndex        =   25
               Top             =   2252
               Width           =   1080
            End
            Begin VB.TextBox txtComp_YearsInOpt 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   5100
               TabIndex        =   27
               Top             =   2252
               Width           =   750
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Paid-Up Capital : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   3
               Left            =   630
               TabIndex        =   28
               Top             =   2730
               Width           =   1440
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Major Product(s) : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   2
               Left            =   585
               TabIndex        =   22
               Top             =   1950
               Width           =   1485
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nature of Business : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   2
               Left            =   345
               TabIndex        =   20
               Top             =   1545
               Width           =   1725
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type of Organization:"
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   360
               TabIndex        =   16
               Top             =   1170
               Width           =   1710
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No(s): "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   6900
               TabIndex        =   12
               Top             =   375
               Width           =   900
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Office Address : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   735
               TabIndex        =   14
               Top             =   750
               Width           =   1335
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Business Name : "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   0
               Left            =   585
               TabIndex        =   10
               Top             =   390
               Width           =   1485
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Establistment Date:"
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   1
               Left            =   450
               TabIndex        =   24
               Top             =   2320
               Width           =   1620
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Years In Operations: "
               ForeColor       =   &H00400000&
               Height          =   225
               Index           =   0
               Left            =   3255
               TabIndex        =   26
               Top             =   2320
               Width           =   1725
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "References"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Index           =   3
            Left            =   60
            TabIndex        =   110
            Top             =   10920
            Width           =   10965
            Begin VB.TextBox txtRef_SupTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8250
               TabIndex        =   121
               Top             =   795
               Width           =   2565
            End
            Begin VB.TextBox txtRef_TradeTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8250
               TabIndex        =   117
               Top             =   420
               Width           =   2565
            End
            Begin VB.TextBox txtRef_LoanTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8250
               TabIndex        =   125
               Top             =   1215
               Width           =   2565
            End
            Begin VB.TextBox txtRef_CreditTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8250
               TabIndex        =   129
               Top             =   1620
               Width           =   2565
            End
            Begin VB.TextBox txtRef_SupAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   120
               Top             =   795
               Width           =   4155
            End
            Begin VB.TextBox txtRef_TradeAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   116
               Top             =   420
               Width           =   4155
            End
            Begin VB.TextBox txtRef_LoanAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   124
               Top             =   1215
               Width           =   4155
            End
            Begin VB.TextBox txtRef_CreditAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   128
               Top             =   1620
               Width           =   4155
            End
            Begin VB.TextBox txtRef_SupName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   960
               TabIndex        =   119
               Top             =   795
               Width           =   2985
            End
            Begin VB.TextBox txtRef_TradeName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   960
               TabIndex        =   115
               Top             =   420
               Width           =   2985
            End
            Begin VB.TextBox txtRef_LoanName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   960
               TabIndex        =   123
               Top             =   1215
               Width           =   2985
            End
            Begin VB.TextBox txtRef_CreditName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   960
               TabIndex        =   127
               Top             =   1620
               Width           =   2985
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   315
               Index           =   3
               Left            =   8250
               TabIndex        =   113
               Top             =   180
               Width           =   2055
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   315
               Index           =   3
               Left            =   3990
               TabIndex        =   112
               Top             =   180
               Width           =   3885
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   3
               Left            =   2235
               TabIndex        =   111
               Top             =   180
               Width           =   525
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trade:"
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   405
               TabIndex        =   114
               Top             =   510
               Width           =   525
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Loan:"
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   465
               TabIndex        =   122
               Top             =   1270
               Width           =   465
            End
            Begin VB.Label Label65 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Suppliers:"
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   2
               Left            =   90
               TabIndex        =   118
               Top             =   890
               Width           =   840
            End
            Begin VB.Label Label64 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Credit :"
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   2
               Left            =   330
               TabIndex        =   126
               Top             =   1650
               Width           =   585
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Bank Account(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2085
            Index           =   1
            Left            =   60
            TabIndex        =   130
            Top             =   13020
            Width           =   10965
            Begin VB.TextBox txtBA_Bank4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   147
               Top             =   1620
               Width           =   3135
            End
            Begin VB.TextBox txtBA_Bank3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   143
               Top             =   1245
               Width           =   3135
            End
            Begin VB.TextBox txtBA_Bank1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   135
               Top             =   480
               Width           =   3135
            End
            Begin VB.TextBox txtBA_Bank2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   139
               Top             =   855
               Width           =   3135
            End
            Begin VB.TextBox txtBA_AcctNo4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   5700
               MaxLength       =   20
               TabIndex        =   149
               Top             =   1620
               Width           =   2625
            End
            Begin VB.TextBox txtBA_AcctNo3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   5700
               MaxLength       =   20
               TabIndex        =   145
               Top             =   1245
               Width           =   2625
            End
            Begin VB.TextBox txtBA_AcctNo1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   5700
               MaxLength       =   20
               TabIndex        =   137
               Top             =   480
               Width           =   2625
            End
            Begin VB.TextBox txtBA_AcctNo2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   5700
               MaxLength       =   20
               TabIndex        =   141
               Top             =   855
               Width           =   2625
            End
            Begin VB.TextBox txtBA_Balance4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8340
               TabIndex        =   150
               Tag             =   "0.00"
               Top             =   1620
               Width           =   2505
            End
            Begin VB.TextBox txtBA_Balance3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8340
               TabIndex        =   146
               Tag             =   "0.00"
               Top             =   1245
               Width           =   2505
            End
            Begin VB.TextBox txtBA_Balance1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8340
               TabIndex        =   138
               Tag             =   "0.00"
               Top             =   480
               Width           =   2505
            End
            Begin VB.TextBox txtBA_Balance2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8340
               TabIndex        =   142
               Tag             =   "0.00"
               Top             =   855
               Width           =   2505
            End
            Begin VB.ComboBox cboBA_TOA1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3405
               TabIndex        =   136
               Top             =   480
               Width           =   2265
            End
            Begin VB.ComboBox cboBA_TOA4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3405
               TabIndex        =   148
               Top             =   1620
               Width           =   2265
            End
            Begin VB.ComboBox cboBA_TOA3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3405
               TabIndex        =   144
               Top             =   1245
               Width           =   2265
            End
            Begin VB.ComboBox cboBA_TOA2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3405
               TabIndex        =   140
               Top             =   855
               Width           =   2265
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bank/Branch"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   1245
               TabIndex        =   131
               Top             =   210
               Width           =   1125
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type of Account"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   3840
               TabIndex        =   132
               Top             =   210
               Width           =   1395
            End
            Begin VB.Label Label67 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Account Number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   6285
               TabIndex        =   133
               Top             =   210
               Width           =   1455
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   9105
               TabIndex        =   134
               Top             =   210
               Width           =   705
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Officers/Directors"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Index           =   2
            Left            =   60
            TabIndex        =   88
            Top             =   8100
            Width           =   10965
            Begin VB.TextBox txtODContactPerson 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   240
               TabIndex        =   107
               Top             =   2280
               Width           =   3705
            End
            Begin VB.TextBox txtODName2 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   240
               TabIndex        =   95
               Top             =   840
               Width           =   3705
            End
            Begin VB.TextBox txtODName1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   240
               TabIndex        =   92
               Top             =   435
               Width           =   3705
            End
            Begin VB.TextBox txtODName3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   240
               TabIndex        =   98
               Top             =   1245
               Width           =   3705
            End
            Begin VB.TextBox txtODName4 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   240
               TabIndex        =   101
               Top             =   1635
               Width           =   3705
            End
            Begin VB.TextBox txtODTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8220
               TabIndex        =   109
               Top             =   2280
               Width           =   2565
            End
            Begin VB.TextBox txtODAddress2 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   96
               Top             =   840
               Width           =   4155
            End
            Begin VB.TextBox txtODAddress1 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   93
               Top             =   435
               Width           =   4155
            End
            Begin VB.TextBox txtODAddress3 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   4020
               MaxLength       =   45
               TabIndex        =   99
               Top             =   1245
               Width           =   4155
            End
            Begin VB.TextBox txtODAddress4 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   102
               Top             =   1635
               Width           =   4155
            End
            Begin VB.ComboBox cboODPosition1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8220
               TabIndex        =   94
               Top             =   435
               Width           =   2565
            End
            Begin VB.ComboBox cboODPosition2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8220
               TabIndex        =   97
               Top             =   840
               Width           =   2565
            End
            Begin VB.ComboBox cboODPosition3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8220
               TabIndex        =   100
               Top             =   1245
               Width           =   2565
            End
            Begin VB.ComboBox cboODPosition4 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   8220
               TabIndex        =   103
               Top             =   1635
               Width           =   2565
            End
            Begin VB.ComboBox cboODDesignation 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3990
               TabIndex        =   108
               Top             =   2280
               Width           =   4155
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No. : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   10
               Left            =   8175
               TabIndex        =   106
               Top             =   2040
               Width           =   750
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Person : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   8
               Left            =   240
               TabIndex        =   104
               Top             =   2040
               Width           =   1455
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   2
               Left            =   9105
               TabIndex        =   91
               Top             =   210
               Width           =   705
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   2
               Left            =   1755
               TabIndex        =   89
               Top             =   210
               Width           =   525
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position/Designation : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   9
               Left            =   3930
               TabIndex        =   105
               Top             =   2040
               Width           =   1875
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   2
               Left            =   5490
               TabIndex        =   90
               Top             =   210
               Width           =   765
            End
         End
         Begin VB.Frame fraStockholder 
            Caption         =   "Major Stockholders (If Corporation)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1995
            Left            =   60
            TabIndex        =   72
            Top             =   6105
            Width           =   10965
            Begin VB.TextBox txtSH_Name4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   85
               Top             =   1530
               Width           =   3705
            End
            Begin VB.TextBox txtSH_Name3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   82
               Top             =   1185
               Width           =   3705
            End
            Begin VB.TextBox txtSH_Name1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   270
               TabIndex        =   76
               Top             =   450
               Width           =   3705
            End
            Begin VB.TextBox txtSH_Name2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   240
               TabIndex        =   79
               Top             =   825
               Width           =   3705
            End
            Begin VB.TextBox txtSH_Adress4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   86
               Top             =   1530
               Width           =   4155
            End
            Begin VB.TextBox txtSH_Adress3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   83
               Top             =   1185
               Width           =   4155
            End
            Begin VB.TextBox txtSH_Adress1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   77
               Top             =   450
               Width           =   4155
            End
            Begin VB.TextBox txtSH_Adress2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   3990
               MaxLength       =   45
               TabIndex        =   80
               Top             =   825
               Width           =   4155
            End
            Begin VB.TextBox txtSH_Amount4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8190
               TabIndex        =   87
               Tag             =   "0.00"
               Top             =   1530
               Width           =   2475
            End
            Begin VB.TextBox txtSH_Amount3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8190
               TabIndex        =   84
               Tag             =   "0.00"
               Top             =   1185
               Width           =   2475
            End
            Begin VB.TextBox txtSH_Amount1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8190
               TabIndex        =   78
               Tag             =   "0.00"
               Top             =   450
               Width           =   2475
            End
            Begin VB.TextBox txtSH_Amount2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   8190
               TabIndex        =   81
               Tag             =   "0.00"
               Top             =   825
               Width           =   2475
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amount of Stocks Owned"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   8325
               TabIndex        =   75
               Top             =   210
               Width           =   2205
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   5340
               TabIndex        =   74
               Top             =   210
               Width           =   765
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   225
               Index           =   1
               Left            =   1755
               TabIndex        =   73
               Top             =   210
               Width           =   525
            End
         End
      End
      Begin VB.VScrollBar ScrollBar1 
         Height          =   2895
         LargeChange     =   500
         Left            =   11160
         Max             =   11160
         SmallChange     =   250
         TabIndex        =   151
         Top             =   30
         Value           =   10
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_ApplicationCorporate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim AddorEdit                                                         As String
Dim rsLoan                                                            As ADODB.Recordset
Dim PROSPECTID                                                        As Long
Dim CUSCDE                                                            As String
Dim ProfileType                                                       As String
Dim AddingLoan                                                        As Boolean
Dim WithEvents FormSearch                                             As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1
Private LoanID                                                        As Long

Function SetStatus(xx As String) As Integer

    If xx = "O" Then
        SetStatus = 0
    ElseIf xx = "A" Then
        SetStatus = 1
    ElseIf xx = "C" Then
        SetStatus = 2
    ElseIf xx = "D" Then
        SetStatus = 3
    ElseIf xx = "P" Then
        SetStatus = 4
    End If
End Function

Function GetStatus(xx As String) As String
    If xx = "O" Then
        GetStatus = "On Process"
    ElseIf xx = "A" Then
        GetStatus = "Approved"
    ElseIf xx = "C" Then
        GetStatus = "Cancelled"
    ElseIf xx = "D" Then
        GetStatus = "Disapproved"
    ElseIf xx = "P" Then
        GetStatus = "Pending"
    End If
End Function

Private Function AORVALUE(Principal, AOR, TERM) As Double
    'UDPATING COCODE    :AXP-526200721:41
    If AOR <= 0 Then: AORVALUE = 0: Exit Function
    If Principal <= 0 Then: AORVALUE = 0: Exit Function
    If TERM <= 0 Then: AORVALUE = 0: Exit Function
    Dim Interest                                                      As Double
    Interest = NumericVal(AOR)
    Interest = AOR / 1200
    AORVALUE = FormatNumber((Principal * Interest / (1 - ((1 / (1 + Interest) ^ TERM)))), 2)
End Function

Public Function ShowLoanApp(IDX As Long) As Boolean
    If IDX > 0 Then
        AddorEdit = "EDIT"
        LoanID = IDX
    Else
        AddorEdit = "ADD"
    End If
End Function

Public Function AddFromProspects(IDXPROSPECTID As Long) As Boolean
    AddingLoan = True
    If IDXPROSPECTID <> 0 Then
        AddingLoan = True
    Else
        Unload Me
    End If
    InitMemVars

    Dim oCusRs                                                        As ADODB.Recordset

    Set oCusRs = gconDMIS.Execute("SELECT  * FROM CRIS_PROSPECTS WHERE PROSPECTID=" & IDXPROSPECTID)
    If oCusRs.EOF = True Or oCusRs.BOF = True Then
        MsgBox " Error Fetching Record"
        Exit Function
    End If

    If IsDate(oCusRs!LogApplication) = True Then

        'AddEditLoan

        'Exit Function

    End If


    Dim Telephone                                                     As String

    AddorEdit = "ADD"
    picLoan.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    txtDateApplied.Text = FormatDateTime(LOGDATE, vbShortDate)
    InitMemVars
    txtApl_No = GenerateCode("SMIS_LOANCORP", "APLCODE", "0000000000")


    CUSCDE = Null2String(oCusRs!CUSCDE)
    PROSPECTID = Null2String(oCusRs!PROSPECTID)
    ProfileType = Null2String(oCusRs!ProspectType)
    '
    '    If IsNull(oCusRs!LogAppointment) = False Then
    '        MsgBox " This Prospect Has Appointment Do You Want Preview Appointment to Load Application "
    '    End If

    txtDateApplied = FormatDateTime(LOGDATE, vbShortDate)
    txtComp_Busname = Null2String(oCusRs.Fields("Acctname"))
    cboLoan_UnitModel = Null2String(oCusRs.Fields("VARIANT"))
    cboLoan_SAENAME = Null2String(oCusRs.Fields("SAE"))
    txtODContactPerson = Null2String(oCusRs!ContactPerson)
    txtComp_OfficeAdd = Null2String(oCusRs!Address)
    Telephone = Null2String(oCusRs!Telephone) & "\" & Null2String(oCusRs!Mobile)

    If Left(Telephone, 1) = "\" Or Right(Telephone, 1) = "\" Then
        txtComp_TelNo = Replace(Telephone, "\", "")
    Else
        txtComp_TelNo = Telephone
    End If
End Function

Sub InitCbo()
    Dim TEMPRS                                                        As ADODB.Recordset
    Dim odPos                                                         As String
    Dim SQL                                                           As String

    Call FillCombo("Select descript from All_Model order by descript asc", -1, 0, cboLoan_UnitModel)
    Call FillCombo("Select NAME from SMIS_vw_Srep order by NAME asc", -1, 0, cboLoan_SAENAME)

    Call FillCombo("Select Distinct PlaceofIssue from SMIS_LoanCorp Order by 1 asc ", -1, 0, cboComp_IssuedAt)
    Call FillCombo("Select Distinct PlaceofUse from SMIS_LoanCorp Order by 1 asc ", -1, 0, cboLoan_PlaceofUse)
    Call FillCombo("Select distinct NatureOfBusiness from SMIS_LoanCorp  Order by 1 asc ", -1, 0, cboComp_NatureOfBusiness)
    Call FillCombo("Select distinct Company from SMIS_FINCOM Order by 1 asc ", -1, 0, cboFinCom)


    With cboLoanStatus
        .AddItem "On Process"
        .AddItem "Approved"
        .AddItem "Disapproved"
        .AddItem "Pending"
        .ListIndex = 0
    End With



    SQL = SQL & "SELECT DISTINCT ODPosition1 Xpos From SMIS_LoanCorp Where LEN(ODPOSITION1)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT ODPosition2  Xpos From SMIS_LoanCorp Where LEN(ODPOSITION2)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT ODPosition3  Xpos From SMIS_LoanCorp Where LEN(ODPOSITION3)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT ODPosition4  Xpos From SMIS_LoanCorp Where LEN(ODPOSITION4)>0"
    SQL = SQL & "ORDER BY 1"
    Set TEMPRS = gconDMIS.Execute(SQL)

    cboODPosition1.Clear
    cboODPosition1.Clear
    cboODPosition2.Clear
    cboODPosition3.Clear
    cboODPosition4.Clear
    cboODDesignation.Clear
    cboBA_TOA1.Clear
    cboBA_TOA2.Clear
    cboBA_TOA3.Clear
    cboBA_TOA4.Clear

    While Not TEMPRS.EOF
        odPos = TEMPRS!XPos
        cboODPosition1.AddItem odPos
        cboODPosition2.AddItem odPos
        cboODPosition3.AddItem odPos
        cboODPosition4.AddItem odPos
        cboODDesignation.AddItem odPos
        TEMPRS.MoveNext
    Wend
    Set TEMPRS = Nothing

    SQL = "SELECT DISTINCT BA_TOA1 Xpos From SMIS_LoanCorp Where LEN(BA_TOA1)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT BA_TOA2  Xpos From SMIS_LoanCorp Where LEN(BA_TOA2)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT BA_TOA3  Xpos From SMIS_LoanCorp Where LEN(BA_TOA3)>0"
    SQL = SQL & "UNION "
    SQL = SQL & "SELECT DISTINCT BA_TOA4  Xpos From SMIS_LoanCorp Where LEN(BA_TOA4)>0"
    SQL = SQL & "ORDER BY 1 "
    Set TEMPRS = gconDMIS.Execute(SQL)
    While Not TEMPRS.EOF
        odPos = TEMPRS!XPos
        cboBA_TOA1.AddItem odPos
        cboBA_TOA2.AddItem odPos
        cboBA_TOA3.AddItem odPos
        cboBA_TOA4.AddItem odPos
        TEMPRS.MoveNext
    Wend
    Set TEMPRS = Nothing




End Sub

Sub InitMemVars()
    '''''
    Dim cntrl                                                         As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            cntrl.Text = cntrl.Tag
        End If
    Next
    optCorporationTypeSingle.Value = True
    optPurposePrivate.Value = True
End Sub

Sub StoreMemVars()
    If Not (rsLoan.EOF Or rsLoan.BOF) Then
        With rsLoan

            labid = Null2String(.Fields("ID"))
            txtApl_No = Null2String(.Fields("APL_NO"))

            CUSCDE = Null2String(.Fields("APLCODE"))
            PROSPECTID = NumericVal(.Fields("ProspectID"))
            txtDateApplied = Null2String(.Fields("DateApplied"))
            txtComp_Busname = Null2String(.Fields("Busname"))
            txtComp_TelNo = Null2String(.Fields("TelNo"))
            txtComp_OfficeAdd = Null2String(.Fields("OfficeAdd"))
            txtComp_TelNo = Null2String(.Fields("TelNo"))

            If Null2String(.Fields("CorporationType")) = "Corporation" Then
                optCorporationTypeCorporation.Value = True
            ElseIf Null2String(.Fields("CorporationType")) = "Partnership" Then
                optCorporationTypePartnership.Value = True
            ElseIf Null2String(.Fields("CorporationType")) = "Single" Then
                optCorporationTypeSingle.Value = True
            End If

            txtComp_DateEstablised = Null2String(.Fields("DateEstablised"))
            txtComp_YearsInOpt = Null2String(.Fields("YearsInOpt"))
            cboComp_NatureOfBusiness = Null2String(.Fields("NatureOfBusiness"))
            txtComp_MajorProduct = Null2String(.Fields("MajorProduct"))
            txtComp_PaidUpCapital = FormatNumber(NumericVal(.Fields("PaidUpCapital")))
            txtComp_TinNO = Null2String(.Fields("TinNo"))
            txtComp_CCINo = Null2String(.Fields("CCINo"))
            txtComp_IssuedOn = Null2String(.Fields("DateofIssue"))
            cboComp_IssuedAt = Null2String(.Fields("PlaceOfIssue"))
            cboComp_IssuedAt = Null2String(.Fields("PlaceOfIssue"))

            cboLoan_UnitModel = Null2String(.Fields("UnitModel"))
            cboLoan_SAENAME = Null2String(.Fields("SAENAME"))
            cboFinCom = Null2String(rsLoan!FINCOM)
            cboLoan_PlaceofUse = Null2String(.Fields("PlaceofUse"))
            txtLoan_UnitCost = FormatNumber(NumericVal(.Fields("NetCostPrice")))
            txtLoan_Downpayment = FormatNumber(NumericVal(.Fields("DownPayment")))
            txtLoan_FinBalAmount = FormatNumber(NumericVal(.Fields("BalanceFianced")))
            txtLoan_AORPercentage = FormatNumber(NumericVal(.Fields("AOR")))
            txtLoan_BankTerms = NumericVal(.Fields("Terms"))
            txtLoan_MonthlyAmortization = FormatNumber(NumericVal(.Fields("MonthlyAmortization")))

            If Null2String(.Fields("Purpose")) = "Business" Then
                optPurposeBusiness.Value = True
            ElseIf Null2String(.Fields("Purpose")) = "Private" Then
                optPurposePrivate.Value = True
            ElseIf Null2String(.Fields("Purpose")) = "Public" Then
                optPurposePublic.Value = True
            End If

            txtSH_Name1 = Null2String(.Fields("StockHolderName1"))
            txtSH_Adress1 = Null2String(.Fields("StockHolderAdd1"))
            txtSH_Amount1 = FormatNumber(NumericVal(.Fields("StockHolderAmount1")))

            txtSH_Name2 = Null2String(.Fields("StockHolderName2"))
            txtSH_Adress2 = Null2String(.Fields("StockHolderAdd2"))
            txtSH_Amount2 = FormatNumber(NumericVal(.Fields("StockHolderAmount2")))

            txtSH_Name3 = Null2String(.Fields("StockHolderName3"))
            txtSH_Adress3 = Null2String(.Fields("StockHolderAdd3"))
            txtSH_Amount3 = FormatNumber(NumericVal(.Fields("StockHolderAmount3")))

            txtSH_Name4 = Null2String(.Fields("StockHolderName4"))
            txtSH_Adress4 = Null2String(.Fields("StockHolderAdd4"))
            txtSH_Amount4 = FormatNumber(NumericVal(.Fields("StockHolderAmount4")))

            txtODName1 = Null2String(.Fields("ODName1"))
            txtODAddress1 = Null2String(.Fields("ODAddress1"))
            cboODPosition1 = Null2String(.Fields("ODPosition1"))

            txtODName2 = Null2String(.Fields("ODName2"))
            txtODAddress2 = Null2String(.Fields("ODAddress2"))
            cboODPosition2 = Null2String(.Fields("ODPosition2"))

            txtODName3 = Null2String(.Fields("ODName3"))
            txtODAddress3 = Null2String(.Fields("ODAddress3"))
            cboODPosition3 = Null2String(.Fields("ODPosition3"))

            txtODName4 = Null2String(.Fields("ODName4"))
            txtODAddress4 = Null2String(.Fields("ODAddress4"))
            cboODPosition4 = Null2String(.Fields("ODPosition4"))

            txtODContactPerson = Null2String(.Fields("ODContactPerson"))
            cboODDesignation = Null2String(.Fields("ODDesignation"))
            txtODTelNo = Null2String(.Fields("ODTelNo"))

            txtSuretyName = Null2String(.Fields("SuretyName"))
            txtSuretyAddress = Null2String(.Fields("SuretyAddress"))
            txtSuretyCompanyIncome = FormatNumber(NumericVal(.Fields("SuretyCompanyIncome")))
            txtSuretyIncome = FormatNumber(NumericVal(.Fields("SuretyIncome")))

            txtRef_TradeName = Null2String(.Fields("Ref_TradeName"))
            txtRef_TradeAdd = Null2String(.Fields("Ref_TradeAdd"))
            txtRef_TradeTelNo = Null2String(.Fields("Ref_TradeTelNo"))

            txtRef_SupName = Null2String(.Fields("Ref_SupName"))
            txtRef_SupAdd = Null2String(.Fields("Ref_SupAdd"))
            txtRef_SupTelNo = Null2String(.Fields("Ref_SupTelNo"))

            txtRef_LoanName = Null2String(.Fields("Ref_LoanName"))
            txtRef_LoanAdd = Null2String(.Fields("Ref_LoanAdd"))
            txtRef_LoanTelNo = Null2String(.Fields("Ref_LoanTelNo"))

            txtRef_CreditName = Null2String(.Fields("Ref_CreditName"))
            txtRef_CreditAdd = Null2String(.Fields("Ref_CreditAdd"))
            txtRef_CreditTelNo = Null2String(.Fields("Ref_CreditTelNo"))

            txtBA_Bank1 = Null2String(.Fields("BA_Bank1"))
            cboBA_TOA1 = Null2String(.Fields("BA_TOA1"))
            txtBA_AcctNo1 = Null2String(.Fields("BA_AcctNo1"))
            txtBA_Balance1 = FormatNumber(NumericVal(.Fields("BA_Balance1")))

            txtBA_Bank2 = Null2String(.Fields("BA_Bank2"))
            cboBA_TOA2 = Null2String(.Fields("BA_TOA2"))
            txtBA_AcctNo2 = Null2String(.Fields("BA_AcctNo2"))
            txtBA_Balance2 = FormatNumber(NumericVal(.Fields("BA_Balance2")))

            txtBA_Bank3 = Null2String(.Fields("BA_Bank3"))
            cboBA_TOA3 = Null2String(.Fields("BA_TOA3"))
            txtBA_AcctNo3 = Null2String(.Fields("BA_AcctNo3"))
            txtBA_Balance3 = FormatNumber(NumericVal(.Fields("BA_Balance3")))


            txtBA_Bank4 = Null2String(.Fields("BA_Bank4"))
            cboBA_TOA4 = Null2String(.Fields("BA_TOA4"))
            txtBA_AcctNo4 = Null2String(.Fields("BA_AcctNo4"))
            txtBA_Balance4 = FormatNumber(NumericVal(.Fields("BA_Balance4")))

        End With

        Dim TStatus, lStatus                                          As String
        TStatus = Null2String(rsLoan!STATUS)
        lStatus = Null2String(rsLoan!lStatus)
        labTStatus = GetStatus(Null2String(rsLoan!lStatus))

        If IsDate(rsLoan!Lastupdated) = True Then
            labLStatus = GetStatus(Null2String(rsLoan!lStatus)) & "-" & FormatDateTime(rsLoan!Lastupdated, vbShortDate)
        Else
            labLStatus = GetStatus(Null2String(rsLoan!lStatus))
        End If



        If Null2String(rsLoan!IsProcessed) = True Then
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPRINT.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = False
            cmdUpdateStatus.Enabled = False
        Else
            If TStatus = "P" Then
                labTStatus.Visible = True
                labTStatus.Caption = "POSTED "
                cmdEdit.Enabled = False
                cmdPost.Enabled = False
                cmdPRINT.Enabled = True
                cmdUnPost.Enabled = True
                cmdUpdateStatus.Enabled = True
            ElseIf TStatus = "C" Then
                labTStatus.Caption = "CANCELLED "
                cmdEdit.Enabled = False
                cmdPost.Enabled = False
                cmdUnPost.Enabled = False
                cmdPRINT.Enabled = False
                cmdCancelCO.Enabled = False
                cmdUnPost.Enabled = False
                cmdUpdateStatus.Enabled = False
            Else
                labTStatus.Visible = False
                labTStatus.Caption = ""
                cmdEdit.Enabled = True
                cmdPost.Enabled = True
                cmdPRINT.Enabled = True
                cmdCancelCO.Enabled = True
                cmdUnPost.Enabled = False
                cmdUpdateStatus.Enabled = False
            End If


        End If

    Else

        If AddingLoan = False Then
            ShowNoRecord
            Select Case MsgBox("There are no Loan Application(s)." & vbCrLf & "  Do You Want To Add New Record", vbYesNo Or vbQuestion Or vbDefaultButton1, App.title)
                Case vbYes
                    cmdAdd.Value = True
                Case vbNo
                    Unload Me
            End Select
        End If
    End If

End Sub

Sub UpdateAmountDetails()
    If AddorEdit = "" Then Exit Sub
    Dim A, b
    A = NumericVal(txtLoan_UnitCost)
    b = NumericVal(txtLoan_Downpayment)
    txtLoan_FinBalAmount = FormatNumber((A - b), 2)
End Sub

Private Sub cboBA_TOA1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboBA_TOA2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboBA_TOA3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboBA_TOA4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboComp_IssuedAt_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub cboComp_NatureOfBusiness_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboLoan_PlaceofUse_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboODDesignation_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboODPosition1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboODPosition2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboODPosition3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboODPosition4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Set FormSearch = New frmSMIS_Mis_SearchMaster
    ' If LOGSAE <> "" Then
    'Call FormSearch.SearchForProspects("((isdate(logso)=0) AND PROSPECTTYPE <>'P' AND isdate(logapplication) =0 and status<>'C') AND ProspectType='P' AND USERCODE='" & LOGSAE & "'")
    'Else
    'Call FormSearch.SearchForProspects(" (isdate(logapplication) =0 and status<>'C') AND ProspectType='P' ")
    'End If

    Call FormSearch.SearchForProspects("(isdate(logso)=0) AND PROSPECTTYPE <>'P'")
    FormSearch.Show 1
    txtDateApplied.Enabled = True





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    picLoan.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Do You want to Cancel this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = ("UPDATE SMIS_LOANCORP SET STATUS='C', LSTATUS='C' WHERE ID=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & txtApl_No, "", ""
    '*********RESET THE SQL_STATEMENT VARIABLE
    SQL_STATEMENT = ""
    '*******************************
    SQL_STATEMENT = ("UPDATE CRIS_PROSPECTS SET LOGAPPLICATION=NULL,APPNO=Null, LOGAPPLICATIONTYPE=NULL WHERE APPNO=" & N2Str2Null(rsLoan!Apl_no) & " AND PROSPECTID=" & PROSPECTID)
    NEW_LogAudit "EE", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & N2Str2Null(txtApl_No), "", ""

    rsRefresh
    rsLoan.Find ("ID=" & labid)
    StoreMemVars
    LogAudit "C", "CORORATE LOAN APPLICATION", txtApl_No & " " & txtComp_Busname
    MessagePop RecSaveOk, "Cancelled", "Record Sucessfully Canncelled", 1000, 2




    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancelSO_Click(Index As Integer)
    ShowHidePictureBox2 picFindLoan, False
End Sub

Private Sub cmdCancelStatus_Click(Index As Integer)
    ShowHidePictureBox2 picStatus, False
End Sub

Private Sub cmdDelete_Click()
    If rsLoan.BOF And rsLoan.EOF Then
        MsgBox "Nothing to delete"
    Else
        If ShowConfirmDelete = True Then
            gconDMIS.Execute ("DELETE FROM SMIS_LOANCORP WHERE ID=" & labid)
            gconDMIS.Execute ("DELETE FROM SMIS_LoanDocument Where AplType='C' And APLCODE=" & N2Str2Null(txtApl_No))
            gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET LOGAPPLICATION=NULL, LOGAPPLICATIONTYPE=NULL WHERE PROSPECTID=" & PROSPECTID)
            InitMemVars
            rsRefresh
            StoreMemVars
            If FormExist("MainForm") Then
                MainForm.ShowData
            End If

        End If
    End If

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If txtApl_No.Text <> "" Then
        AddorEdit = "EDIT"
        picLoan.Enabled = True
        picAdds.Visible = False
        picSaves.Visible = True
        txtDateApplied.Enabled = False
        On Error Resume Next
        txtComp_Busname.SetFocus
    End If
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEditTranDate_Click()
    If AddorEdit = "EDIT" Then
        If Function_Access(LOGID, "ACESS_SYSTEM", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
        txtDateApplied.Enabled = True: txtDateApplied.SetFocus
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

    On Error GoTo ErrorCode:

    Call flex_FillListView(gconDMIS.Execute("SELECT DateApplied, BusName,OfficeAdd , LStatus,ID from SMIS_LOANCORP"), lstLoan)
    cmdSelect.Enabled = False
    ShowHidePictureBox2 picFindLoan, True





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

    If Not rsLoan.BOF Then
        rsLoan.MoveFirst
    Else
        ShowFirstRecordMsg
    End If
    StoreMemVars





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:

    If Not rsLoan.EOF Then
        rsLoan.MoveLast
    Else
        ShowLastRecordMsg
    End If
    StoreMemVars





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    If Not rsLoan.EOF Then
        rsLoan.MoveNext
    End If

    If rsLoan.EOF And rsLoan.RecordCount > 0 Then
        rsLoan.MoveLast

        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox("Do You want to Post this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = ("UPDATE SMIS_LOANCORP SET STATUS='P' WHERE ID=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "P", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & Null2String(txtApl_No), "", ""
    rsRefresh
    rsLoan.Find ("ID=" & labid)
    StoreMemVars
    LogAudit "P", "CORORATE LOAN APPLICATION", txtApl_No & " " & txtComp_Busname

    MessagePop RecSaveOk, "Posted", "Record Sucessfully Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    If Not rsLoan.BOF Then
        rsLoan.MovePrevious
    End If

    If rsLoan.BOF And rsLoan.RecordCount > 0 Then
        rsLoan.MoveFirst
        ShowFirstRecordMsg
    End If

    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtComp_Busname)) = "" Then
        ShowIsRequiredMsg "Company Name"
        On Error Resume Next
        txtComp_Busname.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(cboLoan_SAENAME)) = "" Then
        ShowIsRequiredMsg "Sales Agent Name"
        On Error Resume Next
        cboLoan_SAENAME.SetFocus
        Exit Sub
    End If


    If RTrim(LTrim(cboLoan_UnitModel)) = "" Then
        ShowIsRequiredMsg "Unit Model"
        On Error Resume Next
        cboLoan_UnitModel.SetFocus
        Exit Sub
    End If
    Dim AplCode, Apl_no, DateApplied                                  As String
    Dim Busname, TelNo, OfficeAdd, CorporationType, DateEstablised, YearsInOpt, NatureOfBusiness, MajorProduct, PaidUpCapital As String
    Dim TinNo, CCINo, DateofIssue, PlaceOfIssue                       As String
    Dim UnitModel, SAENAME, Purpose, PlaceofUse, NetCostPrice, DownPayment, BalanceFianced, AOR, Terms, MonthlyAmortization As String
    Dim StockHolderName1, StockHolderAdd1, StockHolderAmount1         As String
    Dim StockHolderName2, StockHolderAdd2, StockHolderAmount2         As String
    Dim StockHolderAmount3, StockHolderName3, StockHolderAdd3         As String
    Dim StockHolderName4, StockHolderAdd4, StockHolderAmount4         As String
    Dim ODName1, ODAddress1, ODPosition1                              As String
    Dim ODName2, ODAddress2, ODPosition2                              As String
    Dim ODName3, ODAddress3, ODPosition3                              As String
    Dim ODName4, ODAddress4, ODPosition4                              As String
    Dim ODContactPerson, ODDesignation, ODTelNo                       As String
    Dim SuretyName, SuretyAddress, SuretyCompanyIncome, SuretyIncome  As String
    Dim Ref_TradeName, Ref_TradeAdd, Ref_TradeTelNo                   As String
    Dim Ref_SupName, Ref_SupAdd, Ref_SupTelNo                         As String
    Dim Ref_LoanName, Ref_LoanAdd, Ref_LoanTelNo                      As String
    Dim Ref_CreditName, Ref_CreditAdd, Ref_CreditTelNo                As String
    Dim BA_Bank1, BA_TOA1, BA_AcctNo1, BA_Balance1                    As String
    Dim BA_Bank2, BA_TOA2, BA_AcctNo2, BA_Balance2                    As String
    Dim BA_Bank3, BA_TOA3, BA_AcctNo3, BA_Balance3                    As String
    Dim BA_Bank4, BA_TOA4, BA_AcctNo4, BA_Balance4                    As String
    Dim SQL                                                           As String
    Dim SACODE                                                        As String
    AplCode = N2Str2Null(CUSCDE)
    Apl_no = N2Str2Null(txtApl_No)
    DateApplied = N2Str2Null(txtDateApplied)
    Busname = N2Str2Null(txtComp_Busname)
    TelNo = N2Str2Null(txtComp_TelNo)
    OfficeAdd = N2Str2Null(txtComp_OfficeAdd)
    If optCorporationTypeCorporation.Value = True Then
        CorporationType = N2Str2Null("Corporation")
    ElseIf optCorporationTypePartnership.Value = True Then
        CorporationType = N2Str2Null("Partnership")
    ElseIf optCorporationTypeSingle.Value = True Then
        CorporationType = N2Str2Null("Single")
    Else
        CorporationType = N2Str2Null("")
    End If


    DateEstablised = N2Str2Null(txtComp_DateEstablised)
    YearsInOpt = N2Str2Null(txtComp_YearsInOpt)

    NatureOfBusiness = N2Str2Null(cboComp_NatureOfBusiness)
    MajorProduct = N2Str2Null(txtComp_MajorProduct)
    PaidUpCapital = NumericVal(txtComp_PaidUpCapital)
    TinNo = N2Str2Null(txtComp_TinNO)
    CCINo = N2Str2Null(txtComp_CCINo)
    DateofIssue = N2Str2Null(txtComp_IssuedOn)
    PlaceOfIssue = N2Str2Null(cboComp_IssuedAt)
    UnitModel = N2Str2Null(cboLoan_UnitModel)
    SAENAME = N2Str2Null(cboLoan_SAENAME)

    If optPurposeBusiness Then
        Purpose = N2Str2Null("Business")
    ElseIf optPurposePrivate Then
        Purpose = N2Str2Null("Private")
    ElseIf optPurposePublic Then
        Purpose = N2Str2Null("Public")
    Else
        Purpose = N2Str2Null("")
    End If


    PlaceofUse = N2Str2Null(cboLoan_PlaceofUse)
    NetCostPrice = NumericVal(txtLoan_UnitCost)
    DownPayment = NumericVal(txtLoan_Downpayment)
    BalanceFianced = NumericVal(txtLoan_FinBalAmount)
    AOR = NumericVal(txtLoan_AORPercentage)
    Terms = NumericVal(txtLoan_BankTerms)
    MonthlyAmortization = NumericVal(txtLoan_MonthlyAmortization)

    StockHolderName1 = N2Str2Null(txtSH_Name1)
    StockHolderAdd1 = N2Str2Null(txtSH_Adress1)
    StockHolderAmount1 = NumericVal(txtSH_Amount1)

    StockHolderName2 = N2Str2Null(txtSH_Name2)
    StockHolderAdd2 = N2Str2Null(txtSH_Adress2)
    StockHolderAmount2 = NumericVal(txtSH_Amount2)

    StockHolderName3 = N2Str2Null(txtSH_Name3)
    StockHolderAdd3 = N2Str2Null(txtSH_Adress3)
    StockHolderAmount3 = NumericVal(txtSH_Amount3)

    StockHolderName4 = N2Str2Null(txtSH_Name4)
    StockHolderAdd4 = N2Str2Null(txtSH_Adress4)
    StockHolderAmount4 = NumericVal(txtSH_Amount4)

    ODName1 = N2Str2Null(txtODName1)
    ODAddress1 = N2Str2Null(txtODAddress1)
    ODPosition1 = N2Str2Null(cboODPosition1)

    ODName2 = N2Str2Null(txtODName2)
    ODAddress2 = N2Str2Null(txtODAddress2)
    ODPosition2 = N2Str2Null(cboODPosition2)

    ODName3 = N2Str2Null(txtODName3)
    ODAddress3 = N2Str2Null(txtODAddress3)
    ODPosition3 = N2Str2Null(cboODPosition3)

    ODName4 = N2Str2Null(txtODName4)
    ODAddress4 = N2Str2Null(txtODAddress4)
    ODPosition4 = N2Str2Null(cboODPosition4)


    ODContactPerson = N2Str2Null(txtODContactPerson)
    ODDesignation = N2Str2Null(cboODDesignation)
    ODTelNo = N2Str2Null(txtODTelNo)

    SuretyName = N2Str2Null(txtSuretyName)
    SuretyAddress = N2Str2Null(txtSuretyAddress)
    SuretyCompanyIncome = NumericVal(txtSuretyCompanyIncome)
    SuretyIncome = NumericVal(txtSuretyIncome)

    Ref_TradeName = N2Str2Null(txtRef_TradeName)
    Ref_TradeAdd = N2Str2Null(txtRef_TradeAdd)
    Ref_TradeTelNo = N2Str2Null(txtRef_TradeTelNo)

    Ref_SupName = N2Str2Null(txtRef_SupName)
    Ref_SupAdd = N2Str2Null(txtRef_SupAdd)
    Ref_SupTelNo = N2Str2Null(txtRef_SupTelNo)

    Ref_LoanName = N2Str2Null(txtRef_LoanName)
    Ref_LoanAdd = N2Str2Null(txtRef_LoanAdd)
    Ref_LoanTelNo = N2Str2Null(txtRef_LoanTelNo)

    Ref_CreditName = N2Str2Null(txtRef_CreditName)
    Ref_CreditAdd = N2Str2Null(txtRef_CreditAdd)
    Ref_CreditTelNo = N2Str2Null(txtRef_CreditTelNo)

    BA_Bank1 = N2Str2Null(txtBA_Bank1)
    BA_TOA1 = N2Str2Null(cboBA_TOA1)
    BA_AcctNo1 = N2Str2Null(txtBA_AcctNo1)
    BA_Balance1 = NumericVal(txtBA_Balance1)

    BA_Bank2 = N2Str2Null(txtBA_Bank2)
    BA_TOA2 = N2Str2Null(cboBA_TOA2)
    BA_AcctNo2 = N2Str2Null(txtBA_AcctNo2)
    BA_Balance2 = NumericVal(txtBA_Balance2)

    BA_Bank3 = N2Str2Null(txtBA_Bank3)
    BA_TOA3 = N2Str2Null(cboBA_TOA3)
    BA_AcctNo3 = N2Str2Null(txtBA_AcctNo3)
    BA_Balance3 = NumericVal(txtBA_Balance3)

    BA_Bank4 = N2Str2Null(txtBA_Bank4)
    BA_TOA4 = N2Str2Null(cboBA_TOA4)
    BA_AcctNo4 = N2Str2Null(txtBA_AcctNo4)
    BA_Balance4 = NumericVal(txtBA_Balance4)


    SACODE = N2Str2Null(GetSAECode(cboLoan_SAENAME))

    Dim rsHanapID                                                     As ADODB.Recordset
    Dim vID                                                           As String

    Set rsHanapID = New ADODB.Recordset

    If AddorEdit = "ADD" Then

        SQL = "INSERT INTO SMIS_LoanCorp( USERCODE,"
        SQL = SQL & "AplCode,Apl_no, ProspectID,DateApplied,  "
        SQL = SQL & "Busname, TelNo, OfficeAdd, CorporationType, DateEstablised, YearsInOpt, NatureOfBusiness, MajorProduct, PaidUpCapital,  "
        SQL = SQL & "TinNo, CCINo, DateofIssue, PlaceOfIssue,  "
        SQL = SQL & "UnitModel, SAEName, Purpose, PlaceofUse, NetCostPrice, Downpayment, BalanceFianced, AOR, Terms, MonthlyAmortization,  "
        SQL = SQL & "StockHolderName1, StockHolderAdd1, StockHolderAmount1,  "
        SQL = SQL & "StockHolderName2, StockHolderAdd2, StockHolderAmount2,  "
        SQL = SQL & "StockHolderAmount3, StockHolderName3, StockHolderAdd3,  "
        SQL = SQL & "StockHolderName4, StockHolderAdd4, StockHolderAmount4,  "
        SQL = SQL & "ODName1, ODAddress1, ODPosition1,  "
        SQL = SQL & "ODName2, ODAddress2, ODPosition2,  "
        SQL = SQL & "ODName3, ODAddress3, ODPosition3,  "
        SQL = SQL & "ODName4, ODAddress4, ODPosition4,  "
        SQL = SQL & "ODContactPerson, ODDesignation, ODTelNo,  "
        SQL = SQL & "SuretyName, SuretyAddress, SuretyCompanyIncome, SuretyIncome,  "
        SQL = SQL & "Ref_TradeName, Ref_TradeAdd, Ref_TradeTelNo,  "
        SQL = SQL & "Ref_SupName, Ref_SupAdd, Ref_SupTelNo,  "
        SQL = SQL & "Ref_LoanName, Ref_LoanAdd, Ref_LoanTelNo,  "
        SQL = SQL & "Ref_CreditName, Ref_CreditAdd, Ref_CreditTelNo,  "
        SQL = SQL & "BA_Bank1, BA_TOA1, BA_AcctNo1, BA_Balance1,  "
        SQL = SQL & "BA_Bank2, BA_TOA2, BA_AcctNo2, BA_Balance2,  "
        SQL = SQL & "BA_Bank3, BA_TOA3, BA_AcctNo3, BA_Balance3,  "
        SQL = SQL & "BA_Bank4, BA_TOA4, BA_AcctNo4, BA_Balance4,  "
        SQL = SQL & "LStatus "
        SQL = SQL & ") values("
        SQL = SQL & SACODE & ","
        SQL = SQL & AplCode & "," & Apl_no & "," & PROSPECTID & "," & DateApplied & ","
        SQL = SQL & Busname & "," & TelNo & "," & OfficeAdd & "," & CorporationType & "," & DateEstablised & "," & YearsInOpt & "," & NatureOfBusiness & "," & MajorProduct & "," & PaidUpCapital & ","
        SQL = SQL & TinNo & "," & CCINo & "," & DateofIssue & "," & PlaceOfIssue & ","
        SQL = SQL & UnitModel & "," & SAENAME & "," & Purpose & "," & PlaceofUse & "," & NetCostPrice & "," & DownPayment & "," & BalanceFianced & "," & AOR & "," & Terms & "," & MonthlyAmortization & ","
        SQL = SQL & StockHolderName1 & "," & StockHolderAdd1 & "," & StockHolderAmount1 & ","
        SQL = SQL & StockHolderName2 & "," & StockHolderAdd2 & "," & StockHolderAmount2 & ","
        SQL = SQL & StockHolderAmount3 & "," & StockHolderName3 & "," & StockHolderAdd3 & ","
        SQL = SQL & StockHolderName4 & "," & StockHolderAdd4 & "," & StockHolderAmount4 & ","
        SQL = SQL & ODName1 & "," & ODAddress1 & "," & ODPosition1 & ","
        SQL = SQL & ODName2 & "," & ODAddress2 & "," & ODPosition2 & ","
        SQL = SQL & ODName3 & "," & ODAddress3 & "," & ODPosition3 & ","
        SQL = SQL & ODName4 & "," & ODAddress4 & "," & ODPosition4 & ","
        SQL = SQL & ODContactPerson & "," & ODDesignation & "," & ODTelNo & ","
        SQL = SQL & SuretyName & "," & SuretyAddress & "," & SuretyCompanyIncome & "," & SuretyIncome & ","
        SQL = SQL & Ref_TradeName & "," & Ref_TradeAdd & "," & Ref_TradeTelNo & ","
        SQL = SQL & Ref_SupName & "," & Ref_SupAdd & "," & Ref_SupTelNo & ","
        SQL = SQL & Ref_LoanName & "," & Ref_LoanAdd & "," & Ref_LoanTelNo & ","
        SQL = SQL & Ref_CreditName & "," & Ref_CreditAdd & "," & Ref_CreditTelNo & ","
        SQL = SQL & BA_Bank1 & "," & BA_TOA1 & "," & BA_AcctNo1 & "," & BA_Balance1 & ","
        SQL = SQL & BA_Bank2 & "," & BA_TOA2 & "," & BA_AcctNo2 & "," & BA_Balance2 & ","
        SQL = SQL & BA_Bank3 & "," & BA_TOA3 & "," & BA_AcctNo3 & "," & BA_Balance3 & ","
        SQL = SQL & BA_Bank4 & "," & BA_TOA4 & "," & BA_AcctNo4 & "," & BA_Balance4 & ","
        SQL = SQL & " 'O' )"
        gconDMIS.Execute SQL

        SQL_STATEMENT = SQL
        NEW_LogAudit "A", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & txtApl_No, "", ""


        LogAudit "A", "CORORATE LOAN APPLICATION", txtApl_No & " " & txtComp_Busname
    Else


        SQL = "UPDATE SMIS_LoanCorp SET "
        SQL = SQL & " DateApplied=" & DateApplied & "  ,  "
        SQL = SQL & " Busname=" & Busname & "  , TelNo=" & TelNo & "  , OfficeAdd=" & OfficeAdd & "  , CorporationType=" & CorporationType & "  , DateEstablised=" & DateEstablised & "  , YearsInOpt=" & YearsInOpt & ", NatureOfBusiness=" & NatureOfBusiness & "  , MajorProduct=" & MajorProduct & "  , PaidUpCapital=" & PaidUpCapital & " ,  "
        SQL = SQL & " TinNo=" & TinNo & "  , CCINo=" & CCINo & "  , DateofIssue=" & DateofIssue & "  , PlaceOfIssue=" & PlaceOfIssue & "  ,  "
        SQL = SQL & " UnitModel=" & UnitModel & "  , SAEName=" & SAENAME & "  , Purpose=" & Purpose & "  , PlaceofUse=" & PlaceofUse & "  ,  "
        SQL = SQL & " NetCostPrice=" & NetCostPrice & "  , Downpayment=" & DownPayment & "  , BalanceFianced=" & BalanceFianced & "  , AOR=" & AOR & "  , Terms=" & Terms & "  , MonthlyAmortization=" & MonthlyAmortization & "  ,  "
        SQL = SQL & " StockHolderName1=" & StockHolderName1 & "  , StockHolderAdd1=" & StockHolderAdd1 & "  , StockHolderAmount1=" & StockHolderAmount1 & "  ,  "
        SQL = SQL & " StockHolderName2=" & StockHolderName2 & "  , StockHolderAdd2=" & StockHolderAdd2 & "  , StockHolderAmount2=" & StockHolderAmount2 & "  ,  "
        SQL = SQL & " StockHolderName3=" & StockHolderName3 & "  , StockHolderAdd3=" & StockHolderAdd3 & "  , StockHolderAmount3=" & StockHolderAmount3 & "  ,  "
        SQL = SQL & " StockHolderName4=" & StockHolderName4 & "  , StockHolderAdd4=" & StockHolderAdd4 & "  , StockHolderAmount4=" & StockHolderAmount4 & "  ,  "
        SQL = SQL & " ODName1=" & ODName1 & "  , ODAddress1=" & ODAddress1 & "  , ODPosition1=" & ODPosition1 & "  ,  "
        SQL = SQL & " ODName2=" & ODName2 & "  , ODAddress2=" & ODAddress2 & "  , ODPosition2=" & ODPosition2 & "  ,  "
        SQL = SQL & " ODName3=" & ODName3 & "  , ODAddress3=" & ODAddress3 & "  , ODPosition3=" & ODPosition3 & "  ,  "
        SQL = SQL & " ODName4=" & ODName4 & "  , ODAddress4=" & ODAddress4 & "  , ODPosition4=" & ODPosition4 & "  ,  "
        SQL = SQL & " ODContactPerson=" & ODContactPerson & "  , ODDesignation=" & ODDesignation & "  , ODTelNo=" & ODTelNo & "  ,  "
        SQL = SQL & " SuretyName=" & SuretyName & "  , SuretyAddress=" & SuretyAddress & "  , SuretyCompanyIncome=" & SuretyCompanyIncome & "  , SuretyIncome=" & SuretyIncome & "  ,  "
        SQL = SQL & " Ref_TradeName=" & Ref_TradeName & "  , Ref_TradeAdd=" & Ref_TradeAdd & "  , Ref_TradeTelNo=" & Ref_TradeTelNo & "  ,  "
        SQL = SQL & " Ref_SupName=" & Ref_SupName & "  , Ref_SupAdd=" & Ref_SupAdd & "  , Ref_SupTelNo=" & Ref_SupTelNo & "  ,  "
        SQL = SQL & " Ref_LoanName=" & Ref_LoanName & "  , Ref_LoanAdd=" & Ref_LoanAdd & "  , Ref_LoanTelNo=" & Ref_LoanTelNo & "  ,  "
        SQL = SQL & " Ref_CreditName=" & Ref_CreditName & "  , Ref_CreditAdd=" & Ref_CreditAdd & "  , Ref_CreditTelNo=" & Ref_CreditTelNo & "  ,  "
        SQL = SQL & " BA_Bank1=" & BA_Bank1 & "  , BA_TOA1=" & BA_TOA1 & "  , BA_AcctNo1=" & BA_AcctNo1 & "  , BA_Balance1=" & BA_Balance1 & "  ,  "
        SQL = SQL & " BA_Bank2=" & BA_Bank2 & "  , BA_TOA2=" & BA_TOA2 & "  , BA_AcctNo2=" & BA_AcctNo2 & "  , BA_Balance2=" & BA_Balance2 & "  ,  "
        SQL = SQL & " BA_Bank3=" & BA_Bank3 & "  , BA_TOA3=" & BA_TOA3 & "  , BA_AcctNo3=" & BA_AcctNo3 & "  , BA_Balance3=" & BA_Balance3 & "  ,  "
        SQL = SQL & " USERCODE=" & SACODE & " , "
        SQL = SQL & " BA_Bank4=" & BA_Bank4 & "  , BA_TOA4=" & BA_TOA4 & "  , BA_AcctNo4=" & BA_AcctNo4 & "  , BA_Balance4=" & BA_Balance4
        SQL = SQL & " WHERE ID= " & labid
        gconDMIS.Execute (SQL)

        SQL_STATEMENT = SQL
        NEW_LogAudit "E", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & txtApl_No, "", ""
        LogAudit "E", "CORORATE LOAN APPLICATION", txtApl_No & " " & txtComp_Busname
    End If

    gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET AppNo=" & Apl_no & " , LOGAPPLICATION=" & DateApplied & " ,LogApplicationType='C' WHERE PROSPECTID=" & PROSPECTID)
    picLoan.Enabled = False
    picAdds.Visible = True
    picSaves.Visible = False
    rsRefresh
    Call rsLoan.Find("APL_NO=" & N2Str2Null(txtApl_No))
    If FormExist("MainForm") Then
        MainForm.ShowData
    End If
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelect_Click()
    If lstLoan.SelectedItem Is Nothing Then Exit Sub
    SearchID (lstLoan.SelectedItem.ListSubItems(4).Text)
    ShowHidePictureBox2 picFindLoan, False

End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "CORPORATE LOAN APPLICATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Do You want to Un Post this Applications", vbInformation + vbYesNo) = vbNo Then: Exit Sub
    SQL_STATEMENT = ("UPDATE SMIS_LOANCORP SET STATUS='U' WHERE ID=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "U", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & Null2String(txtApl_No), "", ""

    LogAudit "U", "CORORATE LOAN APPLICATION", txtApl_No & " " & txtComp_Busname
    rsRefresh
    rsLoan.Find ("ID=" & labid)
    StoreMemVars
    MessagePop RecSaveOk, "Un Posted", "Record Sucessfully Un-Posted", 1000, 2
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdUpdateStatus_Click()
    On Error GoTo ErrorCode:

    cboLoanStatus.ListIndex = SetStatus(Null2String(rsLoan!lStatus))
    txtReasonNote = Null2String(rsLoan!Notes)
    If IsDate(rsLoan!Lastupdated) Then
        DTPicker1 = Null2String(rsLoan!Lastupdated)
    Else
        DTPicker1 = LOGDATE
    End If
    ShowHidePictureBox2 picStatus, True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDocumentCheckList_Click()
    Dim SQL                                                           As String
    Dim lst                                                           As ListItem
    Dim RS                                                            As ADODB.Recordset

    On Error GoTo ErrorCode:

    ListView1.Enabled = False

    SQL = " Select Code, DocumentName , 1 chklist from SMIS_DOCUMENT where code in (select DocumentCode from SMIS_LoanDocument Where AplType='C' and AplCode=" & N2Str2Null(txtApl_No) & "  )" & vbCrLf
    SQL = SQL & " Union " & vbCrLf
    SQL = SQL & " Select Code, DocumentName , 0  chklist  from SMIS_DOCUMENT where code not in (select DocumentCode from SMIS_LoanDocument Where AplType='C' and AplCode=" & N2Str2Null(txtApl_No) & "  )" & vbCrLf

    Set RS = gconDMIS.Execute(SQL)
    ListView1.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        ListView1.Enabled = True
    End If

    While Not RS.EOF
        Set lst = ListView1.ListItems.Add(, , Null2String(RS!CODE))
        Call lst.ListSubItems.Add(, , Null2String(RS!DocumentName))
        Call lst.ListSubItems.Add(, , Null2String(RS!CODE))
        lst.Checked = CBool(RS!Chklist)
        RS.MoveNext
    Wend
    Set RS = Nothing
    ShowHidePictureBox2 picDocumentList, True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command3_Click()
    ShowHidePictureBox2 picDocumentList, False
End Sub

Private Sub Command4_Click()
    Dim Item                                                          As ListItem
    On Error GoTo ErrorCode:

    'gconDMIS.Execute ("Delete from SMIS_loanDocument where aplcode=" & N2Str2Null(txtApl_No) & " AND AplType='C'")
    For Each Item In ListView1.ListItems

        If Item.Checked = True Then
            SQL_STATEMENT = (" insert into SMIS_loanDocument([DocumentCode],[AplCode],[AplType]) values (" & N2Str2Null(Item.Text) & ", " & N2Str2Null(txtApl_No) & ", 'C')")
            gconDMIS.Execute (SQL_STATEMENT)
            NEW_LogAudit "EE", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & txtApl_No, "", ""
        End If
    Next
    MessagePop RecSaveOk, "Updated", "Document Listing Added"
    ShowHidePictureBox2 picDocumentList, False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Command6_Click()
    On Error GoTo ErrorCode:
    SQL_STATEMENT = ("Update smis_loancorp set notes= " & _
                     N2Str2Null(txtReasonNote) & ", LStatus =" & N2Str2Null(Left(cboLoanStatus, 1)) & ",Lastupdated='" & DTPicker1 & "',FINCOM=" & N2Str2Null(cboFinCom) & "  Where Id=" & labid)

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "EE", "CORPORATE LOAN APPLICATION", SQL_STATEMENT, Null2String(PROSPECTID), "", "Application No:" & Null2String(txtApl_No), "", ""

    rsLoan.Requery
    rsLoan.Find ("ID=" & labid)
    StoreMemVars

    ShowHidePictureBox2 picStatus, False

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

Private Sub Form_Activate()
    On Error Resume Next
    txtComp_Busname.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If picStatus.Visible = True Then
            ShowHidePictureBox2 picStatus, False
        ElseIf picFindLoan.Visible = True Then
            ShowHidePictureBox2 picFindLoan, False
        ElseIf picDocumentList.Visible = True Then
            ShowHidePictureBox2 picDocumentList, False
        End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (CORPORATE LOAN APPLICATION)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(PROSPECTID), "CORPORATE LOAN APPLICATION")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    Me.Height = Screen.TwipsPerPixelY * 540
    picMiddles.Height = Me.ScaleHeight - picBottoms.Height - picTops.Height
    ScrollBar1.Height = picMiddles.ScaleHeight - 15
    ScrollBar1.Max = Abs(picMiddles.ScaleHeight - picLoan.Height) + 20
    Call AddColumnHeader("Date, Company Name , Address , Status", lstLoan)
    Call ResizeColumnHeader(lstLoan, "15,40,28,15")
    Call AddColumnHeader("CODE,DOCUMENT NAME ", ListView1)
    Call ResizeColumnHeader(ListView1, "35,65")
    CenterMe frmMain, Me, 1
    picAdds.Visible = True
    picSaves.Visible = False
    InitMemVars
    InitCbo
    rsRefresh
    If LoanID > 0 Then
        rsLoan.Find ("ID='" & LoanID & "'")
    End If
    If AddingLoan = True Then
        Exit Sub
    Else
        StoreMemVars
    End If

End Sub

Private Sub FormSearch_NoSelectionMade()
    If rsLoan.EOF Or rsLoan.BOF Then
        Unload Me
    End If
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Dim Telephone                                                     As String
    AddorEdit = "ADD"
    picLoan.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True

    txtDateApplied.Text = FormatDateTime(LOGDATE, vbShortDate)
    InitMemVars

    txtApl_No = GenerateCode("SMIS_LOANCORP", "APLCODE", "0000000000")
    CUSCDE = Null2String(oCusRs!CUSCDE)
    PROSPECTID = Null2String(oCusRs!PROSPECTID)
    ProfileType = Null2String(oCusRs!ProspectType)

    txtDateApplied = FormatDateTime(LOGDATE, vbShortDate)
    txtComp_Busname = Null2String(oCusRs.Fields("Acctname"))
    cboLoan_UnitModel = Null2String(oCusRs.Fields("VARIANT"))
    cboLoan_SAENAME = Null2String(oCusRs.Fields("SAE"))
    txtODContactPerson = Null2String(oCusRs!ContactPerson)
    txtComp_OfficeAdd = Null2String(oCusRs!Address)


    Telephone = Null2String(oCusRs!Telephone) & "\" & Null2String(oCusRs!Mobile)
    If Left(Telephone, 1) = "\" Or Right(Telephone, 1) = "\" Then
        txtComp_TelNo = Replace(Telephone, "\", "")
    Else
        txtComp_TelNo = Telephone
    End If
    Unload FormSearch
    Set FormSearch = Nothing

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = True Then
        Command4.Enabled = True
        Exit Sub
    End If
    For Each Item In ListView1.ListItems
        If Item.Checked = True Then
            Command4.Enabled = True
            Exit Sub
        End If
    Next
    Command4.Enabled = False

End Sub

Private Sub lstLoan_DblClick()
    cmdSelect_Click
End Sub

Private Sub lstLoan_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdSelect.Enabled = True
End Sub

'Private Sub optCorporationTypeCorporation_Click()
'    Dim cntrl                          As Control
'    fraStockholder.Enabled = True
'End Sub
'
'Private Sub optCorporationTypePartnership_Click()
'    fraStockholder.Enabled = False
'End Sub
'
'Private Sub optCorporationTypeSingle_Click()
'    fraStockholder.Enabled = False
'End Sub

Private Sub rsRefresh()
    Set rsLoan = New Recordset
    Call rsLoan.Open("SElect * from SMIS_LoanCorp ORDER BY ID DESC", gconDMIS, adOpenDynamic, adLockReadOnly)
End Sub

Private Sub ScrollBar1_Change()
    picLoan.Top = 0 - ScrollBar1.Value
End Sub

Private Sub Timer1_Timer()
    If labLStatus.Caption <> "" Then
        If labLStatus.Visible = True Then
            labLStatus.Visible = False
        Else
            labLStatus.Visible = True
        End If
    End If
End Sub

Private Sub txtBA_Balance1_GotFocus()
    If NumericVal(txtBA_Balance1.Text) <= 0 Then txtBA_Balance1 = ""

End Sub

Private Sub txtBA_Balance1_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtBA_Balance1_LostFocus()
    If NumericVal(txtBA_Balance1.Text) <= 0 Then txtBA_Balance1 = "0.00"
    txtBA_Balance1 = FormatNumber(txtBA_Balance1)
End Sub

Private Sub txtBA_Balance2_GotFocus()
    If NumericVal(txtBA_Balance2.Text) <= 0 Then txtBA_Balance2 = ""

End Sub

Private Sub txtBA_Balance2_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtBA_Balance2_LostFocus()
    If NumericVal(txtBA_Balance2.Text) <= 0 Then txtBA_Balance2 = "0.00"
    txtBA_Balance2 = FormatNumber(txtBA_Balance2)
End Sub

Private Sub txtBA_Balance3_GotFocus()
    If NumericVal(txtBA_Balance3.Text) <= 0 Then txtBA_Balance3 = ""

End Sub

Private Sub txtBA_Balance3_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtBA_Balance3_LostFocus()
    If NumericVal(txtBA_Balance3.Text) <= 0 Then txtBA_Balance3 = "0.00"
    txtBA_Balance3 = FormatNumber(txtBA_Balance3)
End Sub

Private Sub txtBA_Balance4_GotFocus()
    If NumericVal(txtBA_Balance4.Text) <= 0 Then txtBA_Balance4 = ""

End Sub

Private Sub txtBA_Balance4_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtBA_Balance4_LostFocus()
    If NumericVal(txtBA_Balance4.Text) <= 0 Then txtBA_Balance4 = "0.00"
    txtBA_Balance4 = FormatNumber(txtBA_Balance4)
End Sub

Private Sub txtBA_Bank1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtBA_Bank2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtBA_Bank3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtBA_Bank4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtComp_CCINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtComp_DateEstablised_LostFocus()
    If IsDate(txtComp_DateEstablised) Then
        txtComp_YearsInOpt = DateDiff("yyyy", txtComp_DateEstablised, LOGDATE)
    ElseIf IsNumeric(txtComp_DateEstablised) Then
        txtComp_YearsInOpt = Year(LOGDATE) - NumericVal(txtComp_DateEstablised)
    End If
End Sub

Private Sub txtComp_IssuedOn_LostFocus()
    If IsDate(txtComp_IssuedOn) = False Then
        txtComp_IssuedOn = ""
    Else
        txtComp_IssuedOn = FormatDateTime(txtComp_IssuedOn, vbShortDate)
    End If
End Sub

Private Sub txtComp_PaidUpCapital_GotFocus()
    If NumericVal(txtComp_PaidUpCapital.Text) <= 0 Then txtComp_PaidUpCapital = ""

End Sub

Private Sub txtComp_PaidUpCapital_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtComp_PaidUpCapital_LostFocus()
    If NumericVal(txtComp_PaidUpCapital.Text) <= 0 Then txtComp_PaidUpCapital = "0.00"
    txtComp_PaidUpCapital = FormatNumber(txtComp_PaidUpCapital)
End Sub

Private Sub txtComp_TinNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)

End Sub

Private Sub txtLoan_AORPercentage_Change()
    txtLoan_FinBalAmount_Change
End Sub

Private Sub txtLoan_AORPercentage_GotFocus()
    If NumericVal(txtLoan_AORPercentage.Text) <= 0 Then txtLoan_AORPercentage = ""

End Sub

Private Sub txtLoan_AORPercentage_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_AORPercentage_LostFocus()
    If NumericVal(txtLoan_AORPercentage.Text) <= 0 Then txtLoan_AORPercentage = "0.00"
    txtLoan_AORPercentage = FormatNumber(txtLoan_AORPercentage)
End Sub

Private Sub txtLoan_BankTerms_Change()
    txtLoan_FinBalAmount_Change
End Sub

Private Sub txtLoan_BankTerms_GotFocus()
    If NumericVal(txtLoan_BankTerms.Text) <= 0 Then txtLoan_BankTerms = ""

End Sub

Private Sub txtLoan_BankTerms_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_BankTerms_LostFocus()
    If NumericVal(txtLoan_BankTerms.Text) <= 0 Then txtLoan_BankTerms = "0"
    txtLoan_BankTerms = FormatNumber(txtLoan_BankTerms)
End Sub

Private Sub txtLoan_Downpayment_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateAmountDetails
End Sub

Private Sub txtLoan_Downpayment_GotFocus()
    If NumericVal(txtLoan_Downpayment.Text) <= 0 Then txtLoan_Downpayment = ""

End Sub

Private Sub txtLoan_Downpayment_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_Downpayment_LostFocus()
    If NumericVal(txtLoan_Downpayment.Text) <= 0 Then txtLoan_Downpayment = "0.00"
    txtLoan_Downpayment = FormatNumber(txtLoan_Downpayment)
End Sub

Private Sub txtLoan_FinBalAmount_Change()
    On Error Resume Next
    If AddorEdit = "" Then Exit Sub
    txtLoan_MonthlyAmortization = AORVALUE(NumericVal(txtLoan_FinBalAmount), NumericVal(txtLoan_AORPercentage), NumericVal(txtLoan_BankTerms))
End Sub

Private Sub txtLoan_UnitCost_Change()
    If AddorEdit = "" Then Exit Sub
    UpdateAmountDetails
End Sub

Private Sub txtLoan_UnitCost_GotFocus()
    If NumericVal(txtLoan_UnitCost.Text) <= 0 Then txtLoan_UnitCost = ""

End Sub

Private Sub txtLoan_UnitCost_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtLoan_UnitCost_LostFocus()
    If NumericVal(txtLoan_UnitCost.Text) <= 0 Then txtLoan_UnitCost = "0.00"
    txtLoan_UnitCost = FormatNumber(txtLoan_UnitCost)
End Sub

Private Sub txtODContactPerson_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtODName1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtODName2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtODName3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtODName4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRef_CreditName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRef_LoanName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRef_SupName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRef_TradeName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSH_Amount1_GotFocus()
    If NumericVal(txtSH_Amount1.Text) <= 0 Then txtSH_Amount1 = ""

End Sub

Private Sub txtSH_Amount1_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSH_Amount1_LostFocus()
    If NumericVal(txtSH_Amount1.Text) <= 0 Then txtSH_Amount1 = "0.00"
    txtSH_Amount1 = FormatNumber(txtSH_Amount1)
End Sub

Private Sub txtSH_Amount2_GotFocus()
    If NumericVal(txtSH_Amount2.Text) <= 0 Then txtSH_Amount2 = ""

End Sub

Private Sub txtSH_Amount2_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSH_Amount2_LostFocus()
    If NumericVal(txtSH_Amount2.Text) <= 0 Then txtSH_Amount2 = "0.00"
    txtSH_Amount2 = FormatNumber(txtSH_Amount2)
End Sub

Private Sub txtSH_Amount3_GotFocus()
    If NumericVal(txtSH_Amount3.Text) <= 0 Then txtSH_Amount3 = ""

End Sub

Private Sub txtSH_Amount3_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSH_Amount3_LostFocus()
    If NumericVal(txtSH_Amount3.Text) <= 0 Then txtSH_Amount3 = "0.00"
    txtSH_Amount3 = FormatNumber(txtSH_Amount3)
End Sub

Private Sub txtSH_Amount4_GotFocus()
    If NumericVal(txtSH_Amount4.Text) <= 0 Then txtSH_Amount4 = ""

End Sub

Private Sub txtSH_Amount4_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSH_Amount4_LostFocus()
    If NumericVal(txtSH_Amount4.Text) <= 0 Then txtSH_Amount4 = "0.00"
    txtSH_Amount4 = FormatNumber(txtSH_Amount4)
End Sub

Private Sub txtSH_Name1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSH_Name2_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSH_Name3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSH_Name4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSuretyCompanyIncome_GotFocus()
    If NumericVal(txtSuretyCompanyIncome.Text) <= 0 Then txtSuretyCompanyIncome = ""

End Sub

Private Sub txtSuretyCompanyIncome_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSuretyCompanyIncome_LostFocus()
    If NumericVal(txtSuretyCompanyIncome.Text) <= 0 Then txtSuretyCompanyIncome = "0.00"
    txtSuretyCompanyIncome = FormatNumber(txtSuretyCompanyIncome)
End Sub

Private Sub txtSuretyIncome_GotFocus()
    If NumericVal(txtSuretyIncome.Text) <= 0 Then txtSuretyIncome = ""

End Sub

Private Sub txtSuretyIncome_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtSuretyIncome_LostFocus()
    If NumericVal(txtSuretyIncome.Text) <= 0 Then txtSuretyIncome = "0.00"
    txtSuretyIncome = FormatNumber(txtSuretyIncome)
End Sub

Private Sub txtSuretyName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Public Sub AddEditLoan(IDX As Long)
    AddorEdit = "EDIT"
    rsLoan.Find ("ID='" & IDX & "'")
    StoreMemVars
    cmdEdit.Value = True
End Sub

Public Sub SearchID(XXX)

    rsLoan.MoveFirst
    rsLoan.Find ("ID=" & XXX)
    StoreMemVars

End Sub

