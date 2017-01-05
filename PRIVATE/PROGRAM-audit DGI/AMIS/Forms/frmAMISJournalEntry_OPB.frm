VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmAMISJournalEntry_OPB 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   7770
   ClientLeft      =   11040
   ClientTop       =   4800
   ClientWidth     =   9855
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAMISJournalEntry_OPB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9855
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4395
      Left            =   9900
      ScaleHeight     =   4395
      ScaleWidth      =   2955
      TabIndex        =   11
      Top             =   0
      Width           =   2955
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4245
         Left            =   90
         ScaleHeight     =   4215
         ScaleWidth      =   2745
         TabIndex        =   12
         Top             =   60
         Width           =   2775
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   30
            TabIndex        =   14
            Top             =   720
            Width           =   2685
            _Version        =   655364
            _ExtentX        =   4736
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   " F4 - Add/View Details"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
            Height          =   375
            Left            =   30
            TabIndex        =   16
            Top             =   1500
            Width           =   2685
            _Version        =   655364
            _ExtentX        =   4736
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   " F11 - Post by Batch"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
            Height          =   375
            Left            =   30
            TabIndex        =   15
            Top             =   1110
            Width           =   2685
            _Version        =   655364
            _ExtentX        =   4736
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   " F9 - Add from Templates"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
            Alignment       =   1
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   345
            Left            =   30
            TabIndex        =   13
            Top             =   360
            Width           =   2685
            _Version        =   655364
            _ExtentX        =   4736
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   " F3 - Add Entries"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            VisualTheme     =   3
            Alignment       =   1
         End
      End
   End
   Begin Crystal.CrystalReport rptAP 
      Left            =   9570
      Top             =   8100
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Accounts Payable Printout"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -210
      ScaleHeight     =   900
      ScaleWidth      =   9735
      TabIndex        =   65
      Top             =   6870
      Width           =   9735
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
         Left            =   8820
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   765
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
         Left            =   8070
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   765
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
         Left            =   7320
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Cancel this Transaction"
         Top             =   30
         Width           =   765
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
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
         Left            =   6540
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Unpost this Transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
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
         Left            =   5790
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":1B5D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":1CAF
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Post this Transaction"
         Top             =   30
         Width           =   765
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
         Left            =   5040
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":1FD4
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   765
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
         Left            =   4290
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":2482
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":25D4
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   765
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
         Left            =   3540
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":28E7
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":2A39
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   765
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
         Left            =   2790
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":2D89
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":2EDB
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Move to First Record"
         Top             =   30
         Width           =   765
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
         Left            =   2040
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":3239
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":338B
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   765
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
         Left            =   1290
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":3685
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":37D7
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   765
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
         Left            =   540
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":3B2F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":3C81
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7860
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   78
      Top             =   6870
      Width           =   1980
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
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":3FE0
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":4132
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   765
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
         Left            =   10
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":4470
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_OPB.frx":45C2
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox picGJ 
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   150
      ScaleHeight     =   6405
      ScaleWidth      =   9555
      TabIndex        =   24
      Top             =   420
      Width           =   9555
      Begin TabDlg.SSTab SSTab1 
         Height          =   5280
         Left            =   60
         TabIndex        =   82
         Top             =   1050
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   9313
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "[<F3> Add &Journal Entries]"
         TabPicture(0)   =   "frmAMISJournalEntry_OPB.frx":4912
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstGJ"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   495
            Left            =   60
            TabIndex        =   84
            Top             =   4410
            Width           =   9195
            Begin VB.PictureBox picDatePosted 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   90
               ScaleHeight     =   375
               ScaleWidth      =   3405
               TabIndex        =   85
               Top             =   90
               Width           =   3405
               Begin VB.Label lblDatePosted 
                  BackStyle       =   0  'Transparent
                  Caption         =   "02/29/2008"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   86
                  Top             =   60
                  Width           =   1500
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackColor       =   &H8000000D&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Date Posted:"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   210
                  Left            =   0
                  TabIndex        =   87
                  Top             =   60
                  Width           =   1245
               End
            End
            Begin VB.TextBox txtGJTotCredit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   345
               Left            =   7710
               MaxLength       =   14
               TabIndex        =   90
               Text            =   "Text1"
               Top             =   90
               Width           =   1485
            End
            Begin VB.TextBox txtGJTotDebit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   345
               Left            =   6150
               MaxLength       =   14
               TabIndex        =   89
               Text            =   "Text1"
               Top             =   90
               Width           =   1515
            End
            Begin VB.TextBox txtGJOutBalance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1320
               MaxLength       =   14
               TabIndex        =   88
               Text            =   "Text1"
               Top             =   90
               Width           =   1515
            End
            Begin VB.Label lblOutBalance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Out of Balance"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   0
               TabIndex        =   91
               Top             =   210
               Width           =   1275
            End
         End
         Begin MSComctlLib.ListView lstGJ 
            Height          =   4335
            Left            =   30
            TabIndex        =   83
            Top             =   60
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   7646
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
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmAMISJournalEntry_OPB.frx":492E
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEM #"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ACCOUNT CODE"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCOUNT DESCRIPTION"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "DEBIT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "CREDIT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox picGJEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   150
         ScaleHeight     =   1965
         ScaleWidth      =   9225
         TabIndex        =   28
         Top             =   2880
         Width           =   9255
         Begin VB.CommandButton cmdGJCancel 
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
            Left            =   8430
            MouseIcon       =   "frmAMISJournalEntry_OPB.frx":4A90
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_OPB.frx":4BE2
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Cancel"
            Top             =   1110
            Width           =   765
         End
         Begin VB.CommandButton cmdGJDelete 
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
            Left            =   60
            MouseIcon       =   "frmAMISJournalEntry_OPB.frx":4F20
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_OPB.frx":5072
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Delete Selected Record"
            Top             =   1110
            Width           =   765
         End
         Begin VB.CommandButton cmdGJSave 
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
            Left            =   7680
            MouseIcon       =   "frmAMISJournalEntry_OPB.frx":539D
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_OPB.frx":54EF
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Save Entry"
            Top             =   1110
            Width           =   765
         End
         Begin VB.ComboBox cboJVSupCust 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1F6F5&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00973640&
            Height          =   330
            Left            =   2340
            TabIndex        =   42
            Text            =   "Combo1"
            Top             =   690
            Width           =   4305
         End
         Begin VB.Frame fraATC2 
            BackColor       =   &H00F5F5F5&
            Height          =   915
            Left            =   2310
            TabIndex        =   44
            Top             =   990
            Width           =   4365
            Begin VB.TextBox txtTaxBase2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   2550
               MaxLength       =   15
               TabIndex        =   50
               Top             =   510
               Width           =   1725
            End
            Begin VB.TextBox txtRATE2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   49
               Top             =   510
               Width           =   615
            End
            Begin VB.ComboBox cboATC2 
               BackColor       =   &H00F1F6F5&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00973640&
               Height          =   330
               Left            =   60
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   510
               Width           =   1425
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Taxbase Amt."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   2550
               TabIndex        =   47
               Top             =   240
               Width           =   1725
            End
            Begin VB.Label Label48 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "RATE"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   1380
               TabIndex        =   46
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "ATC Code"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label Label46 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   2190
               TabIndex        =   51
               Top             =   540
               Width           =   855
            End
         End
         Begin VB.ComboBox cboGJAccountNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   60
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin RichTextLib.RichTextBox txtGJAccountName 
            Height          =   315
            Left            =   2340
            TabIndex        =   35
            Top             =   330
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   15920873
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"frmAMISJournalEntry_OPB.frx":583F
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSMask.MaskEdBox txtGJDebit 
            Height          =   315
            Left            =   6690
            TabIndex        =   36
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15920873
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtGJCredit 
            Height          =   315
            Left            =   7950
            TabIndex        =   38
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   15920873
            ForeColor       =   7347754
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   315
            Left            =   7770
            TabIndex        =   54
            Top             =   1200
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtGJItemNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2580
            MaxLength       =   4
            TabIndex        =   34
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin RichTextLib.RichTextBox txtGJAccountParticulars 
            Height          =   885
            Left            =   2340
            TabIndex        =   56
            Top             =   2250
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1561
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmAMISJournalEntry_OPB.frx":58D2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label labATC 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1170
            TabIndex        =   43
            Top             =   750
            Width           =   1305
         End
         Begin VB.Label labGJID 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   1890
            TabIndex        =   57
            Top             =   2400
            Width           =   2205
         End
         Begin VB.Label Label29 
            BackColor       =   &H00FCFCFC&
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2340
            TabIndex        =   40
            Top             =   420
            Width           =   2685
         End
         Begin VB.Label Label28 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quantity"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2880
            TabIndex        =   41
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label27 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Item No."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6840
            TabIndex        =   37
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label26 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7950
            TabIndex        =   32
            Top             =   60
            Width           =   795
         End
         Begin VB.Label Label25 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   6720
            TabIndex        =   31
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label24 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   90
            TabIndex        =   29
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Item No."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   390
            TabIndex        =   39
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label22 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   2400
            TabIndex        =   30
            Top             =   60
            Width           =   2205
         End
      End
      Begin wizButton.cmd cmdGJEntry 
         Height          =   2115
         Left            =   90
         TabIndex        =   27
         Top             =   2820
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3731
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmAMISJournalEntry_OPB.frx":5969
      End
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   60
         TabIndex        =   26
         Top             =   330
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmAMISJournalEntry_OPB.frx":5985
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   255
         Left            =   90
         TabIndex        =   25
         Top             =   60
         Width           =   1695
      End
   End
   Begin VB.PictureBox picTemplates 
      Height          =   4125
      Left            =   1110
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   59
      Top             =   960
      Width           =   7185
      Begin VB.TextBox txtSearchTemplates 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   60
         MaxLength       =   50
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   61
         Top             =   450
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   5583
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":5A1C
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DESCRIPTION"
            Object.Width           =   11819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FCFCFC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Press <Enter> to Insert Account Entries From Template"
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
         Height          =   285
         Left            =   30
         TabIndex        =   62
         Top             =   3750
         Width           =   7035
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1050
      TabIndex        =   58
      Top             =   900
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7488
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmAMISJournalEntry_OPB.frx":5B7E
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   150
      ScaleHeight     =   1200
      ScaleWidth      =   9555
      TabIndex        =   1
      Top             =   60
      Width           =   9555
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1920
         Top             =   840
      End
      Begin VB.TextBox txtJItemNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   810
         Width           =   855
      End
      Begin VB.TextBox txtJNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7950
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txtVoucherNo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "000226"
         Top             =   60
         Width           =   1005
      End
      Begin VB.TextBox txtJDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7950
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   60
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6855
         TabIndex        =   9
         Top             =   930
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   210
         Left            =   6690
         TabIndex        =   4
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   4110
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label labPosted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "*** POSTED ***"
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
         Height          =   375
         Left            =   2550
         TabIndex        =   3
         Top             =   60
         Width           =   4065
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   5775
      Left            =   150
      TabIndex        =   17
      Top             =   240
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   10186
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "frmAMISJournalEntry_OPB.frx":5B9A
   End
   Begin VB.Frame fraFindAccount 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Chart of Accounts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   240
      TabIndex        =   18
      Top             =   300
      Width           =   9375
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   90
         MaxLength       =   50
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   270
         Width           =   9195
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4515
         Left            =   90
         TabIndex        =   21
         Top             =   660
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   7964
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmAMISJournalEntry_OPB.frx":5BB6
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   11819
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
      End
      Begin VB.CommandButton cmdAddAccount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add Account"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   5850
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3960
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label labAccountCode 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   20
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00EBFAFA&
         BackStyle       =   0  'Transparent
         Caption         =   "[Press <Enter> to Accept]                         [<F8> Change Search]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   23
         Top             =   5310
         Width           =   9225
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Journal No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3810
      TabIndex        =   64
      Top             =   6420
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label TEXTID 
      Caption         =   "fortestingpurpose"
      Height          =   705
      Left            =   1800
      TabIndex        =   81
      Top             =   8010
      Width           =   2295
   End
   Begin VB.Label TESTID 
      Caption         =   "TESTID"
      Height          =   645
      Left            =   270
      TabIndex        =   0
      Top             =   -5730
      Width           =   2535
   End
   Begin VB.Label lblVPJAcctCode 
      Caption         =   "dont delete this "
      Height          =   165
      Left            =   11220
      TabIndex        =   63
      Top             =   1080
      Width           =   1845
   End
End
Attribute VB_Name = "frmAMISJournalEntry_OPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As New ADODB.Recordset
Dim rsJournal_Det                                           As New ADODB.Recordset
Dim rsPV_Detail                                             As New ADODB.Recordset
Dim rsCV_Detail                                             As New ADODB.Recordset
Dim rsCRJ_Detail                                            As New ADODB.Recordset
Dim rsJV_detail                                             As New ADODB.Recordset
Dim rsChartAccount                                          As New ADODB.Recordset
Dim rsJournal_HD2                                           As New ADODB.Recordset
Dim rsProfile                                               As New ADODB.Recordset
Dim rsCheckJournal_HD                                       As New ADODB.Recordset
Dim rsVENDOR                                                As New ADODB.Recordset
Dim rsPayTerm                                               As New ADODB.Recordset
Dim rsBanks                                                 As New ADODB.Recordset
Dim rsCustomer                                              As New ADODB.Recordset
Dim rsInvoiceType                                           As New ADODB.Recordset
Dim rsATC                                                   As New ADODB.Recordset
Dim kcnt, Jcnt                                              As Integer
Attribute Jcnt.VB_VarUserMemId = 1073938448
Dim AddorEdit                                               As String
Attribute AddorEdit.VB_VarUserMemId = 1073938450
Dim SearchBy                                                As String
Attribute SearchBy.VB_VarUserMemId = 1073938451
Public CDJ_CIB                                              As String
Attribute CDJ_CIB.VB_VarUserMemId = 1073938452
Public CDJ_AP                                               As String
Attribute CDJ_AP.VB_VarUserMemId = 1073938453
Dim LocalAcess                                              As String
Attribute LocalAcess.VB_VarUserMemId = 1073938454
Dim TOTDEBIT                                                As Double
Attribute TOTDEBIT.VB_VarUserMemId = 1073938455
Dim TOTCREDIT                                               As Double
Attribute TOTCREDIT.VB_VarUserMemId = 1073938456
Dim TOTTAX                                                  As Double
Attribute TOTTAX.VB_VarUserMemId = 1073938457
Dim OUTBALANCE                                              As Double
Attribute OUTBALANCE.VB_VarUserMemId = 1073938458
Dim TOTAL_AR_AMOUNT                                         As Double
Attribute TOTAL_AR_AMOUNT.VB_VarUserMemId = 1073938459
Dim TOTALPVAMOUNT                                           As Double
Attribute TOTALPVAMOUNT.VB_VarUserMemId = 1073938460
Dim COMP_SJ_OUTPUT_TAX                                      As Double
Attribute COMP_SJ_OUTPUT_TAX.VB_VarUserMemId = 1073938461
Dim PrevJType                                               As String
Attribute PrevJType.VB_VarUserMemId = 1073938462
Dim PrevJNo                                                 As String
Attribute PrevJNo.VB_VarUserMemId = 1073938463
Dim PrevInvoiceType                                         As String
Attribute PrevInvoiceType.VB_VarUserMemId = 1073938464
Dim PrevInvoiceNo                                           As String
Attribute PrevInvoiceNo.VB_VarUserMemId = 1073938465
Dim PrevPV_VoucherNo                                        As String
Attribute PrevPV_VoucherNo.VB_VarUserMemId = 1073938466
Dim PrevPV_Amount                                           As Double
Attribute PrevPV_Amount.VB_VarUserMemId = 1073938467
Dim DirectDisbursementVoucherNo                             As String
Attribute DirectDisbursementVoucherNo.VB_VarUserMemId = 1073938468
Dim CDJ_IS_FROM_AP                                          As Boolean
Attribute CDJ_IS_FROM_AP.VB_VarUserMemId = 1073938469
Dim IsVPJ                                                   As Boolean
Attribute IsVPJ.VB_VarUserMemId = 1073938470
Dim TotalARAmountToPay                                      As Double
Attribute TotalARAmountToPay.VB_VarUserMemId = 1073938471
Dim TOTAL_AP_AMOUNT                                         As Double
Attribute TOTAL_AP_AMOUNT.VB_VarUserMemId = 1073938472
Dim TotalAPAmountToPay                                      As Double
Attribute TotalAPAmountToPay.VB_VarUserMemId = 1073938473
Dim SJVoucherno                                             As String
Attribute SJVoucherno.VB_VarUserMemId = 1073938474
Dim APJInvoiceNo                                            As String
Attribute APJInvoiceNo.VB_VarUserMemId = 1073938475
Dim APJinvoicetype                                          As String
Attribute APJinvoicetype.VB_VarUserMemId = 1073938476
Dim STR_CED_JOURNALTYPE                                     As String
Dim xJOURNALTYPE                                            As String

Sub SetJType(XXX)
    STR_CED_JOURNALTYPE = XXX
End Sub

Function GetVoucherNo(XXX As String) As String
    Dim rsJournal_HD                                        As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = '" & XXX & "' Order by VoucherNo desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function Setacctcode(VVV As Variant) As String
    Dim rsChartAccount2                                     As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctcode = UCase(Null2String(rsChartAccount2!AcctCode))
    Else
        Setacctcode = ""
    End If
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                     As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
End Function

Function SetAcctType(VVV As Variant) As String
    Dim rsChartAccount2                                     As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,AcctType from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctType = SetDebitCredit(Null2String(rsChartAccount2!ACCTTYPE))
    Else
        SetAcctType = ""
    End If
End Function

Function SetBankCode(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankname = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankCode = Null2String(rsBanks!bankcode)
        CDJ_CIB = N2Str2Null(rsBanks!AcctCode)
    Else
        SetBankCode = ""
        CDJ_CIB = "NULL"
    End If
End Function

Function SetBankName(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        rsBanks.Open "Select bankcode,bankname,acctcode from CMIS_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            SetBankName = Null2String(rsBanks!BankName)
            CDJ_CIB = N2Str2Null(rsBanks!AcctCode)
        Else
            SetBankName = ""
            CDJ_CIB = "NULL"
        End If
    Else
        rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            SetBankName = Null2String(rsBanks!BankName)
            CDJ_CIB = N2Str2Null(rsBanks!AcctCode)
        Else
            SetBankName = ""
            CDJ_CIB = "NULL"
        End If
    End If
End Function

Function SetCustomerCode(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    '    rsCustomer.Open "Select cuscde,acctname from ALL_CUSTMASTER_AMIS where acctname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsCustomer.Open "Select custcode,custname from ALL_CUSTMASTER_AMIS where custname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = Null2String(rsCustomer!CUSTCODE)
    Else
        SetCustomerCode = ""
    End If
End Function

Function SetCustomerName(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select custname from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(CCC))
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!CUSTNAME)
    Else
        SetCustomerName = ""
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerCreditTerm(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select CREDITDAYS from ALL_Customer_Table where cuscde = " & N2Str2Null(CCC))
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCreditTerm = Null2String(rsCustomer!CREDITDAYS)
    Else
        SetCustomerCreditTerm = 0
    End If
    Set rsCustomer = Nothing
End Function

Function SetDebitCredit(VVV As Variant) As String
    Dim rsAccountType                                       As ADODB.Recordset
    Set rsAccountType = New ADODB.Recordset
    rsAccountType.Open "Select Code,DebitCredit from AMIS_Acctype where Code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAccountType.EOF And Not rsAccountType.BOF Then
        SetDebitCredit = Null2String(rsAccountType!DebitCredit)
    Else
        SetDebitCredit = ""
    End If
End Function

Function SetInvCode(INV As Variant)
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invtype = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvCode = Null2String(rsInvoiceType!InvCode)
    Else
        SetInvCode = ""
    End If
End Function

Function SetInvType(INV As Variant)
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invcode = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvType = Null2String(rsInvoiceType!INVTYPE)
    Else
        SetInvType = ""
    End If
End Function

Function SetPayCode(VVV As Variant)
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayCode = Null2String(rsPayTerm!pay_Code)
    Else
        SetPayCode = ""
    End If
End Function

Function SetPayDesc(VVV As Variant) As String
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayDesc = Null2String(rsPayTerm!pay_desc)
    Else
        SetPayDesc = ""
    End If
End Function

Function SetPayNoDays(VVV As Variant) As Integer
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_Desc,no_days from ALL_PayTerm where pay_Desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayNoDays = Null2String(rsPayTerm!no_Days)
    Else
        SetPayNoDays = 0
    End If
End Function

Function SetVendorAddress(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,address from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorAddress = Null2String(rsVENDOR!Address)
    Else
        SetVendorAddress = ""
    End If
End Function

Function SetVendorCode(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where nameofvendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!CODE)
    Else
        SetVendorCode = ""
    End If
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function StoreGJEntry(ByVal ID As Variant)
    On Error GoTo ErrorCode
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,JNo,acct_code,acct_name,debit,jitemno,credit,tax,atc,rate,taxbase from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labGJID.Caption = rsJournal_Det!ID
        txtGJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboGJAccountNo.Text = Null2String(rsJournal_Det!ACCT_CODE)
        txtGJAccountName.Text = Null2String(rsJournal_Det!acct_Name)
        txtGJDebit.Text = N2Str2Zero(rsJournal_Det!Debit)
        txtGJCredit.Text = N2Str2Zero(rsJournal_Det!Credit)
        If fraATC2.Visible = True Then
            cboJVSupCust.Text = SetVendorName(Null2String(rsJournal_HD!VendorCode))
            If Null2String(rsJournal_Det!ATC) <> "" Then
                cboATC2.Text = Null2String(rsJournal_Det!ATC)
            Else
                cboATC2.ListIndex = 0
            End If
            txtRATE2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
            txtTaxBase2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!TAXBASE))
        End If
        StoreGJParticulars Null2String(rsJournal_Det!JNo), Null2String(rsJournal_Det!jitemno)
    End If
    Exit Function

ErrorCode:
    Resume Next
End Function

Function StoreGJParticulars(ByVal JNo As Variant, ByVal ItemNo As Variant)
    Set rsJV_detail = New ADODB.Recordset
    rsJV_detail.Open "select JNo,ItemNo,Particulars from AMIS_JV_Detail where JNo = " & N2Str2Null(JNo) & " and ItemNo = " & N2Str2Null(ItemNo), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJV_detail.EOF And Not rsJV_detail.BOF Then
        txtGJAccountParticulars.Text = Null2String(rsJV_detail!Particulars)
    End If
End Function

Function StoreDealerCode(XXX As String) As String
    Dim rsREPORWITHDealer                                   As ADODB.Recordset
    Set rsREPORWITHDealer = New ADODB.Recordset
    Set rsREPORWITHDealer = gconDMIS.Execute("SELECT  dbo.CSMS_SellingDealer.DealerCode AS VEHICLE_DEALER_CODE, dbo.CSMS_Repor.INVOICE FROM dbo.CSMS_Repor INNER JOIN dbo.CSMS_CusVeh ON dbo.CSMS_Repor.PLATE_NO = dbo.CSMS_CusVeh.VCOND_NO INNER JOIN dbo.CSMS_SellingDealer ON dbo.CSMS_CusVeh.SELLING_DEALER = dbo.CSMS_SellingDealer.DealerCode Where dbo.CSMS_Repor.INVOICE = '" & XXX & "'")
    If Not rsREPORWITHDealer.EOF And Not rsREPORWITHDealer.BOF Then
        StoreDealerCode = Null2String(rsREPORWITHDealer!VEHICLE_DEALER_CODE)
    End If
    Set rsREPORWITHDealer = Nothing
End Function

Function ReturnAP_AccountCode(XXX As String) As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AP' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAP_AccountCode = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Sub BringToFrontGJ()
    cmdGJEntry.ZOrder 0
    cmdGJEntry.Visible = True
    picGJEntry.ZOrder 0
    picGJEntry.Visible = True
    picGJEntry.Enabled = True
End Sub

Sub BringToFrontTemplates()
    cmdTemplates.ZOrder 0
    picTemplates.ZOrder 0
    FillTemplates
End Sub

Private Sub cmdNoCharge_Click()
'COMMENTED BY: ACL
'    If xJOURNALTYPE = "SJ" Then
'        ReturnInvoiceNo txtVoucherNo, xJOURNALTYPE
'        With frmAMIS_Payment
'            frmAMIS_Payment.FillPaymentdetail AMIS_Invoiceno, AMIS_Invoicetype
'            frmAMIS_Payment.Show
'        End With
'    End If
'    If xJOURNALTYPE = "APJ" Then
'        With frmAMIS_Payment
'            frmAMIS_Payment.FillPaymentdetail txtVoucherNo, ""
'            frmAMIS_Payment.Show
'        End With
'    End If
End Sub

Sub FillDetails()
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0: TotalARAmountToPay = 0
    TOTAL_AP_AMOUNT = 0: TotalAPAmountToPay = 0: PrevPV_Amount = 0
    Dim J_ITemNo, PV_ITEMNO                                 As Integer

    txtGJTotDebit.Text = ZERO: txtGJTotCredit.Text = ZERO: txtGJOutBalance.Text = ZERO
    lstGJ.Sorted = False: lstGJ.ListItems.Clear
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        Screen.MousePointer = 11
        rsJournal_Det.MoveFirst
        Do While Not rsJournal_Det.EOF
            kcnt = kcnt + 1
            If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
            lstGJ.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
            lstGJ.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!ACCT_CODE)
            lstGJ.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
            lstGJ.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!Debit))
            lstGJ.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!Credit))
            lstGJ.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
            TOTDEBIT = TOTDEBIT + NumericVal(N2Str2Zero(rsJournal_Det!Debit))
            TOTCREDIT = TOTCREDIT + NumericVal(N2Str2Zero(rsJournal_Det!Credit))
            TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJournal_Det!tax))
            rsJournal_Det.MoveNext
        Loop
        lstGJ.Sorted = True: lstGJ.Refresh
        OUTBALANCE = TOTDEBIT - TOTCREDIT
        txtGJTotDebit.Text = ToDoubleNumber(TOTDEBIT)
        txtGJTotCredit.Text = ToDoubleNumber(TOTCREDIT)
        txtGJOutBalance.Text = ToDoubleNumber(Abs(OUTBALANCE))
        Screen.MousePointer = 0
        If labPosted.Caption = "" Then
            If ToDoubleNumber(Abs(OUTBALANCE)) <> 0 Then
                cmdPost.Enabled = False
                lblOutBalance.Visible = True
                txtGJOutBalance.Visible = True
            Else
                cmdPost.Enabled = True
                lblOutBalance.Visible = False
                txtGJOutBalance.Visible = False
            End If
            Screen.MousePointer = 0
        End If
    Else
        cmdPost.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccount2                                     As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccount2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If SearchBy = "NAME" Then
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,upper(Description),Accttype,ID from AMIS_ChartAccount where description like'" & XXX & "%' order by acctcode asc")
    Else
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,UPPER(Description),Accttype,ID from AMIS_ChartAccount where acctcode like'" & XXX & "%' order by acctcode asc")
    End If
    If Not (rsChartAccount2.EOF And rsChartAccount2.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccount2
        lstAccounts.Refresh
        lstAccounts.Enabled = True
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Sub FillSearchTemplates(XXX As String)
    Dim rsTemplate_Header                                   As ADODB.Recordset
    lstTemplates.Enabled = False
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & xJOURNALTYPE & "' AND description like '" & XXX & "%' order by description asc")
    If Not (rsTemplate_Header.EOF And rsTemplate_Header.BOF) Then
        Listview_Loadval Me.lstTemplates.ListItems, rsTemplate_Header
        lstTemplates.Refresh
        lstTemplates.Enabled = True
        lstTemplates.Enabled = True
    Else
        lstTemplates.Enabled = False
    End If

End Sub

Sub FillTemplates()
    Dim rsTemplate_Header                                   As ADODB.Recordset
    lstTemplates.Enabled = False
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & xJOURNALTYPE & "' order by description asc")
    If Not (rsTemplate_Header.EOF And rsTemplate_Header.BOF) Then
        lstTemplates.Enabled = True
        Listview_Loadval Me.lstTemplates.ListItems, rsTemplate_Header
        lstTemplates.Refresh
        lstTemplates.Enabled = True
    Else
        lstTemplates.Enabled = False
    End If

End Sub

Sub FindDupJNo(DDD As String)
    rsJournal_HD.Bookmark = rsFind(rsJournal_HD.Clone, "jno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub ShowInvoiceApp(XXX As String, YYY As String)
    INVOICE_DETAIL_TYPE = XXX
    INVOICE_DETAIL_TRANNO = YYY
    frmInvoiceAppDetail.Show vbModal
End Sub

'Function SetCustomerCode(CCC As Variant)
'Set rsCustomer = New ADODB.Recordset
'    rsCustomer.Open "Select cuscde,acctname from ALL_CUSTMASTER_AMIS where acctname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
'If Not rsCustomer.EOF And Not rsCustomer.BOF Then
'   SetCustomerCode = Null2String(rsCustomer!cuscde)
'Else
'   SetCustomerCode = ""
'End If
'End Function

'Function SetCustomerName(CCC As Variant)
'Set rsCustomer = New ADODB.Recordset
'    rsCustomer.Open "Select cuscde,acctname from ALL_CUSTMASTER_AMIS where cuscde = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
'If Not rsCustomer.EOF And Not rsCustomer.BOF Then
'   SetCustomerName = Null2String(rsCustomer!acctname)
'Else
'   SetCustomerName = ""
'End If
'End Function

Sub InitGJ()
    txtGJItemNo.Text = Format(kcnt + 1, "0000")
    cboGJAccountNo.Text = ""
    txtGJAccountName.Text = ""
    txtGJDebit.Text = ZERO
    txtGJCredit.Text = ZERO
    txtGJAccountParticulars.Text = "Pls. Type Your Remarks Here..."
    txtSearch.Text = ""
    'cboATC2.ListIndex = 0
    'txtRATE2.Text = "0"
    'txtTaxBase2.Text = ZERO
End Sub

Sub InitJournal()
    txtJItemNo.Text = Format(kcnt + 1, "0000")
End Sub

Sub initMemvars()
    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"

    Dim rsJDate                                             As ADODB.Recordset
    Set rsJDate = New ADODB.Recordset
    rsJDate.Open "Select * from ALL_PROFILE", gconDMIS, adOpenKeyset
    If Not rsJDate.EOF And Not rsJDate.BOF Then
        txtJDate.Text = Null2Date(rsJDate!Cut_Off_Date)
    End If
    picDatePosted.Visible = False
    labPosted.Caption = ""
    labPosted.Visible = False
End Sub

Sub InsertAccountEntries(XXX As Variant)
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                       As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                     As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET            As Double
    Dim J_STATUS, J_JITEMNO                                 As String
    Dim rsTemplate_Details                                  As ADODB.Recordset
    Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("Select * from AMIS_Template_Details Where TemplateCode = " & XXX)
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        rsTemplate_Details.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsTemplate_Details.EOF
            kcnt = kcnt + 1
            J_JDATE = N2Date2Null(txtJDate.Text)
            J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
            J_JTYPE = N2Str2Null(xJOURNALTYPE)
            J_JNO = N2Str2Null(txtJNo.Text)
            J_JITEMNO = N2Str2Null(Format(kcnt, "0000"))
            J_ACCT_CODE = N2Str2Null(rsTemplate_Details!AccountCode)
            J_ACCT_NAME = N2Str2Null(rsTemplate_Details!DESCRIPTION)
            J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
            J_STATUS = "'N'"
            gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                             " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                             ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
            rsTemplate_Details.MoveNext
        Loop
        StoreMemVars
        FillDetails

        Screen.MousePointer = 0
    End If
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    If xJOURNALTYPE = "OPB" Then
        cboGJAccountNo.Text = labAccountCode.Caption
        txtGJAccountName.Text = Setacctname(labAccountCode.Caption)
        If cboGJAccountNo.Text <> "" Then
            If SetAcctType(cboGJAccountNo.Text) = "C" Then
                On Error Resume Next
                txtGJCredit.SetFocus
            Else
                On Error Resume Next
                txtGJDebit.SetFocus
            End If
        End If
    End If
    cmdFindAccount.ZOrder 1
    fraFindAccount.ZOrder 1
End Sub

Sub OkAccountSetCursor()
    If xJOURNALTYPE = "OPB" Then
        If cboGJAccountNo.Text <> "" Then
            If SetAcctType(cboGJAccountNo.Text) = "C" Then
                On Error Resume Next
                txtGJCredit.SetFocus
            Else
                On Error Resume Next
                txtGJDebit.SetFocus
            End If
        End If
    End If
End Sub

Sub rsRefresh()
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by ID asc", gconDMIS, adOpenKeyset
End Sub

Sub SearchVoucherNo(XXX As String)
    If XXX <> "" Then
        On Error GoTo ErrorCode
        rsJournal_HD.Bookmark = rsFind(rsJournal_HD.Clone, "voucherno", XXX).Bookmark
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & XXX, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

Sub SendToBack()
    fraFindAccount.ZOrder 1
    cmdFindAccount.ZOrder 1
    fraFindAccount.Visible = False
    cmdFindAccount.Visible = False
End Sub

Sub SendToBackGJ()
    cmdGJEntry.ZOrder 1
    cmdGJEntry.Visible = False
    picGJEntry.ZOrder 1
    picGJEntry.Visible = False
    picGJEntry.Enabled = False
    Picture1.Enabled = True
End Sub

Sub SendToBackTemplates()
    cmdTemplates.ZOrder 1
    picTemplates.ZOrder 1
End Sub

Sub StoreMemVars()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        labID.Caption = rsJournal_HD!ID
        txtJNo.Text = Null2String(rsJournal_HD!JNo)
        txtVoucherNo.Text = Null2String(rsJournal_HD!VOUCHERNO)
        txtJDate.Text = Format(Null2String(rsJournal_HD!JDATE), "DD-MMM-YY")

        If xJOURNALTYPE = "OPB" Then txtParticulars2.Locked = True

        txtParticulars2.Text = Null2String(rsJournal_HD!remarks)

        If Null2String(rsJournal_HD!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "*** CANCELLED *** [" & Null2String(rsJournal_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = False
            picDatePosted.Visible = False
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            labPosted.Visible = True
            labPosted.Caption = "*** POSTED *** [" & Null2String(rsJournal_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPrint.Enabled = True
            picDatePosted.Visible = True
            If LOGLEVEL = "ADM" Then cmdUnPost.Enabled = True Else cmdUnPost.Enabled = False
            If Null2String(rsJournal_HD!DATEPOSTED) = "" Then
                lblDatePosted = Format(rsJournal_HD!JDATE, "DD-MMM-YY")
            Else
                lblDatePosted = Format(rsJournal_HD!DATEPOSTED, "DD-MMM-YY")
            End If
        Else
            labPosted.Caption = ""
            labPosted.Visible = False
            cmdEdit.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = True
            cmdPost.Enabled = True
            cmdPrint.Enabled = False
            picDatePosted.Visible = False
        End If
        FillDetails
    Else
        MsgBox "No Such Record!": If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub StoreSearch(XXX As Variant)
    rsRefresh
    rsJournal_HD.Find "VoucherNo = " & N2Str2Null(XXX)
    StoreMemVars
End Sub

Private Sub cboATC2_Click()
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC2.Text))
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE2.Text = N2Str2Zero(rsATC!Rate)
    End If
    Set rsATC = Nothing
End Sub

Private Sub cboGJAccountNo_Change()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String
    txtGJAccountName.Text = Setacctname(cboGJAccountNo.Text)
    'DEALER INCOME TAX WITHHELD
    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Then
        DEALER_ITW_COMPENSATION = "21-04000-00"
        DEALER_ITW_EXPANDED = "21-04000-00"
    End If
    If COMPANY_CODE = "HBK" Then
        DEALER_ITW_COMPENSATION = "21-05001-00"
        DEALER_ITW_EXPANDED = "21-05002-00"
    End If
    If COMPANY_CODE = "HGC" Then
        DEALER_ITW_COMPENSATION = "21-05002-00"
        DEALER_ITW_EXPANDED = "21-05003-00"
    End If
    If COMPANY_CODE = "HMH" Then
        DEALER_ITW_COMPENSATION = "21-05002-00"
        DEALER_ITW_EXPANDED = "21-05003-00"
    End If
    '    If cboGJAccountNo.Text = DEALER_ITW_COMPENSATION Or cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
    '        fraATC2.Visible = True: labATC.Visible = True: cboJVSupCust.Visible = True
    '        On Error Resume Next
    '        cboATC2.SetFocus
    '    Else
    fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False
    '    End If
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

'TEMPORARY CODE: FOR EDITING OF VOUCHERNO
    If CDate(txtJDate.Text) <= "12/21/2009" Then
        txtVoucherNo.Enabled = True
    Else
        txtVoucherNo.Enabled = False
    End If
    'TEMPORARY CODE: FOR EDITING OF VOUCHERNO

    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    SendToBackGJ
    SendToBackTemplates
    Dim rsProfile                                           As ADODB.Recordset
    Dim AccountingMonth, AccountingYear                     As Integer
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        AccountingMonth = rsProfile!PERIODMONTH
        AccountingYear = rsProfile!PERIODYEAR
    End If
    'TEMPORARY DISABLED BY FML - 2/15/2008 2:45AM
    '    Dim rsDetails                                      As ADODB.Recordset
    '    Set rsDetails = New ADODB.Recordset
    '    Set rsDetails = gconDMIS.Execute("Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit, VoucherNo from AMIS_Journal_Det Where jtype = '" & xJOURNALTYPE & "' and Month(Jdate) = " & AccountingMonth & " and Year(Jdate) = " & AccountingYear & " and Status <> 'C' group by VoucherNo order by VoucherNo asc")
    '    Dim SQL
    '    SQL = "Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit, VoucherNo from AMIS_Journal_Det Where jtype = '" & xJOURNALTYPE & "' and Month(Jdate) = " & AccountingMonth & " and Year(Jdate) = " & AccountingYear & " and Status <> 'C' group by VoucherNo order by VoucherNo asc"
    '    If Not rsDetails.EOF And Not rsDetails.EOF Then
    '        Screen.MousePointer = 11
    '        Do While Not rsDetails.EOF
    '            If Round(rsDetails!TotalDebit, 2) <> Round(rsDetails!Totalcredit, 2) Then
    '                Screen.MousePointer = 0
    '                MsgBox "TOTAL DEBIT: [" & Round(rsDetails!TotalDebit, 2) & "] TOTAL CREDIT: [" & Round(rsDetails!Totalcredit, 2) & "]" & vbCrLf & _
                     '                       "Warning: " & xJOURNALTYPE & "-" & rsDetails!vOUCHERNO & " is still not balance or has zero details" & vbCrLf & _
                     '                     "              Adding Other Entries is not Allowed!", vbCritical, "Unbalanced Entry"
    '                Exit Sub
    '            End If
    '            rsDetails.MoveNext
    '        Loop
    '        Screen.MousePointer = 0
    '    End If

    AddorEdit = "ADD"

    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    FillDetails

    On Error Resume Next
    'txtVoucherNo.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdAddAccount_Click()
    Screen.MousePointer = 11
    REFRESH_ACCOUNT = True
    frmAMISFILESChartOfAccount.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdCancelCO_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_CancelEntry", LocalAcess) = False Then Exit Sub
    '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If
    If MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Cancel Journal") = vbYes Then
        Screen.MousePointer = 11
        '        If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
        '            With FrmCancelTransaction
        '                .lblTransaction_type = xJOURNALTYPE
        '                .LblTransactionNo = txtVoucherNo.Text
        '                FrmCancelTransaction.Show
        '                If CANCEL_ANS = "NO" Then Exit Sub
        '                Screen.MousePointer = 0
        '            End With
        '        End If
        '        If xJOURNALTYPE = "GJ" Then
        '            With FrmCancelTransaction
        '                .lblTransaction_type = xJOURNALTYPE
        '                .LblTransactionNo = txtVoucherNo.Text
        '                FrmCancelTransaction.Show
        '                If CANCEL_ANS = "NO" Then Exit Sub
        '                Screen.MousePointer = 0
        '            End With
        '        End If
        '        Screen.MousePointer = 0
        '        ' UPDATE DUE TO NEW AUDIT : BTT 08282008
        '        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        '        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        '        If xJOURNALTYPE = "APJ" Then
        '            'Update By BTT 06282008
        '            With FrmCancelTransaction
        '                .lblTransaction_type = xJOURNALTYPE
        '                .LblTransactionNo = txtVoucherNo.Text
        '                FrmCancelTransaction.Show
        '                If CANCEL_ANS = "NO" Then Exit Sub
        '                Screen.MousePointer = 0
        '            End With
        '            SQL_STATEMENT = "update AMIS_PV_Detail set status = 'C' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '            gconDMIS.Execute SQL_STATEMENT
        '            NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        '        End If
        '        If xJOURNALTYPE = "CDJ" Then
        '            'Update By BTT 06282008
        '
        '            With FrmCancelTransaction
        '                .lblTransaction_type = xJOURNALTYPE
        '                .LblTransactionNo = txtVoucherNo.Text
        '                FrmCancelTransaction.Show
        '                If CANCEL_ANS = "NO" Then Exit Sub
        '                Screen.MousePointer = 0
        '            End With
        '
        '            Set rsCV_Detail = New ADODB.Recordset
        '            Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where Jtype = 'APJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
        '            If rsCV_Detail.EOF And rsCV_Detail.BOF Then
        '                Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where Jtype = 'VPJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
        '            End If
        '            If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
        '                SQL_STATEMENT = "update AMIS_CV_Detail set status = 'C' where jtype = " & N2Str2Null(rsCV_Detail!jtype) & " and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                gconDMIS.Execute SQL_STATEMENT
        '                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        '            End If
        '        End If
        '        If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        '            'Update By BTT 06282008
        '            With FrmCancelTransaction
        '                .lblTransaction_type = xJOURNALTYPE
        '                .LblTransactionNo = txtVoucherNo.Text
        '                FrmCancelTransaction.Show
        '                If CANCEL_ANS = "NO" Then Exit Sub
        '                Screen.MousePointer = 0
        '            End With
        '            Set rsCRJ_Detail = New ADODB.Recordset
        '            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail Where VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
        '            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
        '                SQL_STATEMENT = "update AMIS_CV_Detail set status = 'C' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                gconDMIS.Execute SQL_STATEMENT
        '                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        '            End If
        '
        '        End If
        'UPDATED BY: ACL
        'DESCRIPTION: CONFIRMATION OF CANCELLED TRANSACTION
        If xJOURNALTYPE = "OPB" Then
            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                'Call FrmCancelTransaction.LoadJournal(xJOURNALTYPE)
                FrmCancelTransaction.Show 1
            End With

            If CANCEL_ANS = "NO" Then Exit Sub
            Screen.MousePointer = 0

            SQL_STATEMENT = "update AMIS_Journal_HD set status = 'C',USERCODE='" & LOGCODE & "',DATECANCELLED='" & LOGDATE & "' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "C", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

            SQL_STATEMENT = "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "C", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If
        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        Screen.MousePointer = 0
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub

    'TEMPORARY CODE: FOR EDITING OF VOUCHERNO
    '        If CDate(txtJDate.Text) <= "12/21/2009" Then
    '            txtVoucherNo.Enabled = True
    '        Else
    '            txtVoucherNo.Enabled = False
    '        End If
    'TEMPORARY CODE: FOR EDITING OF VOUCHERNO

    AddorEdit = "EDIT"
    PrevJType = UCase(xJOURNALTYPE)
    PrevJNo = Format(txtJNo.Text, "000000")
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_HD!ID
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Or xJOURNALTYPE = "PDJ" Or xJOURNALTYPE = "CLO" Then txtParticulars2.Locked = False
    On Error Resume Next
    'txtVoucherNo.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFind_Click()
    On Error GoTo ErrorCode:

    frmAMISSearchGJ.Show vbModal
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdGJCancel_Click()
    SendToBackGJ
    StoreMemVars
End Sub

Private Sub cmdGJDelete_Click()

    If labGJID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If

    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Journal_Det where id = " & labGJID.Caption
        Screen.MousePointer = 11
    Else
        Exit Sub
    End If

    Dim cnt                                                 As Integer
    Dim rsJournalDup                                        As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemNo,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update AMIS_Journal_Det set JItemNo = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            rsJournalDup.MoveNext
            DoEvents
        Loop
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    'cmdGJCancel.Value = True
    Call cmdGJCancel_Click

    If lstGJ.ListItems.Count > 0 And lstGJ.Enabled = True Then
        lstGJ.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdGJEntry_Click()
    SendToBackGJ
    cmdGJEntry.Visible = True: cmdGJEntry.ZOrder 0
    picGJEntry.Visible = True: picGJEntry.ZOrder 0
    picGJEntry.Enabled = True: cmdGJDelete.Visible = False
    AddorEdit = "ADD"
    InitGJ
    On Error Resume Next
    cboGJAccountNo.SetFocus
End Sub

Private Sub cmdGJSave_Click()
    On Error GoTo ErrorCode

    Dim str_MSG                                             As String


    str_MSG = "Error in saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If JournalEntries = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Opening Balance Entry")
        MsgBox str_MSG, vbCritical, "Journal Entry Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Function JournalEntries() As Boolean
    On Error GoTo ErrorCode

    If cboGJAccountNo.Text = "" Then
        MsgBox "Account Code must have a value", vbInformation, "Error Encountered!"
        JournalEntries = True
        Exit Function
    End If

    If NumericVal(txtGJDebit.Text) > 0 And NumericVal(txtGJCredit.Text) > 0 Then
        MsgBox "Invalid Journal Entry! Debit and Credit Amount can not have both Amount!", vbCritical, "Invalid Entry!"
        JournalEntries = True
        Exit Function
    End If

    'If AddorEdit = "ADD" Then
    '    Dim rsJournal_DetClone                         As ADODB.Recordset
    '    Set rsJournal_DetClone = New ADODB.Recordset
    '    rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(xJOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
    '    If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
    '        MsgBox "Account Code already used in this transaction", vbInformation, "Error in Part Number Validation"
    '        exit function
    '    End If
    'End If


    If xJOURNALTYPE = "GJ" Then
        If Left(cboGJAccountNo.Text, 5) = "11-02" Or Left(cboGJAccountNo.Text, 5) = "11-03" Then
            If MsgBox("A/R Codes must have a DM/CM to update the Customer Subsidiary" & vbCrLf & " or use DM/CM for A/R Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/R schedule") = vbYes Then
                'save in audit trail
                MsgBox "Reminder: You must use DM/CM to update A/R Schedule", vbInformation, "Confirmation Logged in Audit Trail."
            Else
                JournalEntries = True
                Exit Function
            End If
        End If
        If Left(cboGJAccountNo.Text, 5) = "21-01" Then
            If MsgBox("A/P Codes must have a DM/CM to update the Vendors Subsidiary" & vbCrLf & " or use DM/CM for A/P Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/P schedule") = vbYes Then
                'save in audit trail
                MsgBox "Reminder: You must use DM/CM to update A/P Schedule", vbInformation, "Confirmation Logged in Audit Trail."
            Else
                JournalEntries = True
                Exit Function
            End If
        End If
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                       As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                     As String
    Dim J_DEBIT, J_CREDIT, J_TAX                            As Double
    Dim J_STATUS, J_JITEMNO                                 As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_JITEMNO = N2Str2Null(txtGJItemNo.Text)
    J_ACCT_CODE = N2Str2Null(cboGJAccountNo.Text)
    J_ACCT_NAME = N2Str2Null(txtGJAccountName.Text)
    J_DEBIT = Round(NumericVal(txtGJDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtGJCredit.Text), 2)

    J_STATUS = "'N'"

    Dim J_SUPCODE, J_ATC                                    As String
    Dim J_RATE, J_TAXBASE                                   As Double
    If cboGJAccountNo.Text = "21-04001-00" Or cboGJAccountNo.Text = "21-04002-00" Then
        J_SUPCODE = N2Str2Null(SetVendorCode(cboJVSupCust.Text))
        J_ATC = N2Str2Null(cboATC2.Text)
        J_RATE = NumericVal(txtRATE2.Text)
        J_TAXBASE = NumericVal(txtTaxBase2.Text)
    Else
        J_SUPCODE = "'999999'"
        J_ATC = "NULL"
        J_RATE = 0
        J_TAXBASE = 0
    End If
    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        If txtGJAccountParticulars.Text <> "" And txtGJAccountParticulars.Text <> "Pls Type Your Remarks Here!" Then
            gconDMIS.Execute "insert into AMIS_JV_Detail " & _
                             "(JNo,VoucherNo,itemno,Particulars,status)" & _
                             " values (" & J_JNO & ", " & J_VOUCHERNO & ", " & J_JITEMNO & _
                             ", " & N2Str2Null(txtGJAccountParticulars.Text) & _
                             ", " & J_STATUS & ")"
        End If
        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                         ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"

        NEW_LogAudit "AA", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        gconDMIS.Execute "update AMIS_Journal_Det set" & _
                         " jdate = " & J_JDATE & "," & _
                         " voucherno = " & J_VOUCHERNO & "," & _
                         " jtype = " & J_JTYPE & "," & _
                         " jno = " & J_JNO & "," & _
                         " jitemno = " & J_JITEMNO & "," & _
                         " acct_code = " & J_ACCT_CODE & "," & _
                         " acct_name = " & J_ACCT_NAME & "," & _
                         " debit = " & J_DEBIT & "," & _
                         " credit = " & J_CREDIT & "," & _
                         " tax = " & J_TAX & "," & _
                         " ATC = " & J_ATC & "," & _
                         " RATE = " & J_RATE & "," & _
                         " TAXBASE = " & J_TAXBASE & "," & _
                         " status = " & J_STATUS & _
                         " where id = " & labGJID.Caption
        gconDMIS.Execute "update AMIS_JV_Detail set" & _
                         " Particulars = " & N2Str2Null(txtGJAccountParticulars.Text) & _
                         " where JNo = " & J_JNO & " and ItemNo = " & J_JITEMNO
        NEW_LogAudit "EE", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " VENDORCODE = " & J_SUPCODE & "," & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    Picture1.Enabled = True
    If AddorEdit = "ADD" Then cmdGJEntry_Click Else cmdGJCancel_Click

    JournalEntries = True
    Exit Function
ErrorCode:
    JournalEntries = False
End Function

'Upating Code       : AXP-0713200713:18
Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:
    rsJournal_HD.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MoveNext
    If rsJournal_HD.EOF Then
        rsJournal_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPost_Click()
    On Error GoTo ErrorCode:

    Dim str_MSG                                             As String


    str_MSG = "Error in Posting @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If JournalPosting = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Opening Balance")
        MsgBox str_MSG, vbCritical, "Posting Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    '    SaveLogFile
    ShowVBError
End Sub

Function JournalPosting() As Boolean
    On Error GoTo ErrorCode


    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then
        JournalPosting = True
        Exit Function
    End If
    If MsgBox("Are you sure you want to Post this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then




        'UPDATED BY: JUN --- DATE UPDATED: 18-19-2009 --- DESCRIPTION: VALIDATE CREDIT AND DEBIT IN LINE ITEM BOTH ARE ZERO
        If COMPANY_CODE <> "HPI" Then
            Dim rsZERO                                      As ADODB.Recordset
            Set rsZERO = New ADODB.Recordset
            rsZERO.Open "Select ACCT_NAME,JITEMNO,DEBIT,CREDIT from AMIS_JOURNAL_DET WHERE DEBIT = 0 AND CREDIT = 0 AND VOUCHERNO = '" & txtVoucherNo.Text & "' AND JTYPE = '" & xJOURNALTYPE & "'", gconDMIS, adOpenKeyset
            If Not rsZERO.EOF And Not rsZERO.BOF Then
                MessagePop InfoFriend, "INFORMATION", "You can't POST this transaction both debit and credit is ZERO." & " " & "LINE #-" & " " & Null2String(rsZERO!jitemno) & "" & " and " & "ACCT NAME-" & " " & "" & Null2String(rsZERO!acct_Name) & ""
                JournalPosting = True
                Exit Function
            End If
            Set rsZERO = Nothing
        End If
        'UPDATED BY: JUN

        '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
        '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
        '        exit function
        '    End If
        'Update by BTT : to update the balance of the SJ
        'If xJOURNALTYPE = "CRJ" Then
        '    UpdateBalanceSJ txtVoucherNo.Text, True
        'End If

        Dim rsCheckDetails                                  As ADODB.Recordset
        Dim rsCheckCRJDetails                               As ADODB.Recordset
        Dim TotalCRJ_Credit                                 As Double

        Screen.MousePointer = 11



        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',USERCODE='" & LOGCODE & "',DATEPOSTED='" & LOGDATE & "' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        'UPDATED BY: JUN --- DATE UPDATED: 05-30-2009 --- DESCRIPTION: VALIDATE IF ALL ENTRY IN AMIS_JOURNAL_DET WAS TAG AS POSTED IF NOT UPDATE THE STATUS INTO POSTED
        Dim rsCHECK_POSTED                                  As ADODB.Recordset
        Set rsCHECK_POSTED = gconDMIS.Execute("SELECT STATUS FROM AMIS_JOURNAL_DET where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " AND STATUS <> 'P'")
        If Not rsCHECK_POSTED.EOF And Not rsCHECK_POSTED.BOF Then
            gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        Else
            'ALL DETAILS ARE POSTED
        End If
        Set rsCHECK_POSTED = Nothing
        'UPDATED BY: JUN---------------------------------------------------------------------------------------------------------------------------------------------------------------------------



        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
    End If

    JournalPosting = True
    Exit Function
ErrorCode:
    JournalPosting = False
End Function

'Upating Code       : AXP-0713200713:18
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MovePrevious
    If rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdPrint_Click()
    Dim Ans                                                 As String
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Print", LocalAcess) = False Then Exit Sub

    Ans = MsgBox("Are you sure do you want to print this Transaction?", vbQuestion + vbYesNo, "Print Transaction")
    If Ans = vbYes Then

        'For Reprint Routin : Update by BTT

        If xJOURNALTYPE = "OPB" Then ShowReport "AccountsOpenningBalance", "OpeningBalance", "{Amis_journal_hd.VOUCHERNO} = '" & txtVoucherNo.Text & "' and {Amis_journal_hd.jtype} = 'OPB'", "OPENING BALANCE", LOGDATE, False
        NEW_LogAudit "PX", "OPENING BALANCE", "", "", "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup, rsProfile                                As ADODB.Recordset

    If IsNull(txtJNo.Text) = True Then
        MessagePop RecSaveError, "Error!", "Journal No. must not be empty"
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' and jno = '" & txtJNo.Text & "' order by jtype,jno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MessagePop RecSaveError, "Error!", "Journal No. already exist!"
                Exit Sub
            End If

        End If
    End If
    If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "Error"
        Exit Sub
    End If
    If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "OPB" And xJOURNALTYPE <> "PDJ" Then
        If COMPANY_CODE = "HMH" Then
            'Updated by: ACL 10202009
            If CheckIfOpen(xJOURNALTYPE, Trim(txtJDate.Text), Year(txtJDate.Text)) = False Then
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                Exit Sub
            End If
        Else
            Set rsProfile = New ADODB.Recordset
            Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
            If Not rsProfile.EOF And Not rsProfile.BOF Then
                If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
                    If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
                        MessagePop RecSaveError, "Error!", "Warning: Journal Date is not valid in Accounting Period!"
                        'MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
                        Exit Sub
                    End If
                Else
                    MessagePop RecSaveError, "Error!", "Warning: Journal Date is not valid in Accounting Period!"
                    'MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
                    Exit Sub
                End If
            End If
        End If
    End If
    '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                       As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE      As String
    Dim J_CUSTOMERNAME                                      As String
    Dim J_DEBIT, J_CREDIT, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_CHECKNO                                 As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_CHECKDATE, J_BANKCODE                             As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS, J_DEALER                                   As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(Format(txtVoucherNo.Text, "000000"))
    J_JTYPE = N2Str2Null(xJOURNALTYPE)


    J_AMOUNTTOPAY = 0
    J_INVOICEDATE = "NULL"
    J_BALANCE = 0
    J_AMOUNTPAID = 0

    J_DUEDATE = "NULL"

    J_PAYTYPE = "NULL"

    J_JNO = N2Str2Null(txtJNo.Text)





    J_STATUS = "'N'"

    J_CHECKNO = "NULL"

    J_TERMS = "NULL"
    J_DEALER = "NULL"


    J_CHECKDATE = "NULL"

    J_BANKCODE = "NULL"

    J_CUSTOMERNAME = "NULL"

    J_VENDORCODE = "'999999'"
    If xJOURNALTYPE = "OPB" Then
        J_CUSTOMERCODE = "'999999'"
        J_CUSTOMERNAME = "NULL"
    End If


    J_INVOICETYPE = "NULL"




    J_INVOICENO = "NULL"

    J_DEBIT = "NULL"
    J_CREDIT = "NULL"
    J_OUTBALANCE = "NULL"
    J_INVOICEAMT = "NULL"
    J_REFNO = "NULL"
    J_REFDATE = "NULL"

    If Trim(txtParticulars2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars2.Text))
    J_PAIDSTATUS = "'N'"
    J_RECEIVESTATUS = "'N'"

    If AddorEdit = "ADD" Then
        Dim rsJournal_HDDup                                 As ADODB.Recordset
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
        J_JNO = N2Str2Null(txtJNo.Text)
        J_VOUCHERNO = N2Str2Null(GetVoucherNo(xJOURNALTYPE))
        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                        " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "')"
        gconDMIS.Execute SQL_STATEMENT

        labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        NEW_LogAudit "A", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " jdate = " & J_JDATE & "," & _
                        " voucherno = " & J_VOUCHERNO & "," & _
                        " jtype = " & J_JTYPE & "," & _
                        " vendorcode = " & J_VENDORCODE & "," & _
                        " customercode = " & J_CUSTOMERCODE & ", customername = " & J_CUSTOMERNAME & "," & _
                        " invoicedate = " & J_INVOICEDATE & "," & _
                        " invoicetype = " & J_INVOICETYPE & "," & _
                        " invoiceno = " & J_INVOICENO & "," & _
                        " invoiceamt = " & J_INVOICEAMT & "," & _
                        " duedate = " & J_DUEDATE & "," & _
                        " paytype = " & J_PAYTYPE & "," & _
                        " refno = " & J_REFNO & "," & _
                        " refdate = " & J_REFDATE & ", terms = " & J_TERMS & ", dealer = " & J_DEALER & "," & _
                        " amounttopay = " & J_AMOUNTTOPAY & ", Balance = " & J_BALANCE & ", AmountPaid = " & J_AMOUNTPAID & "," & _
                        " jno = " & J_JNO & "," & _
                        " debit = " & J_DEBIT & "," & _
                        " credit = " & J_CREDIT & "," & _
                        " outbalance = " & J_OUTBALANCE & "," & _
                        " CheckNo = " & J_CHECKNO & ", " & _
                        " CheckDate = " & J_CHECKDATE & ", " & _
                        " BankCode = " & J_BANKCODE & ", " & _
                        " status = " & J_STATUS & ", PaidStatus = " & J_PAIDSTATUS & ", ReceiveStatus = " & J_RECEIVESTATUS & "," & _
                        " remarks = " & J_REMARKS & ", USERCODE = '" & LOGCODE & "', LASTUPDATE = '" & LOGDATE & "'" & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT

        labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        NEW_LogAudit "E", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        CheckIfthereISCDJ txtVoucherNo
        SQL_STATEMENT = "update AMIS_Journal_Det set" & _
                        " jtype = " & J_JTYPE & "," & _
                        " jdate = " & J_JDATE & "," & _
                        " USERCODE = '" & LOGCODE & "'," & _
                        " LASTUPDATE = '" & LOGDATE & "'," & _
                        " jno = " & J_JNO & _
                        " where jtype = '" & PrevJType & "' and jno = '" & PrevJNo & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    If AddorEdit <> "ADD" Then
        rsJournal_HD.Find "jno = " & J_JNO
        cmdCancel.Value = True
        FillDetails
        If xJOURNALTYPE = "VDJ" Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                            " debit = " & J_DEBIT & "," & _
                            " credit = " & TOTCREDIT & "," & _
                            " outbalance = " & OUTBALANCE & _
                            " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "OPENING BALANCE AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        ElseIf xJOURNALTYPE = "VCJ" Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" _
                            & " debit = " & TOTDEBIT & "," _
                            & " credit = " & J_CREDIT & "," _
                            & " outbalance = " & OUTBALANCE _
                            & " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "OPENING BALANCE AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        Else
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                            " debit = " & TOTDEBIT & "," & _
                            " credit = " & TOTCREDIT & "," & _
                            " tax = " & TOTTAX & "," & _
                            " outbalance = " & OUTBALANCE & _
                            " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "OPENING BALANCE AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If
    End If
    rsRefresh
    rsJournal_HD.Find "jno = " & J_JNO
    cmdCancel.Value = True

    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdUnPost_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub
    If MsgBox("Are you sure you want to Unpost this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then
        If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "OPB" Then
            '            If COMPANY_CODE = "HPI" Then
            'Updated by: ACL 10202009
            If CheckIfOpen(xJOURNALTYPE, Trim(txtJDate.Text), Year(txtJDate.Text)) = False Then
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                Exit Sub
            End If
            '            Else
            '                Set rsProfile = New ADODB.Recordset
            '                Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
            '                If Not rsProfile.EOF And Not rsProfile.BOF Then
            '                    If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
            '                        If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
            '                            MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                            Exit Sub
            '                        End If
            '                    Else
            '                        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                        Exit Sub
            '                    End If
            '                End If
            '            End If
        End If
        '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
        '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
        '        Exit Sub
        '    End If



        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        LogAudit "U", "OPENING BALANCE", txtJNo
        NEW_LogAudit "U", "OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub
Function VALIDATE_UNPOSTING() As Boolean
    Dim rsVALIDATE_UNPOSTING                                As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim rsGET_CRJ_VOUCHERNO                                 As ADODB.Recordset
    xVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
    Set rsVALIDATE_UNPOSTING = New ADODB.Recordset

    rsVALIDATE_UNPOSTING.Open "SELECT X.COMBI_LINK,X.CUSTOMERCODE,X.INVOICETYPE,X.INVOICENO,X.ACCOUNT_CODE FROM " & _
                              "( " & _
                              "SELECT RTRIM(LTRIM(CUSTOMERCODE)) + '-' + RTRIM(LTRIM(INVOICETYPE)) + '-' + RTRIM(LTRIM(INVOICENO)) + '-' + RTRIM(LTRIM(ACCOUNT_CODE)) AS COMBI_LINK, " & _
                              "RTRIM(LTRIM(CUSTOMERCODE)) AS CUSTOMERCODE,RTRIM(LTRIM(INVOICETYPE)) AS INVOICETYPE,RTRIM(LTRIM(INVOICENO)) AS INVOICENO,RTRIM(LTRIM(ACCOUNT_CODE)) AS ACCOUNT_CODE From AMIS_AR WHERE SJVOUCHERNO = '" & xVOUCHERNO & "' " & _
                              ")X  WHERE  X.COMBI_LINK IN(SELECT RTRIM(LTRIM(CUSTOMERCODE)) + '-' + RTRIM(LTRIM(INVOICETYPE)) + '-' + RTRIM(LTRIM(INVOICENO)) + '-' + RTRIM(LTRIM(ACCT_CODE)) FROM AMIS_DETAIL)", gconDMIS, adOpenKeyset
    If Not rsVALIDATE_UNPOSTING.EOF And Not rsVALIDATE_UNPOSTING.BOF Then
        Set rsGET_CRJ_VOUCHERNO = New ADODB.Recordset
        rsGET_CRJ_VOUCHERNO.Open "SELECT JTYPE,VOUCHERNO FROM AMIS_DETAIL WHERE CUSTOMERCODE = " & N2Str2Null(rsVALIDATE_UNPOSTING!CustomerCode) & " AND INVOICENO = " & N2Str2Null(rsVALIDATE_UNPOSTING!INVOICENO) & " AND INVOICETYPE = " & N2Str2Null(rsVALIDATE_UNPOSTING!InvoiceType) & " AND ACCT_CODE = " & N2Str2Null(rsVALIDATE_UNPOSTING!Account_code) & " ", gconDMIS, adOpenKeyset
        If Not rsGET_CRJ_VOUCHERNO.EOF And Not rsGET_CRJ_VOUCHERNO.BOF Then
            MessagePop InfoFriend, "INFORMATION", "You can't un-post this voucher it has a payment please see Cash Receipts Journal " & "" & Null2String(rsGET_CRJ_VOUCHERNO!JTYPE) & "" & " - " & "" & Null2String(rsGET_CRJ_VOUCHERNO!VOUCHERNO) & ""
            VALIDATE_UNPOSTING = True
        End If
        Set rsGET_CRJ_VOUCHERNO = Nothing
    Else
        VALIDATE_UNPOSTING = False
    End If
    Set rsVALIDATE_UNPOSTING = Nothing
End Function
Sub UNPOST_CRJ()
    Dim rsUNPOST_CRJ                                        As ADODB.Recordset
    Dim rsIS_IN_AR                                          As ADODB.Recordset
    Set rsUNPOST_CRJ = New ADODB.Recordset
    rsUNPOST_CRJ.Open "SELECT INVOICENO,INVOICETYPE,CUSTOMERCODE,J_CLASS,VOUCHERNO,CR_TYPE FROM AMIS_CRJ_DETAIL WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
    If Not rsUNPOST_CRJ.EOF And Not rsUNPOST_CRJ.BOF Then
        Do While Not rsUNPOST_CRJ.EOF
            Set rsIS_IN_AR = New ADODB.Recordset
            rsIS_IN_AR.Open "SELECT * FROM AMIS_AR WHERE INVOICENO = '" & rsUNPOST_CRJ!INVOICENO & "' AND INVOICETYPE = '" & rsUNPOST_CRJ!InvoiceType & "' AND ACCOUNT_CODE = '" & rsUNPOST_CRJ!J_CLASS & "' AND CUSTOMERCODE = '" & rsUNPOST_CRJ!CustomerCode & "' ", gconDMIS, adOpenKeyset
            If Not rsIS_IN_AR.EOF And Not rsIS_IN_AR.BOF Then
                gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = 0 , BALANCE = " & NumericVal(rsIS_IN_AR!AMOUNT_TOPAY) & " WHERE INVOICENO = '" & rsUNPOST_CRJ!INVOICENO & "' AND INVOICETYPE = '" & rsUNPOST_CRJ!InvoiceType & "' AND ACCOUNT_CODE = '" & rsUNPOST_CRJ!J_CLASS & "' AND CUSTOMERCODE = '" & rsUNPOST_CRJ!CustomerCode & "'"
                gconDMIS.Execute "DELETE FROM AMIS_DETAIL WHERE INVOICENO = '" & rsUNPOST_CRJ!INVOICENO & "' AND INVOICETYPE = '" & rsUNPOST_CRJ!InvoiceType & "' AND ACCT_CODE = '" & rsUNPOST_CRJ!J_CLASS & "' AND CUSTOMERCODE = '" & rsUNPOST_CRJ!CustomerCode & "' AND VOUCHERNO = '" & rsUNPOST_CRJ!VOUCHERNO & "' AND JTYPE = '" & rsUNPOST_CRJ!CR_type & "'"
            Else
                gconDMIS.Execute "DELETE FROM AMIS_DETAIL WHERE INVOICENO = '" & rsUNPOST_CRJ!INVOICENO & "' AND INVOICETYPE = '" & rsUNPOST_CRJ!InvoiceType & "' AND ACCT_CODE = '" & rsUNPOST_CRJ!J_CLASS & "' AND CUSTOMERCODE = '" & rsUNPOST_CRJ!CustomerCode & "' AND VOUCHERNO = '" & rsUNPOST_CRJ!VOUCHERNO & "' AND JTYPE = '" & rsUNPOST_CRJ!CR_type & "'"
            End If
            rsUNPOST_CRJ.MoveNext
        Loop
    End If
    Set rsUNPOST_CRJ = Nothing
End Sub

Private Sub FillGrid()
    Dim rsChartAccount2                                     As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccount2 = New ADODB.Recordset
    Set rsChartAccount2 = gconDMIS.Execute("select acctcode,UPPER(Description),Accttype,ID from AMIS_ChartAccount order by acctcode asc")
    If Not (rsChartAccount2.EOF And rsChartAccount2.BOF) Then
        lstAccounts.Enabled = True
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccount2
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If

End Sub

Private Sub Command1_Click()
    If Module_Access(LOGID, "SYSTEM SETUP", "SYSTEM") = False Then Exit Sub
    frmAMISProfile.Show
End Sub

Private Sub Command3_Click()
' update by BTT 2/3/2009
'    If xJOURNALTYPE = "SJ" Then
'
'        ReturnInvoiceNo txtVoucherNo, xJOURNALTYPE
'        With frmAMIS_Payment
'            frmAMIS_Payment.FillPaymentdetail AMIS_Invoiceno, AMIS_Invoicetype
'            frmAMIS_Payment.Show
'        End With
'    End If
'    If xJOURNALTYPE = "APJ" Then
'        With frmAMIS_Payment
'            frmAMIS_Payment.FillPaymentdetail txtVoucherNo, ""
'            frmAMIS_Payment.Show
'        End With
'    End If
'    If xJOURNALTYPE = "CRJ" Then
'
'        OR_NUMBER_GLOBAL = txtInvoiceNo.Text
'        frmORPaymentDetail.Show vbModal
'    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorCode

    Select Case KeyCode
    Case vbKeyF1 And Shift = 1:
        If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
        Unload frmALL_AuditInquiry
        frmALL_AuditInquiry.Show
        frmALL_AuditInquiry.ZOrder 0
        frmALL_AuditInquiry.Caption = Me.Caption
        Call frmALL_AuditInquiry.DisplayHistory(labID, xJOURNALTYPE)

    Case vbKeyReturn
        If Me.ActiveControl.Name = "cboGJAccountNo" And cboGJAccountNo.Text = "" Then
            fraFindAccount.Visible = True
            cmdFindAccount.Visible = True
            cmdFindAccount.ZOrder 0
            fraFindAccount.ZOrder 0
            fraFindAccount.Enabled = True
            DoEvents
            txtSearch.SetFocus
        ElseIf Me.ActiveControl.Name = "cboAccount" Then
            OkAccount
        ElseIf Me.ActiveControl.Name = "txtGJCredit" And SetAcctType(cboGJAccountNo.Text) = "C" And Val(txtGJCredit.Text) <= 0 And Val(txtGJDebit.Text) <= 0 Then
            On Error Resume Next
            txtGJCredit.SetFocus
        ElseIf Me.ActiveControl.Name = "txtGJDebit" And SetAcctType(cboGJAccountNo.Text) = "D" And Val(txtGJDebit.Text) <= 0 And Val(txtGJCredit.Text) <= 0 Then
            On Error Resume Next
            txtGJDebit.SetFocus
        Else
            MoveKeyPress KeyCode
        End If
    Case vbKeyEscape
        If fraFindAccount.Visible = True Then
            If Me.ActiveControl.Name = "txtSearch" Then
                SendToBack
                SendToBackGJ
                SendToBackTemplates
                StoreMemVars
            Else
                txtSearch.SetFocus
            End If
        Else
            If Picture1.Visible = True Then
                If Me.ActiveControl.Name = "txtSearchTemplates" Then
                    SendToBackGJ
                    SendToBackTemplates
                    StoreMemVars
                ElseIf Me.ActiveControl.Name = "lstTemplates" Then
                    On Error Resume Next

                    txtSearchTemplates.SetFocus
                Else
                    SendToBackGJ
                    SendToBackTemplates
                    StoreMemVars
                End If
                Picture1.Enabled = True
            End If
        End If
    Case vbKeyF3
        If Picture1.Visible = True Then
            If Null2String(rsJournal_HD!Status) = "C" Then
                'MsgBox "Journals are Already Cancelled" & vbCrLf & _
                 "and cannot be Change", vbInformation, "Edit Not Allowed!"
                MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"
            ElseIf Null2String(rsJournal_HD!Status) = "P" Then
                'MsgBox "Journals are Already Posted" & vbCrLf & _
                 "and cannot be Change", vbInformation, "Edit Not Allowed!"
                MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"
            Else
                If xJOURNALTYPE = "OPB" Then
                    Picture1.Enabled = False
                    cmdGJEntry_Click
                End If
            End If
        End If
    Case vbKeyF4
        If xJOURNALTYPE <> "SJ" Then
            If Picture1.Visible = True Then
                If Null2String(rsJournal_HD!Status) = "C" Then
                    MsgBox "Journals are Already Cancelled" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                ElseIf Null2String(rsJournal_HD!Status) = "P" Then
                    MsgBox "Journals are Already Posted" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                Else
                    Picture1.Enabled = False
                End If
            End If
        End If
    Case vbKeyF5
        If cmdPost.Enabled = True Then
            cmdPost.Value = True
        End If
    Case vbKeyF6
        If cmdUnPost.Enabled = True Then
            cmdUnPost.Value = True
        End If
    Case vbKeyF7
        If cmdCancelCO.Enabled = True Then
            cmdCancelCO.Value = True
        End If
    Case vbKeyF8
        If SearchBy = "NAME" Then
            SearchBy = "CODE": fraFindAccount.Caption = "Search Accounts by Account Code"
        Else
            SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
        End If
    Case vbKeyF9
        If Picture1.Visible = True Then
            If Null2String(rsJournal_HD!Status) = "C" Then
                MsgBox "Journals are Already Cancelled" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            ElseIf Null2String(rsJournal_HD!Status) = "P" Then
                MsgBox "Journals are Already Posted" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            Else
                fraFindAccount.ZOrder 1: cmdFindAccount.ZOrder 1
                fraFindAccount.Visible = False: cmdFindAccount.Visible = False: BringToFrontTemplates
                txtSearchTemplates.SetFocus
            End If
        End If
    Case vbKeyF11
        SendToBackGJ
        SendToBackTemplates

        On Error Resume Next

    Case vbKeyF12
        '        If Null2String(rsJournal_HD!Status) = "C" Then
        '
        '            If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub
        '
        '            If MsgBox("Are you sure you want to Un-Cancel this Transaction?", vbQuestion + vbYesNo, "Un-Cancel Journal") = vbYes Then
        '                Screen.MousePointer = 11
        '                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                rsRefresh
        '                rsJournal_HD.Find "id = " & labID.Caption
        '                StoreMemVars
        '                Screen.MousePointer = 0
        '            End If
        '        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
    If Shift = 1 Then
        If KeyCode = vbKeyF1 Then
            'If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Journal Entry)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "OPENING BALANCE")
        End If
    End If
    If Shift = 2 Then
        If KeyCode = vbKeyA Then cmdAddAccount_Click


        If KeyCode = vbKeyF12 Then

        End If
    End If
    Exit Sub

ErrorCode:
    ShowErrMsg
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBackGJ: SendToBackTemplates
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    'Frame1.Top = 90
    fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False


    If xJOURNALTYPE = "OPB" Then
        LocalAcess = "ACCOUNT OPENING BALANCE"
        Label3.Caption = "Ref. No.": Label5.Caption = "Ref. Date"
        Me.Caption = "OPENING BALANCES"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
    End If

    initMemvars
    txtSearch.Text = "": txtSearchTemplates.Text = ""
    rsRefresh
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveLast
    End If
    'If xJOURNALTYPE = "SJ" Then picInvoiceDet.Visible = True Else picInvoiceDet.Visible = False
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
    LocalAcess = ""
End Sub

Private Sub lstAccounts_DblClick()
    labAccountCode.Caption = lstAccounts.SelectedItem
    OkAccount
End Sub

Private Sub lstAccounts_GotFocus()
    On Error Resume Next
    labAccountCode.Caption = lstAccounts.SelectedItem
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labAccountCode.Caption = Item
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        labAccountCode.Caption = lstAccounts.SelectedItem
        OkAccount
    End If
End Sub

Private Sub lstDetails_DblClick()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"

            'MsgBox "Transactions are Already Cancelled" & vbCrLf & _
             '       "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"

            'MsgBox "Journals are Already Posted" & vbCrLf & _
             '       "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            If kcnt > 0 Then
                AddorEdit = "EDIT"




                On Error Resume Next
                'txtGrossAmt.SetFocus
                OkAccountSetCursor
            End If
        End If
    End If
End Sub

Private Sub lstDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstDetails_DblClick
        If Me.ActiveControl.Name = "txtDebit" Then SendKeys MOVEUP
    End If
End Sub

Private Sub lstGJ_DblClick()
    If Null2String(rsJournal_HD!Status) = "C" Then
        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(rsJournal_HD!Status) = "P" Then
        MsgBox "Journals are Already Posted" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
    Else
        AddorEdit = "EDIT"
        cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub lstGJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstGJ_DblClick
End Sub

Private Sub lstGJ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        AddorEdit = "EDIT"
        cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
        cmdGJDelete_Click
    End If
End Sub

Private Sub lstTemplates_DblClick()
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub lstTemplates_KeyPress(KeyAscii As Integer)
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    If KeyAscii = 13 Then InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub ShortcutCaption2_GotFocus()
    If Picture1.Visible = True Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MsgBox "Journals are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else

            If xJOURNALTYPE = "OPB" Then
                Picture1.Enabled = False
                cmdGJEntry_Click
            End If
        End If
    End If
End Sub


Private Sub ShortcutCaption4_GotFocus()
    If Picture1.Visible = True Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MsgBox "Journals are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            fraFindAccount.ZOrder 1: cmdFindAccount.ZOrder 1
            fraFindAccount.Visible = False: cmdFindAccount.Visible = False: BringToFrontTemplates
            txtSearchTemplates.SetFocus
        End If
    End If
End Sub

Private Sub ShortcutCaption5_GotFocus()
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then labPosted.Visible = False Else labPosted.Visible = True
    End If
End Sub

'Private Sub txtCode_Change()
'cboNameofVendor.Text = SetVendorName(txtCode.Text)
'txtAddress.Caption = SetVendorAddress(txtCode.Text)
'End Sub

'Private Sub txtcustcode_Change()
'cboCustName.Text = SetCustomerName(txtcustcode.Text)
'End Sub

Private Sub txtGJAccountName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtGJAccountParticulars_GotFocus()
    If txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!" Then txtGJAccountParticulars.Text = ""
End Sub

Private Sub txtGJAccountParticulars_LostFocus()
    If txtGJAccountParticulars.Text = "" Then txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!"
End Sub

Private Sub txtGJCredit_GotFocus()
'If Val(txtGJCredit.Text) = 0 Then
'   If OUTBALANCE > 0 And TOTDEBIT > 0 Then
'      txtGJCredit.Text = OUTBALANCE
'      txtGJDebit.Text = ZERO
'   Else
'      txtGJCredit.Text = ""
'   End If
'Else
'   txtGJCredit.Text = NumericVal(txtGJCredit.Text)
'End If
    If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "CLO" Then
        If NumericVal(txtGJDebit.Text) = 0 Then
            If Val(txtGJCredit.Text) = 0 Then
                If OUTBALANCE > 0 And TOTDEBIT > 0 Then
                    txtGJCredit.Text = OUTBALANCE
                    txtGJDebit.Text = ZERO
                Else
                    txtGJCredit.Text = ""
                End If
            Else
                txtGJCredit.Text = NumericVal(txtGJCredit.Text)
            End If
        Else
            txtGJCredit.Text = ZERO
        End If
    End If
End Sub

Private Sub txtGJCredit_LostFocus()
    If txtGJCredit.Text = "" Then txtGJCredit.Text = 0
End Sub

Private Sub txtGJDebit_GotFocus()
'If NumericVal(txtGJDebit.Text) = 0 Then
'   If OUTBALANCE > 0 And TOTCREDIT > 0 Then
'      txtGJCredit.Text = ZERO
'      txtGJDebit.Text = OUTBALANCE
'   Else
'      txtGJDebit.Text = ""
'   End If
'Else
'   txtGJDebit.Text = NumericVal(txtGJDebit.Text)
'End If
    If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "CLO" Then
        If NumericVal(txtGJCredit.Text) = 0 Then
            If NumericVal(txtGJDebit.Text) = 0 Then
                If OUTBALANCE > 0 And TOTCREDIT > 0 Then
                    txtGJCredit.Text = ZERO: txtGJDebit.Text = OUTBALANCE
                Else
                    txtGJDebit.Text = ""
                End If
            Else
                txtGJDebit.Text = NumericVal(txtGJDebit.Text)
            End If
        Else
            txtGJDebit.Text = ZERO
        End If
    End If
End Sub

Private Sub txtGJDebit_LostFocus()
    If txtGJDebit.Text = "" Then txtGJDebit.Text = 0
End Sub

Private Sub txtJDate_GotFocus()
    txtJDate.Text = Format(txtJDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtJDate_LostFocus()
    txtJDate.Text = Format(txtJDate.Text, "DD-MMM-YY")
    txtParticulars2.SetFocus
End Sub

Private Sub txtParticulars2_GotFocus()
    If txtParticulars2.Text = "Pls Type Your Message Here!" Then txtParticulars2.Text = ""
End Sub

Private Sub txtParticulars2_LostFocus()
    If txtParticulars2.Text = "" Then txtParticulars2.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstAccounts.SetFocus
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearchTemplates_Change()
    If Trim(txtSearchTemplates.Text) = "" Then
        FillTemplates
    Else
        FillSearchTemplates (txtSearchTemplates.Text)
    End If
End Sub

Private Sub txtSearchTemplates_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If lstTemplates.Enabled = True Then lstTemplates.SetFocus
    End If
End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Format(txtVoucherNo, "000000")
End Sub

Function rsCHECKINVOICENOandTYPE(xInvoiceType As String, xInvoiceNo As String, XCustomerCode As String) As Boolean
'UPDATED BY: JUN --- DATE UPDATED: 10272008 --- DESCRIPTION: VALIDATE INVOICENO AND INVOICETYPE
    Dim rsExist                                             As ADODB.Recordset
    Set rsExist = gconDMIS.Execute("Select * from AMIS_Journal_hd where INVOICETYPE = '" & xInvoiceType & "' AND INVOICENO = '" & xInvoiceNo & "' AND JTYPE = 'SJ' and CUSTOMERCODE = '" & XCustomerCode & "'")
    If Not rsExist.EOF And Not rsExist.BOF Then
        rsCHECKINVOICENOandTYPE = True                ' yaon
    Else
        rsCHECKINVOICENOandTYPE = False               ' mayo
    End If
    Set rsExist = Nothing
End Function
Function GetSJVoucherNo(ByVal xInvoiceNo As String, ByVal xInvoiceType As String) As Boolean
'Update BTT : 10282008
'To check if the transaction is posted
    Dim RsSJVoucher                                         As New ADODB.Recordset
    Set RsSJVoucher = gconDMIS.Execute("Select Voucherno,invoicetype,invoiceno,Status from Amis_journal_hd where invoiceno=" & xInvoiceNo & " and invoicetype=" & xInvoiceType & " and jtype ='SJ'")
    If Not RsSJVoucher.EOF And Not RsSJVoucher.BOF Then
        If (RsSJVoucher!Status) = "P" Then
            GetSJVoucherNo = True
            SJVoucherno = Null2String(RsSJVoucher!VOUCHERNO)
        Else
            GetSJVoucherNo = False
            MsgBox "Transaction is not posted to sales journal..Please verify", vbExclamation, "WARNING"
        End If
    End If
    Set RsSJVoucher = Nothing
End Function

Sub CheckIfthereISCDJ(XXX As String)
    Dim RSCDJ                                               As New ADODB.Recordset
    Set RSCDJ = gconDMIS.Execute("SELECT amount FROM AMIS_CV_DETAIL where Pv_voucherno='" & XXX & "'")
    If Not RSCDJ.EOF And Not RSCDJ.BOF Then
        gconDMIS.Execute "UPDATE AMIS_journal_hd set balance = " & TOTALPVAMOUNT - TotalAPAmountToPay & "  where voucherno='" & XXX & "' and jtype='APJ'"
    End If
    Set RSCDJ = Nothing
End Sub

Function CHECK_IF_SCHED_ACCNT(xVOUCHERNO As String) As Boolean
    Dim rsCHECK_IF_SCHED_ACCNT                              As ADODB.Recordset
    Dim SHED                                                As Integer
    Dim NOT_SCHED                                           As Integer
    SHED = 0
    NOT_SCHED = 0
    Set rsCHECK_IF_SCHED_ACCNT = New ADODB.Recordset
    rsCHECK_IF_SCHED_ACCNT.Open "Select Acct_Code From Amis_Journal_det where VoucherNo = '" & xVOUCHERNO & "' and Jtype = '" & xJOURNALTYPE & "' and DEBIT <> 0 " & _
                                "AND Acct_Code IN(SELECT AcctCode FROM Amis_ChartAccount where IS_SCHEDULE_ACCNT = 1)", gconDMIS, adOpenKeyset
    If Not rsCHECK_IF_SCHED_ACCNT.EOF And Not rsCHECK_IF_SCHED_ACCNT.BOF Then
        CHECK_IF_SCHED_ACCNT = True
    Else
        CHECK_IF_SCHED_ACCNT = False
    End If
    Set rsCHECK_IF_SCHED_ACCNT = Nothing
End Function

Sub GET_AR_VOUCHERNO()
'UPDATED BY: JUN --- DATE UPDATED: 11/19/2009 --- DESCRIPTION: GET THE AR OF THE PARTICULAR VOUCHERNO
    Dim rsAR_VOUCHER                                        As ADODB.Recordset
    Dim rsCOUNT_CODE                                        As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xBAL                                                As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsCOUNT_CODE = New ADODB.Recordset
    rsCOUNT_CODE.Open "SELECT COUNT(DISTINCT ACCT_CODE) AS COUNT_CODE FROM AMIS_JOURNAL_DET " & _
                      "WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset
    If Not rsCOUNT_CODE.EOF And Not rsCOUNT_CODE.BOF Then
        'THIS IS FOR ACCT_CODE GREATER THAN ONE IN ONE VOUCHERNO
        If NumericVal(rsCOUNT_CODE!COUNT_CODE) > 1 Then
            Set rsAR_VOUCHER = New ADODB.Recordset
            rsAR_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,ACCT_CODE " & _
                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenKeyset
            If Not rsAR_VOUCHER.EOF And Not rsAR_VOUCHER.BOF Then
                Do While Not rsAR_VOUCHER.EOF
                    xVOUCHERNO = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE) & "-" & Null2String(rsAR_VOUCHER!VOUCHERNO))
                    xJdate = N2Str2Null(Null2String(rsAR_VOUCHER!JDATE))
                    xJType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))

                    If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
                        XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!VendorCode))
                        xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAR_VOUCHER!VendorCode)))
                    Else
                        XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!CustomerCode))
                        xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAR_VOUCHER!CustomerCode)))
                    End If

                    xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!INVOICENO))
                    xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!InvoiceType))
                    xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                    xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
                    xAMOUNT_PAID = 0
                    xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                    xACCT_CODE = N2Str2Null(Null2String(rsAR_VOUCHER!ACCT_CODE))
                    xLAST_UPDATED = N2Str2Null(LOGDATE)

                    gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                     "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ")"
                    rsAR_VOUCHER.MoveNext
                Loop
            End If
            Set rsAR_VOUCHER = Nothing
        Else
            Set rsAR_VOUCHER = New ADODB.Recordset
            rsAR_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,ACCT_CODE " & _
                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
            If Not rsAR_VOUCHER.EOF And Not rsAR_VOUCHER.BOF Then
                xVOUCHERNO = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE) & "-" & Null2String(rsAR_VOUCHER!VOUCHERNO))
                xJdate = N2Str2Null(Null2String(rsAR_VOUCHER!JDATE))
                xJType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))

                If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
                    XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!VendorCode))
                    xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAR_VOUCHER!VendorCode)))
                Else
                    XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!CustomerCode))
                    xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAR_VOUCHER!CustomerCode)))
                End If

                xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!InvoiceType))
                xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
                xAMOUNT_PAID = 0
                xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                xACCT_CODE = N2Str2Null(Null2String(rsAR_VOUCHER!ACCT_CODE))
                xLAST_UPDATED = N2Str2Null(LOGDATE)

                gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ")"
            End If
            Set rsAR_VOUCHER = Nothing
        End If
    End If
    Set rsCOUNT_CODE = Nothing
End Sub

Function GET_AR_AMOUNT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'UPDATE BY: JUN --- DATE UPDATED: 11/19/2009 --- DESCRIPTION: THIS IS TO SUM THE AR WITH THE SPECIFIC ACCOUNT CODE

    Dim rsGET_AR_AMOUNT                                     As ADODB.Recordset
    Set rsGET_AR_AMOUNT = New ADODB.Recordset
    rsGET_AR_AMOUNT.Open "SELECT ROUND(SUM(DET.DEBIT),2) AS SUM_DEBIT " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_AR_AMOUNT.EOF And Not rsGET_AR_AMOUNT.BOF Then
        GET_AR_AMOUNT = NumericVal(rsGET_AR_AMOUNT!SUM_DEBIT)
    Else
        GET_AR_AMOUNT = 0
    End If
    Set rsGET_AR_AMOUNT = Nothing
End Function

Sub GET_PAYMENT_VOUCHERNO()
'UPDATED BY: JUN --- DATE UPDATED: 11/19/2009 --- DESCRIPTION: THIS IS TO GET THE PAYMENT OF THE PARTICUALR REFERENCE IN THE SJ OR CUSTOMER OPENING BALANCE
    Dim rsGET_PAYMENT_VOUCHERNO                             As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim XCustomerCode                                       As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xACCT_CODE                                          As String
    Dim xINVOICE_AMT                                        As Double
    Dim xJType                                              As String
    Dim xInvoicedate                                        As String


    Set rsGET_PAYMENT_VOUCHERNO = New ADODB.Recordset
    rsGET_PAYMENT_VOUCHERNO.Open "SELECT HD.JDATE,CRJ.VOUCHERNO,CRJ.CUSTOMERCODE,CRJ.INVOICENO,CRJ.INVOICETYPE,CRJ.J_CLASS,CRJ.INVOICEAMOUNT,CRJ.INVOICEDATE " & _
                                 "FROM AMIS_CRJ_DETAIL CRJ INNER JOIN AMIS_JOURNAL_HD HD ON CRJ.VOUCHERNO = HD.VOUCHERNO AND CRJ.CR_TYPE = HD.JTYPE WHERE CRJ.VOUCHERNO = '" & txtVoucherNo.Text & "' AND CRJ.CR_TYPE = 'CRJ'", gconDMIS, adOpenKeyset
    If Not rsGET_PAYMENT_VOUCHERNO.EOF And Not rsGET_PAYMENT_VOUCHERNO.BOF Then
        Do While Not rsGET_PAYMENT_VOUCHERNO.EOF
            If IsNull(rsGET_PAYMENT_VOUCHERNO!J_CLASS) = True Then
                'THIS IS CREDIT CARD TRANSACTION AR NOT A PAYMENT
            Else
                xVOUCHERNO = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!VOUCHERNO))
                xJdate = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!JDATE))
                XCustomerCode = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!CustomerCode))

                xInvoiceNo = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!InvoiceType))

                xACCT_CODE = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!J_CLASS))
                xINVOICE_AMT = NumericVal(rsGET_PAYMENT_VOUCHERNO!invoiceamount)
                xJType = N2Str2Null("CRJ")
                xInvoicedate = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!invoicedate))

                'INSERT INTO AMIS_AR
                gconDMIS.Execute "INSERT INTO AMIS_DETAIL(INVOICETYPE,INVOICENO,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,VOUCHERNO,JTYPE,INVOICEDATE) " & _
                                 "VALUES(" & xInvoiceType & "," & xInvoiceNo & "," & xINVOICE_AMT & "," & XCustomerCode & "," & xACCT_CODE & "," & xJdate & "," & xVOUCHERNO & "," & xJType & "," & xInvoicedate & ")"

                Dim rsSUM_PAYMENT                           As ADODB.Recordset
                Dim xSUM_PAYMENT                            As Double
                Set rsSUM_PAYMENT = New ADODB.Recordset
                xSUM_PAYMENT = 0
                'SUM THE TOTAL INVOICE AMOUNT IN AMIS DETAIL AND UPDATE THE AMIS_AR AMOUNT_PAID WHICH IS MATCH TO THE REFERENCE
                rsSUM_PAYMENT.Open "SELECT ROUND(SUM(INVOICEAMOUNT),2) AS SUM_BAYAD FROM AMIS_DETAIL WHERE INVOICENO = " & xInvoiceNo & "  AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
                If Not rsSUM_PAYMENT.EOF And Not rsSUM_PAYMENT.BOF Then
                    xSUM_PAYMENT = NumericVal(rsSUM_PAYMENT!SUM_BAYAD)
                Else
                    xSUM_PAYMENT = NumericVal(0)
                End If
                Set rsSUM_PAYMENT = Nothing

                Dim rsGET_SUM_AR                            As ADODB.Recordset
                Dim xSUM_AR                                 As Double
                Dim xAR_BALANCE                             As Double
                xSUM_AR = 0
                xAR_BALANCE = 0
                Set rsGET_SUM_AR = New ADODB.Recordset
                'SUM THE TOTAL AR IN SALES JOURNAL
                rsGET_SUM_AR.Open "SELECT ROUND(SUM(AMOUNT_TOPAY),2) as AMOUNT_TOPAY FROM AMIS_AR WHERE INVOICENO = " & xInvoiceNo & "  AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
                If Not rsGET_SUM_AR.EOF And Not rsGET_SUM_AR.BOF Then
                    Dim rsCHECK_EXIST                       As ADODB.Recordset
                    Dim xSJVOUCHERNO                        As String
                    xSJVOUCHERNO = N2Str2Null(xJType & "-" & Null2String(rsGET_PAYMENT_VOUCHERNO!VOUCHERNO))
                    Set rsCHECK_EXIST = New ADODB.Recordset
                    rsCHECK_EXIST.Open "SELECT * FROM AMIS_AR WHERE INVOICENO = " & xInvoiceNo & "  AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
                    If Not rsCHECK_EXIST.EOF And Not rsCHECK_EXIST.BOF Then
                        xSUM_AR = NumericVal(rsGET_SUM_AR!AMOUNT_TOPAY)
                    Else
                        'AR NOT FOUND IN AMIS_AR OR NO FOUND AR IN SJ BUT HAS A PAYMENT IN CRJ
                        gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,CRJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                         "VALUES(" & xSJVOUCHERNO & ",NULL," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & ",'" & GET_CUST_NAME(Null2String(rsGET_PAYMENT_VOUCHERNO!CustomerCode)) & "',0," & xSUM_PAYMENT & "," & xSUM_PAYMENT & "," & xACCT_CODE & "," & xInvoicedate & "," & LOGDATE & "," & xJdate & ")"
                    End If
                    Set rsCHECK_EXIST = Nothing
                Else
                    xSUM_AR = NumericVal(0)
                End If

                xAR_BALANCE = Round((xSUM_AR - xSUM_PAYMENT), 2)

                Set rsGET_SUM_AR = Nothing
                'UPDATE THE TOTAL AMOUNT PAID AND AR BALANCE TO THE AMIS_AR
                gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & xAR_BALANCE & " WHERE INVOICENO = " & xInvoiceNo & "  AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
            End If
            rsGET_PAYMENT_VOUCHERNO.MoveNext
        Loop
    End If
    Set rsGET_PAYMENT_VOUCHERNO = Nothing
End Sub

Sub GET_AR_CRJ()
'UPDATED BY: JUN --- DATE UPDATED: 11/19/2009 --- DESCRIPTION: THIS IS TO GET THE AR IN CRJ MOSTLY ARE A/R CREDIT CARD TRANSACTION
    Dim rsGET_AR_CRJ                                        As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xBAL                                                As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0


    Set rsGET_AR_CRJ = New ADODB.Recordset
    rsGET_AR_CRJ.Open "SELECT DISTINCT CRJ.INVOICETYPE + '-' + CRJ.INVOICENO,HD.VOUCHERNO,HD.JDATE,HD.JTYPE,HD.BANK,CRJ.INVOICENO,CRJ.INVOICETYPE, " & _
                      "CRJ.INVOICEDATE,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                      "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CRJ_DETAIL CRJ " & _
                      "ON HD.VOUCHERNO = CRJ.VOUCHERNO AND HD.JTYPE = CRJ.CR_TYPE " & _
                      "WHERE HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.JTYPE = 'CRJ' AND DET.DEBIT <> 0 AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset

    If Not rsGET_AR_CRJ.EOF And Not rsGET_AR_CRJ.BOF Then
        Do While Not rsGET_AR_CRJ.EOF
            'If COMPANY_CODE = "HGC" And Null2String(rsGET_AR_CRJ!ACCT_CODE) = "11-02002-00" Then
            If Null2String(rsGET_AR_CRJ!ACCT_CODE) = "11-02002-00" Then
                xVOUCHERNO = N2Str2Null(Null2String(rsGET_AR_CRJ!JTYPE) & "-" & Null2String(rsGET_AR_CRJ!VOUCHERNO))
                xJdate = N2Str2Null(Null2String(rsGET_AR_CRJ!JDATE))
                xJType = N2Str2Null(Null2String(rsGET_AR_CRJ!JTYPE))
                XCustomerCode = N2Str2Null(Null2String(rsGET_AR_CRJ!Bank))
                xCUST_NAME = N2Str2Null(Null2String(GET_CUST_NAME(Null2String(rsGET_AR_CRJ!Bank))))
                xInvoiceNo = N2Str2Null(rsGET_AR_CRJ!INVOICENO)
                xInvoiceType = N2Str2Null(rsGET_AR_CRJ!InvoiceType)
                xInvoicedate = N2Str2Null(rsGET_AR_CRJ!invoicedate)
                xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsGET_AR_CRJ!VOUCHERNO), Null2String(rsGET_AR_CRJ!JTYPE), Null2String(rsGET_AR_CRJ!ACCT_CODE))
                xAMOUNT_PAID = 0
                xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                xACCT_CODE = N2Str2Null(Null2String(rsGET_AR_CRJ!ACCT_CODE))
                xLAST_UPDATED = N2Str2Null(LOGDATE)

                gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ")"
            End If
            rsGET_AR_CRJ.MoveNext
        Loop
    End If
    Set rsGET_AR_CRJ = Nothing
End Sub

Function GET_VEN_NAME(xVENCODE As String) As String
    Dim rsGET_VEN_NAME                                      As ADODB.Recordset
    Set rsGET_VEN_NAME = New ADODB.Recordset
    rsGET_VEN_NAME.Open "SELECT NAMEOFVENDOR FROM ALL_VENDOR WHERE  RTRIM(LTRIM(CODE)) = " & N2Str2Null(xVENCODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_VEN_NAME.EOF And Not rsGET_VEN_NAME.BOF Then
        GET_VEN_NAME = Null2String(rsGET_VEN_NAME!nameofvendor)
    Else
        GET_VEN_NAME = ""
    End If
    Set rsGET_VEN_NAME = Nothing
End Function

Function GET_CUST_NAME(xCUSCODE As String) As String
    Dim rsGET_CUST_NAME                                     As ADODB.Recordset
    Set rsGET_CUST_NAME = New ADODB.Recordset
    rsGET_CUST_NAME.Open "SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE RTRIM(LTRIM(CUSCDE)) = '" & RTrim(LTrim(xCUSCODE)) & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CUST_NAME.EOF And Not rsGET_CUST_NAME.BOF Then
        GET_CUST_NAME = Null2String(rsGET_CUST_NAME!AcctName)
    Else
        GET_CUST_NAME = ""
    End If
    Set rsGET_CUST_NAME = Nothing
End Function

Function VOUCHER_TO_VOUCHER_ADJ() As Boolean
    Dim rsFIND_ADJ                                          As ADODB.Recordset
    Dim rsINFO_ADJ                                          As ADODB.Recordset

    Set rsFIND_ADJ = New ADODB.Recordset
    rsFIND_ADJ.Open "SELECT JTYPE,VOUCHERNO FROM AMIS_JOURNAL_DET WHERE INVOICENO IS NULL AND INVOICETYPE IS NULL AND  ADJ_VOUCHERNO IS NOT NULL AND ADJ_JTYPE IS NOT NULL " & _
                    "AND ADJ_JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND ADJ_VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND ADJ_JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsFIND_ADJ.EOF And Not rsFIND_ADJ.BOF Then
        MessagePop InfoFriend, "INFORMATION", "You can't un-post this voucher it has an adjustment. Please see General Journal " & "" & Null2String(rsFIND_ADJ!JTYPE) & "" & " - " & "" & Null2String(rsFIND_ADJ!VOUCHERNO) & ""
        VOUCHER_TO_VOUCHER_ADJ = True
    Else
        VOUCHER_TO_VOUCHER_ADJ = False
    End If
    Set rsFIND_ADJ = Nothing
End Function

Sub GET_AP_VOUCHERNO()
    Dim rsAP_VOUCHER                                        As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xDUEDATE                                            As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xBAL                                                As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsAP_VOUCHER = New ADODB.Recordset
    rsAP_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,ACCT_CODE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsAP_VOUCHER.EOF And Not rsAP_VOUCHER.BOF Then
        xVOUCHERNO = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE) & "-" & Null2String(rsAP_VOUCHER!VOUCHERNO))
        xJdate = N2Str2Null(Null2String(rsAP_VOUCHER!JDATE))
        xJType = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE))
        xDUEDATE = N2Str2Null(Null2String(rsAP_VOUCHER!DUEDATE))
        If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!VendorCode))
            xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAP_VOUCHER!VendorCode)))
        Else
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!CustomerCode))
            xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAP_VOUCHER!CustomerCode)))
        End If

        xInvoiceNo = N2Str2Null(Null2String(rsAP_VOUCHER!INVOICENO))
        xInvoiceType = N2Str2Null(Null2String(rsAP_VOUCHER!InvoiceType))
        xInvoicedate = N2Str2Null(Null2String(rsAP_VOUCHER!invoicedate))
        xAMOUNT_TO_PAY = GET_AP_AMOUNT(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!JTYPE), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xAMOUNT_PAID = GET_AMOUNT_PAID(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!JTYPE), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
        xACCT_CODE = N2Str2Null(Null2String(rsAP_VOUCHER!ACCT_CODE))
        xLAST_UPDATED = N2Str2Null(LOGDATE)

        SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                        "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
        gconDMIS.Execute SQL_STATEMENT
        If xJOURNALTYPE = "CDJ" Then
            gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid=" & xAMOUNT_PAID & ",Balance = " & xBAL & " where JTYPE ='" & xJOURNALTYPE & "' And VOUCHERNO = " & N2Str2Null(rsAP_VOUCHER!VOUCHERNO)
        End If
    End If
    Set rsAP_VOUCHER = Nothing
End Sub

Function GET_AP_AMOUNT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsGET_AP_AMOUNT                                     As ADODB.Recordset
    Set rsGET_AP_AMOUNT = New ADODB.Recordset
    rsGET_AP_AMOUNT.Open "SELECT ROUND(SUM(DET.CREDIT),2) AS SUM_CREDIT " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_AP_AMOUNT.EOF And Not rsGET_AP_AMOUNT.BOF Then
        GET_AP_AMOUNT = NumericVal(rsGET_AP_AMOUNT!SUM_CREDIT)
    Else
        GET_AP_AMOUNT = 0
    End If
    Set rsGET_AP_AMOUNT = Nothing
End Function

Function GET_AMOUNT_PAID(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
    Dim rsGET_AMOUNT_PAID                                   As ADODB.Recordset
    Set rsGET_AMOUNT_PAID = New ADODB.Recordset
    rsGET_AMOUNT_PAID.Open "SELECT * FROM (SELECT ROUND(SUM(DET.DEBIT),2) AS SUM_DEBIT,HD.VOUCHERNO,HD.JTYPE " & _
                           "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                           "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JTYPE = 'CDJ' AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE VOUCHERNO=" & N2Str2Null(xVOUCHERNO) & ") AND HD.STATUS = 'P' GROUP BY HD.VOUCHERNO,HD.JTYPE) X WHERE VOUCHERNO=" & N2Str2Null(xVOUCHERNO) & "", gconDMIS, adOpenKeyset
    If Not rsGET_AMOUNT_PAID.EOF And Not rsGET_AMOUNT_PAID.BOF Then
        GET_AMOUNT_PAID = rsGET_AMOUNT_PAID!SUM_DEBIT
    Else
        GET_AMOUNT_PAID = 0
    End If
    Set rsGET_AMOUNT_PAID = Nothing
End Function

Sub GET_PAYMENT()
    Dim rsGET_PAYMENT_VOUCHERNO                             As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xPV_VOUCHERNO                                       As String
    Dim xJdate                                              As String
    Dim xVENDORCODE                                         As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xACCT_CODE                                          As String
    Dim xAMOUNT                                             As Double
    Dim xJType                                              As String
    Dim xInvoicedate                                        As String


    Set rsGET_PAYMENT_VOUCHERNO = New ADODB.Recordset
    rsGET_PAYMENT_VOUCHERNO.Open "SELECT HD.JDATE,CV.VOUCHERNO,CV.PV_VOUCHERNO,CV.JTYPE,CV.VENDORCODE,CV.J_CLASS,CV.AMOUNT,CV.DOCDATE " & _
                                 "FROM AMIS_CV_DETAIL CV INNER JOIN AMIS_JOURNAL_HD HD ON CV.VOUCHERNO = HD.VOUCHERNO AND CV.CV_JTYPE = HD.JTYPE WHERE CV.VOUCHERNO = '" & txtVoucherNo.Text & "' AND CV.CV_JTYPE = 'CDJ'", gconDMIS, adOpenKeyset
    If Not rsGET_PAYMENT_VOUCHERNO.EOF And Not rsGET_PAYMENT_VOUCHERNO.BOF Then
        Do While Not rsGET_PAYMENT_VOUCHERNO.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!VOUCHERNO))
            xPV_VOUCHERNO = N2Str2Null(N2String(rsGET_PAYMENT_VOUCHERNO!JTYPE) & "-" & Null2String(rsGET_PAYMENT_VOUCHERNO!pv_voucherno))
            xJdate = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!JDATE))
            xVENDORCODE = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!VendorCode))

            'xINVOICENO = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!INVOICENO))
            'xINVOICETYPE = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!InvoiceType))

            xACCT_CODE = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!J_CLASS))
            xAMOUNT = NumericVal(rsGET_PAYMENT_VOUCHERNO!amount)
            xJType = N2Str2Null("CDJ")

            xInvoicedate = N2Str2Null(Null2String(rsGET_PAYMENT_VOUCHERNO!docdate))

            gconDMIS.Execute "INSERT INTO AMIS_DETAILS(AMOUNTPAID,VENDORCODE,ACCT_CODE,JDATE,VOUCHERNO,JTYPE) " & _
                             "VALUES(" & xAMOUNT & "," & xVENDORCODE & "," & xACCT_CODE & "," & xJdate & "," & xVOUCHERNO & "," & xJType & ")"

            Dim rsSUM_PAYMENT                               As ADODB.Recordset
            Dim xSUM_PAYMENT                                As Double
            Set rsSUM_PAYMENT = New ADODB.Recordset
            xSUM_PAYMENT = 0
            'SUM THE TOTAL INVOICE AMOUNT IN AMIS DETAIL AND UPDATE THE AMIS_AR AMOUNT_PAID WHICH IS MATCH TO THE REFERENCE
            rsSUM_PAYMENT.Open "SELECT ROUND(SUM(AMOUNTPAID),2) AS SUM_BAYAD FROM AMIS_DETAILS WHERE VENDORCODE = " & xVENDORCODE & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            If Not rsSUM_PAYMENT.EOF And Not rsSUM_PAYMENT.BOF Then
                xSUM_PAYMENT = NumericVal(rsSUM_PAYMENT!SUM_BAYAD)
            Else
                xSUM_PAYMENT = NumericVal(0)
            End If
            Set rsSUM_PAYMENT = Nothing

            Dim rsGET_SUM_AP                                As ADODB.Recordset
            Dim xSUM_AP                                     As Double
            Dim xAP_BALANCE                                 As Double
            xSUM_AP = 0
            xAP_BALANCE = 0
            Set rsGET_SUM_AP = New ADODB.Recordset
            'SUM THE TOTAL AP
            'VOUCHERNO = '" & (Null2String(RTrim(LTrim(rsAMIS_APCheck!jtype))) + "-" + Null2String(rsAMIS_APCheck!pv_voucherno)) & "'"

            rsGET_SUM_AP.Open "SELECT ROUND(SUM(AMOUNT2PAY),2) as AMOUNT2PAY FROM AMIS_AP WHERE VENDOR_CODE = " & xVENDORCODE & " AND VOUCHERNO= " & xPV_VOUCHERNO & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset

            If Not rsGET_SUM_AP.EOF And Not rsGET_SUM_AP.BOF Then
                xSUM_AP = NumericVal(rsGET_SUM_AP!AMOUNT2PAY)
            Else
                xSUM_AP = NumericVal(0)
            End If

            xAP_BALANCE = Round((xSUM_AP - xSUM_PAYMENT), 2)

            Set rsGET_SUM_AP = Nothing
            'UPDATE THE TOTAL AMOUNT PAID AND AP BALANCE TO THE AMIS_AP
            gconDMIS.Execute "UPDATE AMIS_AP SET AMOUNTPAID = " & xSUM_PAYMENT & ", BALANCE = " & xAP_BALANCE & " WHERE VENDOR_CODE = " & xVENDORCODE & " AND VOUCHERNO= " & xPV_VOUCHERNO & " AND ACCT_CODE = " & xACCT_CODE & ""
            gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid=" & xSUM_PAYMENT & ",Balance = " & xAP_BALANCE & " where JTYPE =" & N2Str2Null(rsGET_PAYMENT_VOUCHERNO!JTYPE) & " And VOUCHERNO = " & N2Str2Null(rsGET_PAYMENT_VOUCHERNO!pv_voucherno)

            rsGET_PAYMENT_VOUCHERNO.MoveNext
        Loop
    End If
    Set rsGET_PAYMENT_VOUCHERNO = Nothing
End Sub

Sub GET_DIRECT_DISBURSEMENT()
    Dim rsDIRECT_DISBURSEMENT                               As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xVENDORCODE                                         As String
    Dim xACCT_CODE                                          As String
    Dim xAMOUNT                                             As Double
    Dim xJType                                              As String
    Set rsDIRECT_DISBURSEMENT = New ADODB.Recordset
    rsDIRECT_DISBURSEMENT.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.VENDORCODE,HD.AMOUNTPAID,HD.JDATE,HD.JTYPE,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & ")) X WHERE VOUCHERNO= " & N2Str2Null(txtVoucherNo.Text) & " ", gconDMIS, adOpenKeyset
    If Not rsDIRECT_DISBURSEMENT.EOF And Not rsDIRECT_DISBURSEMENT.BOF Then
        xAMOUNT = NumericVal(rsDIRECT_DISBURSEMENT!AMOUNTPAID)
        xVENDORCODE = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!VendorCode))
        xACCT_CODE = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!ACCT_CODE))
        xJdate = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!JDATE))
        xVOUCHERNO = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!VOUCHERNO))
        xJType = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!JTYPE))
        gconDMIS.Execute "INSERT INTO AMIS_DETAILS(AMOUNTPAID,VENDORCODE,ACCT_CODE,JDATE,VOUCHERNO,JTYPE) " & _
                         "VALUES(" & xAMOUNT & "," & xVENDORCODE & "," & xACCT_CODE & "," & xJdate & "," & xVOUCHERNO & "," & xJType & ")"
    End If
End Sub

Sub UNPOST_CDJ()
    Dim rsUNPOST_CDJ                                        As ADODB.Recordset
    Dim rsAMIS_AP                                           As ADODB.Recordset
    Set rsUNPOST_CDJ = New ADODB.Recordset
    rsUNPOST_CDJ.Open "SELECT VENDORCODE,J_CLASS,JTYPE,PV_VOUCHERNO,AMOUNT FROM AMIS_CV_DETAIL WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND CV_JTYPE = 'CDJ'", gconDMIS, adOpenKeyset
    If Not rsUNPOST_CDJ.EOF And Not rsUNPOST_CDJ.BOF Then
        Do While Not rsUNPOST_CDJ.EOF
            Set rsAMIS_AP = New ADODB.Recordset
            rsAMIS_AP.Open "SELECT * FROM AMIS_AP WHERE ACCT_CODE = '" & rsUNPOST_CDJ!J_CLASS & "' AND VENDOR_CODE = '" & rsUNPOST_CDJ!VendorCode & "' ", gconDMIS, adOpenKeyset
            If Not rsAMIS_AP.EOF And Not rsAMIS_AP.BOF Then
                gconDMIS.Execute "UPDATE AMIS_AP SET AMOUNTPAID = 0 , BALANCE = " & NumericVal(rsAMIS_AP!AMOUNT2PAY) & " WHERE ACCT_CODE = '" & rsUNPOST_CDJ!J_CLASS & "' AND VENDOR_CODE = '" & rsUNPOST_CDJ!VendorCode & "' AND VOUCHERNO='" & txtVoucherNo.Text & "'"
                gconDMIS.Execute "DELETE FROM AMIS_DETAILS WHERE ACCT_CODE = '" & rsUNPOST_CDJ!J_CLASS & "' AND VENDORCODE = '" & rsUNPOST_CDJ!VendorCode & "'  AND JTYPE='CDJ' AND VOUCHERNO='" & txtVoucherNo.Text & "'"
                gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid = AmountPaid - " & NumericVal(rsUNPOST_CDJ!amount) & ",Balance = Balance + " & NumericVal(rsUNPOST_CDJ!amount) & " where JTYPE='" & Null2String(rsUNPOST_CDJ!JTYPE) & "' and VOUCHERNO = '" & Null2String(rsUNPOST_CDJ!pv_voucherno) & "'"
            End If
            rsUNPOST_CDJ.MoveNext
        Loop
    End If
    Set rsUNPOST_CDJ = Nothing
End Sub

Sub UNPOST_DIRECT_DISBURSEMENT()
    Dim rsDIRECT_DISBURSEMENT                               As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xVENDORCODE                                         As String
    Dim xACCT_CODE                                          As String
    Dim xAMOUNT                                             As Double
    Dim xJType                                              As String
    Set rsDIRECT_DISBURSEMENT = New ADODB.Recordset
    rsDIRECT_DISBURSEMENT.Open "SELECT * FROM (SELECT HD.VOUCHERNO,HD.VENDORCODE,HD.AMOUNTPAID,HD.JDATE,HD.JTYPE,DET.ACCT_CODE FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO=DET.VOUCHERNO AND HD.JTYPE=DET.JTYPE WHERE HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & ")) X WHERE VOUCHERNO= " & N2Str2Null(txtVoucherNo.Text) & " ", gconDMIS, adOpenKeyset
    If Not rsDIRECT_DISBURSEMENT.EOF And Not rsDIRECT_DISBURSEMENT.BOF Then
        xAMOUNT = NumericVal(rsDIRECT_DISBURSEMENT!AMOUNTPAID)
        xVENDORCODE = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!VendorCode))
        xACCT_CODE = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!ACCT_CODE))
        xJdate = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!JDATE))
        xVOUCHERNO = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!VOUCHERNO))
        xJType = N2Str2Null(Null2String(rsDIRECT_DISBURSEMENT!JTYPE))
        gconDMIS.Execute "DELETE FROM AMIS_DETAILS WHERE ACCT_CODE = " & xACCT_CODE & " AND VENDORCODE = " & xVENDORCODE & " AND JTYPE='CDJ' AND VOUCHERNO =" & xVOUCHERNO & ""
    End If
End Sub

Function AR_SHEDULE_ACCNT(xACCT_CODE As String) As Boolean
    Dim rsAR_ACCT_CODE                                      As ADODB.Recordset
    Set rsAR_ACCT_CODE = New ADODB.Recordset
    rsAR_ACCT_CODE.Open "SELECT * FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT = 1 AND ACCTCODE = " & N2Str2Null(RTrim(LTrim(xACCT_CODE))) & "", gconDMIS, adOpenKeyset
    If Not rsAR_ACCT_CODE.EOF And Not rsAR_ACCT_CODE.BOF Then
        AR_SHEDULE_ACCNT = True
    Else
        AR_SHEDULE_ACCNT = False
    End If
    Set rsAR_ACCT_CODE = Nothing
End Function

Function CheckIfPosted(xVOUCHERNO As String) As Boolean
    Dim RSCRJ                                               As ADODB.Recordset
    Set RSCRJ = New ADODB.Recordset
    RSCRJ.Open "Select VoucherNo,InvoiceNo,InvoiceType from AMIS_CRJ_Detail where VoucherNo ='" & xVOUCHERNO & "'", gconDMIS, adOpenForwardOnly
    If Not RSCRJ.EOF And Not RSCRJ.BOF Then
        Do While Not RSCRJ.EOF
            Dim rsSJPosted                                  As ADODB.Recordset
            Set rsSJPosted = New ADODB.Recordset
            rsSJPosted.Open "Select JType,InvoiceNo,InvoiceType from AMIS_Journal_HD where JType='SJ' and InvoiceType ='" & RSCRJ!InvoiceType & "' and InvoiceNo='" & RSCRJ!INVOICENO & "' and Status='P'", gconDMIS, adOpenForwardOnly
            If Not rsSJPosted.EOF And Not rsSJPosted.BOF Then
                CheckIfPosted = True
            Else
                CheckIfPosted = False
            End If
            RSCRJ.MoveNext
        Loop
    Else
        'No CRJ Detail
        CheckIfPosted = True
    End If
    Set RSCRJ = Nothing
    Set rsSJPosted = Nothing
End Function

Function CheckIfOpen(xJType As String, xAcctMonth, xAcctYear) As Boolean
    Dim rsCheckOpen                                         As ADODB.Recordset
    Set rsCheckOpen = New ADODB.Recordset
    rsCheckOpen.Open "Select * from AMIS_AccountingPeriod where JType = '" & xJType & "' and Month(AcctMonth) = '" & Format(xAcctMonth, "m") & "' and Year(AcctMonth) = '" & Format(xAcctMonth, "yyyy") & "' and Status=0 and CurrPeriod = 1", gconDMIS, adOpenForwardOnly
    If Not rsCheckOpen.EOF And Not rsCheckOpen.BOF Then
        CheckIfOpen = True
    Else
        CheckIfOpen = False
    End If
    Set rsCheckOpen = Nothing
End Function

Sub LOADJOURNAL(XXX As String)
    xJOURNALTYPE = XXX
End Sub
