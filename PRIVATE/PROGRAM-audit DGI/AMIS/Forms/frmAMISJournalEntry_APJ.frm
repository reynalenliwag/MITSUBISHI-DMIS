VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMISJournalEntry_APJ 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   7770
   ClientLeft      =   11040
   ClientTop       =   4800
   ClientWidth     =   9855
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAMISJournalEntry_APJ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9855
   Begin VB.PictureBox picRefCDJ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7065
      ScaleHeight     =   345
      ScaleWidth      =   2535
      TabIndex        =   71
      Top             =   870
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Label RefCDJ 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ref CDJ# 000000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   0
         TabIndex        =   72
         Top             =   0
         Width           =   2775
      End
   End
   Begin VB.PictureBox picBatchImport 
      Height          =   1755
      Left            =   2820
      ScaleHeight     =   1695
      ScaleWidth      =   4095
      TabIndex        =   182
      Top             =   2820
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdBatchPosting 
         Caption         =   "Batch Post"
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
         Left            =   840
         TabIndex        =   184
         Top             =   900
         Width           =   2805
      End
      Begin VB.CommandButton cmdCloseRange 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   183
         Top             =   0
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   405
         Left            =   570
         TabIndex        =   185
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57802753
         CurrentDate     =   40603
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   405
         Left            =   2520
         TabIndex        =   186
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   57802753
         CurrentDate     =   40603
      End
      Begin VB.Label lblVoucherNo 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2370
         TabIndex        =   192
         Top             =   1410
         Width           =   1695
      End
      Begin VB.Label lblPosting 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   191
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   0
         TabIndex        =   190
         Top             =   1410
         Width           =   4125
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   189
         Top             =   0
         Width           =   4155
         _Version        =   655364
         _ExtentX        =   7329
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Select Date Range for Open Journal Period"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8388608
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   60
         TabIndex        =   188
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2190
         TabIndex        =   187
         Top             =   480
         Width           =   225
      End
   End
   Begin TabDlg.SSTab JournalTAB 
      Height          =   4215
      Left            =   120
      TabIndex        =   79
      Top             =   2550
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   7435
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "[<F3> Add &Journal Entries]   [<Ctrl+J> View &Journals]   "
      TabPicture(0)   =   "frmAMISJournalEntry_APJ.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDetails"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddJournal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAddJournal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[<F4> Add &Details]   [<Ctrl+D> View &Details]   "
      TabPicture(1)   =   "frmAMISJournalEntry_APJ.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPV_Detail"
      Tab(1).Control(1)=   "cmdPV_Entry"
      Tab(1).Control(2)=   "picPV_Entry"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picPV_Entry 
         BackColor       =   &H00FF8080&
         Height          =   1575
         Left            =   -74790
         ScaleHeight     =   1515
         ScaleWidth      =   9105
         TabIndex        =   85
         Top             =   750
         Width           =   9165
         Begin VB.CommandButton Command4 
            Caption         =   ".."
            Height          =   345
            Left            =   5010
            TabIndex        =   172
            ToolTipText     =   "Show Invoice Application"
            Top             =   1830
            Width           =   465
         End
         Begin VB.ComboBox cboARTag 
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
            Left            =   3270
            TabIndex        =   165
            Text            =   "Combo1"
            Top             =   690
            Visible         =   0   'False
            Width           =   3825
         End
         Begin VB.CommandButton cmdPVCancel 
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
            Left            =   8400
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":0902
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":0A54
            Style           =   1  'Graphical
            TabIndex        =   98
            Top             =   690
            Width           =   705
         End
         Begin VB.CommandButton cmdPVDelete 
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
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":0D92
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":0EE4
            Style           =   1  'Graphical
            TabIndex        =   96
            Top             =   690
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtMRR_No 
            Height          =   315
            Left            =   1650
            TabIndex        =   92
            ToolTipText     =   "Press Enter to show AP/VPJ transaction"
            Top             =   330
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtPVAmount 
            Height          =   315
            Left            =   7080
            TabIndex        =   95
            Top             =   330
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox txtINV_No 
            Height          =   315
            Left            =   3270
            TabIndex        =   93
            Top             =   330
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   10
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
         Begin MSMask.MaskEdBox txtPO_No 
            Height          =   315
            Left            =   60
            TabIndex        =   91
            Top             =   330
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtProd_No 
            Height          =   315
            Left            =   5070
            TabIndex        =   94
            Top             =   330
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtPVItemNo 
            Height          =   225
            Left            =   510
            TabIndex        =   99
            Top             =   420
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   397
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdPVSave 
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
            Left            =   7710
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":120F
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":1361
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   690
            Width           =   705
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tag AR Type"
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
            Left            =   2040
            TabIndex        =   164
            Top             =   750
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label51 
            BackColor       =   &H00FF8080&
            Caption         =   "INFO:Press Enter key to show APJ/VPJ Transaction"
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
            Height          =   405
            Left            =   1680
            TabIndex        =   163
            Top             =   690
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   8310
            TabIndex        =   90
            Top             =   90
            Width           =   795
         End
         Begin VB.Label labPV1 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "PO Number"
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
            Left            =   90
            TabIndex        =   86
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label labPV2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MRR Number"
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
            Left            =   1680
            TabIndex        =   87
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label labPV3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Number"
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
            Left            =   3270
            TabIndex        =   88
            Top             =   120
            Width           =   1545
         End
         Begin VB.Label labPV4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Product Number"
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
            Left            =   5100
            TabIndex        =   89
            Top             =   120
            Width           =   1875
         End
         Begin VB.Label labPVID 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MRR Number"
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
            Left            =   300
            TabIndex        =   100
            Top             =   420
            Width           =   1305
         End
      End
      Begin VB.PictureBox fraAddJournal 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   240
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   111
         Top             =   690
         Width           =   9135
         Begin VB.CommandButton cmdJournalCancel 
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
            Left            =   8315
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":16B1
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":1803
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   765
            Width           =   705
         End
         Begin VB.TextBox txtCredit 
            Alignment       =   1  'Right Justify
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
            Left            =   7950
            MaxLength       =   15
            TabIndex        =   123
            Top             =   330
            Width           =   1100
         End
         Begin VB.TextBox txtDebit 
            Alignment       =   1  'Right Justify
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
            Left            =   6780
            MaxLength       =   15
            TabIndex        =   121
            Top             =   330
            Width           =   1100
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   2310
            TabIndex        =   118
            Top             =   -30
            Width           =   4425
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   315
               Left            =   30
               TabIndex        =   120
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393217
               BackColor       =   16777215
               Enabled         =   -1  'True
               MultiLine       =   0   'False
               TextRTF         =   $"frmAMISJournalEntry_APJ.frx":1B41
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
            Begin VB.Label Label33 
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
               Left            =   60
               TabIndex        =   119
               Top             =   90
               Width           =   2205
            End
         End
         Begin VB.ComboBox cboAcct_Code 
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
            TabIndex        =   116
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin VB.TextBox txtAcctID 
            BackColor       =   &H00FF0000&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   840
            TabIndex        =   117
            Text            =   "Text1"
            Top             =   330
            Width           =   585
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
            Left            =   690
            MaxLength       =   4
            TabIndex        =   115
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin VB.Frame fraATC 
            Height          =   915
            Left            =   2340
            TabIndex        =   128
            Top             =   660
            Width           =   4365
            Begin VB.ComboBox cboATC 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   60
               TabIndex        =   132
               Top             =   510
               Width           =   1425
            End
            Begin VB.TextBox txtRATE 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   133
               Top             =   510
               Width           =   615
            End
            Begin VB.TextBox txtTaxBase 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   134
               Top             =   510
               Width           =   1725
            End
            Begin VB.Label Label41 
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
               TabIndex        =   135
               Top             =   540
               Width           =   855
            End
            Begin VB.Label Label45 
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
               TabIndex        =   130
               Top             =   240
               Width           =   1365
            End
            Begin VB.Label Label44 
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
               TabIndex        =   129
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label43 
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
               TabIndex        =   131
               Top             =   240
               Width           =   1725
            End
         End
         Begin VB.Frame fraComp 
            Height          =   915
            Left            =   2340
            TabIndex        =   136
            Top             =   660
            Width           =   4365
            Begin VB.TextBox txtNetAmt 
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
               Left            =   2910
               MaxLength       =   10
               TabIndex        =   142
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtTax 
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
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   141
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtGrossAmt 
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
               Left            =   150
               MaxLength       =   10
               TabIndex        =   140
               Top             =   510
               Width           =   1300
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   2910
               TabIndex        =   139
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label labTax 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Output Tax"
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
               Left            =   1560
               TabIndex        =   138
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Amt."
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
               TabIndex        =   137
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.CommandButton cmdJournalSave 
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
            Left            =   7620
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":1BD4
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":1D26
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   765
            Width           =   705
         End
         Begin VB.CommandButton cmdJournalDelete 
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
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":2076
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISJournalEntry_APJ.frx":21C8
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   765
            Width           =   705
         End
         Begin VB.Label Label35 
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
            TabIndex        =   124
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label34 
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
            TabIndex        =   112
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label30 
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
            Left            =   7050
            TabIndex        =   113
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label38 
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
            Left            =   8130
            TabIndex        =   114
            Top             =   60
            Width           =   795
         End
         Begin VB.Label labPrevOrdQty 
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
            TabIndex        =   122
            Top             =   360
            Width           =   855
         End
         Begin VB.Label labDetID 
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
            Left            =   930
            TabIndex        =   126
            Top             =   390
            Width           =   915
         End
         Begin VB.Label labPartNo 
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
            TabIndex        =   125
            Top             =   420
            Width           =   2685
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1845
         Left            =   180
         TabIndex        =   110
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3254
         TX              =   "cmd1"
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
         MICON           =   "frmAMISJournalEntry_APJ.frx":24F3
      End
      Begin VB.PictureBox fraDetails 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   90
         ScaleHeight     =   3435
         ScaleWidth      =   9405
         TabIndex        =   101
         Top             =   120
         Width           =   9405
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   -420
            Top             =   2940
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Enabled         =   0   'False
            Height          =   465
            Left            =   30
            TabIndex        =   103
            Top             =   2940
            Width           =   9345
            Begin VB.PictureBox picDatePosted 
               BorderStyle     =   0  'None
               Height          =   375
               Left            =   30
               ScaleHeight     =   375
               ScaleWidth      =   2925
               TabIndex        =   176
               Top             =   60
               Width           =   2925
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
                  TabIndex        =   178
                  Top             =   60
                  Width           =   1500
               End
               Begin VB.Label label4 
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
                  TabIndex        =   177
                  Top             =   60
                  Width           =   1245
               End
            End
            Begin VB.PictureBox picChat 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   60
               ScaleHeight     =   345
               ScaleWidth      =   5655
               TabIndex        =   104
               Top             =   60
               Visible         =   0   'False
               Width           =   5655
               Begin VB.Label Label40 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Warning: AR Amount is not Balance with Total Journal Details Amount."
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
                  Left            =   0
                  TabIndex        =   105
                  Top             =   60
                  Width           =   5685
               End
            End
            Begin VB.TextBox txtOutBalance 
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
               Height          =   345
               Left            =   1320
               MaxLength       =   14
               TabIndex        =   107
               Text            =   "Text1"
               Top             =   60
               Width           =   1515
            End
            Begin VB.TextBox txtTotDebit 
               Alignment       =   1  'Right Justify
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
               Left            =   6270
               MaxLength       =   15
               TabIndex        =   109
               Text            =   "Text1"
               Top             =   60
               Width           =   1485
            End
            Begin VB.TextBox txtTotCredit 
               Alignment       =   1  'Right Justify
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
               Left            =   7770
               MaxLength       =   15
               TabIndex        =   108
               Text            =   "Text1"
               Top             =   60
               Width           =   1485
            End
            Begin VB.Label labOutBalance 
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
               Left            =   -60
               TabIndex        =   106
               Top             =   90
               Width           =   1275
            End
         End
         Begin MSComctlLib.ListView lstDetails 
            Height          =   2835
            Left            =   30
            TabIndex        =   102
            Top             =   60
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   5001
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
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":250F
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
      Begin wizButton.cmd cmdPV_Entry 
         Height          =   1665
         Left            =   -74835
         TabIndex        =   84
         Top             =   690
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2937
         TX              =   "cmd1"
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
         MICON           =   "frmAMISJournalEntry_APJ.frx":2671
      End
      Begin VB.PictureBox picPV_Detail 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3795
         Left            =   -74940
         ScaleHeight     =   3795
         ScaleWidth      =   9435
         TabIndex        =   80
         Top             =   90
         Width           =   9435
         Begin MSMask.MaskEdBox txtTotalPV_Amount 
            Height          =   345
            Left            =   8010
            TabIndex        =   82
            Top             =   3390
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   7347754
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ListView lstPV_Detail 
            Height          =   3285
            Left            =   60
            TabIndex        =   81
            Top             =   60
            Width           =   9315
            _ExtentX        =   16431
            _ExtentY        =   5794
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
            MouseIcon       =   "frmAMISJournalEntry_APJ.frx":268D
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEM #"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO NUMBER"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MRR NUMBER"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PRODUCT NO."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "AMOUNT"
               Object.Width           =   2823
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total :"
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
            Height          =   285
            Left            =   7320
            TabIndex        =   83
            Top             =   3450
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   120
      ScaleHeight     =   2610
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   60
      Width           =   9600
      Begin VB.PictureBox picNewEntity 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   6600
         ScaleHeight     =   315
         ScaleWidth      =   345
         TabIndex        =   179
         Top             =   480
         Visible         =   0   'False
         Width           =   345
         Begin VB.CommandButton cmdSelect 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   0
            TabIndex        =   180
            Top             =   0
            Width           =   345
         End
         Begin VB.Label lblClass 
            Height          =   195
            Left            =   510
            TabIndex        =   181
            Top             =   90
            Width           =   645
         End
      End
      Begin VB.PictureBox picReceivable 
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   9510
         TabIndex        =   36
         Top             =   3825
         Width           =   9510
         Begin VB.ComboBox cboBankName2 
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
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   4545
            TabIndex        =   52
            Text            =   "Invoice Type"
            Top             =   900
            Width           =   4950
         End
         Begin VB.CheckBox chkNonVat 
            Caption         =   "Non-Vat"
            Height          =   285
            Left            =   1140
            TabIndex        =   48
            Top             =   930
            Width           =   915
         End
         Begin VB.TextBox txtDealer 
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
            Left            =   7710
            MaxLength       =   50
            TabIndex        =   55
            Top             =   930
            Width           =   1755
         End
         Begin RichTextLib.RichTextBox txtRemarks2 
            Height          =   705
            Left            =   4560
            TabIndex        =   61
            Top             =   1350
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   1244
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmAMISJournalEntry_APJ.frx":27EF
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
         Begin VB.TextBox txtRefDate 
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
            Left            =   7710
            MaxLength       =   10
            TabIndex        =   45
            Top             =   540
            Width           =   1755
         End
         Begin VB.TextBox txtRefNo 
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
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   44
            Top             =   540
            Width           =   2085
         End
         Begin VB.ComboBox cboInvoiceType 
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
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   1530
            TabIndex        =   41
            Text            =   "Invoice Type"
            Top             =   510
            Width           =   1500
         End
         Begin VB.ComboBox cboCustName 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   2520
            TabIndex        =   38
            Text            =   "Customer Name"
            Top             =   30
            Width           =   4080
         End
         Begin VB.TextBox txtInvoiceAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   1530
            MaxLength       =   15
            TabIndex        =   62
            Text            =   "0.00"
            Top             =   1710
            Width           =   1485
         End
         Begin VB.TextBox txtInvoiceDate2 
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   57
            Text            =   "88/88/8888"
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtCustCode 
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
            TabIndex        =   39
            Text            =   "000226"
            Top             =   45
            Width           =   1005
         End
         Begin VB.ComboBox cboPayTerm2 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   5460
            TabIndex        =   54
            Text            =   "Invoice Type"
            Top             =   930
            Width           =   1200
         End
         Begin VB.TextBox txtTerms 
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
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   53
            Top             =   930
            Width           =   855
         End
         Begin VB.TextBox txtInvoiceNo 
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   47
            Text            =   "000000"
            Top             =   930
            Width           =   1485
         End
         Begin VB.Label labTerms 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
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
            Height          =   285
            Left            =   3180
            TabIndex        =   50
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label labDealer 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer"
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
            Height          =   285
            Left            =   6720
            TabIndex        =   56
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label RefCRJ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ref CRJ# 000000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   6720
            TabIndex        =   40
            Top             =   60
            Width           =   2775
         End
         Begin VB.Label labBankName 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Height          =   285
            Left            =   3180
            TabIndex        =   51
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label labRefDate 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref. Date"
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
            Height          =   285
            Left            =   6720
            TabIndex        =   46
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label labRefNo 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Reference No."
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
            Height          =   285
            Left            =   3180
            TabIndex        =   43
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label labType 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type"
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
            Height          =   285
            Left            =   150
            TabIndex        =   42
            Top             =   570
            Width           =   1425
         End
         Begin VB.Label Label32 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Cust. Code"
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
            Height          =   285
            Left            =   180
            TabIndex        =   37
            Top             =   60
            Width           =   1935
         End
         Begin VB.Label labParticulars 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3180
            TabIndex        =   59
            Top             =   1350
            Width           =   1695
         End
         Begin VB.Label labAmt 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Amount"
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
            Height          =   285
            Left            =   150
            TabIndex        =   60
            Top             =   1740
            Width           =   1425
         End
         Begin VB.Label labDate 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Date"
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
            Height          =   285
            Left            =   150
            TabIndex        =   58
            Top             =   1350
            Width           =   1425
         End
         Begin VB.Label LabNo 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. No."
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
            Height          =   285
            Left            =   150
            TabIndex        =   49
            Top             =   960
            Width           =   1425
         End
      End
      Begin VB.TextBox txtCode 
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
         TabIndex        =   9
         Text            =   "000226"
         Top             =   465
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
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   60
         Width           =   1545
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
      Begin VB.ComboBox cboNameofVendor 
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
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   2520
         TabIndex        =   6
         Text            =   "cboRecvd_Desc"
         Top             =   450
         Width           =   4080
      End
      Begin VB.TextBox txtDueDate 
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
         Left            =   7950
         MaxLength       =   10
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   450
         Width           =   1545
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
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.PictureBox picPayables 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   90
         ScaleHeight     =   1275
         ScaleWidth      =   9555
         TabIndex        =   14
         Top             =   1110
         Width           =   9555
         Begin VB.TextBox txtPayCode 
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
            Left            =   1410
            MaxLength       =   3
            TabIndex        =   15
            Text            =   "000226"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtAmountToPay 
            Alignment       =   1  'Right Justify
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
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtInvoiceDate 
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "88/88/8888"
            Top             =   450
            Width           =   1695
         End
         Begin VB.ComboBox cboPayType 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1950
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   17
            Text            =   "cboPayType"
            Top             =   60
            Width           =   2325
         End
         Begin RichTextLib.RichTextBox txtRemarks 
            Height          =   765
            Left            =   4380
            TabIndex        =   22
            Top             =   420
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   1349
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmAMISJournalEntry_APJ.frx":2886
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
         Begin VB.Label Label11 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4350
            TabIndex        =   18
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amt. to Pay"
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
            Left            =   225
            TabIndex        =   21
            Top             =   840
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
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
            Left            =   150
            TabIndex        =   19
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type"
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
            Left            =   30
            TabIndex        =   16
            Top             =   90
            Width           =   1335
         End
      End
      Begin VB.PictureBox picDisbursement 
         BorderStyle     =   0  'None
         Height          =   1245
         Left            =   90
         ScaleHeight     =   1245
         ScaleWidth      =   9525
         TabIndex        =   24
         Top             =   1140
         Width           =   9525
         Begin VB.TextBox txtCheckAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   4380
            MaxLength       =   10
            TabIndex        =   35
            Text            =   "000226"
            Top             =   840
            Width           =   1485
         End
         Begin VB.TextBox txtCheckDate 
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   34
            Text            =   "000226"
            Top             =   810
            Width           =   1815
         End
         Begin VB.TextBox txtCheckNo 
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
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   29
            Text            =   "000226"
            Top             =   420
            Width           =   1815
         End
         Begin VB.ComboBox cboBankName 
            Appearance      =   0  'Flat
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
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   4380
            TabIndex        =   27
            Text            =   "cboRecvd_Desc"
            Top             =   30
            Width           =   5070
         End
         Begin VB.TextBox txtBankCode 
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
            Left            =   1410
            MaxLength       =   8
            TabIndex        =   25
            Text            =   "000226"
            Top             =   30
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtParticulars 
            Height          =   795
            Left            =   4380
            TabIndex        =   32
            Top             =   420
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmAMISJournalEntry_APJ.frx":291D
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
         Begin VB.Label labCheckAmt 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amt"
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
            Height          =   285
            Left            =   3270
            TabIndex        =   162
            Top             =   870
            Width           =   1935
         End
         Begin VB.Label Label10 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Height          =   285
            Left            =   3270
            TabIndex        =   28
            Top             =   90
            Width           =   1935
         End
         Begin VB.Label Label14 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3270
            TabIndex        =   33
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
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
            Height          =   285
            Left            =   60
            TabIndex        =   31
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check No."
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
            Height          =   285
            Left            =   60
            TabIndex        =   30
            Top             =   450
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Code"
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
            Height          =   285
            Left            =   60
            TabIndex        =   26
            Top             =   60
            Width           =   1935
         End
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
         TabIndex        =   173
         Top             =   60
         Width           =   4065
      End
      Begin VB.Label labSupplierPayTo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   1935
      End
      Begin VB.Label labDueDate 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
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
         Left            =   6990
         TabIndex        =   10
         Top             =   510
         Width           =   885
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
         TabIndex        =   1
         Top             =   120
         Width           =   1245
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
         Left            =   6855
         TabIndex        =   13
         Top             =   930
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label txtAddress 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   11
         Top             =   840
         Width           =   6465
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4545
      Left            =   15240
      ScaleHeight     =   4545
      ScaleWidth      =   3045
      TabIndex        =   166
      Top             =   30
      Width           =   3045
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4245
         Left            =   90
         ScaleHeight     =   4215
         ScaleWidth      =   2745
         TabIndex        =   167
         Top             =   60
         Width           =   2775
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   30
            TabIndex        =   169
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
            TabIndex        =   171
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
            TabIndex        =   170
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
            TabIndex        =   168
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
   Begin VB.CommandButton cmdPrinting 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Command1"
      Height          =   2445
      Left            =   9690
      TabIndex        =   78
      Top             =   7890
      Width           =   2775
   End
   Begin wizButton.cmd cmdShowPostRange 
      Height          =   2385
      Left            =   9900
      TabIndex        =   77
      Top             =   2790
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   4207
      TX              =   "cmd1"
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
      MICON           =   "frmAMISJournalEntry_APJ.frx":29B1
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -210
      ScaleHeight     =   900
      ScaleWidth      =   9735
      TabIndex        =   148
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":29CD
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":2B1F
         Style           =   1  'Graphical
         TabIndex        =   160
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":2E85
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":2FD7
         Style           =   1  'Graphical
         TabIndex        =   159
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":333D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":348F
         Style           =   1  'Graphical
         TabIndex        =   158
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":37C9
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":391B
         Style           =   1  'Graphical
         TabIndex        =   157
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":3C60
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":3DB2
         Style           =   1  'Graphical
         TabIndex        =   156
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":40D7
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":4229
         Style           =   1  'Graphical
         TabIndex        =   155
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":4585
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":46D7
         Style           =   1  'Graphical
         TabIndex        =   154
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":49EA
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":4B3C
         Style           =   1  'Graphical
         TabIndex        =   153
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":4E8C
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":4FDE
         Style           =   1  'Graphical
         TabIndex        =   152
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":533C
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":548E
         Style           =   1  'Graphical
         TabIndex        =   151
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":5788
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":58DA
         Style           =   1  'Graphical
         TabIndex        =   150
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":5C32
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":5D84
         Style           =   1  'Graphical
         TabIndex        =   149
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
      TabIndex        =   145
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":60E3
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":6235
         Style           =   1  'Graphical
         TabIndex        =   147
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":6573
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_APJ.frx":66C5
         Style           =   1  'Graphical
         TabIndex        =   146
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
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
      Left            =   210
      TabIndex        =   64
      Top             =   390
      Width           =   9405
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
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   270
         Width           =   9225
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4515
         Left            =   90
         TabIndex        =   67
         Top             =   660
         Width           =   9225
         _ExtentX        =   16272
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":6A15
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
         TabIndex        =   68
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
         TabIndex        =   66
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
         TabIndex        =   69
         Top             =   5310
         Width           =   9225
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1080
      TabIndex        =   70
      Top             =   810
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   7488
      TX              =   "cmd1"
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
      MICON           =   "frmAMISJournalEntry_APJ.frx":6B77
   End
   Begin VB.PictureBox picTemplates 
      Height          =   4125
      Left            =   1170
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   73
      Top             =   930
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
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   75
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
         MouseIcon       =   "frmAMISJournalEntry_APJ.frx":6B93
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
         TabIndex        =   76
         Top             =   3750
         Width           =   7035
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   5775
      Left            =   150
      TabIndex        =   63
      Top             =   300
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   10186
      TX              =   "cmd1"
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
      MICON           =   "frmAMISJournalEntry_APJ.frx":6CF5
   End
   Begin VB.Label TEXTID 
      Caption         =   "fortestingpurpose"
      Height          =   705
      Left            =   1800
      TabIndex        =   175
      Top             =   8010
      Width           =   2295
   End
   Begin VB.Label TESTID 
      Caption         =   "TESTID"
      Height          =   645
      Left            =   270
      TabIndex        =   174
      Top             =   -5730
      Width           =   2535
   End
   Begin VB.Label lblVPJAcctCode 
      Caption         =   "dont delete this "
      Height          =   165
      Left            =   11220
      TabIndex        =   161
      Top             =   1080
      Width           =   1845
   End
End
Attribute VB_Name = "frmAMISJournalEntry_APJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As New ADODB.Recordset
Dim rsJournal_Det                                           As New ADODB.Recordset
Dim rsOUTBAL                                                As New ADODB.Recordset
Dim rsPV_Detail                                             As New ADODB.Recordset
Dim rsCV_Detail                                             As New ADODB.Recordset
Dim rsCRJ_Detail                                            As New ADODB.Recordset
Dim rsJV_detail                                             As New ADODB.Recordset
Dim rsChartAccount                                          As New ADODB.Recordset
Dim rsATC                                                   As New ADODB.Recordset
Dim rsJournal_HD2                                           As New ADODB.Recordset
Dim rsProfile                                               As New ADODB.Recordset
Dim rsCheckJournal_HD                                       As New ADODB.Recordset
Dim rsVENDOR                                                As New ADODB.Recordset
Dim rsPayTerm                                               As New ADODB.Recordset
Dim rsBanks                                                 As New ADODB.Recordset
Dim rsCustomer                                              As New ADODB.Recordset
Dim rsInvoiceType                                           As New ADODB.Recordset
Dim kcnt, Jcnt                                              As Integer
Attribute Jcnt.VB_VarUserMemId = 1073938449
Dim AddorEdit                                               As String
Dim SearchBy                                                As String
Public CDJ_CIB                                              As String
Public CDJ_AP                                               As String
Dim LocalAcess                                              As String
Dim TOTDEBIT                                                As Double
Dim TOTCREDIT                                               As Double
Dim TOTTAX                                                  As Double
Dim OUTBALANCE                                              As Double
Dim out                                                     As Integer
Attribute out.VB_VarUserMemId = 1073938450
Dim TOTAL_AR_AMOUNT                                         As Double
Dim TOTALPVAMOUNT                                           As Double
Dim COMP_SJ_OUTPUT_TAX                                      As Double
Dim PrevJType                                               As String
Dim PrevJNo                                                 As String
Dim PrevInvoiceType                                         As String
Dim PrevInvoiceNo                                           As String
Dim PrevPV_VoucherNo                                        As String
Dim PrevPV_Amount                                           As Double
Dim DirectDisbursementVoucherNo                             As String
Dim CDJ_IS_FROM_AP                                          As Boolean
Dim IsVPJ                                                   As Boolean
Dim TotalARAmountToPay                                      As Double
Dim TOTAL_AP_AMOUNT                                         As Double
Dim TotalAPAmountToPay                                      As Double
Dim SJVoucherno                                             As String
Dim APJInvoiceNo                                            As String
Dim APJinvoicetype                                          As String
Dim xJOURNALTYPE                                            As String
Attribute xJOURNALTYPE.VB_VarUserMemId = 1073938451
Dim xEntityClass                                            As String
Dim bSelectEntity                                           As Boolean
Dim CURRENT_VENDORCODE                                      As String
Dim BatchPost                                               As Boolean
Dim WithEvents frmNewEntity                                 As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim JOURNAL_DETID                                           As Long
'Dim WithEvents frmNewAMISJournalEntries                As frmAMISJournalEntries

Sub LOADJOURNAL(XXX As String)
    xJOURNALTYPE = XXX
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
    rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankName = Null2String(rsBanks!BankName)
        CDJ_CIB = N2Str2Null(rsBanks!AcctCode)
    Else
        SetBankName = ""
        CDJ_CIB = "NULL"
    End If
End Function

Function SetCustomerCode(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
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
        If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VCJ" Then
            If txtAcct_Name.Text = "ACCOUNTS PAYABLE - TRADE" Then SetDebitCredit = "D"
        ElseIf xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
            If txtAcct_Name.Text = "ACCOUNTS RECEIVABLE - TRADE" Then SetDebitCredit = "C"
        Else
            SetDebitCredit = Null2String(rsAccountType!DebitCredit)
        End If
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
    If bSelectEntity = True Then
        rsVENDOR.Open "Select CODE,ACCOUNTNAME as nameofvendor from ALL_ENTITY where ENTITYCODE='" & lblClass.Caption & "' AND ACCOUNTNAME = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where nameofvendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!CODE)
    Else
        SetVendorCode = ""
    End If
End Function

Function SetVendorID(XXX As Variant)
    Dim rsVENDORID
    Set rsVENDORID = New ADODB.Recordset
    rsVENDORID.Open "Select ID from ALL_VENDOR_TABLE where nameofvendor = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDORID.EOF And Not rsVENDORID.BOF Then
        SetVendorID = Null2String(rsVENDORID!ID)
    Else
        SetVendorID = ""
    End If
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    If bSelectEntity = True Then
        rsVENDOR.Open "Select CODE,ACCOUNTNAME as nameofvendor from ALL_ENTITY where ENTITYCODE='" & lblClass.Caption & "' AND code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function StoreJournalEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,acct_code,acct_name,debit,jitemno,credit,tax,grossamt,netamt,ATC,RATE,TAXBASE from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labDetID.Caption = rsJournal_Det!ID

        'for testing
        TEXTID.Caption = rsJournal_Det!ID


        labPartNo.Caption = Null2String(rsJournal_Det!ACCT_CODE)
        txtJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboAcct_Code.Text = Null2String(rsJournal_Det!ACCT_CODE)
        txtAcct_Name.Text = Null2String(rsJournal_Det!acct_Name)
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Debit))
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Credit))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!tax))
        txtGrossAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!grossamt))
        txtNetAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!netamt))
        If xJOURNALTYPE = "APJ" And fraATC.Visible = True Then
            If Null2String(rsJournal_Det!ATC) <> "" Then
                cboATC.Text = Null2String(rsJournal_Det!ATC)
            Else
                cboATC.ListIndex = 0
            End If
            txtRATE.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
            txtTaxBase.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!TAXBASE))
        Else
            ' Update By BTT : 09262008
            If Null2String(rsJournal_Det!ATC) <> "" Then
                cboATC.Text = Null2String(rsJournal_Det!ATC)
            End If
            txtRATE.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
            txtTaxBase.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!TAXBASE))
        End If
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

Function StorePVEntry(ByVal ID As Variant)
    If xJOURNALTYPE = "APJ" Then
        Set rsPV_Detail = New ADODB.Recordset
        rsPV_Detail.Open "select * from AMIS_PV_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPV_Detail.EOF And Not rsPV_Detail.BOF Then
            labPVID.Caption = rsPV_Detail!ID
            txtPVItemNo.Text = Null2String(rsPV_Detail!ItemNo)
            txtPO_No.Text = Null2String(rsPV_Detail!po_no)
            txtMRR_No.Text = Null2String(rsPV_Detail!MRR_No)
            txtINV_No.Text = Null2String(rsPV_Detail!INV_NO)
            txtProd_No.Text = Null2String(rsPV_Detail!Prod_No)
            txtPVAmount.Text = N2Str2Zero(rsPV_Detail!amount)
        End If
    ElseIf xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        Set rsCRJ_Detail = New ADODB.Recordset
        rsCRJ_Detail.Open "select * from AMIS_CRJ_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            labPVID.Caption = rsCRJ_Detail!ID
            txtPVItemNo.Text = Null2String(rsCRJ_Detail!ItemNo)
            txtPO_No.Text = txtVoucherNo.Text
            ReturnAccountDescription (Null2String(rsCRJ_Detail!J_CLASS))
            txtPO_No.Enabled = False
            cboARTag.Text = Setacctname(Null2String(rsCRJ_Detail!J_CLASS))
            txtMRR_No.Text = Null2String(rsCRJ_Detail!InvoiceType)
            txtINV_No.Text = Null2String(rsCRJ_Detail!INVOICENO)
            txtProd_No.Text = Null2String(rsCRJ_Detail!invoicedate)
            txtPVAmount.Text = N2Str2Zero(rsCRJ_Detail!invoiceamount)
            PrevInvoiceType = Null2String(rsCRJ_Detail!InvoiceType)
            PrevInvoiceNo = Null2String(rsCRJ_Detail!INVOICENO)
            PrevPV_Amount = N2Str2Zero(rsCRJ_Detail!invoiceamount)
        End If
    Else
        Set rsCV_Detail = New ADODB.Recordset
        rsCV_Detail.Open "select * from AMIS_CV_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
            labPVID.Caption = rsCV_Detail!ID
            txtPVItemNo.Text = Null2String(rsCV_Detail!ItemNo)
            txtPO_No.Text = txtVoucherNo.Text
            txtPO_No.Enabled = False
            txtPO_No.Text = Null2String(rsCV_Detail!JTYPE)
            txtMRR_No.Text = Null2String(rsCV_Detail!pv_voucherno)
            PrevPV_VoucherNo = Null2String(rsCV_Detail!pv_voucherno)
            txtINV_No.Text = Null2String(rsCV_Detail!docdate)
            txtProd_No.Text = Null2String(rsCV_Detail!DUEDATE)
            txtPVAmount.Text = N2Str2Zero(rsCV_Detail!amount)
            txtMRR_No.Enabled = False
            txtProd_No.Enabled = True
            txtINV_No.Enabled = True
            PrevPV_Amount = N2Str2Zero(rsCV_Detail!amount)
        End If
    End If
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

Sub BringToFront()
    Picture1.Enabled = False
    cmdAddJournal.ZOrder 0
    cmdAddJournal.Visible = True
    fraAddJournal.ZOrder 0
    fraAddJournal.Visible = True
    fraAddJournal.Enabled = True
End Sub

Sub BringToFrontPV()
    cmdPV_Entry.ZOrder 0
    cmdPV_Entry.Visible = True
    picPV_Entry.ZOrder 0
    picPV_Entry.Visible = True
    picPV_Entry.Enabled = True
End Sub

Sub BringToFrontTemplates()
    cmdTemplates.ZOrder 0
    picTemplates.ZOrder 0
    FillTemplates
End Sub

Private Sub cmdInternalRO_Click()
    If xJOURNALTYPE <> "SJ" Then
        JournalTAB.Tab = 1
        txtMRR_No.BackColor = &HFFFFFF
        txtINV_No.BackColor = &HFFFFFF
    Else
        ShowInvoiceApp SetInvCode(cboInvoiceType), txtInvoiceNo.Text
    End If
End Sub

Private Sub cboATC_Change()
    Dim RSATCRTE                                            As New ADODB.Recordset
    Dim xTOTAL                                              As Double
    GettheTaxBaseAmnt
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC.Text))
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE.Text = N2Str2Zero(rsATC!Rate)
        If NumericVal(txtRATE.Text) > 0 Then
            txtCredit.Text = Round(NumericVal(txtTaxBase.Text) * (NumericVal(txtRATE.Text) / 100), 2)
        End If
    End If
    Set rsATC = Nothing
End Sub

Private Sub cmdBatchPosting_Click()
    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then
        '        JournalPosting = True
        '        Exit Function
        Exit Sub
    End If

    If dtTo.Value < dtFrom.Value Then
        MessagePop InfoWarning, "Check Date", "Please check selected date."
        dtTo.SetFocus
        Exit Sub
    End If

    If Month(dtFrom.Value) <> Month(dtTo.Value) Then
        MessagePop InfoWarning, "Sysem Message", "Date selected should be of the same month"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to perform Batch Posting?", vbQuestion + vbYesNo, "Message") = vbYes Then
        If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "OPB" Then
            'Updated by: ACL 10202009
            If CheckIfOpen(xJOURNALTYPE, Trim(dtTo.Value), Year(dtTo.Value)) = False Then
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                Exit Sub
            End If
        End If
        BatchPost = True

        Dim rsJournal_HD                                    As ADODB.Recordset
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE='" & xJOURNALTYPE & "' AND STATUS='N' AND JDATE BETWEEN '" & Format(dtFrom.Value, "mm/dd/yyyy") & "' AND '" & Format(dtTo.Value, "mm/dd/yyyy") & "' ORDER BY ID", gconDMIS, adOpenForwardOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
            Do While Not rsJournal_HD.EOF
                lblPosting.Caption = "Posting Voucher No."
                lblVoucherNo.Caption = Null2String(rsJournal_HD!VOUCHERNO)
                Call StoreSearch(Null2String(rsJournal_HD!VOUCHERNO))
                If NumericVal(txtTotDebit) <> NumericVal(txtTotCredit) Then
                    MsgBox "Entry is not balanced. Posting of Entry Not Allowed.", vbInformation
                    Exit Sub
                End If
                If CheckGLSLARDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                If CheckGLSLARCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                If CheckGLSLAPDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                If CheckGLSLAPCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If

                cmdPost_Click
                DoEvents
                rsJournal_HD.MoveNext
            Loop
        End If
        Set rsJournal_HD = Nothing

    End If
End Sub

Private Sub cmdCloseRange_Click()
    Picture1.Enabled = True
    JournalTAB.Enabled = True
    picBatchImport.Visible = False
    picBatchImport.ZOrder 1
    BatchPost = False
End Sub

Sub cmdPVSave_Click()
    On Error GoTo ErrorCode

    Dim str_MSG                                             As String

    If Len(txtINV_No.Text) = 0 Then
        MessagePop InfoWarning, "Required Field", "Please input Invoice No."
        txtINV_No.SetFocus
        Exit Sub
    ElseIf Len(txtPVAmount.Text) = 0 Then
        MessagePop InfoWarning, "Required  Field", "Please input Invoice Amount"
        txtPVAmount.SetFocus
        Exit Sub
    End If

    str_MSG = "Error in saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If DetailPosting = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "APJ Details")
        MsgBox str_MSG, vbCritical, "APJ Detail Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0

ErrorCode:
    ShowVBError
End Sub

Function DetailPosting() As Boolean
    On Error GoTo ErrorCode

    Dim Ans                                                 As String
    txtMRR_No.BackColor = &HFFFFFF
    txtINV_No.BackColor = &HFFFFFF
    If AddorEdit = "ADD" Then
        Dim rsPV_DetailClone                                As ADODB.Recordset
        Set rsPV_DetailClone = New ADODB.Recordset
        rsPV_DetailClone.Open "select * from AMIS_PV_Detail where PO_NO = " & N2Str2Null(txtPO_No.Text) & " and MRR_NO = " & N2Str2Null(txtMRR_No.Text) & " and INV_NO = " & N2Str2Null(txtINV_No.Text), gconDMIS
        If Not rsPV_DetailClone.EOF And Not rsPV_DetailClone.BOF Then
            MsgBox "PO Number : " & txtPO_No.Text & " with MRR Number : " & txtMRR_No.Text & " and Invoice Number : " & txtINV_No.Text & " already used in this transaction", vbInformation, "Error in PO Number, MRR Number, Invoice Number Validation"
            DetailPosting = True
            Exit Function
        End If
    End If
    'UPDATE BY BTT : 11/27/2008

    'UPDATED BY  : JUN -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO              As String
    Dim J_JVOUCHERNO, J_JDATE                               As String
    Dim PV_AMOUNT                                           As Double
    Dim PV_STATUS, PV_ITEMNO                                As String
    Dim PV_VENDORCODE                                       As String

    J_JVOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JDATE = N2Str2Null(txtJDate.Text)
    PV_ITEMNO = N2Str2Null(Format(txtPVItemNo.Text, "0000"))
    PV_PONO = N2Str2Null(txtPO_No.Text)
    PV_MRRNO = N2Str2Null(txtMRR_No.Text)             ' TYPE
    PV_INVNO = N2Str2Null(txtINV_No.Text)             ' NO
    PV_PRODNO = N2Str2Null(txtProd_No.Text)           ' DATE
    PV_AMOUNT = NumericVal(txtPVAmount.Text)          ' AMOUNT
    PV_STATUS = "'N'"
    PV_VENDORCODE = N2Str2Null(txtCode.Text)

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        Dim rsJournal_HD_APJ                                As ADODB.Recordset
        Dim rsPV_Detail_APJ                                 As ADODB.Recordset

        Set rsJournal_HD_APJ = New ADODB.Recordset
        Set rsJournal_HD_APJ = gconDMIS.Execute("Select VoucherNo,VendorCode from AMIS_Journal_HD Where Jtype = 'APJ' and VendorCode = '" & txtCode.Text & "' Order By VoucherNo Asc")
        If Not rsJournal_HD_APJ.EOF And Not rsJournal_HD_APJ.BOF Then
            Do While Not rsJournal_HD_APJ.EOF
                Set rsPV_Detail_APJ = New ADODB.Recordset
                Set rsPV_Detail_APJ = gconDMIS.Execute("Select * from AMIS_PV_Detail Where (Inv_No = " & PV_INVNO & " OR Prod_No = " & PV_PRODNO & ") AND VoucherNo = " & N2Str2Null(rsJournal_HD_APJ!VOUCHERNO))
                If Not rsPV_Detail_APJ.EOF And Not rsPV_Detail_APJ.BOF Then
                    Screen.MousePointer = 0
                    MsgBox "Invoice No. or Prod No. Already Used in PV Number - " & Null2String(rsPV_Detail_APJ!VOUCHERNO)
                    DetailPosting = True
                    Exit Function
                Else
                    rsJournal_HD_APJ.MoveNext
                End If
            Loop
        End If
        'FOR APJ
        If xJOURNALTYPE = "APJ" Then
            SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                            "(JTYPE,VoucherNo,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status,JOURNAL_HD_ID)" & _
                            " values ('" & xJOURNALTYPE & "'," & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                            ", " & PV_STATUS & "," & labID.Caption & ")"
            gconDMIS.Execute SQL_STATEMENT
            JOURNAL_DETID = FindNewID(J_JVOUCHERNO, "VOUCHERNO", "AMIS_PV_DETAIL", N2Str2Null(xJOURNALTYPE), "JTYPE")
            NEW_LogAudit "AA", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "PV", txtVoucherNo, xJOURNALTYPE, CStr(JOURNAL_DETID)
        End If
    Else
        ' End of Balance computation
        'EDIT MODE
        If xJOURNALTYPE = "APJ" Then
            SQL_STATEMENT = "update AMIS_PV_Detail set" & _
                            " VoucherNo = " & J_JVOUCHERNO & "," & _
                            " JType = '" & xJOURNALTYPE & "'," & _
                            " itemno = " & PV_ITEMNO & "," & _
                            " PO_No = " & PV_PONO & "," & _
                            " MRR_No = " & PV_MRRNO & "," & _
                            " INV_No = " & PV_INVNO & "," & _
                            " PROD_No = " & PV_PRODNO & "," & _
                            " Amount = " & PV_AMOUNT & "," & _
                            " status = " & PV_STATUS & _
                            " where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "PV", txtVoucherNo, xJOURNALTYPE, labPVID.Caption
        End If
    End If                                            ' end of ADD

    '    If AMIS_AP_ENTRY = False Then
    '        Exit Function
    '    End If

    'FillDetails
    rsRefresh
    'On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then
        cmdPV_Entry_Click
    Else
        cmdPVCancel.Value = True
    End If

    '   AXP 1062008 FOR CHECKING: FOR KIM CHECKING
    '   SHOULD ALSO BE INCLUDED IN DELETE BUTTON TOO
    '   UPDATE AMIS JOURNAL HEADER BALANCE AFTER PAYMENT HAS BEEN DONE IN DISBURSMENT

    Dim RS_TOTALPAY                                         As ADODB.Recordset
    If txtPO_No = "APJ" Then
        Set RS_TOTALPAY = gconDMIS.Execute("SELECT isnull(SUM(AMOUNT),0)  AS TOTALPAYMENTS  FROM AMIS_CV_Detail WHERE CV_JTYPE='CDJ' AND JTYPE='VPJ' AND PV_VOUCHERNO=" & N2Str2Null(txtMRR_No))
        If Not RS_TOTALPAY.EOF Or Not RS_TOTALPAY.BOF Then
            gconDMIS.Execute ("update amis_journal_hd  set  BALANCE=AMOUNTTOPAY-" & RS_TOTALPAY!TOTALPAYMENTS & " WHERE JTYPE='APJ' AND VOUCHERNO= " & N2Str2Null(txtMRR_No))
        End If
    End If

    JournalTAB.TabEnabled(0) = True
    Picture1.Enabled = True
    cboARTag.BackColor = &H80000005

    DetailPosting = True
    Exit Function

ErrorCode:
    DetailPosting = False
End Function

Sub FillDetails()
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0: TotalARAmountToPay = 0
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE: TOTAL_AP_AMOUNT = 0: TotalAPAmountToPay = 0: PrevPV_Amount = 0
    Dim J_ITemNo, PV_ITEMNO                                 As Integer
    If xJOURNALTYPE <> "GJ" And xJOURNALTYPE <> "OPB" And xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "CLO" Then
        lstDetails.Sorted = False: lstDetails.ListItems.Clear
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            Screen.MousePointer = 11
            rsJournal_Det.MoveFirst: TOTAL_AR_AMOUNT = 0
            Do While Not rsJournal_Det.EOF
                kcnt = kcnt + 1
                If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
                lstDetails.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
                lstDetails.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!ACCT_CODE)
                lstDetails.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
                lstDetails.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!Debit))

                'COMMENTED BY: JUN | DESCRIPTION: COMPUTATION OF AR AMOUNT WILL BE CHECK AND BASE ONLY FOR ACCOUNT TAG AS AR SCHEDULE ACCOUNT
                'If Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "11-02" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "11-03" Then
                '    TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + N2Str2Zero(rsJournal_Det!CREDIT)
                '    TotalARAmountToPay = TotalARAmountToPay + N2Str2Zero(rsJournal_Det!DEBIT)
                'End If

                'UPDATED BY: JUN | DATE UPDATED: 01072010 | DESCRIPTION: CHECK IF ACCT CODE IS AR SCHEDULE ACCOUNT
                If AR_SHEDULE_ACCNT(Null2String(rsJournal_Det!ACCT_CODE)) = True Then
                    TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + N2Str2Zero(rsJournal_Det!Credit)
                    TotalARAmountToPay = TotalARAmountToPay + N2Str2Zero(rsJournal_Det!Debit)
                End If
                'UPDATED BY: JUN----------------------------------------------------------------------------------

                '21-07 AS SCHEDULED ACCOUNT NOT YET INCLUDED
                If Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-01" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-02" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-06" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-07" Then
                    TOTAL_AP_AMOUNT = TOTAL_AP_AMOUNT + N2Str2Zero(rsJournal_Det!Credit)
                    TotalAPAmountToPay = TotalAPAmountToPay + N2Str2Zero(rsJournal_Det!Debit)
                End If
                lstDetails.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!Credit))
                lstDetails.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
                If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then COMP_SJ_OUTPUT_TAX = 0
                TOTDEBIT = TOTDEBIT + Round(NumericVal(N2Str2Zero(rsJournal_Det!Debit)), 2)
                TOTCREDIT = TOTCREDIT + Math.Round(N2Str2Zero(rsJournal_Det!Credit), 2)
                TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJournal_Det!tax))
                rsJournal_Det.MoveNext


            Loop
            lstDetails.Sorted = True: lstDetails.Refresh
            txtTotDebit.Text = ToDoubleNumber(TOTDEBIT)
            txtTotCredit.Text = ToDoubleNumber(TOTCREDIT)
            OUTBALANCE = Round(TOTDEBIT - TOTCREDIT, 2)

            If labPosted.Caption = "" Then
                If Abs(OUTBALANCE) <> 0 Then
                    txtOutBalance.Text = Abs(OUTBALANCE)
                    cmdPost.Enabled = False
                    labOutBalance.Visible = True
                    txtOutBalance.Visible = True
                Else
                    txtOutBalance.Text = Abs(OUTBALANCE)
                    cmdPost.Enabled = True
                    labOutBalance.Visible = False
                    txtOutBalance.Visible = False
                End If
            End If
            Screen.MousePointer = 0
        Else
            Screen.MousePointer = 0
            If xJOURNALTYPE = "VDJ" Or xJOURNALTYPE = "VCJ" Or xJOURNALTYPE = "CSJ" Or xJOURNALTYPE = "CCM" Then
            Else
                cmdPost.Enabled = False
            End If
        End If

        'DISPLAY JOURNAL DETAILS
        Jcnt = 0
        TOTALPVAMOUNT = 0
        txtTotalPV_Amount.Text = ZERO
        If xJOURNALTYPE = "APJ" Then
            lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
            Set rsPV_Detail = New ADODB.Recordset
            Set rsPV_Detail = gconDMIS.Execute("select * from AMIS_PV_Detail where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = '" & xJOURNALTYPE & "' order by itemno asc")
            If Not rsPV_Detail.EOF And Not rsPV_Detail.BOF Then
                Screen.MousePointer = 11
                rsPV_Detail.MoveFirst: TOTALPVAMOUNT = 0
                Do While Not rsPV_Detail.EOF
                    Jcnt = Jcnt + 1
                    If Null2String(rsPV_Detail!ItemNo) = "" Then PV_ITEMNO = Jcnt Else PV_ITEMNO = Null2String(rsPV_Detail!ItemNo)
                    lstPV_Detail.ListItems.Add Jcnt, , Format(PV_ITEMNO, "0000")
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 1, , Null2String(rsPV_Detail!po_no)
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 2, , Null2String(rsPV_Detail!MRR_No)
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 3, , Null2String(rsPV_Detail!INV_NO)
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 4, , Null2String(rsPV_Detail!Prod_No)
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsPV_Detail!amount))
                    lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 6, , rsPV_Detail!ID
                    TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsPV_Detail!amount))
                    rsPV_Detail.MoveNext
                Loop
                lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
                txtTotalPV_Amount.Text = TOTALPVAMOUNT
                Screen.MousePointer = 0
            End If
        End If
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

Sub InitCbo()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select acctcode from AMIS_ChartAccount order by acctcode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
    Set rsChartAccount = Nothing
    'If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VDJ" Or xJOURNALTYPE = "VCJ" Then TO BE DELETED
    If xJOURNALTYPE = "APJ" Then
        Set rsATC = New ADODB.Recordset
        Set rsATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
        If Not rsATC.EOF And Not rsATC.BOF Then
            'Combo_Loadval cboATC, rsATC
            rsATC.MoveFirst: cboATC.AddItem ""
            Do While Not rsATC.EOF
                cboATC.AddItem Null2String(rsATC!ATC)
                rsATC.MoveNext
            Loop
        End If
        Set rsATC = Nothing

        Set rsVENDOR = New ADODB.Recordset
        Set rsVENDOR = gconDMIS.Execute("select nameofvendor from ALL_Vendor order by nameofvendor asc")
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            Combo_Loadval cboNameofVendor, rsVENDOR
        End If
        Set rsVENDOR = Nothing

        Set rsBanks = New ADODB.Recordset
        Set rsBanks = gconDMIS.Execute("select bankname from ALL_Banks order by bankname asc")
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            Combo_Loadval cboBankName, rsBanks
        End If
        Set rsBanks = Nothing
    End If

    Set rsPayTerm = New ADODB.Recordset
    Set rsPayTerm = gconDMIS.Execute("select pay_desc from ALL_PayTerm order by pay_desc asc")
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        Combo_Loadval cboPayType, rsPayTerm
    End If
    Set rsPayTerm = Nothing

    Dim rsAR_Accounts                                       As ADODB.Recordset
    Set rsAR_Accounts = New ADODB.Recordset
    Set rsAR_Accounts = gconDMIS.Execute("Select Description from AMIS_ChartAccount where Titles in('1102' ,'1103','1102','1204','2102','2107')ORDER BY Description")
    If Not rsAR_Accounts.EOF And Not rsAR_Accounts.BOF Then
        rsAR_Accounts.MoveFirst: cboARTag.Clear
        Do While Not rsAR_Accounts.EOF
            cboARTag.AddItem Null2String(rsAR_Accounts!DESCRIPTION)
            rsAR_Accounts.MoveNext
        Loop
    End If
End Sub
Sub InitCustomer()
    Set rsCustomer = New ADODB.Recordset
    'Set rsCustomer = gconDMIS.Execute("select acctname from ALL_CUSTMASTER_AMIS where (acctname <> '' and acctname is not null) order by acctname asc")
    Set rsCustomer = gconDMIS.Execute("select custname from ALL_CUSTMASTER_AMIS where (custname <> '' and custname is not null) order by custname asc")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Combo_Loadval cboCustName, rsCustomer
    End If
    Set rsCustomer = Nothing
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

Sub initGrid()
    If xJOURNALTYPE = "CDJ" Then
        With lstPV_Detail
            .ColumnHeaders(2).Text = "PV Number"
            .ColumnHeaders(2).Width = 2900
            .ColumnHeaders(3).Text = "Doc. Date"
            .ColumnHeaders(3).Width = 2000
            .ColumnHeaders(4).Text = "Due Date"
            .ColumnHeaders(4).Alignment = lvwColumnLeft
            .ColumnHeaders(4).Width = 2000
            .ColumnHeaders(5).Width = 1
            txtMRR_No.MaxLength = 6
        End With
    End If
End Sub

Sub InitJournal()
    txtJItemNo.Text = Format(kcnt + 1, "0000")
    cboAcct_Code.Text = ""
    txtAcct_Name.Text = ""
    txtDebit.Text = ZERO
    txtCredit.Text = ZERO
    txtTax.Text = ZERO
    txtGrossAmt.Text = ZERO
    txtNetAmt.Text = ZERO
    txtSearch.Text = ""
    labDetID.Caption = ""
    If xJOURNALTYPE = "APJ" Then
        cboATC.Text = ""
        txtRATE.Text = "0"
        txtTaxBase.Text = ZERO
    End If
End Sub

Sub initMemvars()
    Call GetNewVoucherNo
    txtJDate.Text = LOGDATE:

    CDJ_CIB = ""
    CDJ_AP = ""
    picDatePosted.Visible = False
    'Accounts Payable Module'
    txtCode.Text = ""
    txtAddress.Caption = "":
    txtInvoiceDate.Text = LOGDATE
    txtDueDate.Text = LOGDATE:
    txtBankCode.Text = ""
    picRefCDJ.ZOrder 0: picRefCDJ.Visible = False
    txtRemarks.Text = "Pls Type Your Message Here!"
    cboPayType.Text = ""
    '---------------------------'
    'Cash Disbursement Module'
    txtCheckNo.Text = "": txtCheckDate.Text = LOGDATE: txtPayCode.Text = ""
    cboNameofVendor.Text = ""
    txtTotDebit.Text = ZERO: txtTotCredit.Text = ZERO
    txtAmountToPay.Text = ZERO: txtOutBalance.Text = ZERO
    txtCheckAmt.Text = ZERO
    txtParticulars.Text = "Pls Type Your Message Here!"
    '    '---------------------------'
    '    'Accounts Receivable Module'
    txtCustCode.Text = ""
    cboCustName.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate2.Text = LOGDATE
    txtInvoiceAmt.Text = ZERO
    txtRefNo.Text = ""
    txtRefDate.Text = LOGDATE
    txtRemarks2.Text = "Pls Type Your Message Here!"
    '---------------------------'

    txtTotalPV_Amount.Text = ZERO
    labPosted.Caption = ""
    labPosted.Visible = False
    labOutBalance.Visible = False
    txtOutBalance.Visible = False
    'InitGrid
    SendToBack
End Sub

Sub InitPV_Detail()
    txtPVItemNo.Text = Format(Jcnt + 1, "0000")
    txtMRR_No.Text = ""
    If xJOURNALTYPE = "APJ" Then
        txtPO_No.Text = "": txtINV_No.Text = "": txtProd_No.Text = ""
        txtPVAmount.Text = ZERO
    ElseIf xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        txtPO_No.Text = txtVoucherNo.Text: txtINV_No.Text = ""
        txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
        txtPVAmount.Text = ZERO
    Else
        labPV1.Caption = "Voucher No": txtPO_No.Text = txtVoucherNo.Text: txtPO_No.Enabled = False
        labPV2.Caption = "PV Voucher No.": labPV3.Caption = "Doc. Date": labPV4.Caption = "Due Date"
        txtINV_No.Text = LOGDATE: txtINV_No.Format = "dd-mmm-yy"
        txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
        txtPVAmount.Text = ZERO
        txtProd_No.Enabled = True: txtMRR_No.Enabled = True: txtINV_No.Enabled = True
    End If
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
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,JOURNAL_HD_ID)" & _
                             " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                             ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & "," & labID.Caption & ")"
            rsTemplate_Details.MoveNext
        Loop
        StoreMemVars
        FillDetails
        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
            lstDetails.SetFocus
        End If

        Screen.MousePointer = 0
    End If
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    cboAcct_Code.Text = labAccountCode.Caption
    If cboAcct_Code.Text <> "" Then
        If SetAcctType(cboAcct_Code.Text) = "C" Then
            On Error Resume Next
            txtCredit.SetFocus
        Else
            On Error Resume Next
            txtDebit.SetFocus
        End If
    End If
    cmdFindAccount.ZOrder 1
    fraFindAccount.ZOrder 1
End Sub

Sub OkAccountSetCursor()
    If cboAcct_Code.Text <> "" Then
        If SetAcctType(cboAcct_Code.Text) = "C" Then
            txtCredit.SetFocus
        Else
            txtDebit.SetFocus
        End If
    End If
End Sub

Sub rsRefresh()
    If xJOURNALTYPE = "APJ" Then Me.Caption = "ACCOUNTS PAYABLE JOURNAL ENTRY"
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
    cmdAddJournal.ZOrder 1
    cmdAddJournal.Visible = False
    fraAddJournal.ZOrder 1
    fraAddJournal.Visible = False
    fraAddJournal.Enabled = False
    fraFindAccount.ZOrder 1
    cmdFindAccount.ZOrder 1
    fraFindAccount.Visible = False
    cmdFindAccount.Visible = False
    cmdShowPostRange.Visible = False
    cmdPrinting.ZOrder 1
End Sub

Sub SendToBackPV()
    cmdPV_Entry.ZOrder 1
    cmdPV_Entry.Visible = False
    picPV_Entry.ZOrder 1
    picPV_Entry.Visible = False
    picPV_Entry.Enabled = False
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
        txtInvoiceDate.Text = Format(Null2String(rsJournal_HD!invoicedate), "DD-MMM-YY")
        txtDueDate.Text = Format(Null2String(rsJournal_HD!DUEDATE), "DD-MMM-YY")
        'If Null2String(rsJournal_HD!paytype) = "" Or (Null2String(rsJournal_HD!paytype) <> Vendor_PayCode(Null2String(rsJournal_HD!VendorCode))) Then
        '    txtPayCode.Text = Vendor_PayCode(Null2String(rsJournal_HD!VendorCode))
        'Else
        txtPayCode.Text = Null2String(rsJournal_HD!paytype)
        'End If
        txtTerms.Text = Null2String(rsJournal_HD!TERMS)
        'If SetPayDesc(Null2String(rsJournal_HD!paytype)) = "" Then
        '    cboPayType.ListIndex = -1
        'Else
        cboPayType.Text = SetPayDesc(txtPayCode.Text)
        'End If
        If xJOURNALTYPE = "APJ" Then
            txtCode.Text = Null2String(rsJournal_HD!VendorCode)
            lblClass.Caption = Null2String(rsJournal_HD!Entity_Class)
            cboNameofVendor.Text = SetVendorName(txtCode.Text)
            CURRENT_VENDORCODE = Null2String(rsJournal_HD!VendorCode)
            txtAddress.Caption = SetVendorAddressNew(txtCode.Text, lblClass.Caption)
            cboBankName.Text = SetBankName(Null2String(rsJournal_HD!bankcode))
            txtCheckAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!AMOUNTPAID))
            APJInvoiceNo = GetAPJInvoiceNo(Null2String(txtVoucherNo.Text), xJOURNALTYPE)
            APJinvoicetype = GetAPJInvoiceType(Null2String(txtVoucherNo.Text), xJOURNALTYPE)
        End If

        If xJOURNALTYPE = "APJ" Then
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where JTYPE = '" & xJOURNALTYPE & "' AND PV_VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                picRefCDJ.ZOrder 0: picRefCDJ.Visible = True
                RefCDJ.Caption = "Ref CDJ# " & Null2String(rsCRJ_Detail!VOUCHERNO)
            Else
                picRefCDJ.ZOrder 1: picRefCDJ.Visible = False
                RefCDJ.Caption = ""
            End If
        End If

        txtBankCode.Text = Null2String(rsJournal_HD!bankcode)
        txtCheckNo.Text = Null2String(rsJournal_HD!CheckNo)
        txtCheckDate.Text = Null2String(rsJournal_HD!CheckDate)
        txtParticulars.Text = Null2String(rsJournal_HD!remarks)
        txtTotDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!Debit))
        txtTotCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!Credit))
        txtOutBalance.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!OUTBALANCE))
        txtAmountToPay.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!AMOUNTTOPAY))
        txtRemarks.Text = Null2String(rsJournal_HD!remarks)
        txtRemarks2.Text = Null2String(rsJournal_HD!remarks)
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

Private Sub cboAcct_Code_Change()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String

    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    GettheTaxBaseAmnt
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        cboATC.Text = VendorATC(txtCode.Text)
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
    End If
End Sub

Private Sub cboAcct_Code_Click()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cboATC_Click()
'Update By BTT: 09262008
    Dim RSATCRTE                                            As New ADODB.Recordset
    Dim xTOTAL                                              As Double
    GettheTaxBaseAmnt
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC.Text))
    If Not rsATC.EOF And Not rsATC.BOF Then
        txtRATE.Text = N2Str2Zero(rsATC!Rate)
        If NumericVal(txtRATE.Text) > 0 Then
            txtCredit.Text = Round(NumericVal(txtTaxBase.Text) * (NumericVal(txtRATE.Text) / 100), 2)
        End If
    End If
    Set rsATC = Nothing
End Sub

Private Sub cboBankName_Click()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cboBankName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboBankName_LostFocus()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cboBankName2_Click()
    txtBankCode.Text = SetBankCode(cboBankName2.Text)
End Sub

Private Sub cboBankName2_LostFocus()
    txtBankCode.Text = SetBankCode(cboBankName2.Text)
End Sub

Private Sub cboCustName_Change()
    txtCustCode.Text = SetCustomerCode(cboCustName.Text)
End Sub

Private Sub cboCustName_Click()
    txtCustCode.Text = SetCustomerCode(cboCustName.Text)
End Sub

Private Sub cboCustName_GotFocus()
    VBComBoBoxDroppedDown cboCustName
End Sub

Private Sub cboNameofVendor_Change()
    If bSelectEntity = True Then
        txtCode.Text = SetVendorCode(cboNameofVendor.Text)
        txtAddress.Caption = SetVendorAddressNew(txtCode.Text, lblClass.Caption)
    Else
        txtCode.Text = SetVendorCode(cboNameofVendor.Text)
        txtAddress.Caption = SetVendorAddress(txtCode.Text)
        lblClass.Caption = "V"
    End If
    If AddorEdit = "ADD" Or AddorEdit = "EDIT" Then
        txtPayCode.Text = Vendor_PayCode(txtCode.Text)
        cboPayType.Text = SetPayDesc(txtPayCode.Text)
    End If
End Sub

Private Sub cboNameofVendor_Click()
    If bSelectEntity = True Then
        txtCode.Text = SetVendorCode(cboNameofVendor.Text)
        txtAddress.Caption = SetVendorAddressNew(txtCode.Text, lblClass.Caption)
    Else
        txtCode.Text = SetVendorCode(cboNameofVendor.Text)
        txtAddress.Caption = SetVendorAddress(txtCode.Text)
        lblClass.Caption = "V"
    End If
    If AddorEdit = "ADD" Or AddorEdit = "EDIT" Then
        txtPayCode.Text = Vendor_PayCode(txtCode.Text)
        cboPayType.Text = SetPayDesc(txtPayCode.Text)
    End If
End Sub

Private Sub cboNameofVendor_GotFocus()
    VBComBoBoxDroppedDown cboNameofVendor
End Sub

Private Sub cboNameofVendor_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboPayTerm2_Change()
    txtPayCode.Text = SetPayCode(cboPayTerm2.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayTerm2.Text), txtInvoiceDate2.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayTerm2_Click()
    txtPayCode.Text = SetPayCode(cboPayTerm2.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayTerm2.Text), txtInvoiceDate2.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayTerm2_LostFocus()
    txtPayCode.Text = SetPayCode(cboPayTerm2.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayTerm2.Text), txtInvoiceDate2.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayType_Change()
    txtPayCode.Text = SetPayCode(cboPayType.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayType_Click()
    On Error Resume Next
    txtPayCode.Text = SetPayCode(cboPayType)
    If cboPayType <> "" Then
        txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType), txtInvoiceDate.Text), "DD-MMM-YY")
    End If
End Sub

Private Sub cboPayType_LostFocus()
    On Error Resume Next
    txtPayCode.Text = SetPayCode(cboPayType.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    SendToBack
    SendToBackPV
    'SendToBackGJ TO BE DELETED
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
    'SendToBack
    initMemvars
    FillDetails

    lstDetails.Enabled = False
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

Private Sub cmdAddJournal_Click()
    SendToBack
    cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
    fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
    fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
    AddorEdit = "ADD"
    InitJournal
    On Error Resume Next
    cboAcct_Code.SetFocus

    '    Set frmNewAMISJournalEntries = New frmAMISJournalEntries
    '    frmNewAMISJournalEntries.Caption = frmAMISJournalEntry_APJ.Caption
    '    Call frmNewAMISJournalEntries.LOADJOURNAL("APJ")
    '    frmNewAMISJournalEntries.Show 1
    '    On Error Resume Next
    '    frmNewAMISJournalEntries.cboAcct_Code.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDetails.Enabled = True
    AddorEdit = ""
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdCancelCO_Click()
    On Error GoTo ErrorCode:
    CANCEL_ANS = ""
    If Function_Access(LOGID, "Acess_CancelEntry", LocalAcess) = False Then Exit Sub
    '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If
    If MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Cancel Journal") = vbYes Then
        If xJOURNALTYPE = "APJ" Then
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
            NEW_LogAudit "C", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

            SQL_STATEMENT = "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "C", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

            'SQL_STATEMENT = "update AMIS_PV_Detail set status = 'C' where JTYPE= '" & xJOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            Dim rsPVDETAIL                                  As ADODB.Recordset
            Set rsPVDETAIL = New ADODB.Recordset
            rsPVDETAIL.Open "SELECT ID FROM AMIS_PV_DETAIL WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenForwardOnly
            If Not rsPVDETAIL.EOF And Not rsPVDETAIL.BOF Then
                Do While Not rsPVDETAIL.EOF
                    gconDMIS.Execute "UPDATE AMIS_PV_DETAIL SET STATUS='C' WHERE ID=" & rsPVDETAIL!ID & ""
                    rsPVDETAIL.MoveNext
                Loop
            End If
            Set rsPVDETAIL = Nothing
            NEW_LogAudit "C", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

            SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'C' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
            gconDMIS.Execute SQL_STATEMENT

            SQL_STATEMENT = "UPDATE AMIS_DETAIL SET STATUS = 'C' WHERE JTYPE='" & xJOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT

            SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'C' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
            gconDMIS.Execute SQL_STATEMENT

            SQL_STATEMENT = "UPDATE AMIS_DETAILS SET STATUS = 'C' WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO ='" & txtVoucherNo.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
        End If

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
    End If
    Screen.MousePointer = 0
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
    lstDetails.Enabled = False
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_HD!ID
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
    If xJOURNALTYPE = "APJ" Then
        frmAMISSearchAPJ.Show vbModal
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

'UPDATED BY: JUN-----------------------------------------------------
'DATE UPDATED: 08182009
'DESCRIPTION: NAVIGATIONAL ERROR THAT CAUSE TRIAL BALANCE NOT BALANCE
    SendToBack
    SendToBackPV
    'UPDATED BY: JUN-----------------------------------------------------

    rsJournal_HD.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdJournalCancel_Click()
    SendToBack
    StoreMemVars
    Picture1.Enabled = True
    JournalTAB.TabEnabled(1) = True
End Sub

Private Sub cmdJournalDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", LocalAcess) = False Then Exit Sub
    Dim xACCT_CODE                                          As String
    If labDetID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        xACCT_CODE = GetAcctcode(labDetID.Caption)

        If CheckARDetails(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), labDetID.Caption) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the AR details"
            Exit Sub
        End If

        If CheckAPDetails(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), labDetID.Caption) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the AP details"
            Exit Sub
        End If

        If CheckARPaymentDetails(xJOURNALTYPE, txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), labDetID.Caption) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the Payment details"
            Exit Sub
        End If

        If CheckAPPaymentDetails(xJOURNALTYPE, txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), labDetID.Caption) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the Payment details"
            Exit Sub
        End If

        SQL_STATEMENT = "delete from AMIS_Journal_Det where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "DT", txtVoucherNo, xJOURNALTYPE, labDetID.Caption
    End If
    Dim cnt                                                 As Integer
    Dim rsJournalDup                                        As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemno,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            'UPDATE DUE TO NEW AUDIT : BTT 08292008
            SQL_STATEMENT = "update AMIS_Journal_Det set JItemno = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            gconDMIS.Execute SQL_STATEMENT
            rsJournalDup.MoveNext
            '            NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        Loop
    End If
    FillDetails
    If xJOURNALTYPE = "VDJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    ElseIf xJOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    cmdJournalCancel.Value = True
    JournalTAB.TabEnabled(1) = True
    Picture1.Enabled = True
    If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
        lstDetails.SetFocus
    End If
End Sub

Private Sub cmdJournalSave_Click()
    On Error GoTo ErrorCode

    Dim str_MSG                                             As String
    Dim rsChartAccount                                      As ADODB.Recordset
    Dim xSJVOUCHERNO                                        As String

    str_MSG = "Error in saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If JournalEntries = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Accounts Payable Journal")
        MsgBox str_MSG, vbCritical, "Journal Entry Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    gconDMIS.CommitTrans

    If AddorEdit = "ADD" Then cmdAddJournal_Click Else cmdJournalCancel_Click
    If AddorEdit = "EDIT" Then
        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
            lstDetails.SetFocus
        End If
    End If
    JournalTAB.TabEnabled(1) = True
    Picture1.Enabled = True

    Screen.MousePointer = 0

    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function JournalEntries() As Boolean
    On Error GoTo ErrorCode
    Dim rsAPAccounts                                        As ADODB.Recordset
    Dim ValidateAccount                                     As New ADODB.Recordset
    'Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                                 As String
    If cboAcct_Code.Text = "" Or Setacctname(cboAcct_Code.Text) = "" Then
        MsgBox "Account Code and Description must have a value", vbInformation, "Error Encountered!"
        JournalEntries = True
        Exit Function
    End If
    'Updated by: ACL
    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        If cboATC.Text = "" Then
            MsgBox "ATC Code must have a value", vbInformation, "System Message!"
            JournalEntries = True
            Exit Function
        End If
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                       As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                     As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET            As Double
    Dim J_STATUS, J_JITEMNO                                 As String
    Dim J_ATC                                               As String
    Dim J_RATE, J_TAXBASE                                   As Double
    Dim xSJVOUCHERNO                                        As String

    xSJVOUCHERNO = xJOURNALTYPE + "-" + txtVoucherNo.Text
    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(Format(txtVoucherNo.Text, "000000"))
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_JITEMNO = N2Str2Null(Format(txtJItemNo.Text, "0000"))

    J_ACCT_CODE = N2Str2Null(cboAcct_Code.Text)
    J_ACCT_NAME = N2Str2Null(txtAcct_Name.Text)
    J_DEBIT = Round(NumericVal(txtDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtCredit.Text), 2)

    J_TAX = Round(NumericVal(txtTax.Text), 2)
    J_GROSS = Round(NumericVal(txtGrossAmt.Text), 2)
    J_NET = Round(NumericVal(txtNetAmt.Text), 2)
    J_STATUS = "'N'"
    J_ATC = N2Str2Null(cboATC.Text)
    J_RATE = NumericVal(txtRATE.Text)
    J_TAXBASE = NumericVal(txtTaxBase.Text)

    ' Update by BTT
'    If AddorEdit = "ADD" Then
'        Set ValidateAccount = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_dET WHERE ACCT_CODE=" & J_ACCT_CODE & " AND VOUCHERNO =" & J_VOUCHERNO & " AND JTYPE=" & J_JTYPE & "")
'        If ValidateAccount(0) = 1 Then
'            MsgBox "Duplicate Account entry is not allowed..", vbInformation, "Please verify your entry!"
'            JournalEntries = True
'            Exit Function
'        End If
'    End If

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,USERCODE,LASTUPDATE,ATC,RATE,TAXBASE,JOURNAL_HD_ID)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & LOGCODE & "','" & LOGDATE & "'," & J_ATC & "," & J_RATE & "," & J_TAXBASE & "," & labID.Caption & ")"
        gconDMIS.Execute SQL_STATEMENT
        JOURNAL_DETID = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_DET", J_JTYPE, "JTYPE")
        NEW_LogAudit "AA", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "DT", txtVoucherNo, xJOURNALTYPE, CStr(JOURNAL_DETID)


        If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
            SQL_STATEMENT = "UPDATE AMIS_JOURNAL_DET SET CREDIT = '" & NumericVal(txtTotDebit.Text) - NumericVal(txtCredit.Text) & "' WHERE VOUCHERNO='" & txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "' AND LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1)"
            gconDMIS.Execute SQL_STATEMENT
        End If

    Else
        SQL_STATEMENT = "update AMIS_Journal_Det set" & _
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
                        " grossamt = " & J_GROSS & "," & _
                        " netamt = " & J_NET & "," & _
                        " ATC = " & J_ATC & "," & _
                        " RATE = " & J_RATE & "," & _
                        " TAXBASE = " & J_TAXBASE & "," & _
                        " USERCODE = '" & LOGCODE & "'," & _
                        " LASTUPDATE = '" & LOGDATE & "'," & _
                        " status = " & J_STATUS & _
                        " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "DT", txtVoucherNo, xJOURNALTYPE, labDetID.Caption

        'SQL_STATEMENT = "UPDATE AMIS_AR SET ACCOUNT_CODE = '" & cboAcct_Code.Text & "' WHERE SJVOUCHERNO = '" & xSJVOUCHERNO & "'"
        'gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
            SQL_STATEMENT = "UPDATE AMIS_JOURNAL_DET SET CREDIT = '" & NumericVal(txtTotDebit.Text) - NumericVal(txtCredit.Text) & "' WHERE VOUCHERNO='" & txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "' AND LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND ACCT_CODE IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1)"
            gconDMIS.Execute SQL_STATEMENT
        End If
        'labDetID.Caption = ""
    End If
    CHECK_DETAILS
    FillDetails
    If xJOURNALTYPE = "VDJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    ElseIf xJOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        '        NEW_LogAudit "EE", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    JournalEntries = True

    Exit Function
ErrorCode:
    JournalEntries = False
End Function

'Upating Code       : AXP-0713200713:18
Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:

'UPDATED BY: JUN-----------------------------------------------------
'DATE UPDATED: 08182009
'DESCRIPTION: NAVIGATIONAL ERROR THAT CAUSE TRIAL BALANCE NOT BALANCE
    SendToBack
    SendToBackPV
    'UPDATED BY: JUN-----------------------------------------------------

    rsJournal_HD.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

'UPDATED BY: JUN-----------------------------------------------------
'DATE UPDATED: 08182009
'DESCRIPTION: NAVIGATIONAL ERROR THAT CAUSE TRIAL BALANCE NOT BALANCE
    SendToBack
    SendToBackPV
    'UPDATED BY: JUN-----------------------------------------------------

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
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Accounts Payable Journal")
        MsgBox str_MSG, vbCritical, "Posting Error "
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

    Dim KimyDKid                                            As Integer
    For KimyDKid = 1 To lstDetails.ListItems.Count
        If lstDetails.ListItems(KimyDKid).ListSubItems(2).Text = "" Then
            MsgBox "Warning: Invalid Account Description Encountered!", vbCritical, "Can not Post!"
            JournalPosting = True
            Exit Function
        End If
    Next

    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then
        JournalPosting = True
        Exit Function
    End If

    If BatchPost = True Then
        GoTo PerformBatchPosting
    End If
    If MsgBox("Are you sure you want to Post this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then

PerformBatchPosting:
        If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "OPB" Then
            'If COMPANY_CODE = "HPI" Then
            'Updated by: ACL 10202009
            If CheckIfOpen(xJOURNALTYPE, Trim(txtJDate.Text), Year(txtJDate.Text)) = False Then
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                JournalPosting = True
                Exit Function
            End If
        End If

        If NumericVal(txtTotDebit) <> NumericVal(txtTotCredit) Then
            MsgBox "Entry is not balanced. Posting of Entry Not Allowed.", vbInformation
            JournalPosting = True
            Exit Function
        End If

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
        Screen.MousePointer = 11

        If xJOURNALTYPE = "APJ" Then
            If Trim(cboNameofVendor.Text) = "" Then
                MsgBox "Warning: Posting is Not Allowed! Vendor Name is Required!", vbInformation, "Missing Fields"
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If
            If Trim(txtPayCode.Text) = "" Or Trim(cboPayType.Text) = "" Then
                MsgBox "Warning: Payment Term/Type is Required!", vbInformation, "Missing Fields"
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If
            'Dim DEALER_ITW_COMPENSATION                   As String
            Dim DEALER_ITW_EXPANDED                         As String
            'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
            DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

            If CheckWithholding = DEALER_ITW_EXPANDED Then
                If ReturnATCCode(CheckWithholding) = "" Then
                    MsgBox "ATC Code must have a value", vbInformation, "System Message!"
                    Screen.MousePointer = 0
                    JournalPosting = True
                    Exit Function
                End If
            End If

            If CheckGLSLARDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If

            If CheckGLSLARCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If

            If CheckGLSLAPDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If

            If CheckGLSLAPCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                Screen.MousePointer = 0
                JournalPosting = True
                Exit Function
            End If

            Dim rsCheckDetails                              As ADODB.Recordset
            Dim rsCheckCRJDetails                           As ADODB.Recordset
            Dim TotalCRJ_Credit                             As Double
            Set rsCheckDetails = New ADODB.Recordset
            Set rsCheckDetails = gconDMIS.Execute("Select Acct_Code, Credit from AMIS_Journal_Det DET INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE Where LEFT(AC.ACCTCODE,2)='21' AND IS_SCHEDULE_ACCNT=1 and Jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCheckDetails.EOF And Not rsCheckDetails.BOF Then
                rsCheckDetails.MoveFirst
                TotalCRJ_Credit = 0
                Do While Not rsCheckDetails.EOF
                    TotalCRJ_Credit = TotalCRJ_Credit + N2Str2Zero(rsCheckDetails!Credit)
                    rsCheckDetails.MoveNext
                Loop
                If TotalCRJ_Credit > 0 Then
                    If Round(NumericVal(txtAmountToPay.Text), 2) = Round(TotalCRJ_Credit, 2) Then
                        GoTo PostJournal
                    Else
                        Screen.MousePointer = 0
                        MsgBox "Warning: A/P Credit is not equal to Amount to Pay", vbCritical, "Error!"
                        JournalPosting = True
                        Exit Function
                    End If
                Else
                    GoTo PostJournal
                End If
            Else
                GoTo PostJournal
            End If
        End If
        Screen.MousePointer = 0
        'LogAudit "P", "JOURNAL ENTRY", txtJNo
        NEW_LogAudit "P", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
PostJournal:
        If NumericVal(txtTotDebit.Text) <> NumericVal(txtTotCredit.Text) Then
            MsgBox "Warning: Total Debit is not equal to Total Credit", vbCritical, "Cannot be Posted!"
            JournalPosting = True
            Exit Function
        End If
        If xJOURNALTYPE = "APJ" Then
            If TotalAPAmountToPay = 0 Then TotalAPAmountToPay = NumericVal(txtAmountToPay.Text)
            SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',USERCODE='" & LOGCODE & "',DATEPOSTED='" & LOGDATE & "',Balance = " & TotalAPAmountToPay & "- AmountPaid where jtype = 'APJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "P", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If

        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'P' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAIL SET STATUS = 'P' WHERE JTYPE='" & xJOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'P' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAILS SET STATUS = 'P' WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO ='" & txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        'NEW_LogAudit "P", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
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

        If xJOURNALTYPE = "APJ" Then
            Dim rsPVDETAIL                                  As ADODB.Recordset
            Set rsPVDETAIL = New ADODB.Recordset
            rsPVDETAIL.Open "SELECT ID FROM AMIS_PV_DETAIL WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenForwardOnly
            If Not rsPVDETAIL.EOF And Not rsPVDETAIL.BOF Then
                Do While Not rsPVDETAIL.EOF
                    gconDMIS.Execute "UPDATE AMIS_PV_DETAIL SET STATUS='P' WHERE ID=" & rsPVDETAIL!ID & ""
                    rsPVDETAIL.MoveNext
                Loop
            End If
            Set rsPVDETAIL = Nothing
            'NEW_LogAudit "P", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If

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
    SendToBack
    SendToBackPV
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
        If xJOURNALTYPE = "APJ" Then SaveReprintInformation xJOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If xJOURNALTYPE = "APJ" Then ShowReport "AccountsPayable", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "ACCOUNTS PAYABLE JOURNAL PRINTOUT", LOGDATE, False
        NEW_LogAudit "PX", "ACCOUNTS PAYABLE JOURNAL", "", "", "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPV_Entry_Click()
    SendToBackPV
    BringToFrontPV
    AddorEdit = "ADD"
    cmdPVDelete.Visible = False
    InitPV_Detail
    On Error Resume Next
    If xJOURNALTYPE = "APJ" Then
        On Error Resume Next
        txtPO_No.SetFocus
    Else
        On Error Resume Next
        txtMRR_No.SetFocus
    End If
End Sub

Private Sub cmdPVCancel_Click()
    SendToBackPV
    StoreMemVars
    JournalTAB.TabEnabled(0) = True
    Picture1.Enabled = True
End Sub

Private Sub cmdPVDelete_Click()

    If labPVID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If xJOURNALTYPE = "APJ" Then
        If MsgBox("Delete This PV Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            SQL_STATEMENT = "delete from AMIS_PV_Detail where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "PV", txtVoucherNo, xJOURNALTYPE, labPVID.Caption

            '            SQL_STATEMENT = "DELETE FROM AMIS_AP WHERE PV_ID = " & labPVID.Caption
            '            gconDMIS.Execute SQL_STATEMENT
        End If
    End If

    FillDetails
    rsRefresh
    On Error Resume Next
    JournalTAB.TabEnabled(0) = True
    rsJournal_HD.Find "id = " & labID.Caption
    cmdPVCancel.Value = True
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
                'MessagePop RecSaveError, "Error!", "Journal No. already exist!"
                'Exit Sub
                Call GetNewVoucherNo
            End If
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where invoiceno = '" & txtInvoiceNo.Text & "' and invoicedate = '" & CDate(txtInvoiceDate2.Text) & "' and invoicetype = '" & cboInvoiceType.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MessagePop RecSaveError, "Error!", "Invoice Transaction already Encoded!"
                Exit Sub
            End If
        End If
    End If
    If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "Error"
        txtJDate.SetFocus
        Exit Sub
    ElseIf txtInvoiceDate.Text = "" Or IsDate(txtInvoiceDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "Error"
        txtInvoiceDate.SetFocus
        Exit Sub
    End If
    If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "OPB" And xJOURNALTYPE <> "PDJ" Then
        '        If COMPANY_CODE = "HPI" Then
        'Updated by: ACL 10202009
        If CheckIfOpen(xJOURNALTYPE, Trim(txtJDate.Text), Year(txtJDate.Text)) = False Then
            MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            Exit Sub
        End If
        '        Else
        '            Set rsProfile = New ADODB.Recordset
        '            Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
        '            If Not rsProfile.EOF And Not rsProfile.BOF Then
        '                If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
        '                    If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
        '                        MessagePop RecSaveError, "Error!", "Warning: Journal Date is not valid in Accounting Period!"
        '                        'MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
        '                        Exit Sub
        '                    End If
        '                Else
        '                    MessagePop RecSaveError, "Error!", "Warning: Journal Date is not valid in Accounting Period!"
        '                    'MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
        '                    Exit Sub
        '                End If
        '            End If
        '        End If
    End If

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

    'If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "VDJ" Then
    If xJOURNALTYPE = "APJ" Then
        J_INVOICEDATE = N2Str2Null(txtInvoiceDate.Text)
        J_BALANCE = NumericVal(txtAmountToPay.Text)
        J_INVOICENO = N2Str2Null(GetAPJInvoiceNo(Null2String(txtVoucherNo.Text), xJOURNALTYPE))
        J_INVOICETYPE = N2Str2Null(GetAPJInvoiceType(Null2String(txtVoucherNo.Text), xJOURNALTYPE))
        J_AMOUNTPAID = 0
    End If
    J_DUEDATE = N2Str2Null(txtDueDate.Text)
    J_PAYTYPE = N2Str2Null(txtPayCode.Text)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_DEBIT = NumericVal(txtTotDebit.Text)
    J_CREDIT = NumericVal(txtTotCredit.Text)
    J_OUTBALANCE = NumericVal(txtOutBalance.Text)
    J_AMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
    J_STATUS = "'N'"

    J_CHECKNO = N2Str2Null(txtCheckNo.Text)
    J_TERMS = NumericVal(TERMS(N2Str2Null(txtPayCode.Text)))
    J_DEALER = "NULL"
    J_CHECKDATE = "NULL"
    J_BANKCODE = N2Str2Null(txtBankCode.Text)

    J_CUSTOMERNAME = "NULL"
    If xJOURNALTYPE = "APJ" Then
        J_CUSTOMERCODE = "'999999'"
        If Trim(txtCode.Text) = "" Then
            MsgBox "Invalid Supplier!", vbCritical, "Error"

            Exit Sub
        End If
        J_VENDORCODE = N2Str2Null(txtCode.Text)
    End If

    J_INVOICEAMT = NumericVal(txtInvoiceAmt.Text)
    J_REFNO = N2Str2Null(txtRefNo.Text)
    J_REFDATE = N2Date2Null(txtRefDate.Text)
    'If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "VDJ" Then
    If xJOURNALTYPE = "APJ" Then
        If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks.Text))
    End If
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
                        " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE,ENTITY_CLASS)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "','" & lblClass.Caption & "')"
        gconDMIS.Execute SQL_STATEMENT

        labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        NEW_LogAudit "A", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, labID.Caption
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
                        " remarks = " & J_REMARKS & ", USERCODE = '" & LOGCODE & "', LASTUPDATE = '" & LOGDATE & "', ENTITY_CLASS = '" & lblClass.Caption & "'" & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT

        'labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        NEW_LogAudit "E", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        SQL_STATEMENT = "update AMIS_Journal_Det set" & _
                        " jtype = " & J_JTYPE & "," & _
                        " jdate = " & J_JDATE & "," & _
                        " USERCODE = '" & LOGCODE & "'," & _
                        " LASTUPDATE = '" & LOGDATE & "'," & _
                        " jno = " & J_JNO & _
                        " where jtype = '" & PrevJType & "' and jno = '" & PrevJNo & "'"
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "E", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        SQL_STATEMENT = "update AMIS_PV_DETAIL set" & _
                        " jdate = " & J_JDATE & "" & _
                        " where jtype = '" & xJOURNALTYPE & "' and VOUCHERNO = '" & txtVoucherNo & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AR SET JDATE = " & J_JDATE & " where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAIL SET JDATE = " & J_JDATE & " WHERE JTYPE='" & xJOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AP SET JDATE = " & J_JDATE & " where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAILS SET JDATE = " & J_JDATE & " WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO ='" & txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

    End If
    If AddorEdit <> "ADD" Then
        rsJournal_HD.Find "jno = " & J_JNO
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "ACCOUNTS PAYABLE JOURNAL AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    rsRefresh
    rsJournal_HD.Find "jno = " & J_JNO
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    bSelectEntity = True
    Set frmNewEntity = New frmEntity
    Call frmNewEntity.LOADJOURNAL("APJ")
    frmNewEntity.Show 1
End Sub

Private Sub dtTo_LostFocus()
    If dtTo.Value < dtFrom.Value Then
        MessagePop InfoWarning, "Check Date", "Please check selected date."
        dtTo.SetFocus
        Exit Sub
    End If
End Sub

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    lblClass.Caption = strEntityClass
    txtCode.Text = strCode
    cboNameofVendor.Text = strAccountName
    txtAddress.Caption = SetVendorAddressNew(strCode, strEntityClass)
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


        If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
            Dim rsCRJ_Detail                                As ADODB.Recordset
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail where INVOICETYPE = '" & SetInvCode(cboInvoiceType.Text) & "' AND INVOICENO = '" & txtInvoiceNo.Text & "' AND INVOICEDATE = '" & txtInvoiceDate.Text & "' and status <> 'C'")
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                MsgBox "Warning: This Sales Journal is already link to Cash Receipts Voucher No. " & Null2String(rsCRJ_Detail!VOUCHERNO) & vbCrLf & _
                       "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
                Exit Sub
            End If
            'UPDATED BY: JUN --- DATE UPDATE: 11/17/2009 --- DESCRIPTION: THIS IS VALIDATE IF IT HAS A PAYMENT
            'If COMPANY_CODE <> "HGC" Then
            Call VALIDATE_UNPOSTING
            'End If
            'UPDATED BY: JUN
        End If
        If xJOURNALTYPE = "APJ" Then
            Dim rsCV_Detail                                 As ADODB.Recordset
            Set rsCV_Detail = New ADODB.Recordset
            Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail where PV_VoucherNo = '" & txtVoucherNo.Text & "' AND JTYPE= '" & xJOURNALTYPE & "' and status <> 'C'")
            If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
                MsgBox "Warning: This AP Journal is already link to Cash Disbursement Voucher No. " & Null2String(rsCV_Detail!VOUCHERNO) & vbCrLf & _
                       "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
                Exit Sub
            End If
        End If
        Screen.MousePointer = 11
        'UPDATED BY: JUN --- DATE UPDATE: 11/17/2009 --- DESCRIPTION: THIS IS TO VALIDATE IF IT HAS A PAYMENT
        'If COMPANY_CODE <> "HGC" Then
        Dim xVOUCHERNO                                      As String
        xVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text

        If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
            If VOUCHER_TO_VOUCHER_ADJ = True Then
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If

        SQL_STATEMENT = "update AMIS_JOURNAL_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "U", "ACCOUNTS PAYABLE JOURNAL", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        SQL_STATEMENT = "update AMIS_JOURNAL_DET set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        'SQL_STATEMENT = "update AMIS_PV_DETAIL set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        'gconDMIS.Execute SQL_STATEMENT
        Dim rsPVDETAIL                                      As ADODB.Recordset
        Set rsPVDETAIL = New ADODB.Recordset
        rsPVDETAIL.Open "SELECT ID FROM AMIS_PV_DETAIL WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenForwardOnly
        If Not rsPVDETAIL.EOF And Not rsPVDETAIL.BOF Then
            Do While Not rsPVDETAIL.EOF
                gconDMIS.Execute "UPDATE AMIS_PV_DETAIL SET STATUS='N' WHERE ID=" & rsPVDETAIL!ID & ""
                rsPVDETAIL.MoveNext
            Loop
        End If
        Set rsPVDETAIL = Nothing

        SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'N' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAIL SET STATUS = 'N' WHERE JTYPE='" & xJOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'N' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_DETAILS SET STATUS = 'N' WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO ='" & txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        Screen.MousePointer = 0
        'LogAudit "U", "ACCOUNTS PAYABLE JOURNAL", txtJNo
        Exit Sub
    End If
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

Private Sub Command4_Click()
    If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VCJ" Then
        If Trim(txtMRR_No.Text) = "" Then frmAMISSearchAPJ2.Show vbModal

    End If
    If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        SEARCH_TAB = 0
        If Trim(txtMRR_No.Text) = "" Then frmAMISSearchSJ2.Show vbModal
    End If
End Sub

Private Sub Command5_Click()
'    frmAMIS_UNAPPLIED_PAYMENT.Combo1.Text = "Customer Name"
'    frmAMIS_UNAPPLIED_PAYMENT.txtSearch = RTrim(LTrim(cboCustName.Text))
'    frmAMIS_UNAPPLIED_PAYMENT.Show
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
        If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
            fraFindAccount.Visible = True
            cmdFindAccount.Visible = True
            cmdFindAccount.ZOrder 0
            fraFindAccount.ZOrder 0
            fraFindAccount.Enabled = True
            DoEvents
            On Error Resume Next
            txtSearch.SetFocus
        ElseIf Me.ActiveControl.Name = "cboAccount" Then
            OkAccount
        ElseIf Me.ActiveControl.Name = "txtPO_No" And txtPO_No.Text = "" Then
            On Error Resume Next
            txtPO_No.SetFocus
        ElseIf Me.ActiveControl.Name = "txtCredit" And SetAcctType(cboAcct_Code.Text) = "C" And Val(txtCredit.Text) <= 0 And Val(txtDebit.Text) <= 0 Then
            On Error Resume Next
            txtCredit.SetFocus
        ElseIf Me.ActiveControl.Name = "txtDebit" And SetAcctType(cboAcct_Code.Text) = "D" And Val(txtDebit.Text) <= 0 And Val(txtCredit.Text) <= 0 Then
            On Error Resume Next
            txtDebit.SetFocus
        ElseIf Me.ActiveControl.Name = "txtGrossAmt" And NumericVal(txtGrossAmt.Text) <= 0 Then
            On Error Resume Next
            txtGrossAmt.SetFocus
        Else
            MoveKeyPress KeyCode
        End If
    Case vbKeyEscape
        If fraFindAccount.Visible = True Then
            If Me.ActiveControl.Name = "txtSearch" Then
                SendToBack
                SendToBackPV
                SendToBackTemplates
                StoreMemVars
            Else
                txtSearch.SetFocus
            End If
        Else
            If Picture1.Visible = True Then
                If Me.ActiveControl.Name = "txtSearchTemplates" Then
                    SendToBack
                    SendToBackPV
                    SendToBackTemplates
                    StoreMemVars
                ElseIf Me.ActiveControl.Name = "lstTemplates" Then
                    On Error Resume Next

                    txtSearchTemplates.SetFocus
                Else
                    SendToBack
                    SendToBackPV
                    SendToBackTemplates
                    StoreMemVars
                End If
                JournalTAB.TabEnabled(0) = True
                Picture1.Enabled = True
            End If
        End If
    Case vbKeyF2

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
                JournalTAB.Tab = 0
                JournalTAB.TabEnabled(1) = False
                Picture1.Enabled = False
                cmdAddJournal_Click
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
                    If Screen.ActiveForm.Name = "frmAMISJournalEntry_APJ" Then
                        JournalTAB.Tab = 1
                        cmdPV_Entry_Click
                        JournalTAB.TabEnabled(0) = False
                        Picture1.Enabled = False
                        txtMRR_No.BackColor = &HFFFFFF
                        txtINV_No.BackColor = &HFFFFFF
                    End If
                End If
            End If
        Else
            ShowInvoiceApp SetInvCode(cboInvoiceType), txtInvoiceNo.Text
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
                JournalTAB.Tab = 0
                fraFindAccount.ZOrder 1: cmdFindAccount.ZOrder 1
                fraFindAccount.Visible = False: cmdFindAccount.Visible = False: BringToFrontTemplates
                txtSearchTemplates.SetFocus
            End If
        End If
    Case vbKeyF11
        SendToBack
        SendToBackPV
        SendToBackTemplates
        lblVoucherNo.Caption = ""
        Picture1.Enabled = False
        JournalTAB.Enabled = False
        picBatchImport.Visible = True
        picBatchImport.ZOrder 0
        dtFrom.Value = Format(JOURNALFIRSTTRANS(xJOURNALTYPE), "mm/dd/yyyy")
        dtTo.Value = Format(JOURNALLASTTRANS(xJOURNALTYPE), "mm/dd/yyyy")
        On Error Resume Next
    Case vbKeyF12
        '        If Null2String(rsJournal_HD!Status) = "C" Then

        '            If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub

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
            Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCOUNTS PAYABLE JOURNAL")
        End If
    End If
    If Shift = 2 Then
        If KeyCode = vbKeyA Then cmdAddAccount_Click
        If KeyCode = vbKeyJ Then
            If JournalTAB.Tab = 1 Then JournalTAB.Tab = 0
        End If
        If KeyCode = vbKeyD Then
            If JournalTAB.Tab = 0 Then JournalTAB.Tab = 1
        End If
        If KeyCode = vbKeyF12 Then
            ' TEMPORARY to close AP for HAI
            '            If xJOURNALTYPE = "APJ" Then
            '                If MsgBox("Set this AP Transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AP") = vbYes Then
            '                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtAmountToPay) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='APJ'")
            '                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
            '                End If
            '            End If
            '
            '            ' TEMPORARY to close AR for HAI
            '            If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
            '                If MsgBox("Set this AR transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AR") = vbYes Then
            '                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtInvoiceAmt) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='SJ'")
            '                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
            '                End If
            '            End If
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
    Frame1.Enabled = False: SendToBack: SendToBackPV: SendToBackTemplates
    'SendToBackGJ: To be deleted
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    picPayables.Top = 1200
    picDisbursement.Top = 1200
    picReceivable.Top = 420
    'Frame1.Top = 90
    fraATC.Visible = False
    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HQA" Or COMPANY_CODE = "HNE" Then
        bSelectEntity = True
        picNewEntity.Visible = True
    Else
        bSelectEntity = False
    End If
    labCheckAmt.Visible = False: txtCheckAmt.Visible = False: txtParticulars.Height = 795

    JournalTAB.Tab = 0
    If xJOURNALTYPE = "APJ" Then
        LocalAcess = "ACCOUNTS PAYABLE JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False
        Me.Caption = "ACCOUNTS PAYABLE JOURNAL DATA ENTRY"
        labSupplierPayTo = "Supplier Code"
        picPayables.Visible = True: picPayables.ZOrder 0: picPayables.Enabled = True
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labPV1.Caption = "PO Number": labPV2.Caption = "MRR Number"
        labPV3.Caption = "Invoice Number": labPV4.Caption = "Product Number"
        labTax.Caption = "Input Tax": RefCRJ.Visible = False
    End If
    'InitGrid
    InitCbo
    initMemvars
    txtSearch.Text = "": txtSearchTemplates.Text = ""
    rsRefresh
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveLast
    End If
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
    LocalAcess = ""
End Sub

Private Sub JournalTAB_Click(PreviousTab As Integer)
    If Picture1.Visible = True Then
        If JournalTAB.Tab = 0 Then

            If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
                lstDetails.SetFocus
            End If
        End If
        If JournalTAB.Tab = 1 Then
            If lstPV_Detail.ListItems.Count > 0 And lstPV_Detail.Enabled = True Then
                lstPV_Detail.SetFocus
            End If
        End If
    End If
End Sub

Private Sub Label4_Click()
    ShowInvoiceApp SetInvCode(cboInvoiceType), txtInvoiceNo.Text
End Sub

Private Sub labF7_Click()

End Sub

Private Sub labF10_Click()

End Sub

Private Sub lstAccounts_DblClick()
    labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
    OkAccount
End Sub

Private Sub lstAccounts_GotFocus()
    On Error Resume Next
    labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labAccountCode.Caption = Item: cboAcct_Code.Text = Item
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
        OkAccount
    End If
End Sub

Private Sub lstDetails_DblClick()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"
        Else
            If kcnt > 0 Then
                AddorEdit = "EDIT"
                cmdJournalDelete.Visible = True
                BringToFront

                Call StoreJournalEntry(lstDetails.SelectedItem.SubItems(5))
                '                Set frmNewAMISJournalEntries = New frmAMISJournalEntries
                '                frmNewAMISJournalEntries.Caption = frmAMISJournalEntry_APJ.Caption
                '                Call frmNewAMISJournalEntries.LOADJOURNAL("APJ")
                '                Call frmNewAMISJournalEntries.StoreJournalEntry(lstDetails.SelectedItem.SubItems(5))
                '                frmNewAMISJournalEntries.Show 1
                On Error Resume Next
                'txtGrossAmt.SetFocus
                OkAccountSetCursor
            End If
        End If
    End If
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labDetID.Caption = lstDetails.SelectedItem.SubItems(5)
End Sub

Private Sub lstDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstDetails_DblClick
        If Me.ActiveControl.Name = "txtDebit" Then SendKeys MOVEUP
    End If
End Sub

Private Sub lstDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If kcnt > 0 Then
            If Null2String(rsJournal_HD!Status) = "C" Then
                MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            ElseIf Null2String(rsJournal_HD!Status) = "P" Then
                MsgBox "Journals are Already Posted" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            Else
                AddorEdit = "EDIT"
                cmdJournalDelete.Visible = True
                BringToFront
                StoreJournalEntry (lstDetails.SelectedItem.SubItems(5))
                cmdJournalDelete_Click
            End If
        End If
    ElseIf KeyCode = vbKeyF2 Then
        LOAD_ARAP_DETAILS
    End If
End Sub

Private Sub lstPV_Detail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPV_Detail
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

Private Sub lstPV_Detail_DblClick()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"
        Else
            If Jcnt > 0 Then
                AddorEdit = "EDIT"
                cmdPVDelete.Visible = True
                BringToFrontPV
                StorePVEntry (lstPV_Detail.SelectedItem.SubItems(6))
                Picture1.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lstPV_Detail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstPV_Detail_DblClick
End Sub

Private Sub lstTemplates_DblClick()
    SendToBack
    SendToBackPV
    SendToBackTemplates
    On Error Resume Next
    InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub lstTemplates_KeyPress(KeyAscii As Integer)
    SendToBack
    SendToBackPV
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
            JournalTAB.Tab = 0
            JournalTAB.TabEnabled(1) = False
            Picture1.Enabled = False
            cmdAddJournal_Click
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
            JournalTAB.Tab = 0
            fraFindAccount.ZOrder 1: cmdFindAccount.ZOrder 1
            fraFindAccount.Visible = False: cmdFindAccount.Visible = False: BringToFrontTemplates
            txtSearchTemplates.SetFocus
        End If
    End If
End Sub

Private Sub ShortcutCaption5_GotFocus()
    SendToBack
    SendToBackPV
    SendToBackTemplates
    cmdShowPostRange.Visible = True
    cmdShowPostRange.ZOrder 0
    On Error Resume Next
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then labPosted.Visible = False Else labPosted.Visible = True
        '        If lblDatePosted.Visible = True Then lblDatePosted.Visible = False Else lblDatePosted.Visible = True
    End If
End Sub

Private Sub txtAmountToPay_GotFocus()
    If Val(txtAmountToPay.Text) = 0 Then txtAmountToPay.Text = "" Else txtAmountToPay.Text = Format(txtAmountToPay.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtAmountToPay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtAmountToPay_LostFocus()
    If txtAmountToPay.Text = "" Then txtAmountToPay.Text = "0.00" Else txtAmountToPay.Text = Format(txtAmountToPay.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtBankCode_Change()
    If xJOURNALTYPE = "APJ" Then cboBankName.Text = SetBankName(txtBankCode.Text)
End Sub

Private Sub txtBankCode_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCheckDate_GotFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCheckDate_LostFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtCredit_GotFocus()
    If NumericVal(txtDebit.Text) = 0 Then
        If Val(txtCredit.Text) = 0 Then
            If NumericVal(txtNetAmt.Text) > 0 Then
                txtDebit.Text = ZERO
                txtCredit.Text = NumericVal(txtNetAmt.Text)
            Else
                If OUTBALANCE > 0 And TOTDEBIT > 0 Then
                    txtCredit.Text = OUTBALANCE
                    txtDebit.Text = ZERO
                Else
                    txtCredit.Text = ""
                End If
            End If
        Else
            txtCredit.Text = NumericVal(txtCredit.Text)
        End If
    Else
        txtCredit.Text = ZERO
    End If
End Sub

Private Sub txtCredit_LostFocus()
    If txtCredit.Text = "" Then txtCredit.Text = 0
End Sub

Private Sub txtDebit_GotFocus()
    If NumericVal(txtCredit.Text) = 0 Then
        If NumericVal(txtDebit.Text) = 0 Then
            If NumericVal(txtNetAmt.Text) > 0 Then
                txtDebit.Text = NumericVal(txtNetAmt.Text)
            Else
                If txtAcct_Name.Text = "OUTPUT TAX" And xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
                    txtDebit.Text = ZERO: txtCredit.Text = OUTBALANCE
                Else
                    If OUTBALANCE > 0 And TOTCREDIT > 0 Then
                        txtCredit.Text = ZERO: txtDebit.Text = OUTBALANCE
                    Else
                        txtDebit.Text = ""
                    End If
                End If
            End If
        Else
            txtDebit.Text = NumericVal(txtDebit.Text)
        End If
    Else
        txtDebit.Text = ZERO
    End If
End Sub

Private Sub txtDebit_LostFocus()
    If txtDebit.Text = "" Then txtDebit.Text = 0
End Sub

Private Sub txtGrossAmt_Change()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtTax.Text = Round((NumericVal(txtGrossAmt.Text) / 1.12) * 0.12, 2)
        txtNetAmt.Text = NumericVal(txtGrossAmt.Text) - NumericVal(txtTax.Text)
    Else
        txtTax.Text = 0: txtNetAmt.Text = 0
    End If
End Sub

Private Sub txtGrossAmt_GotFocus()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtGrossAmt.Text = NumericVal(txtGrossAmt.Text)
    Else
        txtGrossAmt.Text = ""
    End If
End Sub

Private Sub txtGrossAmt_LostFocus()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtGrossAmt.Text = ToDoubleNumber(txtGrossAmt.Text)
    End If
End Sub

Private Sub txtINV_No_GotFocus()
    On Error Resume Next
    If xJOURNALTYPE = "CDJ" Then
        If txtMRR_No.Text = "" Then txtMRR_No.SetFocus
    End If
End Sub

Private Sub txtINV_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtInvoiceAmt_GotFocus()
    txtInvoiceAmt.Text = NumericVal(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtInvoiceAmt_LostFocus()
    txtInvoiceAmt.Text = ToDoubleNumber(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceDate_Change()
    On Error Resume Next
    If IsDate(txtInvoiceDate.Text) = True Then
        txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
    End If
End Sub

Private Sub txtInvoiceDate_GotFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate2_GotFocus()
    txtInvoiceDate2.Text = Format(txtInvoiceDate2.Text, "MM-DD-YYYY")
End Sub

Private Sub txtInvoiceDate2_LostFocus()
    If txtInvoiceDate2.Text <> "" Then
        If IsDate(txtInvoiceDate2.Text) = True Then
            txtInvoiceDate2.Text = Format(txtInvoiceDate2.Text, "DD-MMM-YY")
        Else
            'MsgBoxXP "Invalid Invoice Date!", "Error", XP_OKOnly, msg_Exclamation
            MessagePop RecSaveError, "Error", "Invalid Invoice Date!"
            On Error Resume Next
            txtInvoiceDate.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtJDate_GotFocus()
    txtJDate.Text = Format(txtJDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtJDate_LostFocus()
    txtJDate.Text = Format(txtJDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtMRR_No_Change()
'    Dim theJtype                                       As String
'    Dim SQL                                            As String
'    Dim RS                                             As New ADODB.Recordset
'    If xJOURNALTYPE = "CDJ" Then
'        Set rsJournal_HD2 = New ADODB.Recordset
'        'Update BTT : 09012008
'        If txtPO_No.Text = "VPJ" Then
'            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'VPJ'")
'        Else
'            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'APJ'and status='P'")
'        End If
'        If Not rsJournal_HD2.EOF And Not rsJournal_HD2.BOF Then
'            theJtype = Null2String(rsJournal_HD2!jtype)
'            txtINV_No.Text = Null2String(rsJournal_HD2!JDATE)
'            txtProd_No.Text = Null2String(rsJournal_HD2!DUEDATE)
'            txtPVAmount.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD2!BALANCE))
'            If theJtype = "VPJ" Then
'                Set RS = New ADODB.Recordset
'                Set RS = gconDMIS.Execute("SELECT acct_code from AMIS_journal_det where voucherno='" & txtMRR_No.Text & "' and jtype ='VPJ'")
'                If Not RS.EOF And Not RS.BOF Then
'                    CDJ_AP = N2Str2Null(RS!ACCT_CODE)
'                    IsVPJ = True
'                End If
'            Else
'                CDJ_AP = ReturnAP_AccountCode("AP")
'                CDJ_IS_FROM_AP = True
'                IsVPJ = False
'            End If
'        Else
'            txtINV_No.Text = ""
'            txtProd_No.Text = ""
'            txtPVAmount.Text = ZERO
'            CDJ_IS_FROM_AP = False
'        End If
'    End If
End Sub

Private Sub txtMRR_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VCJ" Then
        If KeyAscii = 13 Then
            If Trim(txtMRR_No.Text) = "" Then
                frmAMISSearchAPJ2.LOADJOURNAL (xJOURNALTYPE)
                frmAMISSearchAPJ2.Show vbModal
            End If
        End If
    End If
    If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        If KeyAscii = 13 Then
            SEARCH_TAB = 0
            If Trim(txtMRR_No.Text) = "" Then
                frmAMISSearchSJ2.LOADJOURNAL (xJOURNALTYPE)
                frmAMISSearchSJ2.Show vbModal
            End If
        End If
    End If
End Sub

Private Sub txtMRR_No_LostFocus()
    If xJOURNALTYPE = "CRJ" And AddorEdit <> "EDIT" Then
        Dim rsAR_Accounts                                   As New ADODB.Recordset
        Set rsAR_Accounts = New ADODB.Recordset
        Set rsAR_Accounts = gconDMIS.Execute("select Acct_Code from AMIS_Journal_Det Where (Left(Acct_Code,5) = '11-02' or Left(Acct_Code,5) = '11-03' or Left(Acct_Code,5) = '11-04') and  VoucherNo = '" & txtVoucherNo.Text & "' AND Jtype = '" & xJOURNALTYPE & "'")
        If Not rsAR_Accounts.EOF And Not rsAR_Accounts.BOF Then
            cboARTag.Text = Setacctname(rsAR_Accounts!ACCT_CODE)
        End If
        Set rsAR_Accounts = Nothing
    End If
End Sub

Private Sub txtParticulars_GotFocus()
    If txtParticulars.Text = "Pls Type Your Message Here!" Then txtParticulars.Text = ""
End Sub

Private Sub txtParticulars_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtParticulars.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtParticulars_LostFocus()
    If txtParticulars.Text = "" Then txtParticulars.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtPayCode_Change()
    If SetPayDesc(txtPayCode.Text) = "" Then
        cboPayType.ListIndex = -1
    Else
        cboPayType.Text = SetPayDesc(txtPayCode.Text)
    End If
End Sub

Private Sub txtPayCode_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPayCode_LostFocus()
    If SetPayDesc(txtPayCode.Text) = "" Then
        cboPayType.ListIndex = -1
    Else
        cboPayType.Text = SetPayDesc(txtPayCode.Text)
    End If
End Sub

Private Sub txtPO_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtProd_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtProd_No_LostFocus()
    If xJOURNALTYPE = "CRJ" Then
        If IsDate(txtProd_No) = False Then
            MsgBox "Invalid date!", vbExclamation, "WARNING"
            txtProd_No.Text = ""
        End If
    End If
End Sub

Private Sub txtPVAmount_GotFocus()
    If NumericVal(txtPVAmount.Text) = 0 Then txtPVAmount.Text = ""
End Sub

Private Sub txtPVAmount_LostFocus()
    If NumericVal(txtPVAmount.Text) > 0 Then txtPVAmount.Text = ToDoubleNumber(txtPVAmount.Text)
End Sub

Private Sub txtRefDate_GotFocus()
    txtRefDate.Text = Format(txtRefDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtRefDate_LostFocus()
    If txtRefDate.Text <> "" Then
        If IsDate(txtRefDate.Text) = True Then
            txtRefDate.Text = Format(txtRefDate.Text, "DD-MMM-YY")
        Else
            MessagePop RecSaveError, "Error", "Invalid Reference Date!"
            On Error Resume Next
            txtRefDate.SetFocus
            Exit Sub
        End If
    End If
    If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        On Error Resume Next
        cboBankName2.SetFocus
    End If
End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtRemarks.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtRemarks_LostFocus()
    If txtRemarks.Text = "" Then txtRemarks.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtRemarks2_GotFocus()
    If txtRemarks2.Text = "Pls Type Your Message Here!" Then txtRemarks2.Text = ""
End Sub

Private Sub txtRemarks2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtRemarks2.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtRemarks2_LostFocus()
    If txtRemarks2.Text = "" Then txtRemarks2.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        Call FillSearchGrid(txtSearch.Text)
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

Private Sub txtTax_Change()
    txtNetAmt.Text = ToDoubleNumber(NumericVal(txtGrossAmt.Text) - NumericVal(txtTax.Text))
End Sub

Private Sub txtTax_GotFocus()
    If Val(txtTax.Text) = 0 Then txtTax.Text = ""
End Sub

Private Sub txtTax_LostFocus()
    If SetAcctType(cboAcct_Code.Text) = "C" Then
        txtDebit.Text = ZERO
        txtCredit.Text = ToDoubleNumber(txtNetAmt.Text)
        On Error Resume Next
        txtCredit.SetFocus
    Else
        If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" And txtAcct_Name.Text = "OUTPUT TAX" Then
            txtDebit.Text = ZERO
            txtCredit.Text = ToDoubleNumber(txtNetAmt.Text)
            On Error Resume Next
            txtCredit.SetFocus
        Else
            txtCredit.Text = ZERO
            txtDebit.Text = ToDoubleNumber(txtNetAmt.Text)
            On Error Resume Next
            txtDebit.SetFocus
        End If
    End If
End Sub

Private Sub txtTaxBase_Change()
' Update By BTT : 09262008
    If NumericVal(txtRATE.Text) > 0 Then
        txtCredit.Text = Round(NumericVal(txtTaxBase.Text) * (NumericVal(txtRATE.Text) / 100), 2)
    End If

End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Format(txtVoucherNo, "000000")
End Sub

Sub GettheTaxBaseAmnt()
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    If xJOURNALTYPE = "APJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'APJ'"
    Else
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'CDJ'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTaxBase.Text = NumericVal(RS!SumDebit)
    End If
    Set RS = Nothing
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
Function ReturnAccountDescription(XXX As String)
    Dim RSACCT                                              As New ADODB.Recordset
    Set RSACCT = gconDMIS.Execute("SELECT * from AMIS_chartaccount where acctcode='" & XXX & "'")
    With RSACCT
        If Not .EOF And Not .BOF Then
            cboARTag.AddItem Null2String(RSACCT!DESCRIPTION)
        End If
    End With
    Set RSACCT = Nothing
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
                      "WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND LEFT(ACCT_CODE,5) IN('11-02','11-03','11-04')", gconDMIS, adOpenKeyset
    If Not rsCOUNT_CODE.EOF And Not rsCOUNT_CODE.BOF Then
        'THIS IS FOR ACCT_CODE GREATER THAN ONE IN ONE VOUCHERNO
        If NumericVal(rsCOUNT_CODE!COUNT_CODE) > 1 Then
            Set rsAR_VOUCHER = New ADODB.Recordset
            rsAR_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,ACCT_CODE " & _
                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03','11-04') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenKeyset
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

                    xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!VOUCHERNO))
                    xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))
                    xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                    xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
                    xAMOUNT_PAID = GET_AR_AMOUNT_CREDIT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
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
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03','11-04') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
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

                xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))
                xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
                xAMOUNT_PAID = GET_AR_AMOUNT_CREDIT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!JTYPE), Null2String(rsAR_VOUCHER!ACCT_CODE))
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

Function GET_AR_AMOUNT_CREDIT(xVOUCHERNO As String, xJType As String, xACCT_CODE As String) As Double
'UPDATE BY: JUN --- DATE UPDATED: 11/19/2009 --- DESCRIPTION: THIS IS TO SUM THE AR WITH THE SPECIFIC ACCOUNT CODE

    Dim rsGET_AR_AMOUNT                                     As ADODB.Recordset
    Set rsGET_AR_AMOUNT = New ADODB.Recordset
    rsGET_AR_AMOUNT.Open "SELECT ROUND(SUM(DET.CREDIT),2) AS SUM_CREDIT " & _
                         "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                         "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsGET_AR_AMOUNT.EOF And Not rsGET_AR_AMOUNT.BOF Then
        GET_AR_AMOUNT_CREDIT = NumericVal(rsGET_AR_AMOUNT!SUM_CREDIT)
    Else
        GET_AR_AMOUNT_CREDIT = 0
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
                'gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & xAR_BALANCE & " WHERE INVOICENO = " & xINVOICENO & "  AND INVOICETYPE = " & xINVOICETYPE & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
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
                      "WHERE HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.JTYPE = 'CRJ' AND DET.DEBIT <> 0 AND LEFT(ACCT_CODE,5) IN('11-02','11-03','11-04')", gconDMIS, adOpenKeyset

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

Function GET_AP_VOUCHERNO() As Boolean
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
    'rsAP_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,ACCT_CODE " & _
     "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
     "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    rsAP_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.CHECKNO AS INVOICENO,HD.CHECKDATE AS INVOICEDATE,HD.DUEDATE,ACCT_CODE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN (SELECT JTYPE,VOUCHERNO,ACCT_CODE FROM AMIS_JOURNAL_DET DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1 AND TRANTYPE1='UNPRESENTED') DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsAP_VOUCHER.EOF And Not rsAP_VOUCHER.BOF Then
        'Do While Not rsAP_VOUCHER.EOF
        xVOUCHERNO = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE) & "-" & Null2String(rsAP_VOUCHER!VOUCHERNO))
        xJdate = N2Str2Null(Null2String(rsAP_VOUCHER!JDATE))
        xJType = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE))
        xDUEDATE = N2Str2Null(Null2String(rsAP_VOUCHER!DUEDATE))
        If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!VendorCode))
            'xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAP_VOUCHER!VendorCode)))
            xCUST_NAME = N2Str2Null(SetVendorName(Null2String(rsAP_VOUCHER!VendorCode)))
        Else
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!CustomerCode))
            'xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAP_VOUCHER!CustomerCode)))
            xCUST_NAME = N2Str2Null(SetVendorName(Null2String(rsAP_VOUCHER!CustomerCode)))
        End If

        xInvoiceNo = N2Str2Null(Null2String(rsAP_VOUCHER!INVOICENO))
        xInvoiceType = "OI"
        xInvoicedate = N2Str2Null(Null2String(rsAP_VOUCHER!invoicedate))
        xAMOUNT_TO_PAY = GET_AP_AMOUNT(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!JTYPE), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xAMOUNT_PAID = GET_AMOUNT_PAID(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!JTYPE), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
        xACCT_CODE = N2Str2Null(Null2String(rsAP_VOUCHER!ACCT_CODE))
        xLAST_UPDATED = N2Str2Null(LOGDATE)

        SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                        "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
        gconDMIS.Execute SQL_STATEMENT
        '            If xJOURNALTYPE = "CDJ" Then
        '                gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid=" & xAMOUNT_PAID & ",Balance = " & xBAL & " where JTYPE ='" & xJOURNALTYPE & "' And VOUCHERNO = " & N2Str2Null(rsAP_VOUCHER!VOUCHERNO)
        '            End If
        '    rsAP_VOUCHER.MoveNext
        'Loop
        GET_AP_VOUCHERNO = True
    Else
        GET_AP_VOUCHERNO = False
    End If
    Set rsAP_VOUCHER = Nothing
End Function

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
                           "WHERE DET.ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " AND HD.JTYPE = " & N2Str2Null(xJType) & " AND HD.VOUCHERNO NOT IN (SELECT VOUCHERNO FROM AMIS_CV_DETAIL WHERE VOUCHERNO=HD.VOUCHERNO AND CV_JTYPE=HD.JTYPE) AND HD.STATUS = 'P' GROUP BY HD.VOUCHERNO,HD.JTYPE) X WHERE VOUCHERNO=" & N2Str2Null(xVOUCHERNO) & " AND JTYPE=" & N2Str2Null(xJType) & "", gconDMIS, adOpenKeyset
    If Not rsGET_AMOUNT_PAID.EOF And Not rsGET_AMOUNT_PAID.BOF Then
        GET_AMOUNT_PAID = rsGET_AMOUNT_PAID!SUM_DEBIT
    Else
        GET_AMOUNT_PAID = 0
    End If
    Set rsGET_AMOUNT_PAID = Nothing
End Function

'Function GET_AMOUNT_PAID(xVOUCHERNO As String, xJtype As String, xACCT_CODE As String) As Double
'    Dim rsGET_AMOUNT_PAID As ADODB.Recordset
'    Set rsGET_AMOUNT_PAID = New ADODB.Recordset
'    rsGET_AMOUNT_PAID.Open "SELECT AMIS_JOURNAL_HD.VOUCHERNO,AMIS_JOURNAL_HD.JTYPE,AMIS_JOURNAL_HD.STATUS,ISNULL((SELECT ROUND(SUM(AMOUNT),2) FROM AMIS_CV_DETAIL WHERE VOUCHERNO='" & xVOUCHERNO & "' AND CV_JTYPE='" & xJtype & "' AND J_CLASS='" & xACCT_CODE & "'),0) AS AMOUNT FROM AMIS_JOURNAL_HD WHERE AMIS_JOURNAL_HD.VOUCHERNO='" & xVOUCHERNO & "' AND AMIS_JOURNAL_HD.JTYPE='" & xJtype & "' AND AMIS_JOURNAL_HD.STATUS = 'P'", gconDMIS, adOpenKeyset
'    If Not rsGET_AMOUNT_PAID.EOF And Not rsGET_AMOUNT_PAID.BOF Then
'        GET_AMOUNT_PAID = rsGET_AMOUNT_PAID!amount
'    Else
'        GET_AMOUNT_PAID = 0
'    End If
'    Set rsGET_AMOUNT_PAID = Nothing
'End Function

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
            'gconDMIS.Execute "UPDATE AMIS_AP SET AMOUNTPAID = " & xSUM_PAYMENT & ", BALANCE = " & xAP_BALANCE & " WHERE VENDOR_CODE = " & xVENDORCODE & " AND VOUCHERNO= " & xPV_VOUCHERNO & " AND ACCT_CODE = " & xACCT_CODE & ""
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

Function ReturnWithholdingTax(XXX As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnWithholdingTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Function ReturnInPutTax()
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Private Sub GetNewVoucherNo()
    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
End Sub

Function ReturnATCCode(XXX As String)
    Dim rsATC                                               As ADODB.Recordset
    Set rsATC = New ADODB.Recordset
    Set rsATC = gconDMIS.Execute("SELECT ATC FROM AMIS_JOURNAL_DET WHERE ACCT_CODE ='" & XXX & "' AND VOUCHERNO='" & txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "'")
    If Not rsATC.EOF And Not rsATC.BOF Then
        ReturnATCCode = Null2String(rsATC!ATC)
    End If
    Set rsATC = Nothing
End Function
Function OUTBAL()
    Dim rsOUTBAL                                            As ADODB.Recordset
    Set rsOUTBAL = New ADODB.Recordset
    Set rsOUTBAL = gconDMIS.Execute("SELECT OUTBALANCE FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "'")
    If Not rsOUTBAL.EOF And Not rsOUTBAL.BOF Then
        OUTBAL = Null2String(rsOUTBAL!OUTBALANCE)
    End If
    Set OUTBAL = Nothing
End Function

Function CheckWithholding()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    Dim rsCheckWithholding                                  As ADODB.Recordset
    Set rsCheckWithholding = New ADODB.Recordset
    Set rsCheckWithholding = gconDMIS.Execute("SELECT DISTINCT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE VOUCHERNO='" & txtVoucherNo.Text & "' AND JTYPE='" & xJOURNALTYPE & "' AND ACCT_CODE = '" & DEALER_ITW_EXPANDED & "'")
    If Not rsCheckWithholding.EOF And Not rsCheckWithholding.BOF Then
        CheckWithholding = Null2String(rsCheckWithholding!ACCT_CODE)
    End If
    Set rsCheckWithholding = Nothing
End Function

Function GetAPJInvoiceNo(XXX As String, YYY As String) As String
    Dim rsGetAPJInvoiceNo                                   As ADODB.Recordset
    Set rsGetAPJInvoiceNo = New ADODB.Recordset
    rsGetAPJInvoiceNo.Open "SELECT INVOICENO FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & XXX & "' AND JTYPE = '" & YYY & "'", gconDMIS, adOpenForwardOnly
    If Not rsGetAPJInvoiceNo.EOF And Not rsGetAPJInvoiceNo.BOF Then
        GetAPJInvoiceNo = Null2String(rsGetAPJInvoiceNo!INVOICENO)
    Else
        GetAPJInvoiceNo = ""
    End If
    Set rsGetAPJInvoiceNo = Nothing
End Function

Function GetAPJInvoiceType(XXX As String, YYY As String) As String
    Dim rsGetAPJInvoiceType                                 As ADODB.Recordset
    Set rsGetAPJInvoiceType = New ADODB.Recordset
    rsGetAPJInvoiceType.Open "SELECT INVOICETYPE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO = '" & XXX & "' AND JTYPE = '" & YYY & "'", gconDMIS, adOpenForwardOnly
    If Not rsGetAPJInvoiceType.EOF And Not rsGetAPJInvoiceType.BOF Then
        GetAPJInvoiceType = Null2String(rsGetAPJInvoiceType!InvoiceType)
    Else
        GetAPJInvoiceType = ""
    End If
    Set rsGetAPJInvoiceType = Nothing
End Function

Function Vendor_PayCode(XXX As String) As String
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "SELECT B.PAY_CODE FROM ALL_VENDOR_TABLE A INNER JOIN ALL_PAYTERM B ON A.TERMS=B.NO_DAYS WHERE CODE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        Vendor_PayCode = rsVENDOR!pay_Code
    Else
        If XXX <> "" Then
            Vendor_PayCode = "CSH"
        End If
    End If
    Set rsVENDOR = Nothing
End Function

Function TERMS(XXX As String) As String
    Dim rsTERMS                                             As ADODB.Recordset
    Set rsTERMS = New ADODB.Recordset
    rsTERMS.Open "SELECT NO_DAYS FROM ALL_PAYTERM WHERE PAY_CODE = " & XXX & "", gconDMIS, adOpenForwardOnly
    If Not rsTERMS.EOF And Not rsTERMS.BOF Then
        TERMS = Null2String(rsTERMS!no_Days)
    Else
        TERMS = ""
    End If
    Set rsTERMS = Nothing
End Function

Function SetVendorAddressNew(VVV As Variant, XXX As Variant)
    Set rsVENDOR = New ADODB.Recordset
    'rsVENDOR.Open "Select code,address from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsVENDOR.Open "Select Code,Address from ALL_ENTITY where code = " & N2Str2Null(VVV) & " AND ENTITYCODE=" & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorAddressNew = Null2String(rsVENDOR!Address)
    Else
        SetVendorAddressNew = ""
    End If
End Function

Public Sub frmNewAMISJournalEntries_NewJournalEntries(AcctCode As String, DESCRIPTION As String, Debit As Double, Credit As Double, DetID As Long)
    cboAcct_Code.Text = AcctCode
    txtAcct_Name.Text = DESCRIPTION
    txtDebit.Text = Debit
    txtCredit.Text = Credit
    labDetID.Caption = DetID
End Sub

Function AMIS_AP_ENTRY() As Boolean
    On Error GoTo ErrorCode
    Dim xVOUCHERNO                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xVENDORCODE                                         As String
    Dim xVENDORNAME                                         As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xBAL                                                As Double
    Dim xACCT_CODE                                          As String
    Dim xInvoicedate                                        As String
    Dim xLAST_UPDATED                                       As String
    Dim xJdate                                              As String
    Dim xDUEDATE                                            As String
    Dim xREFCODE                                            As String
    Dim xREMARKS                                            As String
    Dim xPVID                                               As Integer

    If txtINV_No.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtInvoiceNo.SetFocus
        Exit Function
    ElseIf txtPVAmount.Text = "" Or txtPVAmount.Text = "0.00" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtPVAmount.SetFocus
        Exit Function
    ElseIf txtInvoiceDate.Text = "" Or IsDate(txtInvoiceDate.Text) = False Then
        MessagePop RecSaveError, "System Message", "Invalid is date!"
        txtInvoiceDate.SetFocus
        Exit Function
    End If
    'If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then xREMARKS = "NULL" Else xREMARKS = N2Str2Null(Trim(txtRemarks.Text))

    xVOUCHERNO = N2Str2Null(xJOURNALTYPE + "-" + txtVoucherNo.Text)
    xInvoiceType = "'OI'"
    xInvoiceNo = N2Str2Null(txtINV_No.Text)
    xVENDORCODE = N2Str2Null(txtCode.Text)
    xVENDORNAME = N2Str2Null(cboNameofVendor.Text)
    xAMOUNT_TO_PAY = NumericVal(txtPVAmount.Text)
    xAMOUNT_PAID = 0
    xBAL = NumericVal(txtPVAmount.Text)
    xACCT_CODE = N2Str2Null(GetAPAcctCode(xJOURNALTYPE, txtVoucherNo.Text))
    xInvoicedate = N2Str2Null(txtInvoiceDate.Text)
    xLAST_UPDATED = N2Str2Null(LOGDATE)
    xJdate = N2Str2Null(txtJDate.Text)
    xDUEDATE = N2Str2Null(txtDueDate.Text)
    xREFCODE = N2Str2Null(lblClass + txtCode.Text)
    xREMARKS = "NULL"

    xPVID = GetPVID(xJOURNALTYPE, txtVoucherNo.Text, txtINV_No.Text)
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,REMARKS,ENTITYCODE,REFCODE,DUEDATE,PV_ID) " & _
                        "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & xVENDORCODE & "," & xVENDORNAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xREMARKS & ",'" & lblClass & "'," & xREFCODE & "," & xDUEDATE & "," & xPVID & ")"
        gconDMIS.Execute SQL_STATEMENT

    Else
        SQL_STATEMENT = "UPDATE AMIS_AP SET VENDOR_CODE=" & xVENDORCODE & ",VENDOR_NAME=" & xVENDORNAME & ",INVOICENO=" & xInvoiceNo & ",AMOUNT2PAY=" & xAMOUNT_TO_PAY & ",BALANCE=" & xBAL & ",INVOICEDATE=" & xInvoicedate & ",REMARKS=" & xREMARKS & ",ENTITYCODE='" & lblClass & "',REFCODE=" & xREFCODE & ",DUEDATE=" & xDUEDATE & " WHERE PV_ID=" & labPVID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT
    End If
    AMIS_AP_ENTRY = True
    Exit Function

ErrorCode:
    AMIS_AP_ENTRY = False
End Function

Function GetAPAcctCode(JTYPE As String, VOUCHERNO As String) As String
    Dim rsJournalDT                                         As ADODB.Recordset
    Set rsJournalDT = New ADODB.Recordset
    rsJournalDT.Open "SELECT * FROM (SELECT ACCT_CODE FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-06','21-07') AND JTYPE='" & JTYPE & "' AND VOUCHERNO='" & VOUCHERNO & "' AND CREDIT > 0) DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.ACCT_CODE=AC.ACCTCODE WHERE IS_SCHEDULE_ACCNT=1", gconDMIS, adOpenForwardOnly
    If Not rsJournalDT.EOF And Not rsJournalDT.BOF Then
        GetAPAcctCode = rsJournalDT!ACCT_CODE
    Else
        GetAPAcctCode = ""
    End If
    Set rsJournalDT = Nothing
End Function

Function GetPVID(JTYPE As String, VOUCHERNO As String, INVOICENO As String) As Integer
    Dim rsdetail                                            As ADODB.Recordset
    Set rsdetail = New ADODB.Recordset
    rsdetail.Open "SELECT ID FROM AMIS_PV_DETAIL WHERE JTYPE='" & JTYPE & "' AND VOUCHERNO='" & VOUCHERNO & "' AND INV_NO='" & INVOICENO & "'", gconDMIS, adOpenForwardOnly
    If Not rsdetail.EOF And Not rsdetail.BOF Then
        GetPVID = rsdetail!ID
    Else
        GetPVID = ""
    End If
    Set rsdetail = Nothing
End Function

Sub LOAD_ARAP_DETAILS()
    If Picture1.Visible = True Then
        If lstDetails.ListItems.Count = 0 Then Exit Sub
        If CheckIfARDebitNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), lstDetails.SelectedItem.SubItems(3)) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, lstDetails.SelectedItem.SubItems(3), lblClass, txtCode.Text, lstDetails.SelectedItem.SubItems(3), labDetID.Caption)
            If Null2String(rsJournal_HD!Status) = "N" Then
                Call DetailsTrueFalse(True)
            Else
                Call DetailsTrueFalse(False)
            End If
            frmAMISJournalEntry_Details.Show 1
        ElseIf CheckIfAPDebitNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), lstDetails.SelectedItem.SubItems(3)) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, lstDetails.SelectedItem.SubItems(3), lblClass, txtCode.Text, lstDetails.SelectedItem.SubItems(3), labDetID.Caption)
            If Null2String(rsJournal_HD!Status) = "N" Then
                Call DetailsPaymentTrueFalse(True)
            Else
                Call DetailsPaymentTrueFalse(False)
            End If
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfARCreditNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), lstDetails.SelectedItem.SubItems(4)) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, lstDetails.SelectedItem.SubItems(4), lblClass, txtCode.Text, lstDetails.SelectedItem.SubItems(3), labDetID.Caption)
            If Null2String(rsJournal_HD!Status) = "N" Then
                Call DetailsPaymentTrueFalse(True)
            Else
                Call DetailsPaymentTrueFalse(False)
            End If
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfAPCreditNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), lstDetails.SelectedItem.SubItems(4)) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, lstDetails.SelectedItem.SubItems(4), lblClass, txtCode.Text, lstDetails.SelectedItem.SubItems(3), labDetID.Caption)
            If Null2String(rsJournal_HD!Status) = "N" Then
                Call DetailsTrueFalse(True)
            Else
                Call DetailsTrueFalse(False)
            End If
            frmAMISJournalEntry_Details.Show 1
        End If
    End If
End Sub

Sub CHECK_DETAILS()
    If AddorEdit = "ADD" Then
        If CheckIfARDebitNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtDebit.Text) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, cboAcct_Code.Text, txtJDate.Text, txtDebit.Text, lblClass, txtCode.Text, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_Details.Show 1
        ElseIf CheckIfAPDebitNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtDebit.Text) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, cboAcct_Code.Text, txtJDate.Text, txtDebit.Text, lblClass, txtCode.Text, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfARCreditNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtCredit.Text) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, cboAcct_Code.Text, txtJDate.Text, txtCredit.Text, lblClass, txtCode.Text, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfAPCreditNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtCredit.Text) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, cboAcct_Code.Text, txtJDate.Text, txtCredit.Text, lblClass, txtCode.Text, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_Details.Show 1
        End If
    Else
        If CheckIfSameAccount(lstDetails.SelectedItem.SubItems(1), cboAcct_Code.Text, xJOURNALTYPE + "-" + txtVoucherNo.Text, txtDebit.Text, txtCredit.Text) = False Then
            MessagePop InfoWarning, "System Message", "Action not allowed. " & Chr(13) & "See details for account code " & lstDetails.SelectedItem.SubItems(1)
            Exit Sub
        Else
            If CheckIfARDebitNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), txtDebit.Text) = True Then
                If CheckIfBalanceAR(lstDetails.SelectedItem.SubItems(1), NumericVal(txtDebit.Text), xJOURNALTYPE + "-" + txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, txtDebit.Text, lblClass, txtCode.Text, txtDebit.Text, labDetID.Caption)
                    frmAMISJournalEntry_Details.Show 1
                End If
            ElseIf CheckIfAPDebitNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), txtDebit.Text) = True Then
                If CheckIfBalanceAPDetails(lstDetails.SelectedItem.SubItems(1), NumericVal(txtDebit.Text), xJOURNALTYPE + "-" + txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, txtDebit.Text, lblClass, txtCode.Text, txtDebit.Text, labDetID.Caption)
                    frmAMISJournalEntry_DetailPayment.Show 1
                End If
            ElseIf CheckIfARCreditNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), txtCredit.Text) = True Then
                If CheckIfBalanceARDetails(lstDetails.SelectedItem.SubItems(1), NumericVal(txtCredit.Text), xJOURNALTYPE + "-" + txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(txtVoucherNo.Text, xJOURNALTYPE, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, txtCredit.Text, lblClass, txtCode.Text, txtDebit.Text, labDetID.Caption)
                    frmAMISJournalEntry_DetailPayment.Show 1
                End If
            ElseIf CheckIfAPCreditNotZero(lstDetails.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(lstDetails.SelectedItem.SubItems(1))), txtCredit.Text) = True Then
                If CheckIfBalanceAP(lstDetails.SelectedItem.SubItems(1), NumericVal(txtCredit.Text), xJOURNALTYPE + "-" + txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + txtVoucherNo.Text, lstDetails.SelectedItem.SubItems(1), txtJDate.Text, txtCredit.Text, lblClass, txtCode.Text, txtDebit.Text, labDetID.Caption)
                    frmAMISJournalEntry_Details.Show 1
                End If
            End If
        End If
    End If
    Screen.MousePointer = 0
End Sub

