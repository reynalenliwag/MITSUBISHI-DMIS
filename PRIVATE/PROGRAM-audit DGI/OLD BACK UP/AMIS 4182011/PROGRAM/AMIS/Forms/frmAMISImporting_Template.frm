VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAMISImporting_Template 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IMPORTING TEMPLATES"
   ClientHeight    =   7770
   ClientLeft      =   11040
   ClientTop       =   4800
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAMISImporting_Template.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9855
   Begin VB.PictureBox picDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   120
      ScaleHeight     =   5175
      ScaleWidth      =   9540
      TabIndex        =   25
      Top             =   1530
      Width           =   9600
      Begin VB.PictureBox fraAddJournal 
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
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   180
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   32
         Top             =   1770
         Width           =   9135
         Begin VB.PictureBox picTranType 
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   930
            ScaleHeight     =   735
            ScaleWidth      =   6585
            TabIndex        =   80
            Top             =   870
            Width           =   6585
            Begin VB.PictureBox picTranType2 
               BorderStyle     =   0  'None
               Height          =   705
               Left            =   2880
               ScaleHeight     =   705
               ScaleWidth      =   3855
               TabIndex        =   92
               Top             =   90
               Width           =   3855
               Begin VB.ComboBox cboTranType2 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  ItemData        =   "frmAMISImporting_Template.frx":1082
                  Left            =   0
                  List            =   "frmAMISImporting_Template.frx":1095
                  Style           =   2  'Dropdown List
                  TabIndex        =   94
                  Top             =   210
                  Width           =   1785
               End
               Begin VB.ComboBox cboTranType3 
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  ItemData        =   "frmAMISImporting_Template.frx":10C7
                  Left            =   1920
                  List            =   "frmAMISImporting_Template.frx":10D4
                  Style           =   2  'Dropdown List
                  TabIndex        =   93
                  Top             =   210
                  Width           =   1785
               End
               Begin VB.Label Label6 
                  Caption         =   "RO Details"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   0
                  TabIndex        =   96
                  Top             =   0
                  Width           =   1365
               End
               Begin VB.Label Label7 
                  Caption         =   "Job Type"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   1920
                  TabIndex        =   95
                  Top             =   0
                  Width           =   1365
               End
            End
            Begin VB.ComboBox cboTranType1 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               ItemData        =   "frmAMISImporting_Template.frx":10E5
               Left            =   120
               List            =   "frmAMISImporting_Template.frx":10E7
               Style           =   2  'Dropdown List
               TabIndex        =   81
               Top             =   300
               Width           =   2625
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Account Class"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   82
               Top             =   90
               Width           =   1350
            End
         End
         Begin VB.TextBox txtPercent 
            Height          =   315
            Left            =   8010
            TabIndex        =   50
            Top             =   330
            Width           =   915
         End
         Begin VB.ComboBox cboDRCR 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":10E9
            Left            =   6780
            List            =   "frmAMISImporting_Template.frx":10F3
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   330
            Width           =   1185
         End
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
            MouseIcon       =   "frmAMISImporting_Template.frx":1106
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISImporting_Template.frx":1258
            Style           =   1  'Graphical
            TabIndex        =   33
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
            MouseIcon       =   "frmAMISImporting_Template.frx":1596
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISImporting_Template.frx":16E8
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   765
            Width           =   705
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
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
            Height          =   705
            Left            =   2310
            TabIndex        =   37
            Top             =   -30
            Width           =   4455
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   345
               Left            =   30
               TabIndex        =   38
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   609
               _Version        =   393217
               BackColor       =   16777215
               Enabled         =   -1  'True
               MultiLine       =   0   'False
               TextRTF         =   $"frmAMISImporting_Template.frx":1A13
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
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   60
               TabIndex        =   39
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
            TabIndex        =   36
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
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
            Left            =   6540
            MaxLength       =   4
            TabIndex        =   35
            Text            =   "Text1"
            Top             =   1260
            Visible         =   0   'False
            Width           =   855
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
            MouseIcon       =   "frmAMISImporting_Template.frx":1AA6
            MousePointer    =   99  'Custom
            Picture         =   "frmAMISImporting_Template.frx":1BF8
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   765
            Width           =   705
         End
         Begin VB.Label labModelID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5310
            TabIndex        =   57
            Top             =   1260
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label labACCTTYPE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4380
            TabIndex        =   51
            Top             =   1260
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label labAcctID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   930
            TabIndex        =   49
            Top             =   1260
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label labAccountCode 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2970
            TabIndex        =   46
            Top             =   1260
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label Label35 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Item No."
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   390
            TabIndex        =   45
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label34 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   90
            TabIndex        =   44
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label30 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Debit/Credit"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   6780
            TabIndex        =   43
            Top             =   60
            Width           =   1155
         End
         Begin VB.Label Label38 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Percent"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8100
            TabIndex        =   42
            Top             =   60
            Width           =   795
         End
         Begin VB.Label labDetID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2010
            TabIndex        =   41
            Top             =   1260
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1785
         Left            =   120
         TabIndex        =   31
         Top             =   1710
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   3149
         TX              =   "cmd1"
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
         MICON           =   "frmAMISImporting_Template.frx":1F48
      End
      Begin MSComctlLib.ListView lstDetails 
         Height          =   3285
         Left            =   30
         TabIndex        =   26
         Top             =   1650
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   5794
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
         MouseIcon       =   "frmAMISImporting_Template.frx":1F64
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
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "DR / CR"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.PictureBox picSALESVEHICLE 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   30
         ScaleHeight     =   1185
         ScaleWidth      =   9465
         TabIndex        =   88
         Top             =   60
         Width           =   9495
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00F5D8BC&
            Height          =   555
            Left            =   3480
            ScaleHeight     =   495
            ScaleWidth      =   5895
            TabIndex        =   114
            Top             =   540
            Width           =   5955
            Begin VB.CheckBox chkDown 
               BackColor       =   &H00F5D8BC&
               Caption         =   "Down Pymt"
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
               Left            =   4560
               TabIndex        =   120
               Top             =   120
               Width           =   1305
            End
            Begin VB.CheckBox chkDisc 
               BackColor       =   &H00F5D8BC&
               Caption         =   "Disc"
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
               Left            =   3840
               TabIndex        =   119
               Top             =   120
               Width           =   705
            End
            Begin VB.CheckBox chkFreebies 
               BackColor       =   &H00F5D8BC&
               Caption         =   "Freebies"
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
               Left            =   2820
               TabIndex        =   118
               Top             =   120
               Width           =   1035
            End
            Begin VB.CheckBox chkChattel 
               BackColor       =   &H00F5D8BC&
               Caption         =   "Chattel"
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
               Left            =   1860
               TabIndex        =   117
               Top             =   120
               Width           =   915
            End
            Begin VB.CheckBox chkLTO 
               BackColor       =   &H00F5D8BC&
               Caption         =   "LTO"
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
               Left            =   1200
               TabIndex        =   116
               Top             =   120
               Width           =   645
            End
            Begin VB.CheckBox chkInsurance2 
               BackColor       =   &H00F5D8BC&
               Caption         =   "Insurance"
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
               Left            =   30
               TabIndex        =   115
               Top             =   120
               Width           =   1155
            End
         End
         Begin VB.CommandButton cmdSaveSettings 
            Caption         =   "Save Settings"
            Height          =   345
            Left            =   1950
            TabIndex        =   108
            Top             =   660
            Width           =   1485
         End
         Begin VB.ComboBox cboTo 
            Height          =   330
            Left            =   1020
            TabIndex        =   107
            Top             =   660
            Width           =   825
         End
         Begin VB.ComboBox cboFrom 
            Height          =   330
            Left            =   120
            TabIndex        =   106
            Top             =   660
            Width           =   825
         End
         Begin VB.CheckBox chkCWT 
            BackColor       =   &H00F5D8BC&
            Caption         =   "CWT"
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
            Left            =   7680
            TabIndex        =   105
            Top             =   150
            Width           =   1395
         End
         Begin VB.CheckBox chkZeroRated 
            BackColor       =   &H00F5D8BC&
            Caption         =   "Zero Rated"
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
            Left            =   2550
            TabIndex        =   91
            Top             =   120
            Width           =   1395
         End
         Begin VB.ComboBox cboModel 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":20C6
            Left            =   150
            List            =   "frmAMISImporting_Template.frx":20D3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   90
            Width           =   2295
         End
         Begin VB.ComboBox cboTerm 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":20FA
            Left            =   5190
            List            =   "frmAMISImporting_Template.frx":210A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   90
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Term:"
            Height          =   210
            Index           =   3
            Left            =   4620
            TabIndex        =   113
            Top             =   120
            Width           =   540
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00FAF1DC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00F5D8BC&
            FillStyle       =   0  'Solid
            Height          =   435
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   9405
         End
      End
      Begin VB.PictureBox picPURCHASES 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   9465
         TabIndex        =   52
         Top             =   60
         Width           =   9495
         Begin VB.PictureBox picMode 
            BackColor       =   &H00F5D8BC&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   2700
            ScaleHeight     =   375
            ScaleWidth      =   3225
            TabIndex        =   110
            Top             =   90
            Width           =   3225
            Begin VB.ComboBox cboModeOfPayment 
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
               ItemData        =   "frmAMISImporting_Template.frx":2134
               Left            =   600
               List            =   "frmAMISImporting_Template.frx":2147
               TabIndex        =   111
               Top             =   30
               Width           =   2340
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mode:"
               Height          =   210
               Index           =   2
               Left            =   0
               TabIndex        =   112
               Top             =   90
               Width           =   585
            End
         End
         Begin VB.ComboBox cboSource 
            Height          =   330
            Left            =   7560
            TabIndex        =   104
            Top             =   120
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CheckBox chkGenuine 
            BackColor       =   &H00F5D8BC&
            Caption         =   "Genuine"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            TabIndex        =   54
            Top             =   120
            Width           =   1125
         End
         Begin VB.CheckBox chkNonVAT 
            BackColor       =   &H00F5D8BC&
            Caption         =   "Non-VAT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1470
            TabIndex        =   53
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Source:"
            Height          =   210
            Index           =   1
            Left            =   6780
            TabIndex        =   103
            Top             =   150
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Shape Shape 
            BackColor       =   &H00FAF1DC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00F5D8BC&
            FillStyle       =   0  'Solid
            Height          =   435
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   9405
         End
      End
      Begin VB.PictureBox picOTH 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   9465
         TabIndex        =   97
         Top             =   60
         Width           =   9495
         Begin VB.ComboBox cboPayOption 
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
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":2191
            Left            =   5640
            List            =   "frmAMISImporting_Template.frx":2193
            Style           =   2  'Dropdown List
            TabIndex        =   101
            Top             =   120
            Width           =   1335
         End
         Begin VB.CheckBox chkNONVATOR 
            BackColor       =   &H00F5D8BC&
            Caption         =   "Non-VAT"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   8250
            TabIndex        =   100
            Top             =   120
            Width           =   1065
         End
         Begin VB.ComboBox cboOTH 
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
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":2195
            Left            =   1620
            List            =   "frmAMISImporting_Template.frx":2197
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   98
            Top             =   120
            Width           =   3975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Transaction:"
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
            Height          =   210
            Left            =   210
            TabIndex        =   99
            Top             =   180
            Width           =   1350
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00FAF1DC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00F5D8BC&
            FillStyle       =   0  'Solid
            Height          =   435
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   9405
         End
      End
      Begin VB.PictureBox picSALESPARTS 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   9465
         TabIndex        =   55
         Top             =   60
         Width           =   9495
         Begin VB.CheckBox chkCWTParts 
            BackColor       =   &H00F5D8BC&
            Caption         =   "CWT"
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
            Left            =   4560
            TabIndex        =   121
            Top             =   120
            Width           =   1395
         End
         Begin VB.ComboBox cboPayType 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":2199
            Left            =   150
            List            =   "frmAMISImporting_Template.frx":21A6
            Sorted          =   -1  'True
            TabIndex        =   87
            Top             =   90
            Width           =   825
         End
         Begin VB.ComboBox cboPayClass 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "frmAMISImporting_Template.frx":21B8
            Left            =   1080
            List            =   "frmAMISImporting_Template.frx":21C5
            Sorted          =   -1  'True
            TabIndex        =   86
            Top             =   90
            Width           =   2025
         End
         Begin VB.CheckBox chkDiscount 
            BackColor       =   &H00F5D8BC&
            Caption         =   "Discount"
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
            Left            =   3330
            TabIndex        =   56
            Top             =   120
            Width           =   1185
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FAF1DC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00F5D8BC&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   30
            Width           =   9405
         End
      End
      Begin VB.PictureBox picSALESSERVICE 
         Height          =   1515
         Left            =   60
         ScaleHeight     =   1455
         ScaleWidth      =   9345
         TabIndex        =   58
         Top             =   60
         Width           =   9405
         Begin VB.ComboBox cboInternal 
            Height          =   330
            Left            =   3960
            TabIndex        =   102
            Text            =   "cboInternal"
            Top             =   1020
            Width           =   2355
         End
         Begin VB.Frame fraType 
            BackColor       =   &H00F5D8BC&
            Caption         =   "TERM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   8070
            TabIndex        =   83
            Top             =   180
            Width           =   1155
            Begin VB.OptionButton Option1 
               BackColor       =   &H00F5D8BC&
               Caption         =   "CHARGE"
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
               Left            =   120
               TabIndex        =   85
               Top             =   720
               Width           =   1005
            End
            Begin VB.OptionButton optCash 
               BackColor       =   &H00F5D8BC&
               Caption         =   "CASH"
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
               Left            =   120
               TabIndex        =   84
               Top             =   390
               Width           =   1005
            End
         End
         Begin VB.CheckBox chkSublet 
            BackColor       =   &H00F5D8BC&
            Caption         =   "SUBLET"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   78
            Top             =   480
            Width           =   1455
         End
         Begin VB.Frame fraDiscount 
            BackColor       =   &H00F5D8BC&
            Caption         =   "DISCOUNT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   6390
            TabIndex        =   74
            Top             =   180
            Width           =   1605
            Begin VB.CheckBox chkDiscMaterials 
               BackColor       =   &H00F5D8BC&
               Caption         =   "MATERIALS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   79
               Top             =   930
               Width           =   1455
            End
            Begin VB.CheckBox chkDiscLabor 
               BackColor       =   &H00F5D8BC&
               Caption         =   "LABOR"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   77
               Top             =   210
               Width           =   1455
            End
            Begin VB.CheckBox chkDiscParts 
               BackColor       =   &H00F5D8BC&
               Caption         =   "PARTS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   76
               Top             =   450
               Width           =   1455
            End
            Begin VB.CheckBox chkDiscAcc 
               BackColor       =   &H00F5D8BC&
               Caption         =   "ACCESSORIES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   75
               Top             =   690
               Width           =   1455
            End
         End
         Begin VB.CheckBox chkMaterials 
            BackColor       =   &H00F5D8BC&
            Caption         =   "MATERIALS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   70
            Top             =   750
            Width           =   1455
         End
         Begin VB.Frame frmAccessories 
            BackColor       =   &H00F5D8BC&
            Caption         =   "ACC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   2760
            TabIndex        =   71
            Top             =   420
            Width           =   1065
            Begin VB.CheckBox chkAccBP 
               BackColor       =   &H00F5D8BC&
               Caption         =   "BP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   450
               Width           =   915
            End
            Begin VB.CheckBox chkAccGJ 
               BackColor       =   &H00F5D8BC&
               Caption         =   "GJ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   72
               Top             =   210
               Width           =   915
            End
         End
         Begin VB.Frame fraParts 
            BackColor       =   &H00F5D8BC&
            Caption         =   "PARTS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   1530
            TabIndex        =   67
            Top             =   420
            Width           =   1065
            Begin VB.CheckBox chkPartsGJ 
               BackColor       =   &H00F5D8BC&
               Caption         =   "GJ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   69
               Top             =   210
               Width           =   765
            End
            Begin VB.CheckBox chkPartsBP 
               BackColor       =   &H00F5D8BC&
               Caption         =   "BP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   150
               TabIndex        =   68
               Top             =   450
               Width           =   765
            End
         End
         Begin VB.Frame fraLabor 
            BackColor       =   &H00F5D8BC&
            Caption         =   "LABOR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   330
            TabIndex        =   63
            Top             =   420
            Width           =   1065
            Begin VB.CheckBox chkLaborPMS 
               BackColor       =   &H00F5D8BC&
               Caption         =   "PMS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   66
               Top             =   690
               Width           =   915
            End
            Begin VB.CheckBox chkLaborBP 
               BackColor       =   &H00F5D8BC&
               Caption         =   "BP"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   65
               Top             =   450
               Width           =   915
            End
            Begin VB.CheckBox chkLaborGJ 
               BackColor       =   &H00F5D8BC&
               Caption         =   "GJ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   64
               Top             =   210
               Width           =   915
            End
         End
         Begin VB.CheckBox chkInsurance 
            BackColor       =   &H00F5D8BC&
            Caption         =   "INSURANCE"
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
            Left            =   4530
            TabIndex        =   62
            Top             =   90
            Width           =   1275
         End
         Begin VB.CheckBox chkWarranty 
            BackColor       =   &H00F5D8BC&
            Caption         =   "WARRANTY"
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
            Left            =   3120
            TabIndex        =   61
            Top             =   90
            Width           =   1245
         End
         Begin VB.CheckBox chkInternal 
            BackColor       =   &H00F5D8BC&
            Caption         =   "INTERNAL"
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
            Left            =   1860
            TabIndex        =   60
            Top             =   90
            Width           =   1125
         End
         Begin VB.CheckBox chkCustomer 
            BackColor       =   &H00F5D8BC&
            Caption         =   "CUSTOMER"
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
            Left            =   450
            TabIndex        =   59
            Top             =   90
            Width           =   1275
         End
         Begin VB.Shape Shape4 
            BorderColor     =   &H000000C0&
            Height          =   345
            Left            =   330
            Shape           =   4  'Rounded Rectangle
            Top             =   60
            Width           =   5625
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00FAF1DC&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00F5D8BC&
            FillStyle       =   0  'Solid
            Height          =   1455
            Left            =   30
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   9285
         End
      End
      Begin VB.Label Label53 
         Caption         =   "F3 - Add Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   4950
         Width           =   1815
      End
   End
   Begin VB.PictureBox picHeader 
      BorderStyle     =   0  'None
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
      Height          =   1380
      Left            =   120
      ScaleHeight     =   1380
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   60
      Width           =   9600
      Begin VB.CheckBox chkEWT 
         BackColor       =   &H00F5D8BC&
         Caption         =   "Withholding Tax Agent"
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
         Left            =   7230
         TabIndex        =   109
         Top             =   150
         Width           =   2295
      End
      Begin VB.ComboBox cboTransactionType 
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
         Height          =   330
         ItemData        =   "frmAMISImporting_Template.frx":21EC
         Left            =   1920
         List            =   "frmAMISImporting_Template.frx":21EE
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   960
         Width           =   2235
      End
      Begin VB.ComboBox cboType 
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
         Height          =   330
         ItemData        =   "frmAMISImporting_Template.frx":21F0
         Left            =   4860
         List            =   "frmAMISImporting_Template.frx":21F2
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1920
         TabIndex        =   5
         Text            =   "VIP"
         Top             =   525
         Width           =   7545
      End
      Begin VB.TextBox txtCode 
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
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "000001"
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   210
         Left            =   4260
         TabIndex        =   47
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   1020
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         TabIndex        =   3
         Top             =   570
         Width           =   1155
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
         Left            =   8820
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Top             =   180
         Width           =   1245
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FAF1DC&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00F5D8BC&
         FillStyle       =   0  'Solid
         Height          =   1275
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   60
         Width           =   9525
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
      Height          =   6585
      Left            =   210
      TabIndex        =   7
      Top             =   150
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
         TabIndex        =   8
         Text            =   "txtSearch"
         Top             =   270
         Width           =   9225
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   5565
         Left            =   90
         TabIndex        =   9
         Top             =   660
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   9816
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
         MouseIcon       =   "frmAMISImporting_Template.frx":21F4
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00EBFAFA&
         BackStyle       =   0  'Transparent
         Caption         =   "[Press <Enter> to Accept]      [Press <Ctrl> + <A> to Add Account]       [<F8> Change Search]"
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
         TabIndex        =   10
         Top             =   6300
         Width           =   9225
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   6735
      Left            =   150
      TabIndex        =   6
      Top             =   60
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   11880
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
      MICON           =   "frmAMISImporting_Template.frx":2356
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
      Height          =   870
      Left            =   1050
      ScaleHeight     =   870
      ScaleWidth      =   7665
      TabIndex        =   14
      Top             =   6840
      Width           =   7665
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
         Left            =   6840
         MouseIcon       =   "frmAMISImporting_Template.frx":2372
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":24C4
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   6090
         MouseIcon       =   "frmAMISImporting_Template.frx":282A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":297C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   765
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
         Left            =   5340
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmAMISImporting_Template.frx":2CE2
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":2E34
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel this Transaction"
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
         Left            =   4590
         MouseIcon       =   "frmAMISImporting_Template.frx":316E
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":32C0
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   3840
         MouseIcon       =   "frmAMISImporting_Template.frx":361C
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":376E
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   3090
         MouseIcon       =   "frmAMISImporting_Template.frx":3A81
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":3BD3
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   2340
         MouseIcon       =   "frmAMISImporting_Template.frx":3F23
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":4075
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   1590
         MouseIcon       =   "frmAMISImporting_Template.frx":43D3
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":4525
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   840
         MouseIcon       =   "frmAMISImporting_Template.frx":481F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":4971
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   90
         MouseIcon       =   "frmAMISImporting_Template.frx":4CC9
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":4E1B
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   765
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
      Height          =   870
      Left            =   1050
      ScaleHeight     =   870
      ScaleWidth      =   7665
      TabIndex        =   12
      Top             =   6840
      Width           =   7665
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
         Left            =   6840
         MouseIcon       =   "frmAMISImporting_Template.frx":517A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":52CC
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   6090
         MouseIcon       =   "frmAMISImporting_Template.frx":560A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISImporting_Template.frx":575C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmAMISImporting_Template"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsChartAccount                                     As New ADODB.Recordset
Dim AddorEdit                                          As String
Dim SearchBy                                           As String
Dim SQL_STATEMENT                                      As String
Dim rsAMIS_TEMPLATES                                   As ADODB.Recordset
Dim cntDetails                                         As Integer

Private Sub cboAcct_Code_Change()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    labAcctID.Caption = SetAcctID(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_Click()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    labAcctID.Caption = SetAcctID(cboAcct_Code.Text)
End Sub

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!DESCRIPTION))
    Else
        Setacctname = ""
    End If
    Set rsChartAccount2 = Nothing
End Function

Function SetAcctID(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select ID from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctID = Null2String(rsChartAccount2!ID)
    Else
        SetAcctID = ""
    End If
    Set rsChartAccount2 = Nothing
End Function

Private Sub cboDRCR_Click()
    picTranType.Visible = True
    picTranType2.Visible = False
End Sub

Private Sub cboFrom_Click()
    cboTo.Clear
    Dim rsTO As ADODB.Recordset
    Set rsTO = New ADODB.Recordset
    rsTO.Open "SELECT ID FROM AMIS_IMPORTTEMPLATE_HD WHERE TRANTYPE_ID=2 AND TYPE_ID=6", gconDMIS, adOpenForwardOnly
    If Not rsTO.EOF And Not rsTO.BOF Then
        Do While Not rsTO.EOF
            cboTo.AddItem rsTO!ID
            rsTO.MoveNext
        Loop
    End If
    Set rsTO = Nothing
End Sub

Private Sub cboModel_Change()
    labModelID.Caption = SetModelID(cboModel.Text)
End Sub

Private Sub cboModel_Click()
    labModelID.Caption = SetModelID(cboModel.Text)
End Sub

Private Sub cboTransactionType_Click()
'    If cboTransactionType.Text = "RECEIPTS" Then
'        cboType.Clear
'        With cboType
'        End With
'
'    End If
End Sub

Private Sub cboTranType1_Change()
    If cboTranType1.Text = "OUTPUT TAX" Or cboTranType1.Text = "ACCOUNTS RECEIVABLE" Or cboTranType1.Text = "SHOP SUPPLIES" Or cboTranType1.Text = "INTERNAL" Then
        Label6.Visible = False
        Label7.Visible = False
        cboTranType2.Visible = False
        cboTranType3.Visible = False
        cboTranType2.ListIndex = -1
        cboTranType3.ListIndex = -1
    Else
        Label6.Visible = True
        cboTranType2.Visible = True
        '        Label7.Visible = True
        '        cboTranType3.Visible = True
    End If
End Sub

Private Sub cboTranType1_Click()
    If cboTranType1.Text = "OUTPUT TAX" Or cboTranType1.Text = "ACCOUNTS RECEIVABLE" Or cboTranType1.Text = "SHOP SUPPLIES" Or cboTranType1.Text = "INTERNAL" Then
        Label6.Visible = False
        Label7.Visible = False
        cboTranType2.Visible = False
        cboTranType3.Visible = False
        cboTranType2.ListIndex = -1
        cboTranType3.ListIndex = -1
    Else
        Label6.Visible = True
        cboTranType2.Visible = True
        '        Label7.Visible = True
        '        cboTranType3.Visible = True
    End If

End Sub

Private Sub cboTranType2_Change()
    If cboTranType2.Text = "MATERIALS" Or cboTranType2.Text = "SUBLET" Then
        Label7.Visible = False
        cboTranType3.ListIndex = -1
        cboTranType3.Visible = False
    Else
        Label7.Visible = True
        cboTranType3.ListIndex = -1
        cboTranType3.Visible = True
    End If
End Sub

Private Sub cboTranType2_Click()
    If cboTranType2.Text = "MATERIALS" Or cboTranType2.Text = "SUBLET" Then
        Label7.Visible = False
        cboTranType3.ListIndex = -1
        cboTranType3.Visible = False
    ElseIf cboTranType2.Text <> "" Then
        Label7.Visible = True
        cboTranType3.ListIndex = -1
        cboTranType3.Visible = True
    End If
End Sub

Private Sub cboType_Click()
    cboPayType.Text = ""
    cboPayClass.Text = ""
    cboPayOption.ListIndex = -1
    cboModeOfPayment.Text = ""
    chkDiscount.Value = 0
    chkEWT.Value = 0
    chkZeroRated.Value = 0
    chkGenuine.Value = 0
    chkNonVat.Value = 0
    chkCustomer.Value = 0
    chkInternal.Value = 0
    chkWarranty.Value = 0
    chkInsurance.Value = 0
    chkLaborGJ.Value = 0
    chkLaborBP.Value = 0
    chkLaborPMS.Value = 0
    chkPartsGJ.Value = 0
    chkPartsBP.Value = 0
    chkSublet.Value = 0
    chkMaterials.Value = 0
    chkAccGJ.Value = 0
    chkAccBP.Value = 0
    chkDiscLabor.Value = 0
    chkDiscParts.Value = 0
    chkDiscMaterials.Value = 0
    chkDiscAcc.Value = 0
    If cboTransactionType.Text = "PURCHASES" And cboType.Text = "PARTS" Then
        picPURCHASES.Visible = True
        chkGenuine.Visible = True
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = False
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
        picMode.Visible = False
    ElseIf cboTransactionType.Text = "PURCHASES" And cboType.Text = "VEHICLES" Then
        picPURCHASES.Visible = True
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = False
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
        chkGenuine.Value = 0
        chkGenuine.Visible = False
        picMode.Visible = True
    ElseIf cboTransactionType.Text = "PURCHASES" And (cboType.Text = "ACCESSORIES" Or cboType.Text = "MATERIALS" Or cboType.Text = "SUBLET") Then
        picPURCHASES.Visible = True
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = False
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
        chkGenuine.Value = 0
        chkGenuine.Visible = False
        picMode.Visible = False
    ElseIf cboTransactionType.Text = "SALES" And (cboType.Text = "PARTS" Or cboType.Text = "ACCESSORIES" Or cboType.Text = "MATERIALS" Or cboType.Text = "SUBLET") Then
        picPURCHASES.Visible = False
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = True
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
        cboPayType.Text = ""
        cboPayClass.Text = ""
    ElseIf cboTransactionType.Text = "SALES" And cboType.Text = "VEHICLES" Then
        picPURCHASES.Visible = False
        picSALESPARTS.Visible = False
        picSALESVEHICLE.Visible = True
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
        cboModel.ListIndex = -1
        cboFrom.Clear
        Dim rsFROM As ADODB.Recordset
        Set rsFROM = New ADODB.Recordset
        rsFROM.Open "SELECT ID FROM AMIS_IMPORTTEMPLATE_HD WHERE TRANTYPE_ID=2 AND TYPE_ID=6", gconDMIS, adOpenForwardOnly
        If Not rsFROM.EOF And Not rsFROM.BOF Then
            Do While Not rsFROM.EOF
                cboFrom.AddItem rsFROM!ID
                rsFROM.MoveNext
            Loop
        End If
        Set rsFROM = Nothing
    ElseIf cboTransactionType.Text = "SALES" And cboType.Text = "SERVICE" Then
        picPURCHASES.Visible = False
        picSALESPARTS.Visible = False
        picSALESVEHICLE.Visible = False
        picSALESSERVICE.Visible = True
        picOTH.Visible = False
    ElseIf cboTransactionType.Text = "RECEIPTS" Then
        cboOTH.ListIndex = -1
        picPURCHASES.Visible = False
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = False
        picSALESSERVICE.Visible = False
        picOTH.Visible = True
        If cboType.Text = "OTHERS" Then
            cboOTH.Enabled = True
        Else
            cboOTH.Enabled = False
        End If
    Else
        picPURCHASES.Visible = False
        picSALESVEHICLE.Visible = False
        picSALESPARTS.Visible = False
        picSALESSERVICE.Visible = False
        picOTH.Visible = False
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    AddorEdit = "ADD"
    Picture1.Visible = False
    Picture2.Visible = True
    picHeader.Enabled = True
    initMemvars
    txtCode.Text = Get_Code
    txtDescription.SetFocus
End Sub

Private Sub cmdAddJournal_Click()
    cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
    fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
    fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
    AddorEdit = "ADD"
    InitJournal
    On Error Resume Next
    cboAcct_Code.SetFocus
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    picHeader.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If labID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Details."
        Exit Sub
    End If
    If MsgBox("Are you sure you want to DELETE this template?", vbQuestion + vbYesNo, "Delete Importing Template") = vbYes Then
        gconDMIS.Execute "DELETE FROM AMIS_IMPORTTEMPLATE_HD WHERE ID = " & labID.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsAMIS_TEMPLATES.MoveLast
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    Picture1.Visible = False
    Picture2.Visible = True
    picHeader.Enabled = True
    txtDescription.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode
    rsAMIS_TEMPLATES.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & vbCrLf & error, vbOKOnly, "Error"
End Sub

Private Sub cmdJournalCancel_Click()
    SendToBack
    Picture1.Enabled = True
    lstDetails.Enabled = True
End Sub

Private Sub cmdJournalDelete_Click()
    If labDetID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Details."
        Exit Sub
    End If
    If MsgBox("Are you sure you want to DELETE this detail?", vbQuestion + vbYesNo, "Delete Detail Entry") = vbYes Then
        gconDMIS.Execute "DELETE FROM AMIS_IMPORTTEMPLATE_DT WHERE ID = " & labDetID.Caption
    End If
    FillDetails
    rsRefresh
    On Error Resume Next
    rsAMIS_TEMPLATES.Find "ID = " & labID.Caption
    cmdJournalCancel.Value = True
    If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
        lstDetails.SetFocus
    End If
End Sub

Private Sub cmdJournalSave_Click()
'    SendToBack
    Picture1.Enabled = True
    lstDetails.Enabled = True
    Dim J_CHARTACCOUNT_ID                              As Integer
    Dim J_HD_ID                                        As Integer
    Dim J_DEBIT                                        As Boolean
    Dim J_TRANTYPE1                                    As String
    Dim J_TRANTYPE2                                    As String
    Dim J_TRANTYPE3                                    As String


    If cboAcct_Code.Text = "" Or Setacctname(cboAcct_Code.Text) = "" Then
        MsgBox "Account Code and Description must have a value.", vbInformation, "System Message"
        cboAcct_Code.SetFocus
        Exit Sub
    ElseIf cboDRCR.Text = "" Then
        MsgBox "Debit or Credit must have a value.", vbInformation, "System Message"
        cboDRCR.SetFocus
        Exit Sub
    End If

    J_CHARTACCOUNT_ID = labAcctID.Caption
    J_HD_ID = labID.Caption
    If cboDRCR.Text = "DEBIT" Then
        J_DEBIT = 1
    Else
        J_DEBIT = 0
    End If

    J_TRANTYPE1 = N2Str2Null(cboTranType1.Text)

    If J_TRANTYPE1 = "'OUTPUT TAX'" Or J_TRANTYPE1 = "'ACCOUNTS RECEIVABLE'" Or J_TRANTYPE1 = "'SHOP SUPPLIES'" Then
        J_TRANTYPE2 = "NULL"
        J_TRANTYPE3 = "NULL"
    Else
        J_TRANTYPE2 = N2Str2Null(cboTranType2.Text)
        J_TRANTYPE3 = N2Str2Null(cboTranType3.Text)
    End If

    If J_TRANTYPE2 = "MATERIALS" Then
        J_TRANTYPE3 = "NULL"
    Else
        J_TRANTYPE3 = N2Str2Null(cboTranType3.Text)
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO AMIS_IMPORTTEMPLATE_DT " & _
                        "(CHARTACCOUNT_ID,TEMPLATE_HD_ID,DEBIT,TRANTYPE1,TRANTYPE2,TRANTYPE3) VALUES " & _
                        "('" & J_CHARTACCOUNT_ID & "','" & J_HD_ID & "','" & J_DEBIT & "'," & J_TRANTYPE1 & "," & J_TRANTYPE2 & "," & J_TRANTYPE3 & ")"
        gconDMIS.Execute SQL_STATEMENT
    Else
        SQL_STATEMENT = "UPDATE AMIS_IMPORTTEMPLATE_DT SET " & _
                        "CHARTACCOUNT_ID='" & J_CHARTACCOUNT_ID & "',TEMPLATE_HD_ID='" & J_HD_ID & "',DEBIT='" & J_DEBIT & "',TRANTYPE1=" & J_TRANTYPE1 & ",TRANTYPE2=" & J_TRANTYPE2 & ",TRANTYPE3=" & J_TRANTYPE3 & " WHERE ID = '" & labDetID.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT
    End If
    FillDetails
    InitJournal
    If AddorEdit = "EDIT" Then
        SendToBack
    Else
        cboAcct_Code.SetFocus
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdLast_Click()
    On Error GoTo ErrorCode
    rsAMIS_TEMPLATES.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & vbCrLf & error, vbOKOnly, "Error"
End Sub

Private Sub cmdNext_Click()
    On Error GoTo ErrorCode
    rsAMIS_TEMPLATES.MoveNext
    If rsAMIS_TEMPLATES.EOF Then
        rsAMIS_TEMPLATES.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & vbCrLf & error, vbOKOnly, "Error"
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode
    rsAMIS_TEMPLATES.MovePrevious
    If rsAMIS_TEMPLATES.BOF Then
        rsAMIS_TEMPLATES.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & vbCrLf & error, vbOKOnly, "Error"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode

    If txtDescription.Text = "" Then
        MsgBox "Description cannot be empty.", vbExclamation, "System Message"
        txtDescription.SetFocus
        Exit Sub
    ElseIf cboTransactionType.Text = "" Then
        MsgBox "Please select Transaction Type.", vbExclamation, "System Message"
        cboTransactionType.SetFocus
        Exit Sub
    ElseIf cboType.Text = "" Then
        MsgBox "Please select Type.", vbExclamation, "System Message"
        cboType.SetFocus
        Exit Sub
    End If

    Dim J_CODE                                         As String
    Dim J_DESCRIPTION                                  As String
    Dim J_TRANTYPE                                     As String
    Dim J_TYPE                                         As String
    Dim J_GENUINE                                      As String
    Dim J_NON_VAT                                      As String
    Dim J_WTAX                                         As String
    Dim J_PAYCLASS                                     As String
    Dim J_DISCOUNT                                     As String
    Dim J_PAYTYPE                                      As String
    Dim J_MODEL                                        As String
    Dim J_CUSTOMER                                     As String
    Dim J_INTERNAL                                     As String
    Dim J_WARRANTY                                     As String
    Dim J_INSURANCE                                    As String
    Dim J_LABOR_GJ                                     As String
    Dim J_LABOR_BP                                     As String
    Dim J_LABOR_PMS                                    As String
    Dim J_SUBLET                                       As String
    Dim J_PARTS_GJ                                     As String
    Dim J_PARTS_BP                                     As String
    Dim J_MATERIALS                                    As String
    Dim J_ACC_GJ                                       As String
    Dim J_ACC_BP                                       As String
    Dim J_DISC_LABOR                                   As String
    Dim J_DISC_PARTS                                   As String
    Dim J_DISC_MAT                                     As String
    Dim J_DISC_ACC                                     As String
    Dim J_TERMS                                        As String
    Dim J_CWT                                          As String
    Dim J_MODEOFPAYMENT                                As String
    Dim J_INSURANCE2                                   As String
    Dim J_LTO                                          As String
    Dim J_CHATTEL                                      As String
    Dim J_FREEBIES                                     As String
    Dim J_DOWNPAYMENT                                  As String

    J_CODE = N2Str2Null(Format(txtCode.Text, "000000"))
    J_DESCRIPTION = N2Str2Null(txtDescription.Text)
    J_TRANTYPE = NumericVal(GET_TRANTYPE_ID(cboTransactionType.Text))
    J_TYPE = NumericVal(GET_TYPE_ID(cboType.Text))
    '    If chkGenuine.Value = True Then
    J_GENUINE = chkGenuine.Value
    '    Else
    '        J_GENUINE = "NULL"
    '    End If
    '    If chkNonVAT.Value = True Then
    J_INSURANCE2 = chkInsurance2.Value
    J_LTO = chkLTO.Value
    J_CHATTEL = chkChattel.Value
    
    J_TERMS = N2Str2Null(SetTERM(cboTerm.Text))
    J_MODEOFPAYMENT = N2Str2Null(SetModeOfPayment(cboModeOfPayment.Text))
    
    If cboType.Text = "OTHERS" Then
        J_NON_VAT = chkNONVATOR.Value
    Else
        J_NON_VAT = chkNonVat.Value
    End If
    '    Else
    '        J_NON_VAT = "NULL"
    '    End If
    '    If chkEWT.Value = True Then
    J_WTAX = chkEWT.Value
    J_DOWNPAYMENT = chkDown.Value
    '    Else
    '        J_WTAX = "NULL"
    '    End If
    If cboPayClass.Text = "CUSTOMER PAID" Then
        J_PAYCLASS = "'C'"
    ElseIf cboPayClass.Text = "WARRANTY" Then
        J_PAYCLASS = "'W'"
    ElseIf cboPayClass.Text = "INTERNAL" Then
        J_PAYCLASS = "'I'"
    Else
        J_PAYCLASS = "NULL"
    End If
    '    If chkDiscount.Value = True Then
    If cboType.Text = "PARTS" Then
        J_DISCOUNT = chkDiscount.Value
        J_CWT = chkCWTParts.Value
    ElseIf cboType.Text = "VEHICLES" Then
        J_DISCOUNT = chkDisc.Value
        J_CWT = chkCWT.Value
    End If
    '    Else
    '        J_DISCOUNT = "NULL"
    '    End If
    If cboPayType.Text = "" And cboPayOption.Text = "" Then
        J_PAYTYPE = "NULL"
    Else
        If cboTransactionType.Text = "RECEIPTS" Then
            J_PAYTYPE = N2Str2Null(cboPayOption.Text)
        Else
            J_PAYTYPE = N2Str2Null(cboPayType.Text)
        End If
    End If
    If cboModel.Text = "" Then
        J_MODEL = "NULL"
    Else
        J_MODEL = N2Str2Null(cboModel.Text)
    End If
    '    If chkCustomer.Value = True Then
    J_CUSTOMER = chkCustomer.Value
    '    Else
    '        J_CUSTOMER = "NULL"
    '    End If
    '    If chkInternal.Value = True Then
    J_INTERNAL = chkInternal.Value
    '    Else
    '        J_INTERNAL = "NULL"
    '    End If
    '    If chkWarranty.Value = True Then
    J_WARRANTY = chkWarranty.Value
    '    Else
    '        J_WARRANTY = "NULL"
    '    End If
    '    If chkInsurance.Value = True Then
    J_INSURANCE = chkInsurance.Value
    '    Else
    '        J_INSURANCE = "NULL"
    '    End If
    '    If chkLaborGJ.Value = True Then
    J_LABOR_GJ = chkLaborGJ.Value
    '    Else
    '        J_LABOR_GJ = "NULL"
    '    End If
    '    If chkLaborBP.Value = True Then
    J_LABOR_BP = chkLaborBP.Value
    '    Else
    '        J_LABOR_BP = "NULL"
    '    End If
    '    If chkLaborPMS.Value = True Then
    J_LABOR_PMS = chkLaborPMS.Value
    '    Else
    '        J_LABOR_PMS = "NULL"
    '    End If
    '    If chkSublet.Value = True Then
    J_SUBLET = chkSublet.Value
    '    Else
    '        J_SUBLET = "NULL"
    '    End If
    '    If chkPartsGJ.Value = True Then
    J_PARTS_GJ = chkPartsGJ.Value
    '    Else
    '        J_PARTS_GJ = "NULL"
    '    End If
    '    If chkPartsBP.Value = True Then
    J_PARTS_BP = chkPartsBP.Value
    '    Else
    '        J_PARTS_BP = "NULL"
    '    End If
    '    If chkMaterials.Value = True Then
    J_MATERIALS = chkMaterials.Value
    '    Else
    '        J_MATERIALS = "NULL"
    '    End If
    '    If chkAccGJ.Value = True Then
    J_ACC_GJ = chkAccGJ.Value
    '    Else
    '        J_ACC_GJ = "NULL"
    '    End If
    '    If chkAccBP.Value = True Then
    J_ACC_BP = chkAccBP.Value
    '    Else
    '        J_ACC_BP = "NULL"
    '    End If
    '    If chkDiscLabor.Value = True Then
    J_DISC_LABOR = chkDiscLabor.Value
    '    Else
    '        J_DISC_LABOR = "NULL"
    '    End If
    '    If chkDiscParts.Value = True Then
    J_DISC_PARTS = chkDiscParts.Value
    '    Else
    '        J_DISC_PARTS = "NULL"
    '    End If
    '    If chkDiscMaterials.Value = True Then
    J_DISC_MAT = chkDiscMaterials.Value
    '    Else
    '        J_DISC_MAT = "NULL"
    '    End If
    '    If chkDiscAcc.Value = True Then
    J_DISC_ACC = chkDiscAcc.Value
    '    Else
    '        J_DISC_ACC = "NULL"
    '    End If
    If cboType.Text = "OTHERS" And cboOTH.Text = "" Then
        MessagePop RecSaveWarning, "Other Transaction", "Other Transaction tagging is required."
        cboOTH.SetFocus
        Exit Sub
    Else
        If cboType.Text = "OTHERS" Then
            J_MODEL = N2Str2Null(GET_OTHERCODE(cboOTH.Text))
        End If
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "INSERT INTO AMIS_IMPORTTEMPLATE_HD " & _
                        "(CODE,DESCRIPTION,TRANTYPE_ID,TYPE_ID,GENUINE,PAY_CLASS,NONVAT,DISCOUNT,PAY_TYPE,MODEL_ID,CUSTOMER,INTERNAL,WARRANTY,INSURANCE,LABOR_GJ,LABOR_BP,LABOR_PMS,SUBLET,PARTS_GJ,PARTS_BP,MATERIALS,ACC_GJ,ACC_BP,DISC_LABOR,DISC_PARTS,DISC_MAT,DISC_ACC,TERMS,CWT,MODEOFPAYMENT,INSURANCE2,LTO,CHATTEL,FREEBIES,DOWNPAYMENT) VALUES " & _
                        "(" & J_CODE & "," & J_DESCRIPTION & "," & J_TRANTYPE & "," & J_TYPE & ",'" & J_GENUINE & "'," & J_PAYCLASS & ",'" & J_NON_VAT & "','" & J_DISCOUNT & "'," & J_PAYTYPE & "," & J_MODEL & ",'" & J_CUSTOMER & "','" & J_INTERNAL & "','" & J_WARRANTY & "','" & J_INSURANCE & "','" & J_LABOR_GJ & "','" & J_LABOR_BP & "','" & J_LABOR_PMS & "','" & J_SUBLET & "','" & J_PARTS_GJ & "','" & J_PARTS_BP & "','" & J_MATERIALS & "','" & J_ACC_GJ & "','" & J_ACC_BP & "','" & J_DISC_LABOR & "','" & J_DISC_PARTS & "','" & J_DISC_MAT & "','" & J_DISC_ACC & "'," & J_TERMS & ",'" & J_CWT & "'," & J_MODEOFPAYMENT & ",'" & J_INSURANCE2 & "','" & J_LTO & "','" & J_CHATTEL & "','" & J_FREEBIES & "','" & J_DOWNPAYMENT & "')"
        gconDMIS.Execute SQL_STATEMENT

        'labID.Caption = FindNewID(J_CODE, "CODE", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        'NEW_LogAudit "A", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "UPDATE AMIS_IMPORTTEMPLATE_HD SET "
        SQL_STATEMENT = SQL_STATEMENT & "DESCRIPTION=" & J_DESCRIPTION & ","
        SQL_STATEMENT = SQL_STATEMENT & "TRANTYPE_ID=" & J_TRANTYPE & ","
        SQL_STATEMENT = SQL_STATEMENT & "TYPE_ID=" & J_TYPE & ","
        SQL_STATEMENT = SQL_STATEMENT & "GENUINE='" & J_GENUINE & "',"
        SQL_STATEMENT = SQL_STATEMENT & "PAY_CLASS=" & J_PAYCLASS & ","
        SQL_STATEMENT = SQL_STATEMENT & "NONVAT='" & J_NON_VAT & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DISCOUNT='" & J_DISCOUNT & "',"
        SQL_STATEMENT = SQL_STATEMENT & "PAY_TYPE=" & J_PAYTYPE & ","
        SQL_STATEMENT = SQL_STATEMENT & "MODEL_ID=" & J_MODEL & ","
        SQL_STATEMENT = SQL_STATEMENT & "CUSTOMER='" & J_CUSTOMER & "',"
        SQL_STATEMENT = SQL_STATEMENT & "INTERNAL='" & J_INTERNAL & "',"
        SQL_STATEMENT = SQL_STATEMENT & "WARRANTY='" & J_WARRANTY & "',"
        SQL_STATEMENT = SQL_STATEMENT & "INSURANCE='" & J_INSURANCE & "',"
        SQL_STATEMENT = SQL_STATEMENT & "LABOR_GJ='" & J_LABOR_GJ & "',"
        SQL_STATEMENT = SQL_STATEMENT & "LABOR_BP='" & J_LABOR_BP & "',"
        SQL_STATEMENT = SQL_STATEMENT & "LABOR_PMS='" & J_LABOR_PMS & "',"
        SQL_STATEMENT = SQL_STATEMENT & "SUBLET='" & J_SUBLET & "',"
        SQL_STATEMENT = SQL_STATEMENT & "PARTS_GJ='" & J_PARTS_GJ & "',"
        SQL_STATEMENT = SQL_STATEMENT & "PARTS_BP='" & J_PARTS_BP & "',"
        SQL_STATEMENT = SQL_STATEMENT & "MATERIALS='" & J_MATERIALS & "',"
        SQL_STATEMENT = SQL_STATEMENT & "ACC_GJ='" & J_ACC_GJ & "',"
        SQL_STATEMENT = SQL_STATEMENT & "ACC_BP='" & J_ACC_BP & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DISC_LABOR='" & J_DISC_LABOR & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DISC_PARTS='" & J_DISC_PARTS & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DISC_MAT='" & J_DISC_MAT & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DISC_ACC='" & J_DISC_ACC & "',"
        SQL_STATEMENT = SQL_STATEMENT & "TERMS=" & J_TERMS & ","
        SQL_STATEMENT = SQL_STATEMENT & "CWT='" & J_CWT & "',"
        SQL_STATEMENT = SQL_STATEMENT & "MODEOFPAYMENT=" & J_MODEOFPAYMENT & ","
        SQL_STATEMENT = SQL_STATEMENT & "INSURANCE2='" & J_INSURANCE2 & "',"
        SQL_STATEMENT = SQL_STATEMENT & "LTO='" & J_LTO & "',"
        SQL_STATEMENT = SQL_STATEMENT & "CHATTEL='" & J_CHATTEL & "',"
        SQL_STATEMENT = SQL_STATEMENT & "FREEBIES='" & J_FREEBIES & "',"
        SQL_STATEMENT = SQL_STATEMENT & "DOWNPAYMENT='" & J_DOWNPAYMENT & "'"
        SQL_STATEMENT = SQL_STATEMENT & "WHERE ID='" & labID.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE ALL_PROFILE SET WTAXAGENT='" & J_WTAX & "' WHERE MODULENAME='AMIS'"
        gconDMIS.Execute SQL_STATEMENT
    End If
    rsRefresh
    rsAMIS_TEMPLATES.Find "CODE = " & txtCode.Text
    Picture1.Visible = True
    Picture2.Visible = False
    picHeader.Enabled = False
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    MsgBox "Error:" & Err & " " & vbCrLf & error, vbOKOnly, "Error"
    'Call ErrHandler(gconDMIS)
    'SaveLogFile
    Exit Sub
End Sub

Function GET_TRANTYPE_ID(XXX As String) As Integer
    Dim rsGET_TRANTYPE_ID                              As ADODB.Recordset
    Set rsGET_TRANTYPE_ID = New ADODB.Recordset
    rsGET_TRANTYPE_ID.Open "SELECT ID FROM AMIS_IMPORT_TRANTYPE WHERE TRANTYPE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGET_TRANTYPE_ID.EOF And Not rsGET_TRANTYPE_ID.BOF Then
        GET_TRANTYPE_ID = rsGET_TRANTYPE_ID!ID
    End If
    Set rsGET_TRANTYPE_ID = Nothing
End Function

Function GET_TYPE_ID(XXX As String) As Integer
    Dim rsGET_TYPE_ID                                  As ADODB.Recordset
    Set rsGET_TYPE_ID = New ADODB.Recordset
    rsGET_TYPE_ID.Open "SELECT ID FROM AMIS_IMPORT_TYPE WHERE TYPE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGET_TYPE_ID.EOF And Not rsGET_TYPE_ID.BOF Then
        GET_TYPE_ID = rsGET_TYPE_ID!ID
    End If
    Set rsGET_TYPE_ID = Nothing
End Function

Sub InitJournal()
'txtJItemNo.Text = Format(cntDetails + 1, "0000")
    cboAcct_Code.Text = ""
    txtAcct_Name.Text = ""
    cboDRCR.ListIndex = -1
    cboTranType1.ListIndex = -1
    txtSearch.Text = ""
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
End Sub

Sub BringToFront()
    Picture1.Enabled = False
    cmdAddJournal.ZOrder 0
    cmdAddJournal.Visible = True
    fraAddJournal.ZOrder 0
    fraAddJournal.Visible = True
    fraAddJournal.Enabled = True
End Sub

Private Sub cmdSaveSettings_Click()
Dim SQL As String
SQL = "INSERT INTO AMIS_IMPORTTEMPLATE_DT(CHARTACCOUNT_ID,DEBIT,ACCTYPE,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4,TEMPLATE_HD_ID)"
SQL = SQL & "SELECT CHARTACCOUNT_ID,DEBIT,ACCTYPE,TRANTYPE1,TRANTYPE2,TRANTYPE3,TRANTYPE4," & cboTo.Text & " FROM AMIS_IMPORTTEMPLATE_DT WHERE TEMPLATE_HD_ID=" & cboFrom.Text & ""
gconDMIS.Execute SQL
FillDetails
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
            cmdFindAccount.Visible = True
            fraFindAccount.Visible = True
            cmdFindAccount.ZOrder 0
            fraFindAccount.ZOrder 0
            fraFindAccount.Enabled = True
            DoEvents
            On Error Resume Next
            txtSearch.SetFocus
        End If
    Case vbKeyF3
        If Picture1.Visible = True Then
            Picture1.Enabled = False
            lstDetails.Enabled = False
            cmdAddJournal_Click
        End If
    Case vbKeyF8
        If SearchBy = "NAME" Then
            SearchBy = "CODE": fraFindAccount.Caption = "Search Accounts by Account Code"
        Else
            SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
        End If
    Case vbKeyEscape
        If fraFindAccount.Visible = True Then
            If Me.ActiveControl.Name = "txtSearch" Then
                SendToBack
                Picture1.Enabled = True
            Else
                txtSearch.SetFocus
            End If
        Else
            If Picture1.Visible = True Then
                SendToBack
                lstDetails.Enabled = True
                Picture1.Enabled = True
                'storememvars
            End If
        End If
    Case Else

    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    SendToBack
    InitCbo
    txtSearch.Text = ""
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    initMemvars
    rsRefresh
    If Not rsAMIS_TEMPLATES.EOF And Not rsAMIS_TEMPLATES.BOF Then
        rsAMIS_TEMPLATES.MoveLast
    End If
    StoreMemVars
End Sub

Sub InitCbo()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select acctcode from AMIS_ChartAccount order by acctcode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
    Set rsChartAccount = Nothing

    Dim rsTranType                                     As ADODB.Recordset
    Set rsTranType = New ADODB.Recordset
    rsTranType.Open "SELECT TRANTYPE FROM AMIS_IMPORT_TRANTYPE ORDER BY TRANTYPE", gconDMIS, adOpenForwardOnly
    If Not rsTranType.EOF And Not rsTranType.BOF Then
        Combo_Loadval cboTransactionType, rsTranType
    End If
    Set rsTranType = Nothing

    Dim rsType                                         As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    rsType.Open "SELECT TYPE FROM AMIS_IMPORT_TYPE ORDER BY TYPE", gconDMIS, adOpenForwardOnly
    If Not rsType.EOF And Not rsType.BOF Then
        Combo_Loadval cboType, rsType
    End If
    Set rsType = Nothing

    Dim rsOTH                                          As ADODB.Recordset
    Set rsOTH = New ADODB.Recordset
    rsOTH.Open "SELECT UPPER(DESCNAME) FROM CMIS_SBOOK WHERE BOOK='D' ORDER BY DESCNAME", gconDMIS, adOpenForwardOnly
    If Not rsOTH.EOF And Not rsOTH.BOF Then
        Combo_Loadval cboOTH, rsOTH
    End If
    Set rsOTH = Nothing

    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    rsModel.Open "SELECT MODEL FROM ALL_MODEL WHERE ID IN (SELECT MIN(ID) FROM ALL_MODEL B WHERE ALL_MODEL.MODEL=B.MODEL) ORDER BY MODEL", gconDMIS, adOpenForwardOnly
    If Not rsModel.EOF And Not rsModel.BOF Then
        'Combo_Loadval cboModel, rsModel
        cboModel.Clear
        cboModel.AddItem ""
        Do While Not rsModel.EOF
            cboModel.AddItem Null2String(rsModel!Model)
            rsModel.MoveNext
        Loop
    End If
    Set rsModel = Nothing

    With cboPayOption
        .AddItem "CASH"
        .AddItem "CHECK"
        .AddItem "CARD"
    End With
End Sub

Private Sub lstAccounts_DblClick()
    labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
    labAcctID.Caption = lstAccounts.SelectedItem.SubItems(3)
    OkAccount
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    cboAcct_Code.Text = labAccountCode.Caption
    cmdFindAccount.ZOrder 1
    fraFindAccount.ZOrder 1
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labAccountCode.Caption = Item: cboAcct_Code.Text = Item
    labAcctID.Caption = lstAccounts.SelectedItem.SubItems(3)
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
        OkAccount
    End If
End Sub

Private Sub lstDetails_DblClick()
    If cntDetails > 0 Then
        AddorEdit = "EDIT"
        cmdJournalDelete.Visible = True
        BringToFront
        Call StoreTemplateDetails(lstDetails.SelectedItem.SubItems(4))
        On Error Resume Next
    End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        Call FillSearchGrid(txtSearch.Text)
    End If
End Sub

Private Sub FillGrid()
    Dim rsChartAccount2                                As ADODB.Recordset
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
    Set rsChartAccount2 = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccount2                                As ADODB.Recordset
    lstAccounts.Enabled = False
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccount2 = New ADODB.Recordset
    'XXX = Repleys(LTrim(RTrim(XXX)))
    If SearchBy = "NAME" Then
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,upper(Description),Accttype,ID from AMIS_ChartAccount where description like '%" & XXX & "%' order by acctcode asc")
    Else
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,UPPER(Description),Accttype,ID from AMIS_ChartAccount where acctcode like '%" & XXX & "%' order by acctcode asc")
    End If
    If Not (rsChartAccount2.EOF And rsChartAccount2.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccount2
        lstAccounts.Refresh
        lstAccounts.Enabled = True
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
    Set rsChartAccount2 = Nothing
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstAccounts.SetFocus
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub rsRefresh()
    Set rsAMIS_TEMPLATES = New ADODB.Recordset
    rsAMIS_TEMPLATES.Open "SELECT * FROM AMIS_IMPORTTEMPLATE_HD ORDER BY ID", gconDMIS, adOpenForwardOnly
End Sub

Private Sub StoreMemVars()
    If Not rsAMIS_TEMPLATES.EOF And Not rsAMIS_TEMPLATES.BOF Then
        labID.Caption = rsAMIS_TEMPLATES!ID
        txtCode.Text = Null2String(rsAMIS_TEMPLATES!Code)
        txtDescription.Text = Null2String(rsAMIS_TEMPLATES!DESCRIPTION)
        cboTransactionType.Text = Null2String(GET_TRANTYPE(rsAMIS_TEMPLATES!TRANTYPE_ID))
        Call TRANTYPE_DETAIL(rsAMIS_TEMPLATES!TRANTYPE_ID)
        cboType.Text = Null2String(GET_TYPE(rsAMIS_TEMPLATES!Type_ID))
        cboModeOfPayment.Text = GetModeOfPayment(Null2String(rsAMIS_TEMPLATES!MODEOFPAYMENT))
        
        If N2Str2Null(rsAMIS_TEMPLATES!MODEL_ID) <> "NULL" And cboType.Text = "OTHERS" Then
            cboOTH.Text = Null2String(GET_DESCNAME(rsAMIS_TEMPLATES!MODEL_ID))
        End If
        If Null2Bit(rsAMIS_TEMPLATES!genuine) = 1 Then
            chkGenuine.Value = 1
        Else
            chkGenuine.Value = 0
        End If
        If Null2Bit(rsAMIS_TEMPLATES!NONVAT) = 1 Then
            chkNonVat.Value = 1
        Else
            chkNonVat.Value = 0
        End If

        If Null2Bit(rsAMIS_TEMPLATES!DISCOUNT) = 1 Then
            chkDiscount.Value = 1
        Else
            chkDiscount.Value = 0
        End If

        If cboTransactionType = "SALES" And cboType.Text = "VEHICLES" Then
            If N2Str2Null(rsAMIS_TEMPLATES!MODEL_ID) <> "NULL" Then
                cboModel.Text = Null2String(rsAMIS_TEMPLATES!MODEL_ID)
            End If
            'SetModel (Null2String(rsAMIS_TEMPLATES!MODEL_ID))
            cboPayType.ListIndex = -1
            cboPayClass.ListIndex = -1
            picTranType.Visible = False
            cboTranType2.ListIndex = -1
            cboTranType3.ListIndex = -1
            picPURCHASES.Visible = False
            picSALESPARTS.Visible = False
            picSALESSERVICE.Visible = False
            picSALESVEHICLE.Visible = True
            'If Null2String(rsAMIS_TEMPLATES!TERMS) = "BPO" Then
                cboTerm.Text = GetTERM(Null2String(rsAMIS_TEMPLATES!TERMS))
'            ElseIf Null2String(rsAMIS_TEMPLATES!TERMS) = "COD" Then
'                cboTerm.Text = "CASH"
'            ElseIf Null2String(rsAMIS_TEMPLATES!TERMS) = "CPO" Then
'                cboTerm.Text = "COMPANY PO"
'            ElseIf Null2String(rsAMIS_TEMPLATES!TERMS) = "F" Then
'                cboTerm.Text = "FINANCING"
'            End If
            If Null2Bit(rsAMIS_TEMPLATES!CWT) = 1 Then
                chkCWT.Value = 1
            Else
                chkCWT.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!INSURANCE2) = 1 Then
                chkInsurance2.Value = 1
            Else
                chkInsurance2.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!LTO) = 1 Then
                chkLTO.Value = 1
            Else
                chkLTO.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!CHATTEL) = 1 Then
                chkChattel.Value = 1
            Else
                chkChattel.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!FREEBIES) = 1 Then
                chkFreebies.Value = 1
            Else
                chkFreebies.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!DISCOUNT) = 1 Then
                chkDisc.Value = 1
            Else
                chkDisc.Value = 0
            End If
            
            If Null2Bit(rsAMIS_TEMPLATES!DOWNPAYMENT) = 1 Then
                chkDown.Value = 1
            Else
                chkDown.Value = 0
            End If
        ElseIf cboTransactionType = "SALES" And (cboType.Text = "PARTS" Or cboType.Text = "MATERIALS" Or cboType.Text = "ACCESSORIES") Then
            cboModel.ListIndex = -1
            cboPayClass.ListIndex = -1
            cboPayType.Text = Null2String(rsAMIS_TEMPLATES!PAY_TYPE)
            If Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "C" Then
                cboPayClass.Text = "CUSTOMER PAID"
            ElseIf Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "I" Then
                cboPayClass.Text = "INTERNAL"
            ElseIf Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "W" Then
                cboPayClass.Text = "WARRANTY"
            End If
            If Null2Bit(rsAMIS_TEMPLATES!CWT) = 1 Then
                chkCWTParts.Value = 1
            Else
                chkCWTParts.Value = 0
            End If
            picTranType.Visible = False
            cboTranType2.ListIndex = -1
            cboTranType3.ListIndex = -1
            picPURCHASES.Visible = False
            picSALESSERVICE.Visible = False
            picSALESVEHICLE.Visible = False
            picSALESPARTS.Visible = True
            '        ElseIf cboTransactionType = "SALES" And (cboType.Text = "MATERIALS") Then
            '            cboModel.ListIndex = -1
            '            cboPayClass.ListIndex = -1
            '            cboPayType.Text = Null2String(rsAMIS_TEMPLATES!PAY_TYPE)
            '            If Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "C" Then
            '                cboPayClass.Text = "CUSTOMER PAID"
            '            ElseIf Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "I" Then
            '                cboPayClass.Text = "INTERNAL"
            '            ElseIf Null2String(rsAMIS_TEMPLATES!PAY_CLASS) = "W" Then
            '                cboPayClass.Text = "WARRANTY"
            '            End If
            '            picTranType.Visible = False
            '            cboTranType2.ListIndex = -1
            '            cboTranType3.ListIndex = -1
            '            picPURCHASES.Visible = False
            '            picSALESSERVICE.Visible = False
            '            picSALESPARTS.Visible = True
        ElseIf cboTransactionType = "SALES" And cboType.Text = "SERVICE" Then
            picPURCHASES.Visible = False
            picSALESPARTS.Visible = False
            picSALESSERVICE.Visible = True
            picSALESVEHICLE.Visible = False
            picTranType.Visible = True
            cboModel.ListIndex = -1
            picSALESVEHICLE.Visible = False
            cboPayType.ListIndex = -1
            cboPayClass.ListIndex = -1
            cboModel.ListIndex = -1
        ElseIf cboTransactionType.Text = "PURCHASES" Then
            picPURCHASES.Visible = True
            picSALESPARTS.Visible = False
            picSALESSERVICE.Visible = False
            picSALESVEHICLE.Visible = False
            picTranType.Visible = True
            picTranType2.Visible = False
        ElseIf cboTransactionType.Text = "RECEIPTS" Then
            cboPayOption.Text = Null2String(rsAMIS_TEMPLATES!PAY_TYPE)
        Else
            cboPayType.ListIndex = -1
            cboPayClass.ListIndex = -1
            cboModel.ListIndex = -1
            picSALESSERVICE.Visible = False
            picPURCHASES.Visible = False
            picSALESPARTS.Visible = False
            picSALESVEHICLE.Visible = False
            picTranType.Visible = False
            cboTranType2.ListIndex = -1
            cboTranType3.ListIndex = -1
        End If
        ServiceSetupDetails
        WithholdingTaxAgent
        FillDetails
    Else
        MsgBox "No Such Record!": If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Function GET_TRANTYPE(XXX As Integer) As String
    Dim rsGET_TRANTYPE                                 As ADODB.Recordset
    Set rsGET_TRANTYPE = New ADODB.Recordset
    rsGET_TRANTYPE.Open "SELECT TRANTYPE FROM AMIS_IMPORT_TRANTYPE WHERE ID = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGET_TRANTYPE.EOF And Not rsGET_TRANTYPE.BOF Then
        GET_TRANTYPE = Null2String(rsGET_TRANTYPE!TranType)
    End If
    Set rsGET_TRANTYPE = Nothing
End Function

Function GET_TYPE(XXX As Integer) As String
    Dim rsGET_TYPE                                     As ADODB.Recordset
    Set rsGET_TYPE = New ADODB.Recordset
    rsGET_TYPE.Open "SELECT TYPE FROM AMIS_IMPORT_TYPE WHERE ID = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGET_TYPE.EOF And Not rsGET_TYPE.BOF Then
        GET_TYPE = Null2String(rsGET_TYPE!Type)
    End If
    Set rsGET_TYPE = Nothing
End Function

Private Sub initMemvars()
    txtCode.Text = ""
    txtDescription.Text = ""
    cboTransactionType.ListIndex = -1
    cboType.ListIndex = -1
    cboTranType1.ListIndex = -1
    cboTranType2.ListIndex = -1
    cboTranType3.ListIndex = -1
    cboTerm.ListIndex = -1
    cboOTH.ListIndex = -1
    cboModel.ListIndex = -1
    lstDetails.ListItems.Clear
    chkGenuine.Value = 0
    chkNonVat.Value = 0
    chkCustomer.Value = 0
    chkInternal.Value = 0
    chkWarranty.Value = 0
    chkInsurance.Value = 0
    chkLaborGJ.Value = 0
    chkLaborBP.Value = 0
    chkLaborPMS.Value = 0
    chkPartsGJ.Value = 0
    chkPartsBP.Value = 0
    chkSublet.Value = 0
    chkMaterials.Value = 0
    chkAccGJ.Value = 0
    chkAccBP.Value = 0
    chkDiscLabor.Value = 0
    chkDiscParts.Value = 0
    chkDiscMaterials.Value = 0
    chkDiscAcc.Value = 0
    chkCWT.Value = 0
    chkInsurance2.Value = 0
    chkLTO.Value = 0
    chkChattel.Value = 0
    chkFreebies.Value = 0
    chkDisc.Value = 0
    chkDown.Value = 0
    picSALESPARTS.Visible = False
    picPURCHASES.Visible = False
    picSALESVEHICLE.Visible = False
    picSALESPARTS.Visible = False
    picSALESSERVICE.Visible = False
End Sub

Function Get_Code() As String
    Dim rsGet_Code                                     As ADODB.Recordset
    Set rsGet_Code = New ADODB.Recordset
    rsGet_Code.Open "SELECT ISNULL(MAX(CODE),0) + 1 AS MAXCODE FROM AMIS_IMPORTTEMPLATE_HD", gconDMIS, adOpenForwardOnly
    If Not rsGet_Code.EOF And Not rsGet_Code.BOF Then
        Get_Code = Format(rsGet_Code!MAXCODE, "000000")
    Else
        Get_Code = "000001"
    End If
    Set rsGet_Code = Nothing
End Function

Sub ErrHandler(objCon As Object)
    Dim ADOErr                                         As ADODB.error
    Dim strError                                       As String

    For Each ADOErr In objCon.Errors
        strError = "Error #: " & ADOErr.Number & vbCrLf & _
                   "Error Description : " & ADOErr.DESCRIPTION
    Next

    MsgBox strError, vbCritical, "Error"
    objCon.Errors.Clear
End Sub

Sub FillDetails()
    Dim rsTemplate_Details                             As ADODB.Recordset
    Dim XDETAILS                                       As ListItem
    Set rsTemplate_Details = New ADODB.Recordset
    rsTemplate_Details.Open "SELECT DT.ID,AC.ACCTCODE,AC.DESCRIPTION,DT.DEBIT FROM AMIS_IMPORTTEMPLATE_DT DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.CHARTACCOUNT_ID=AC.ID INNER JOIN AMIS_IMPORTTEMPLATE_HD HD ON HD.ID=DT.TEMPLATE_HD_ID WHERE DT.TEMPLATE_HD_ID='" & labID.Caption & "' ORDER BY DT.ID", gconDMIS, adOpenForwardOnly
    lstDetails.ListItems.Clear
    cntDetails = 0
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        cntDetails = 1
        Do While Not rsTemplate_Details.EOF
            Set XDETAILS = lstDetails.ListItems.Add(, , Format(cntDetails, "0000"))
            XDETAILS.SubItems(1) = rsTemplate_Details!ACCTCODE
            XDETAILS.SubItems(2) = rsTemplate_Details!DESCRIPTION
            XDETAILS.SubItems(3) = DebitorCredit(rsTemplate_Details!DEBIT)
            XDETAILS.SubItems(4) = rsTemplate_Details!ID
            rsTemplate_Details.MoveNext
            cntDetails = cntDetails + 1
        Loop
    End If
    Set rsTemplate_Details = Nothing
End Sub

Function DebitorCredit(XXX As Boolean) As String
    Dim rsDebitorCredit                                As ADODB.Recordset
    Set rsDebitorCredit = New ADODB.Recordset
    rsDebitorCredit.Open "SELECT DEBIT FROM AMIS_IMPORTTEMPLATE_DT WHERE DEBIT = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsDebitorCredit.EOF And Not rsDebitorCredit.BOF Then
        If rsDebitorCredit!DEBIT = True Then
            DebitorCredit = "DEBIT"
        Else
            DebitorCredit = "CREDIT"
        End If
    End If
    Set rsDebitorCredit = Nothing
End Function

Function StoreTemplateDetails(xID As Integer)
    Dim rsTemplate_Details                             As ADODB.Recordset
    Dim XDETAILS                                       As ListItem
    Dim cntDetails                                     As Integer
    Set rsTemplate_Details = New ADODB.Recordset
    rsTemplate_Details.Open "SELECT DT.ID,AC.ACCTCODE,AC.DESCRIPTION,DT.DEBIT,DT.CHARTACCOUNT_ID,DT.TRANTYPE1,DT.TRANTYPE2,DT.TRANTYPE3 FROM AMIS_IMPORTTEMPLATE_DT DT INNER JOIN AMIS_CHARTACCOUNT AC ON DT.CHARTACCOUNT_ID=AC.ID INNER JOIN AMIS_IMPORTTEMPLATE_HD HD ON HD.ID=DT.TEMPLATE_HD_ID WHERE DT.ID='" & xID & "' ORDER BY DT.ID", gconDMIS, adOpenForwardOnly
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        labDetID.Caption = rsTemplate_Details!ID
        labAcctID.Caption = rsTemplate_Details!CHARTACCOUNT_ID
        cboAcct_Code.Text = rsTemplate_Details!ACCTCODE
        txtAcct_Name.Text = rsTemplate_Details!DESCRIPTION
        cboDRCR.Text = DebitorCredit(rsTemplate_Details!DEBIT)
        picTranType.Visible = True
        If IsNull(rsTemplate_Details!TRANTYPE1) = False Then
            cboTranType1.Text = Null2String(rsTemplate_Details!TRANTYPE1)
        Else
            cboTranType1.ListIndex = -1
        End If
        If IsNull(rsTemplate_Details!Trantype2) = False Then
            cboTranType2.Text = Null2String(rsTemplate_Details!Trantype2)
        Else
            cboTranType2.ListIndex = -1
        End If
        If IsNull(rsTemplate_Details!Trantype3) = False Then
            cboTranType3.Text = Null2String(rsTemplate_Details!Trantype3)
        Else
            cboTranType3.ListIndex = -1
        End If
    End If

    If cboTransactionType.Text = "PURCHASES" Then
        picTranType2.Visible = False
    ElseIf cboTransactionType.Text = "SALES" And cboType.Text = "SERVICE" Then
        picTranType2.Visible = True
    ElseIf cboTransactionType.Text = "RECEIPTS" Then
        picTranType2.Visible = False
    Else
        picTranType2.Visible = False
    End If
End Function

Sub WithholdingTaxAgent()
    Dim rsWithholdingTaxAgent                          As ADODB.Recordset
    Set rsWithholdingTaxAgent = New ADODB.Recordset
    rsWithholdingTaxAgent.Open "SELECT WTAXAGENT FROM ALL_PROFILE WHERE MODULENAME='AMIS'", gconDMIS, adOpenForwardOnly
    If Not rsWithholdingTaxAgent.EOF And Not rsWithholdingTaxAgent.BOF Then
        If Null2Bit(rsWithholdingTaxAgent!WTAXAGENT) = 1 Then
            chkEWT.Value = 1
        Else
            chkEWT.Value = 0
        End If
    End If
    Set rsWithholdingTaxAgent = Nothing
End Sub

Function SetModelID(XXX As String) As Integer
    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    rsModel.Open "SELECT ID FROM ALL_MODEL WHERE MODEL = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsModel.EOF And Not rsModel.BOF Then
        SetModelID = rsModel!ID
    End If
    Set rsModel = Nothing
End Function

Function SetModel(XXX As Integer) As String
    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    rsModel.Open "SELECT MODEL FROM ALL_MODEL WHERE ID = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsModel.EOF And Not rsModel.BOF Then
        SetModel = rsModel!Model
    End If
    Set rsModel = Nothing
End Function

Sub ServiceSetupDetails()
    If Null2Bit(rsAMIS_TEMPLATES!Customer) = 1 Then
        chkCustomer.Value = 1
    Else
        chkCustomer.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Internal) = 1 Then
        chkInternal.Value = 1
    Else
        chkInternal.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Warranty) = 1 Then
        chkWarranty.Value = 1
    Else
        chkWarranty.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!INSURANCE) = 1 Then
        chkInsurance.Value = 1
    Else
        chkInsurance.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Labor_GJ) = 1 Then
        chkLaborGJ.Value = 1
    Else
        chkLaborGJ.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Labor_BP) = 1 Then
        chkLaborBP.Value = 1
    Else
        chkLaborBP.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Labor_PMS) = 1 Then
        chkLaborPMS.Value = 1
    Else
        chkLaborPMS.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Parts_GJ) = 1 Then
        chkPartsGJ.Value = 1
    Else
        chkPartsGJ.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Parts_BP) = 1 Then
        chkPartsBP.Value = 1
    Else
        chkPartsBP.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!MATERIALS) = 1 Then
        chkMaterials.Value = 1
    Else
        chkMaterials.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Acc_GJ) = 1 Then
        chkAccGJ.Value = 1
    Else
        chkAccGJ.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Acc_BP) = 1 Then
        chkAccBP.Value = 1
    Else
        chkAccBP.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!SUBLET) = 1 Then
        chkSublet.Value = 1
    Else
        chkSublet.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Disc_Labor) = 1 Then
        chkDiscLabor.Value = 1
    Else
        chkDiscLabor.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Disc_Parts) = 1 Then
        chkDiscParts.Value = 1
    Else
        chkDiscParts.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Disc_Mat) = 1 Then
        chkDiscMaterials.Value = 1
    Else
        chkDiscMaterials.Value = 0
    End If

    If Null2Bit(rsAMIS_TEMPLATES!Disc_Acc) = 1 Then
        chkDiscAcc.Value = 1
    Else
        chkDiscAcc.Value = 0
    End If
End Sub

Sub TRANTYPE_DETAIL(XXX As Integer)
    Dim rsTranTypeClass                                As ADODB.Recordset
    Set rsTranTypeClass = New ADODB.Recordset
    rsTranTypeClass.Open "SELECT TRANTYPE1 FROM AMIS_IMPORT_CLASS WHERE TRANTYPE_ID='" & XXX & "' ORDER BY TRANTYPE1", gconDMIS, adOpenForwardOnly
    If Not rsTranTypeClass.EOF And Not rsTranTypeClass.BOF Then
        Combo_Loadval cboTranType1, rsTranTypeClass
    End If
    Set rsTranTypeClass = Nothing
End Sub

Function GET_OTHERCODE(XXX As String) As String
    Dim rsOTH                                          As ADODB.Recordset
    Set rsOTH = New ADODB.Recordset
    rsOTH.Open "SELECT CODE FROM CMIS_SBOOK WHERE BOOK='D' AND DESCNAME = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsOTH.EOF And Not rsOTH.BOF Then
        GET_OTHERCODE = Null2String(rsOTH!Code)
    End If
    Set rsOTH = Nothing
End Function

Function GET_DESCNAME(XXX As String) As String
    Dim rsOTH                                          As ADODB.Recordset
    Set rsOTH = New ADODB.Recordset
    rsOTH.Open "SELECT DESCNAME FROM CMIS_SBOOK WHERE BOOK='D' AND CODE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsOTH.EOF And Not rsOTH.BOF Then
        GET_DESCNAME = Null2String(rsOTH!DESCNAME)
    End If
    Set rsOTH = Nothing
End Function

Function GetModeOfPayment(XXX As String)
    If XXX = "LC" Then
        GetModeOfPayment = "Letter of Credit"
    ElseIf XXX = "CA" Then
        GetModeOfPayment = "Cash"
    ElseIf XXX = "OA" Then
        GetModeOfPayment = "Open Account"
    ElseIf XXX = "PN" Then
        GetModeOfPayment = "Promissory Note"
    ElseIf XXX = "FC" Then
        GetModeOfPayment = "Financing Co."
    End If
End Function

Function SetModeOfPayment(XXX As String)
    XXX = UCase(XXX)
    If XXX = UCase("Letter of Credit") Then
        SetModeOfPayment = "LC"
    ElseIf XXX = UCase("Open Account") Then
        SetModeOfPayment = "OA"
    ElseIf XXX = UCase("Promissory Note") Then
        SetModeOfPayment = "PN"
    ElseIf XXX = UCase("Financing Co.") Then
        SetModeOfPayment = "FC"
    ElseIf XXX = UCase("Cash") Then
        SetModeOfPayment = "CA"
    Else
        SetModeOfPayment = "NULL"
    End If
End Function

Function GetTERM(XXX As String) As String
    If XXX = "CPO" Then
        GetTERM = "COMPANY PO"
    ElseIf XXX = "F" Then
        GetTERM = "FINANCING"
    ElseIf XXX = "COD" Then
        GetTERM = "CASH"
    ElseIf XXX = "BPO" Then
        GetTERM = "BANK PO"
    End If
End Function

Function SetTERM(XXX As String) As String
    XXX = UCase(XXX)
    If XXX = "COMPANY PO" Then
        SetTERM = "CPO"
    ElseIf XXX = "FINANCING" Then
        SetTERM = "F"
    ElseIf XXX = "CASH" Then
        SetTERM = "COD"
    ElseIf XXX = "BANK PO" Then
        SetTERM = "BPO"
    Else
        SetTERM = "NULL"
    End If
End Function
