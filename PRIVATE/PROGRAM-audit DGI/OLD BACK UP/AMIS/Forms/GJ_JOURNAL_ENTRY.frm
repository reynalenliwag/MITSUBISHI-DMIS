VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMISJournalEntry_GJ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   7770
   ClientLeft      =   11040
   ClientTop       =   4800
   ClientWidth     =   9885
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GJ_JOURNAL_ENTRY.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   9885
   Begin VB.PictureBox picGJ 
      BorderStyle     =   0  'None
      Height          =   6465
      Left            =   150
      ScaleHeight     =   6465
      ScaleWidth      =   9555
      TabIndex        =   70
      Top             =   420
      Width           =   9555
      Begin VB.TextBox txtGJ_Remarks 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   225
         Top             =   1080
         Width           =   9405
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   465
         Left            =   0
         TabIndex        =   205
         Top             =   6000
         Width           =   9525
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
            Left            =   3840
            MaxLength       =   14
            TabIndex        =   208
            Text            =   "Text1"
            Top             =   90
            Width           =   1515
         End
         Begin VB.TextBox txtGJTotDebit 
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
            Left            =   6450
            MaxLength       =   14
            TabIndex        =   207
            Text            =   "Text1"
            Top             =   90
            Width           =   1515
         End
         Begin VB.TextBox txtGJTotCredit 
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
            Left            =   8010
            MaxLength       =   14
            TabIndex        =   206
            Text            =   "Text1"
            Top             =   90
            Width           =   1485
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
            Left            =   90
            TabIndex        =   218
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Out of Balance"
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
            Height          =   285
            Left            =   870
            TabIndex        =   209
            Top             =   120
            Width           =   2925
         End
         Begin VB.Label labGJID 
            Caption         =   "labGJID"
            Height          =   375
            Left            =   2970
            TabIndex        =   214
            Top             =   180
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   390
         Top             =   4740
      End
      Begin MSComctlLib.ListView lstGJ 
         Height          =   4425
         Left            =   90
         TabIndex        =   73
         Top             =   1590
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   7805
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":08CA
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
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   90
         TabIndex        =   72
         Top             =   1980
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"GJ_JOURNAL_ENTRY.frx":0A2C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblINVOICE_DETAIL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   8010
         TabIndex        =   228
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label lblINVOICETYPE_DETAIL 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   7140
         TabIndex        =   227
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
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
         Height          =   495
         Left            =   -180
         TabIndex        =   226
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label labcode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1260
         TabIndex        =   224
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label labJournalType 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   4740
         TabIndex        =   223
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label labInvoiceNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   345
         Left            =   1260
         TabIndex        =   222
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label lbljournaltype 
         Alignment       =   2  'Center
         Caption         =   "Journal Type"
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
         Height          =   375
         Left            =   3150
         TabIndex        =   221
         Top             =   540
         Width           =   1815
      End
      Begin VB.Label lblinvoiceno 
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
         Height          =   345
         Left            =   120
         TabIndex        =   220
         Top             =   510
         Width           =   2385
      End
      Begin VB.Label labName 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2520
         TabIndex        =   219
         Top             =   60
         Width           =   6975
      End
      Begin VB.Label lblname 
         Alignment       =   1  'Right Justify
         Caption         =   "Cust. Name"
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
         Height          =   345
         Left            =   -60
         TabIndex        =   217
         Top             =   90
         Width           =   1245
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
         Left            =   2490
         TabIndex        =   216
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label labDET 
         Caption         =   "labDET"
         Height          =   285
         Left            =   1530
         TabIndex        =   215
         Top             =   30
         Visible         =   0   'False
         Width           =   2295
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
         Left            =   630
         TabIndex        =   71
         Top             =   1290
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Detail"
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
         Height          =   345
         Left            =   6510
         TabIndex        =   229
         Top             =   510
         Width           =   2385
      End
   End
   Begin TabDlg.SSTab JournalTAB 
      Height          =   4215
      Left            =   180
      TabIndex        =   99
      Top             =   2490
      Width           =   9525
      _ExtentX        =   16801
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
      TabCaption(0)   =   "[<F3> Add &Journals]   [<Ctrl> + <J> View &Journals]   "
      TabPicture(0)   =   "GJ_JOURNAL_ENTRY.frx":0ABF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdAddJournal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAddJournal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetails"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[<F4> Add &Details]   [<Ctrl> + <D> View &Details]   "
      TabPicture(1)   =   "GJ_JOURNAL_ENTRY.frx":0ADB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPV_Entry"
      Tab(1).Control(1)=   "cmdPV_Entry"
      Tab(1).Control(2)=   "picPV_Detail"
      Tab(1).ControlCount=   3
      Begin VB.PictureBox fraDetails 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   90
         ScaleHeight     =   3435
         ScaleWidth      =   9405
         TabIndex        =   120
         Top             =   120
         Width           =   9405
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   30
            Top             =   3000
         End
         Begin MSComctlLib.ListView lstDetails 
            Height          =   2835
            Left            =   30
            TabIndex        =   121
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":0AF7
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
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Enabled         =   0   'False
            Height          =   465
            Left            =   30
            TabIndex        =   122
            Top             =   2940
            Width           =   9345
            Begin VB.PictureBox picChat 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   60
               ScaleHeight     =   345
               ScaleWidth      =   6195
               TabIndex        =   123
               Top             =   60
               Visible         =   0   'False
               Width           =   6195
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
                  TabIndex        =   124
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
               TabIndex        =   126
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
               TabIndex        =   128
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
               TabIndex        =   127
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
               TabIndex        =   125
               Top             =   90
               Width           =   1275
            End
         End
      End
      Begin VB.PictureBox fraAddJournal 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   130
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
            Left            =   8325
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":0C59
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":0DAB
            Style           =   1  'Graphical
            TabIndex        =   163
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":10E9
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":123B
            Style           =   1  'Graphical
            TabIndex        =   146
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
            TabIndex        =   142
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
            TabIndex        =   140
            Top             =   330
            Width           =   1100
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   2310
            TabIndex        =   137
            Top             =   -30
            Width           =   4425
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   315
               Left            =   30
               TabIndex        =   139
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393217
               BackColor       =   16777215
               MultiLine       =   0   'False
               TextRTF         =   $"GJ_JOURNAL_ENTRY.frx":1566
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
               TabIndex        =   138
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
            TabIndex        =   135
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
            TabIndex        =   136
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
            TabIndex        =   134
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin VB.Frame fraATC 
            Height          =   915
            Left            =   2340
            TabIndex        =   147
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
               Style           =   2  'Dropdown List
               TabIndex        =   151
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
               TabIndex        =   152
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
               TabIndex        =   153
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
               TabIndex        =   154
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
               TabIndex        =   149
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
               TabIndex        =   148
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
               TabIndex        =   150
               Top             =   240
               Width           =   1725
            End
         End
         Begin VB.Frame fraComp 
            Height          =   915
            Left            =   2340
            TabIndex        =   155
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
               TabIndex        =   161
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
               TabIndex        =   160
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
               TabIndex        =   159
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
               TabIndex        =   158
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
               TabIndex        =   157
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
               TabIndex        =   156
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":15F9
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":174B
            Style           =   1  'Graphical
            TabIndex        =   162
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
            TabIndex        =   143
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
            TabIndex        =   131
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
            TabIndex        =   132
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
            TabIndex        =   133
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
            TabIndex        =   141
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
            Left            =   2880
            TabIndex        =   145
            Top             =   420
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
            TabIndex        =   144
            Top             =   420
            Width           =   2685
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1845
         Left            =   120
         TabIndex        =   129
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
         MICON           =   "GJ_JOURNAL_ENTRY.frx":1A9B
      End
      Begin VB.PictureBox picPV_Entry 
         BackColor       =   &H00FF8080&
         Height          =   1575
         Left            =   -74790
         ScaleHeight     =   1515
         ScaleWidth      =   9105
         TabIndex        =   105
         Top             =   750
         Width           =   9165
         Begin VB.CommandButton Command4 
            Caption         =   ".."
            Height          =   345
            Left            =   5010
            TabIndex        =   211
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
            TabIndex        =   189
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":1AB7
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":1C09
            Style           =   1  'Graphical
            TabIndex        =   119
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":1F47
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":2099
            Style           =   1  'Graphical
            TabIndex        =   117
            Top             =   690
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtMRR_No 
            Height          =   315
            Left            =   1650
            TabIndex        =   1
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
            TabIndex        =   116
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
            TabIndex        =   114
            Top             =   330
            Width           =   1785
            _ExtentX        =   3149
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
         Begin MSMask.MaskEdBox txtPO_No 
            Height          =   315
            Left            =   60
            TabIndex        =   111
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
            TabIndex        =   115
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
            TabIndex        =   112
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":23C4
            MousePointer    =   99  'Custom
            Picture         =   "GJ_JOURNAL_ENTRY.frx":2516
            Style           =   1  'Graphical
            TabIndex        =   118
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
            TabIndex        =   188
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
            TabIndex        =   187
            Top             =   690
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
            TabIndex        =   110
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
            TabIndex        =   106
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
            TabIndex        =   107
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
            TabIndex        =   108
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
            TabIndex        =   109
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
            TabIndex        =   113
            Top             =   420
            Width           =   1305
         End
      End
      Begin wizButton.cmd cmdPV_Entry 
         Height          =   1635
         Left            =   -74820
         TabIndex        =   104
         Top             =   720
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   2884
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
         MICON           =   "GJ_JOURNAL_ENTRY.frx":2866
      End
      Begin VB.PictureBox picPV_Detail 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3795
         Left            =   -74940
         ScaleHeight     =   3795
         ScaleWidth      =   9405
         TabIndex        =   100
         Top             =   90
         Width           =   9405
         Begin MSComctlLib.ListView lstPV_Detail 
            Height          =   3285
            Left            =   60
            TabIndex        =   101
            Top             =   90
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
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":2882
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
         Begin MSMask.MaskEdBox txtTotalPV_Amount 
            Height          =   345
            Left            =   8010
            TabIndex        =   102
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
            TabIndex        =   103
            Top             =   3450
            Width           =   1275
         End
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   10020
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   204
      Top             =   6840
      Width           =   2775
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption56 
         Height          =   825
         Left            =   -30
         TabIndex        =   212
         Top             =   0
         Width           =   2745
         _Version        =   655364
         _ExtentX        =   4842
         _ExtentY        =   1455
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8388608
         GradientColorDark=   16711680
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   9900
      ScaleHeight     =   6765
      ScaleWidth      =   3945
      TabIndex        =   190
      Top             =   0
      Width           =   3945
      Begin VB.PictureBox pic3 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   4245
         Left            =   90
         ScaleHeight     =   4215
         ScaleWidth      =   2745
         TabIndex        =   198
         Top             =   30
         Width           =   2775
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   30
            TabIndex        =   201
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
            TabIndex        =   203
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
            TabIndex        =   202
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
            TabIndex        =   200
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   199
            Top             =   0
            Width           =   2805
            _Version        =   655364
            _ExtentX        =   4948
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Journal Options"
            ForeColor       =   14606302
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            GradientColorLight=   12582912
            GradientColorDark=   8388608
            ForeColor       =   14606302
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   2385
         Left            =   90
         ScaleHeight     =   2355
         ScaleWidth      =   2745
         TabIndex        =   191
         Top             =   4290
         Width           =   2775
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "System Setup"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   90
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":29E4
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   195
            Top             =   1500
            Width           =   2625
         End
         Begin VB.CommandButton cmdInternalRO 
            BackColor       =   &H00C0E0FF&
            Caption         =   "View Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   90
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":2B36
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   192
            Top             =   960
            Width           =   2625
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "View payment(s)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   90
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":2C88
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   210
            Top             =   420
            Width           =   2625
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Set to WARRANTY RO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":2DDA
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   196
            Top             =   1560
            Width           =   90
         End
         Begin VB.CommandButton cmdROVatExempt 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   330
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":2F2C
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   194
            Top             =   2370
            Width           =   2625
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   60
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":307E
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   193
            Top             =   2610
            Width           =   2625
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   345
            Index           =   1
            Left            =   0
            TabIndex        =   197
            Top             =   0
            Width           =   2775
            _Version        =   655364
            _ExtentX        =   4895
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "Journal Advance Options"
            ForeColor       =   14606302
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            GradientColorLight=   12582912
            GradientColorDark=   8388608
            ForeColor       =   14606302
         End
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   150
      ScaleHeight     =   2610
      ScaleWidth      =   9570
      TabIndex        =   0
      Top             =   0
      Width           =   9570
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
            ScrollBars      =   2
            TextRTF         =   $"GJ_JOURNAL_ENTRY.frx":31D0
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
            TabIndex        =   186
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
      Begin VB.PictureBox picReceivable 
         BorderStyle     =   0  'None
         Height          =   2235
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   9510
         TabIndex        =   36
         Top             =   3825
         Width           =   9510
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
            ScrollBars      =   2
            TextRTF         =   $"GJ_JOURNAL_ENTRY.frx":3264
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   60
         Width           =   1545
      End
      Begin VB.TextBox txtVoucherNo 
         Alignment       =   2  'Center
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
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   3
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
         TabIndex        =   7
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
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
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
            ScrollBars      =   2
            TextRTF         =   $"GJ_JOURNAL_ENTRY.frx":32FB
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
         TabIndex        =   213
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
         TabIndex        =   6
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal Date"
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
         Left            =   6630
         TabIndex        =   5
         Top             =   120
         Width           =   1260
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
         Left            =   30
         TabIndex        =   2
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
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   210
      ScaleHeight     =   615
      ScaleWidth      =   9435
      TabIndex        =   180
      Top             =   6060
      Width           =   9495
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F11 - Post Journals by Batch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4710
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":3392
         MousePointer    =   99  'Custom
         TabIndex        =   184
         Top             =   300
         Width           =   4605
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "F9 - Add Journal Entries from Templates"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4710
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":369C
         MousePointer    =   99  'Custom
         TabIndex        =   183
         Top             =   30
         Width           =   4605
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Journal Entries"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":39A6
         MousePointer    =   99  'Custom
         TabIndex        =   182
         Top             =   30
         Width           =   4605
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Add/View Journal Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   60
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":3CB0
         MousePointer    =   99  'Custom
         TabIndex        =   181
         Top             =   300
         Width           =   4605
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
   Begin VB.PictureBox picRefCDJ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6795
      ScaleHeight     =   345
      ScaleWidth      =   2775
      TabIndex        =   75
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
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
         TabIndex        =   76
         Top             =   0
         Width           =   2775
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1170
      TabIndex        =   74
      Top             =   930
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
      MICON           =   "GJ_JOURNAL_ENTRY.frx":3FBA
   End
   Begin VB.PictureBox picTemplates 
      Height          =   4125
      Left            =   1260
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   77
      Top             =   990
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
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   79
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":3FD6
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
         TabIndex        =   80
         Top             =   3750
         Width           =   7035
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
      Left            =   240
      TabIndex        =   64
      Top             =   360
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
         TabIndex        =   65
         Text            =   "Text1"
         Top             =   270
         Width           =   9195
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4515
         Left            =   90
         TabIndex        =   67
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":4138
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
         Left            =   60
         TabIndex        =   69
         Top             =   5310
         Width           =   9225
      End
   End
   Begin VB.CommandButton cmdPrinting 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Command1"
      Height          =   2445
      Left            =   3450
      TabIndex        =   90
      Top             =   1830
      Width           =   2775
   End
   Begin wizButton.cmd cmdShowPostRange 
      Height          =   2385
      Left            =   3540
      TabIndex        =   81
      Top             =   1830
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
      MICON           =   "GJ_JOURNAL_ENTRY.frx":429A
   End
   Begin VB.PictureBox picShowPostRange 
      Height          =   2235
      Left            =   3600
      ScaleHeight     =   2175
      ScaleWidth      =   2535
      TabIndex        =   82
      Top             =   1920
      Width           =   2595
      Begin VB.CommandButton cmdPostRange 
         Caption         =   "POST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   88
         Top             =   1350
         Width           =   2295
      End
      Begin wizProgBar.Prg prgPostRange 
         Height          =   285
         Left            =   90
         TabIndex        =   89
         Top             =   1800
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         Picture         =   "GJ_JOURNAL_ENTRY.frx":42B6
         ForeColor       =   0
         BarPicture      =   "GJ_JOURNAL_ENTRY.frx":42D2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.TextBox txtToVNo 
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
         Left            =   870
         MaxLength       =   10
         TabIndex        =   87
         Top             =   870
         Width           =   1485
      End
      Begin VB.TextBox txtFromVNo 
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
         Left            =   870
         MaxLength       =   10
         TabIndex        =   85
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Post By Range"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   83
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To     :"
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
         Left            =   120
         TabIndex        =   86
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
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
         Left            =   120
         TabIndex        =   84
         Top             =   480
         Width           =   735
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   5775
      Left            =   210
      TabIndex        =   63
      Top             =   240
      Width           =   9435
      _ExtentX        =   16642
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
      MICON           =   "GJ_JOURNAL_ENTRY.frx":42EE
   End
   Begin VB.PictureBox picPrinting 
      Height          =   2265
      Left            =   3600
      ScaleHeight     =   2205
      ScaleWidth      =   2535
      TabIndex        =   91
      Top             =   1920
      Width           =   2595
      Begin VB.PictureBox picPrintCheck 
         Enabled         =   0   'False
         Height          =   885
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   2355
         TabIndex        =   93
         Top             =   450
         Width           =   2415
         Begin VB.OptionButton optSECBANK 
            Caption         =   "EASTWEST BANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   94
            Top             =   -30
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optPRUDBANK 
            Caption         =   "EPCI BANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   95
            Top             =   240
            Width           =   2445
         End
         Begin VB.OptionButton optCHINBANK 
            Caption         =   "CHINABANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   96
            Top             =   510
            Width           =   2355
         End
      End
      Begin VB.CommandButton cmdOkPrint 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   390
         TabIndex        =   98
         Top             =   1830
         Width           =   1725
      End
      Begin VB.OptionButton optPrintVoucher 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRINT VOUCHER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1380
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optPrintCheck 
         Caption         =   "PRINT CHECK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   60
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -150
      ScaleHeight     =   900
      ScaleWidth      =   9735
      TabIndex        =   167
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":430A
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":445C
         Style           =   1  'Graphical
         TabIndex        =   179
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":47C2
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":4914
         Style           =   1  'Graphical
         TabIndex        =   178
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
         Left            =   7380
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":4C7A
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":4DCC
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Cancel this Transaction"
         Top             =   30
         Width           =   705
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
         Left            =   6600
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":5106
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":5258
         Style           =   1  'Graphical
         TabIndex        =   176
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
         Left            =   5850
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":559D
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":56EF
         Style           =   1  'Graphical
         TabIndex        =   175
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
         Left            =   5100
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":5A14
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":5B66
         Style           =   1  'Graphical
         TabIndex        =   174
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
         Left            =   4350
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":5EC2
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":6014
         Style           =   1  'Graphical
         TabIndex        =   173
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
         Left            =   3595
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":6327
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":6479
         Style           =   1  'Graphical
         TabIndex        =   172
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
         Left            =   2850
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":67C9
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":691B
         Style           =   1  'Graphical
         TabIndex        =   171
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
         Left            =   2100
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":6C79
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":6DCB
         Style           =   1  'Graphical
         TabIndex        =   170
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
         Left            =   1345
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":70C5
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":7217
         Style           =   1  'Graphical
         TabIndex        =   169
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
         Left            =   595
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":756F
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":76C1
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7920
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   164
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":7A20
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":7B72
         Style           =   1  'Graphical
         TabIndex        =   166
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
         MouseIcon       =   "GJ_JOURNAL_ENTRY.frx":7EB0
         MousePointer    =   99  'Custom
         Picture         =   "GJ_JOURNAL_ENTRY.frx":8002
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.Label lblVPJAcctCode 
      Caption         =   "dont delete this "
      Height          =   165
      Left            =   11220
      TabIndex        =   185
      Top             =   1080
      Width           =   1845
   End
End
Attribute VB_Name = "frmAMISJournalEntry_GJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As New ADODB.Recordset
Dim rsJournal_Det                                 As New ADODB.Recordset
Dim rsPV_Detail                                   As New ADODB.Recordset
Dim rsCV_Detail                                   As New ADODB.Recordset
Dim rsCRJ_Detail                                  As New ADODB.Recordset
Dim rsJV_detail                                   As New ADODB.Recordset
Dim rsChartAccount                                As New ADODB.Recordset
Dim rsJournal_HD2                                 As New ADODB.Recordset
Dim rsProfile                                     As New ADODB.Recordset
Dim rsCheckJournal_HD                             As New ADODB.Recordset
Dim rsVENDOR                                      As New ADODB.Recordset
Dim rsPayTerm                                     As New ADODB.Recordset
Dim rsBanks                                       As New ADODB.Recordset
Dim rsCustomer                                    As New ADODB.Recordset
Dim rsInvoiceType                                 As New ADODB.Recordset
Dim rsATC                                         As New ADODB.Recordset
Dim kcnt, Jcnt                                    As Integer
Attribute Jcnt.VB_VarUserMemId = 1073938448
Dim AddorEdit                                     As String
Attribute AddorEdit.VB_VarUserMemId = 1073938450
Dim SearchBy                                      As String
Public CDJ_CIB                                    As String
Attribute CDJ_CIB.VB_VarUserMemId = 1073938452
Public CDJ_AP                                     As String
Attribute CDJ_AP.VB_VarUserMemId = 1073938453
Dim LocalAcess                                    As String
Dim TOTDEBIT, TOTCREDIT, TOTTAX, OUTBALANCE, TOTAL_AR_AMOUNT, TOTALPVAMOUNT, COMP_SJ_OUTPUT_TAX As Double
Attribute TOTDEBIT.VB_VarUserMemId = 1073938454
Attribute TOTCREDIT.VB_VarUserMemId = 1073938454
Attribute TOTTAX.VB_VarUserMemId = 1073938454
Attribute OUTBALANCE.VB_VarUserMemId = 1073938454
Attribute TOTAL_AR_AMOUNT.VB_VarUserMemId = 1073938454
Attribute TOTALPVAMOUNT.VB_VarUserMemId = 1073938454
Attribute COMP_SJ_OUTPUT_TAX.VB_VarUserMemId = 1073938454
Dim PrevJType                                     As String
Attribute PrevJType.VB_VarUserMemId = 1073938461
Dim PrevJNo                                       As String
Dim PrevInvoiceType                               As String
Attribute PrevInvoiceType.VB_VarUserMemId = 1073938463
Dim PrevInvoiceNo                                 As String
Dim PrevPV_VoucherNo                              As String
Dim PrevPV_Amount                                 As Double
Dim DirectDisbursementVoucherNo                   As String
Dim CDJ_IS_FROM_AP                                As Boolean
Dim IsVPJ                                         As Boolean
Dim TotalARAmountToPay                            As Double
Dim TOTAL_AP_AMOUNT                               As Double
Dim TotalAPAmountToPay                            As Double
Dim SJVoucherno                                   As String
Dim APJInvoiceNo                                  As String
Dim APJinvoicetype                                As String
Dim xJOURNALTYPE                                  As String

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Function GetVoucherNo(XXX As String) As String
    Dim rsJournal_HD                              As ADODB.Recordset
    Set rsJournal_HD = New ADODB.Recordset
    Set rsJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = '" & XXX & "' Order by VoucherNo desc")
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_HD!VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Private Sub GetNewVoucherNo()
    Dim rsJournal_HDDup                           As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
End Sub

Function Setacctcode(VVV As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctcode = UCase(Null2String(rsChartAccount2!ACCTCODE))
    Else
        Setacctcode = ""
    End If
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function SetAcctType(VVV As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
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
        CDJ_CIB = N2Str2Null(rsBanks!ACCTCODE)
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
            CDJ_CIB = N2Str2Null(rsBanks!ACCTCODE)
        Else
            SetBankName = ""
            CDJ_CIB = "NULL"
        End If
    Else
        rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            SetBankName = Null2String(rsBanks!BankName)
            CDJ_CIB = N2Str2Null(rsBanks!ACCTCODE)
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
    Dim rsAccountType                             As ADODB.Recordset
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
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where nameofvendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!Code)
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

'commented by: JUN 07272009---------------------------------------------
'Function StoreGJEntry(ByVal ID As Variant)
'    On Error GoTo Errorcode
'    Set rsJournal_Det = New ADODB.Recordset
'    rsJournal_Det.Open "select id,JNo,acct_code,acct_name,debit,jitemno,credit,tax,atc,rate,taxbase from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
'
''        labGJID.Caption = rsJournal_Det!ID
''        txtGJItemNo.Text = Null2String(rsJournal_Det!jitemno)
''        cboGJAccountNo.Text = Null2String(rsJournal_Det!ACCT_CODE)
''        txtGJAccountName.Text = Null2String(rsJournal_Det!acct_Name)
''        txtGJDebit.Text = N2Str2Zero(rsJournal_Det!DEBIT)
''        txtGJCredit.Text = N2Str2Zero(rsJournal_Det!CREDIT)
''
''        If fraATC2.Visible = True Then
''            cboJVSupCust.Text = SetVendorName(Null2String(rsJournal_hd!VendorCode))
''            If Null2String(rsJournal_Det!ATC) <> "" Then
''                cboATC2.Text = Null2String(rsJournal_Det!ATC)
''            Else
''                cboATC2.ListIndex = 0
''            End If
''            txtRATE2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
''            txtTaxBase2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!taxbase))
''        End If
''        StoreGJParticulars Null2String(rsJournal_Det!JNo), Null2String(rsJournal_Det!jitemno)
'
'    End If
'    Exit Function
'
'Errorcode:
'    Resume Next
'End Function
'
'Function StoreGJParticulars(ByVal JNo As Variant, ByVal ItemNo As Variant)
'    Set rsJV_detail = New ADODB.Recordset
'    rsJV_detail.Open "select JNo,ItemNo,Particulars from AMIS_JV_Detail where JNo = " & N2Str2Null(JNo) & " and ItemNo = " & N2Str2Null(ItemNo), gconDMIS, adOpenForwardOnly, adLockReadOnly
'    If Not rsJV_detail.EOF And Not rsJV_detail.BOF Then
'        txtGJAccountParticulars.Text = Null2String(rsJV_detail!Particulars)
'    End If
'End Function
'commented by: JUN 07272009---------------------------------------------
Function StoreJournalEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,acct_code,acct_name,debit,jitemno,credit,tax,grossamt,netamt,ATC,RATE,TAXBASE from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labDetID.Caption = rsJournal_Det!ID
        labPartNo.Caption = Null2String(rsJournal_Det!Acct_code)
        txtJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboAcct_Code.Text = Null2String(rsJournal_Det!Acct_code)
        txtAcct_Name.Text = Null2String(rsJournal_Det!acct_Name)
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
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
            txtTaxBase.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!taxbase))
        Else
            ' Update By BTT : 09262008
            If Null2String(rsJournal_Det!ATC) <> "" Then
                cboATC.Text = Null2String(rsJournal_Det!ATC)
            End If
            txtRATE.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
            txtTaxBase.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!taxbase))
        End If
    End If
End Function

Function StoreDealerCode(XXX As String) As String
    Dim rsREPORWITHDealer                         As ADODB.Recordset
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
            txtPO_No.Text = Null2String(rsCV_Detail!jtype)
            txtMRR_No.Text = Null2String(rsCV_Detail!pv_voucherno)
            PrevPV_VoucherNo = Null2String(rsCV_Detail!pv_voucherno)
            txtINV_No.Text = Null2String(rsCV_Detail!docdate)
            txtProd_No.Text = Null2String(rsCV_Detail!duedate)
            txtPVAmount.Text = N2Str2Zero(rsCV_Detail!amount)
            txtMRR_No.Enabled = False
            txtProd_No.Enabled = True
            txtINV_No.Enabled = True
            PrevPV_Amount = N2Str2Zero(rsCV_Detail!amount)
        End If
    End If
End Function

Function ReturnAP_AccountCode(XXX As String) As String
    Dim rsChartAccount                            As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AP' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAP_AccountCode = Null2String(rsChartAccount!ACCTCODE)
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

Sub BringToFrontGJ()

'picGJEntry.ZOrder 0
'picGJEntry.Visible = True
'picGJEntry.Enabled = True
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

Private Sub cboGJAccountNo_Click()
    cboGJAccountNo_Change
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
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE: TOTAL_AP_AMOUNT = 0: TotalAPAmountToPay = 0: PrevPV_Amount = 0
    Dim J_ITemNo, PV_ITEMNO                       As Integer
    txtGJTotDebit.Text = ZERO: txtGJTotCredit.Text = ZERO: txtGJOutBalance.Text = ZERO
    lstGJ.Sorted = False: lstGJ.ListItems.Clear
    Set rsJournal_Det = New ADODB.Recordset
    'UPDATED BY: JUN
    'DATE UPDATED: 06-08-2009
    'DESCRIPTION: CHANGES OF FILTERATION BY JNO TO VOUCHERNO
    Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
    'Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
    'UPDATED BY: JUN

    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        Screen.MousePointer = 11
        rsJournal_Det.MoveFirst
        Do While Not rsJournal_Det.EOF
            kcnt = kcnt + 1
            If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
            lstGJ.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
            lstGJ.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!Acct_code)
            lstGJ.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
            lstGJ.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
            lstGJ.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
            lstGJ.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
            TOTDEBIT = TOTDEBIT + NumericVal(N2Str2Zero(rsJournal_Det!DEBIT))
            TOTCREDIT = TOTCREDIT + NumericVal(N2Str2Zero(rsJournal_Det!CREDIT))
            TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJournal_Det!tax))
            rsJournal_Det.MoveNext
        Loop
        'lstGJ.Sorted = True:
        'lstGJ.Refresh
        OUTBALANCE = TOTDEBIT - TOTCREDIT
        txtGJTotDebit.Text = ToDoubleNumber(TOTDEBIT)
        txtGJTotCredit.Text = ToDoubleNumber(TOTCREDIT)
        txtGJOutBalance.Text = ToDoubleNumber(Abs(OUTBALANCE))
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

        Call DISPAY_INFO

        Screen.MousePointer = 0
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccount2                           As ADODB.Recordset
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
    Dim rsTemplate_Header                         As ADODB.Recordset
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
    Dim rsTemplate_Header                         As ADODB.Recordset
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

    Set rsPayTerm = New ADODB.Recordset
    Set rsPayTerm = gconDMIS.Execute("select pay_desc from ALL_PayTerm order by pay_desc asc")
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        Combo_Loadval cboPayType, rsPayTerm
    End If
    Set rsPayTerm = Nothing

    If xJOURNALTYPE = "GJ" Then
        Set rsATC = New ADODB.Recordset
        Set rsATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
        If Not rsATC.EOF And Not rsATC.BOF Then
            'Combo_Loadval cboATC, rsATC
            'rsATC.MoveFirst: cboATC2.Clear: cboATC2.AddItem ""
            'Do While Not rsATC.EOF
            '    cboATC2.AddItem Null2String(rsATC!ATC)
            '    rsATC.MoveNext
            'Loop
        End If
        Set rsATC = Nothing

        Set rsVENDOR = New ADODB.Recordset
        Set rsVENDOR = gconDMIS.Execute("select nameofvendor from ALL_Vendor order by nameofvendor asc")
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            'Combo_Loadval cboJVSupCust, rsVENDOR
        End If
        Set rsVENDOR = Nothing
    End If
    Dim rsAR_Accounts                             As ADODB.Recordset


    Set rsAR_Accounts = New ADODB.Recordset
    Set rsAR_Accounts = gconDMIS.Execute("Select Description from AMIS_ChartAccount where Titles in('1102' ,'1103','1102','1204','2102','2107')ORDER BY Description")
    If Not rsAR_Accounts.EOF And Not rsAR_Accounts.BOF Then
        rsAR_Accounts.MoveFirst: cboARTag.Clear
        Do While Not rsAR_Accounts.EOF
            cboARTag.AddItem Null2String(rsAR_Accounts!Description)
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

Sub InitGJ()
'txtGJItemNo.Text = Format(kcnt + 1, "0000")
'cboGJAccountNo.Text = ""
'txtGJAccountName.Text = ""
'txtGJDebit.Text = ZERO
'txtGJCredit.Text = ZERO
'txtGJAccountParticulars.Text = "Pls. Type Your Remarks Here..."
    txtSearch.Text = ""
    'cboATC2.ListIndex = 0
    'txtRATE2.Text = "0"
    'txtTaxBase2.Text = ZERO
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
    If xJOURNALTYPE = "APJ" Then
        cboATC.ListIndex = 0
        txtRATE.Text = "0"
        txtTaxBase.Text = ZERO
    End If
End Sub

Sub initMemvars()
    Dim rsJournal_HDDup                           As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
    txtJDate.Text = LOGDATE:

    CDJ_CIB = ""
    CDJ_AP = ""

    txtTotalPV_Amount.Text = ZERO
    labPosted.Caption = ""
    labPosted.Visible = False
    labOutBalance.Visible = False
    txtOutBalance.Visible = False
    '    InitGrid
    SendToBack
End Sub

'Sub InitPV_Detail()
'    txtPVItemNo.Text = Format(Jcnt + 1, "0000")
'    txtMRR_No.Text = ""
'    If xJOURNALTYPE = "APJ" Then
'        txtPO_No.Text = "": txtINV_No.Text = "": txtProd_No.Text = ""
'        txtPVAmount.Text = ZERO
'    ElseIf xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
'        txtPO_No.Text = txtVoucherNo.Text: txtINV_No.Text = ""
'        txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
'        txtPVAmount.Text = ZERO
'    Else
'        labPV1.Caption = "Voucher No": txtPO_No.Text = txtVoucherNo.Text: txtPO_No.Enabled = False
'        labPV2.Caption = "PV Voucher No.": labPV3.Caption = "Doc. Date": labPV4.Caption = "Due Date"
'        txtINV_No.Text = LOGDATE: txtINV_No.Format = "dd-mmm-yy"
'        txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
'        txtPVAmount.Text = ZERO
'        txtProd_No.Enabled = True: txtMRR_No.Enabled = True: txtINV_No.Enabled = True
'    End If
'End Sub

Sub InsertAccountEntries(XXX As Variant)
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET  As Double
    Dim J_STATUS, J_JITEMNO                       As String
    Dim rsTemplate_Details                        As ADODB.Recordset
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
            J_ACCT_NAME = N2Str2Null(rsTemplate_Details!Description)
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
        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
            lstDetails.SetFocus
        End If

        Screen.MousePointer = 0
    End If
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Or xJOURNALTYPE = "PDJ" Or xJOURNALTYPE = "CLO" Then
        'cboGJAccountNo.Text = labAccountCode.Caption
        'txtGJAccountName.Text = Setacctname(labAccountCode.Caption)
        'If cboGJAccountNo.Text <> "" Then
        '    If SetAcctType(cboGJAccountNo.Text) = "C" Then
        '        On Error Resume Next
        '        txtGJCredit.SetFocus
        '    Else
        '        On Error Resume Next
        '        txtGJDebit.SetFocus
        '    End If
        'End If
    Else
        cboAcct_Code.Text = labAccountCode.Caption
        'If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "SJ" Then
        '   txtGrossAmt.SetFocus
        If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
            On Error Resume Next
            txtGrossAmt.SetFocus
        Else
            If cboAcct_Code.Text <> "" Then
                If SetAcctType(cboAcct_Code.Text) = "C" Then
                    On Error Resume Next
                    txtCredit.SetFocus
                Else
                    On Error Resume Next
                    txtDebit.SetFocus
                End If
            End If
        End If
    End If
    cmdFindAccount.ZOrder 1
    fraFindAccount.ZOrder 1
End Sub

Sub OkAccountSetCursor()
'    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Or xJOURNALTYPE = "PDJ" Or xJOURNALTYPE = "CLO" Then
'        'If cboGJAccountNo.Text <> "" Then
'        '    If SetAcctType(cboGJAccountNo.Text) = "C" Then
'                On Error Resume Next
'                txtGJCredit.SetFocus
'            Else
'                On Error Resume Next
'                txtGJDebit.SetFocus
'            End If
'        End If
'    End If
End Sub

Sub rsRefresh()
    If xJOURNALTYPE = "GJ" Then Me.Caption = "GENERAL JOURNAL DATA ENTRY"
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by JDATE asc", gconDMIS, adOpenKeyset
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
    picShowPostRange.Visible = False
    cmdPrinting.ZOrder 1
    picPrinting.ZOrder 1
End Sub

Sub SendToBackGJ()


'picGJEntry.ZOrder 1
'picGJEntry.Visible = False
'picGJEntry.Enabled = False
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
        txtJDate.Text = Format(Null2String(rsJournal_HD!JDate), "DD-MMM-YY")
        txtInvoiceDate.Text = Format(Null2String(rsJournal_HD!invoicedate), "DD-MMM-YY")
        txtDueDate.Text = Format(Null2String(rsJournal_HD!duedate), "DD-MMM-YY")
        txtPayCode.Text = Null2String(rsJournal_HD!paytype)
        txtTerms.Text = Null2String(rsJournal_HD!TERMS)
        If SetPayDesc(Null2String(rsJournal_HD!paytype)) = "" Then
            cboPayType.ListIndex = -1
        Else
            cboPayType.Text = SetPayDesc(Null2String(rsJournal_HD!paytype))
        End If

        If xJOURNALTYPE = "GJ" Then txtParticulars2.Locked = True


        txtBankCode.Text = Null2String(rsJournal_HD!bankcode)
        txtCheckNo.Text = Null2String(rsJournal_HD!CheckNo)
        txtCheckDate.Text = Null2String(rsJournal_HD!CheckDate)
        txtParticulars.Text = Null2String(rsJournal_HD!remarks)
        txtParticulars2.Text = Null2String(rsJournal_HD!remarks)
        txtTotDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!DEBIT))
        txtTotCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!CREDIT))
        txtOutBalance.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!OUTBALANCE))
        txtAmountToPay.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!amounttopay))
        txtRemarks.Text = Null2String(rsJournal_HD!remarks)
        txtRemarks2.Text = Null2String(rsJournal_HD!remarks)
        If Null2String(rsJournal_HD!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "*** CANCELLED ***"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            labPosted.Visible = True
            labPosted.Caption = "*** POSTED ***"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPrint.Enabled = True
            If LOGLEVEL = "ADM" Then cmdUnPost.Enabled = True Else cmdUnPost.Enabled = False
        Else
            labPosted.Caption = ""
            labPosted.Visible = False
            cmdEdit.Enabled = True
            cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = True
            cmdPost.Enabled = True
            cmdPrint.Enabled = False
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
    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    'DEALER INCOME TAX WITHHELD
    If COMPANY_CODE = "HAI" Then
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
        GettheTaxBaseAmnt
    End If
    If COMPANY_CODE = "HMH" Then
        DEALER_ITW_COMPENSATION = "21-05002-00"
        DEALER_ITW_EXPANDED = "21-05003-00"
    End If
    If cboAcct_Code.Text = DEALER_ITW_COMPENSATION Or cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
        If xJOURNALTYPE = "CLO" Then
            Dim rsJournal_HDDet                   As ADODB.Recordset
            Set rsJournal_HDDet = New ADODB.Recordset
            rsJournal_HDDet.Open "select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from AMIS_vw_vLEDGER where Jdate <= '" & txtJDate.Text & "' and Acct_Code = '" & cboAcct_Code.Text & "'", gconDMIS
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                If N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) > 0 Then
                    'txtGJDebit.Text = ZERO
                    'txtGJCredit.Text = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                Else
                    'txtGJDebit.Text = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                    'txtGJCredit.Text = ZERO
                End If
            End If
            Set rsJournal_HDDet = Nothing
        End If
    End If
End Sub

Private Sub cboAcct_Code_Click()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cboATC_Click()
'Update By BTT: 09262008
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

Private Sub cboATC2_Click()
    Set rsATC = New ADODB.Recordset
    'Set rsATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC2.Text))
    If Not rsATC.EOF And Not rsATC.BOF Then
        'txtRATE2.Text = N2Str2Zero(rsATC!Rate)
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

Private Sub cboGJAccountNo_Change()
    Dim DEALER_ITW_COMPENSATION                   As String
    Dim DEALER_ITW_EXPANDED                       As String
    'txtGJAccountName.Text = Setacctname(cboGJAccountNo.Text)
    'DEALER INCOME TAX WITHHELD
    If COMPANY_CODE = "HAI" Then
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
    '        fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False
    '    End If
End Sub

Private Sub cboNameofVendor_Change()
    txtCode.Text = SetVendorCode(cboNameofVendor.Text)
    txtAddress.Caption = SetVendorAddress(txtCode.Text)
End Sub

Private Sub cboNameofVendor_Click()
    txtCode.Text = SetVendorCode(cboNameofVendor.Text)
    txtAddress.Caption = SetVendorAddress(txtCode.Text)
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

    INIT_LAB_LBL

    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    Dim rsProfile                                 As ADODB.Recordset
    Dim AccountingMonth, AccountingYear           As Integer
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


    txtParticulars2.Locked = False
    txtParticulars2.Text = ""
    AddorEdit = "ADD"

    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    SendToBack
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
'    If xJOURNALTYPE = "CDJ" Then
'        If DirectDisbursementVoucherNo <> txtVoucherNo.Text Then
'            If MsgBox("Add from Accounts Payable?", vbQuestion + vbYesNo, "Disbursement for Purchases") = vbYes Then
'                If MsgBox("Warning: This Disbursement will have the default entry of AP and Cash In Bank, Continue?", vbQuestion + vbYesNo) = vbYes Then
'                    SendToBackPV
'                    BringToFrontPV
'                    AddorEdit = "ADD"
'                    cmdPVDelete.Visible = False
'                    InitPV_Detail
'                    CDJ_IS_FROM_AP = True
'                    frmAMISSearchAPJ2.Show vbModal
'                    JournalTAB.Tab = 1
'                    'cmdPVSave_Click
'                Else
'                    GoTo CDJ_ISDirectDisbursement
'                End If
'            Else
'                GoTo CDJ_ISDirectDisbursement
'            End If
'        Else
'            SendToBack
'            cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
'            fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
'            fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
'            AddorEdit = "ADD"
'            InitJournal
'            On Error Resume Next
'            cboAcct_Code.SetFocus
'        End If
'    Else
    SendToBack
    cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
    fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
    fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
    AddorEdit = "ADD"
    InitJournal
    On Error Resume Next
    cboAcct_Code.SetFocus
    '    End If
    '    Exit Sub

CDJ_ISDirectDisbursement:
    CDJ_IS_FROM_AP = False
    DirectDisbursementVoucherNo = txtVoucherNo.Text
    SendToBack
    cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
    fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
    fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
    AddorEdit = "ADD"
    InitJournal
    On Error Resume Next
    cboAcct_Code.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDetails.Enabled = True
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
        If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
        End If
        If xJOURNALTYPE = "GJ" Then
            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
        End If
        Screen.MousePointer = 0
        ' UPDATE DUE TO NEW AUDIT : BTT 08282008
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        If xJOURNALTYPE = "APJ" Then
            'Update By BTT 06282008
            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
            SQL_STATEMENT = "update AMIS_PV_Detail set status = 'C' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If
        If xJOURNALTYPE = "CDJ" Then
            'Update By BTT 06282008

            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With

            Set rsCV_Detail = New ADODB.Recordset
            Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where Jtype = 'APJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If rsCV_Detail.EOF And rsCV_Detail.BOF Then
                Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where Jtype = 'VPJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            End If
            If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
                SQL_STATEMENT = "update AMIS_CV_Detail set status = 'C' where jtype = " & N2Str2Null(rsCV_Detail!jtype) & " and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
            End If
        End If
        If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
            'Update By BTT 06282008
            With FrmCancelTransaction
                .lblTransaction_type = xJOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail Where VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                SQL_STATEMENT = "update AMIS_CV_Detail set status = 'C' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
            End If

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

    AddorEdit = "EDIT"
    PrevJType = UCase(xJOURNALTYPE)
    PrevJNo = Format(txtJNo.Text, "000000")
    lstDetails.Enabled = False
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_HD!ID
    If xJOURNALTYPE = "GJ" Then txtParticulars2.Locked = False
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
    Call frmAMISSearchGJ.LoadJournalType("GJ")
    frmAMISSearchGJ.Show vbModal
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:
    INIT_LAB_LBL
    'FOR NAVIGATIONAL CONTROL-----
    Unload frmAMISJournalEntry_GJDetails
    '-----------------------------
    rsJournal_HD.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdGJCancel_Click()
    SendToBackGJ
    StoreMemVars
    Picture1.Enabled = True
End Sub

Private Sub cmdGJDelete_Click()

    If labGJID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If

    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Journal_Det where id = " & labGJID.Caption
    Else
        Exit Sub
    End If

    Dim cnt                                       As Integer
    Dim rsJournalDup                              As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemNo,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update AMIS_Journal_Det set JItemNo = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            rsJournalDup.MoveNext
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

    If lstGJ.ListItems.Count > 0 And lstGJ.Enabled = True Then
        lstGJ.SetFocus
    End If
End Sub

Private Sub cmdGJEntry_Click()
    SendToBackGJ
    'picGJEntry.Visible = True: picGJEntry.ZOrder 0
    'picGJEntry.Enabled = True: cmdGJDelete.Visible = False
    AddorEdit = "ADD"
    InitGJ
    'UPDATED BY: JUN
    'DATE UPDATED: 06/08/2009
    'DESCRIPTION: INITIALIZE GJ COMBO BOX
    ' Init_cboGJAccountNo
    'UPDATED BY: JUN
    On Error Resume Next
    'cboGJAccountNo.SetFocus
End Sub

'Private Sub cmdGJSave_Click()
'
'    On Error GoTo Errorcode
'    If cboGJAccountNo.Text = "" Then
'        MsgBox "Account Code must have a value", vbInformation, "Error Encountered!"
'        Exit Sub
'    End If
'
'    If NumericVal(txtGJDebit.Text) > 0 And NumericVal(txtGJCredit.Text) > 0 Then
'        MsgBox "Invalid Journal Entry! Debit and Credit Amount can not have both Amount!", vbCritical, "Invalid Entry!"
'        Exit Sub
'    End If
'
'    'If AddorEdit = "ADD" Then
'    '    Dim rsJournal_DetClone                         As ADODB.Recordset
'    '    Set rsJournal_DetClone = New ADODB.Recordset
'    '    rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(xJOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
'    '    If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
'    '        MsgBox "Account Code already used in this transaction", vbInformation, "Error in Part Number Validation"
'    '        Exit Sub
'    '    End If
'    'End If
'
'
'    If xJOURNALTYPE = "GJ" Then
'        If Left(cboGJAccountNo.Text, 5) = "11-02" Or Left(cboGJAccountNo.Text, 5) = "11-03" Then
'            If MsgBox("A/R Codes must have a DM/CM to update the Customer Subsidiary" & vbCrLf & " or use DM/CM for A/R Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/R schedule") = vbYes Then
'                'save in audit trail
'                MsgBox "Reminder: You must use DM/CM to update A/R Schedule", vbInformation, "Confirmation Logged in Audit Trail."
'            Else
'                Exit Sub
'            End If
'        End If
'        If Left(cboGJAccountNo.Text, 5) = "21-01" Then
'            If MsgBox("A/P Codes must have a DM/CM to update the Vendors Subsidiary" & vbCrLf & " or use DM/CM for A/P Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/P schedule") = vbYes Then
'                'save in audit trail
'                MsgBox "Reminder: You must use DM/CM to update A/P Schedule", vbInformation, "Confirmation Logged in Audit Trail."
'            Else
'                Exit Sub
'            End If
'        End If
'    End If
'
'    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                  As String
'    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                As String
'    Dim J_DEBIT, J_CREDIT, J_TAX                       As Double
'    Dim J_STATUS, J_JITEMNO                            As String
'
'    J_JDATE = N2Date2Null(txtJDate.Text)
'    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
'    J_JTYPE = N2Str2Null(xJOURNALTYPE)
'    J_JNO = N2Str2Null(txtJNo.Text)
'    J_JITEMNO = N2Str2Null(txtGJItemNo.Text)
'    J_ACCT_CODE = N2Str2Null(cboGJAccountNo.Text)
'    J_ACCT_NAME = N2Str2Null(txtGJAccountName.Text)
'    J_DEBIT = Round(NumericVal(txtGJDebit.Text), 2)
'    J_CREDIT = Round(NumericVal(txtGJCredit.Text), 2)
'    J_TAX = Round(NumericVal(txtTax.Text), 2)
'    J_STATUS = "'N'"
'
'    Dim J_SUPCODE, J_ATC                               As String
'    Dim J_RATE, J_TAXBASE                              As Double
'    If cboGJAccountNo.Text = "21-04001-00" Or cboGJAccountNo.Text = "21-04002-00" Then
'        J_SUPCODE = N2Str2Null(SetVendorCode(cboJVSupCust.Text))
'        J_ATC = N2Str2Null(cboATC2.Text)
'        J_RATE = NumericVal(txtRATE2.Text)
'        J_TAXBASE = NumericVal(txtTaxBase2.Text)
'    Else
'        J_SUPCODE = "'999999'"
'        J_ATC = "NULL"
'        J_RATE = 0
'        J_TAXBASE = 0
'    End If
'    Screen.MousePointer = 11
'    If AddorEdit = "ADD" Then
'        If txtGJAccountParticulars.Text <> "" And txtGJAccountParticulars.Text <> "Pls Type Your Remarks Here!" Then
'            gconDMIS.Execute "insert into AMIS_JV_Detail " & _
             '                             "(JNo,VoucherNo,itemno,Particulars,status)" & _
             '                           " values (" & J_JNO & ", " & J_VOUCHERNO & ", " & J_JITEMNO & _
             '                             ", " & N2Str2Null(txtGJAccountParticulars.Text) & _
             '                             ", " & J_STATUS & ")"
'        End If
'        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
         '                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE)" & _
         '                       " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
         '                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
         '                         ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
'    Else
'        gconDMIS.Execute "update AMIS_Journal_Det set" & _
         '                       " jdate = " & J_JDATE & "," & _
         '                       " voucherno = " & J_VOUCHERNO & "," & _
         '                       " jtype = " & J_JTYPE & "," & _
         '                       " jno = " & J_JNO & "," & _
         '                       " jitemno = " & J_JITEMNO & "," & _
         '                       " acct_code = " & J_ACCT_CODE & "," & _
         '                       " acct_name = " & J_ACCT_NAME & "," & _
         '                       " debit = " & J_DEBIT & "," & _
         '                       " credit = " & J_CREDIT & "," & _
         '                       " tax = " & J_TAX & "," & _
         '                       " ATC = " & J_ATC & "," & _
         '                       " RATE = " & J_RATE & "," & _
         '                       " TAXBASE = " & J_TAXBASE & "," & _
         '                       " status = " & J_STATUS & _
         '                       " where id = " & labGJID.Caption
'        gconDMIS.Execute "update AMIS_JV_Detail set" & _
         '                       " Particulars = " & N2Str2Null(txtGJAccountParticulars.Text) & _
         '                       " where JNo = " & J_JNO & " and ItemNo = " & J_JITEMNO
'    End If
'    FillDetails
'    gconDMIS.Execute "update AMIS_Journal_HD set" & _
     '                   " VENDORCODE = " & J_SUPCODE & "," & _
     '                   " debit = " & TOTDEBIT & "," & _
     '                   " credit = " & TOTCREDIT & "," & _
     '                   " tax = " & TOTTAX & "," & _
     '                   " outbalance = " & OUTBALANCE & _
     '                   " where id = " & labID.Caption
'    rsRefresh
'    Picture1.Enabled = True
'    On Error Resume Next
'    rsJournal_hd.Find "id = " & labID.Caption
'    StoreMemvars
'    If AddorEdit = "ADD" Then cmdGJEntry_Click Else cmdGJCancel_Click
'    Screen.MousePointer = 0
'    Exit Sub
'
'Errorcode:
'    Screen.MousePointer = 0
'    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
'    Exit Sub
'End Sub

Private Sub cmdJournalCancel_Click()
    SendToBack
    StoreMemVars
    Picture1.Enabled = True
    JournalTAB.TabEnabled(1) = True
End Sub

Private Sub cmdJournalDelete_Click()

    If labDetID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Journal_Det where id = " & labDetID.Caption
    End If
    Dim cnt                                       As Integer
    Dim rsJournalDup                              As ADODB.Recordset
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
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
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
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    ElseIf xJOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
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
    Dim ValidateAccount                           As New ADODB.Recordset
    On Error GoTo ErrorCode
    If cboAcct_Code.Text = "" Or Setacctname(cboAcct_Code.Text) = "" Then
        MsgBox "Account Code and Description must have a value", vbInformation, "Error Encountered!"
        Exit Sub
    End If


    'NOT TO ALLOW INPUT OF SAME ACCOUNT CODE
    '    If AddorEdit = "ADD" Then
    '        Dim rsJournal_DetClone                         As ADODB.Recordset
    '        Set rsJournal_DetClone = New ADODB.Recordset
    '        rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(xJOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
    '        If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
    '            MsgBox "Account Code already used in this transaction", vbInformation, "Error in Account Code Validation"
    '            Exit Sub
    '        End If
    '    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET  As Double
    Dim J_STATUS, J_JITEMNO                       As String
    Dim J_ATC                                     As String
    Dim J_RATE, J_TAXBASE                         As Double

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
    If AddorEdit = "ADD" Then
        Set ValidateAccount = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_JOURNAL_dET WHERE ACCT_CODE=" & J_ACCT_CODE & " AND VOUCHERNO =" & J_VOUCHERNO & " AND JTYPE=" & J_JTYPE & "")
        If ValidateAccount(0) = 1 Then
            MsgBox "Duplicate Account entry is not allowed..", vbInformation, "Please verify your entry!"
            Exit Sub
        End If
    End If

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,USERCODE,LASTUPDATE,ATC,RATE,TAXBASE)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                        ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & LOGCODE & "','" & LOGDATE & "'," & J_ATC & "," & J_RATE & "," & J_TAXBASE & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
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
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    FillDetails
    If xJOURNALTYPE = "VDJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    ElseIf xJOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
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
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:
    INIT_LAB_LBL
    'FOR NAVIGATIONAL CONTROL-----
    Unload frmAMISJournalEntry_GJDetails
    '-----------------------------

    rsJournal_HD.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    INIT_LAB_LBL

    'FOR NAVIGATIONAL CONTROL-----
    Unload frmAMISJournalEntry_GJDetails
    '-----------------------------
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

Private Sub cmdOkPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optPrintVoucher.Value = True Then
        Screen.MousePointer = 11
        picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
        ShowReport "CashDisbursement", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CASH DISBURSEMENT JOURNAL PRINTOUT", LOGDATE, False
        Screen.MousePointer = 0
    Else
        If optSECBANK.Value = True Then
            If MsgBox("Please Insert Security Bank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Security Bank Check
                ShowReport "SecurityBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
        If optPRUDBANK.Value = True Then
            If MsgBox("Please Insert Prudential Bank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Prudential Bank Check
                ShowReport "PrudentialBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
        If optCHINBANK.Value = True Then
            If MsgBox("Please Insert Chinabank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Chinabank Check
                ShowReport "ChinaBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
    End If
End Sub

Private Sub cmdPost_Click()
On Error GoTo ErrorCode:

Dim str_MSG                                   As String


    str_MSG = "Error Appear In During @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf
    
    gconDMIS.BeginTrans
    If JournalPosting = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "General Journal")
        MsgBox str_MSG, vbCritical, "Posting Error "
        cmdExit.Enabled = True
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    gconDMIS.CommitTrans
    Screen.MousePointer = 0

ErrorCode:
'    SaveLogFile
    ShowVBError
End Sub

Function JournalPosting() As Boolean
On Error GoTo ErrorCode

    Dim KimyDKid                                  As Integer

    'UPDATED BY: JUN
    'DATE UPDATED: 08122009 BABYFEMZ BIRTHDAY
    'DESCRIPTION: DO NOT ALLOW THE USER TO POST A TRANSACTION IF IT HAS NO DETAIL
    If lstGJ.ListItems.Count = 0 Then
        MessagePop InfoFriend, "INFORMATION", "You cannot POST this transaction it must have a detail"
        JournalPosting = True
        Exit Function
    End If

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
    If MsgBox("Are you sure you want to Post this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then
        If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "OPB" Then
            '        If COMPANY_CODE = "HPI" Then
            'Updated by: ACL 10202009
            If CheckIfOpen(xJOURNALTYPE, Trim(txtJDate.Text), Year(txtJDate.Text)) = False Then
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                JournalPosting = True
                Exit Function
            End If
            '        Else
            '            Set rsProfile = New ADODB.Recordset
            '            Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
            '            If Not rsProfile.EOF And Not rsProfile.BOF Then
            '                If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
            '                    If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
            '                        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                        exit function
            '                    End If
            '                Else
            '                    MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                    exit function
            '                End If
            '            End If
            '        End If
        End If
        '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
        '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
        '        exit function
        '    End If
        'Update by BTT : to update the balance of the SJ
        'If xJOURNALTYPE = "CRJ" Then
        '    UpdateBalanceSJ txtVoucherNo.Text, True
        'End If

        Dim rsCheckDetails                        As ADODB.Recordset
        Dim rsCheckCRJDetails                     As ADODB.Recordset
        Dim TotalCRJ_Credit                       As Double

        Screen.MousePointer = 11
        
        Screen.MousePointer = 0
 
        If NumericVal(txtTotDebit.Text) <> NumericVal(txtTotCredit.Text) Then
            MsgBox "Warning: Total Debit is not equal to Total Credit", vbCritical, "Cannot be Posted!"
            JournalPosting = True
            Exit Function
        End If

        'If COMPANY_CODE <> "HGC" Then
        If VALIDATE_POSTING = True Then
            JournalPosting = True
            Exit Function
        End If
        
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        'UPDATED BY: JUN
        'DATE UPDATED: 05-30-2009
        'DESCRIPTION: VALIDATE IF ALL ENTRY IN AMIS_JOURNAL_DET WAS TAG AS POSTED IF NOT UPDATE THE STATUS INTO POSTED
        Dim rsCHECK_POSTED                        As ADODB.Recordset
        Set rsCHECK_POSTED = gconDMIS.Execute("SELECT STATUS FROM AMIS_JOURNAL_DET where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " AND STATUS <> 'P'")
        If Not rsCHECK_POSTED.EOF And Not rsCHECK_POSTED.BOF Then
            gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        Else
            'ALL DETAILS ARE POSTED
        End If
        Set rsCHECK_POSTED = Nothing

        'UPDATED BY: JUN --- DATE UPDATED: 11/17/2009 --- DESCRITPION: INSERT SUM OF AR IN AMIS_AR
        'If COMPANY_CODE <> "HGC" Then
        Call GET_AR_GJ
        Call GET_GJ_PAYMENT

        Call GET_AP_GJ
        'Call GET_AP_GJ2
        Call GET_AP_GJ_PAYMENT

        'Call GJ_REMARKS_XXX
        'End If
        'UPDATED BY: JUN-------------------------


        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
    End If

    JournalPosting = True
    Exit Function

ErrorCode:
'    SaveLogFile
    JournalPosting = False
End Function


'Upating Code       : AXP-0713200713:18
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    INIT_LAB_LBL

    'FOR NAVIGATIONAL CONTROL-----
    Unload frmAMISJournalEntry_GJDetails
    '-----------------------------

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
    Dim Ans                                       As String
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Print", LocalAcess) = False Then Exit Sub

    Ans = MsgBox("Are you sure do you want to print this Transaction?", vbQuestion + vbYesNo, "Print Transaction")
    If Ans = vbYes Then

        'For Reprint Routin : Update by BTT
        If xJOURNALTYPE = "GJ" Then SaveReprintInformation xJOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub

        If xJOURNALTYPE = "GJ" Then ShowReport "GeneralJournal", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "GENERAL JOURNAL PRINTOUT", LOGDATE, False
        NEW_LogAudit "PX", "JOURNAL ENTRY", "", "", "", txtVoucherNo, xJOURNALTYPE, txtJNo
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
    '    InitPV_Detail
    On Error Resume Next
    '    If xJOURNALTYPE = "APJ" Then
    '        On Error Resume Next
    '        txtPO_No.SetFocus
    '    Else
    On Error Resume Next
    txtMRR_No.SetFocus
    '    End If
End Sub

Private Sub cmdPVCancel_Click()
    SendToBackPV
    StoreMemVars
    JournalTAB.TabEnabled(0) = True
    Picture1.Enabled = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup, rsProfile                      As ADODB.Recordset


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
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where invoiceno = '" & txtInvoiceNo.Text & "' and invoicedate = " & N2Date2Null(txtInvoiceDate2.Text) & " and invoicetype = '" & cboInvoiceType.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MessagePop RecSaveError, "Error!", "Invoice Transaction already Encoded!"
                Exit Sub
            End If
        End If
    End If
    If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "Error"
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
    '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If

    Dim J_JDATE                                   As String
    Dim J_VOUCHERNO, J_JTYPE                      As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_CUSTOMERNAME                            As String
    Dim J_DEBIT, J_CREDIT, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_CHECKNO                       As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE       As String
    Dim J_INVOICETYPE, J_INVOICENO                As String
    Dim J_CHECKDATE, J_BANKCODE                   As String
    Dim J_REFNO, J_REFDATE                        As String
    Dim J_TERMS, J_DEALER                         As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS             As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(Format(txtVoucherNo.Text, "000000"))
    J_JTYPE = N2Str2Null(xJOURNALTYPE)

    J_INVOICEDATE = "NULL"
    J_BALANCE = 0
    J_AMOUNTPAID = 0

    J_DUEDATE = N2Date2Null(txtDueDate.Text)
    J_PAYTYPE = N2Str2Null(txtPayCode.Text)

    J_JNO = N2Str2Null(txtJNo.Text)
    J_DEBIT = NumericVal(txtTotDebit.Text)
    J_CREDIT = NumericVal(txtTotCredit.Text)
    J_OUTBALANCE = NumericVal(txtOutBalance.Text)
    J_AMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
    J_STATUS = "'N'"

    J_CHECKNO = N2Str2Null(txtCheckNo.Text)
    J_TERMS = "NULL"
    J_DEALER = "NULL"
    J_CHECKDATE = "NULL"

    J_BANKCODE = N2Str2Null(txtBankCode.Text)

    J_CUSTOMERNAME = "NULL"
    J_VENDORCODE = "'999999'"
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Or xJOURNALTYPE = "PDJ" Or xJOURNALTYPE = "CLO" Then
        J_CUSTOMERCODE = "'999999'"
        J_CUSTOMERNAME = "NULL"
    Else
        If Trim(txtCustCode.Text) = "" Then
            MsgBox "Please Check the Customer Information!", vbInformation, "Innformation"
            Exit Sub
        End If
        J_CUSTOMERCODE = N2Str2Null(txtCustCode.Text)
        J_CUSTOMERNAME = N2Str2Null(cboCustName.Text)
    End If

    If xJOURNALTYPE <> "APJ" Then                          ' update by BTT
        J_INVOICETYPE = N2Str2Null(SetInvCode(cboInvoiceType.Text))
    End If

    If xJOURNALTYPE <> "APJ" Then                          ' update by BTT
        J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
    End If

    J_INVOICEAMT = NumericVal(txtInvoiceAmt.Text)
    J_REFNO = N2Str2Null(txtRefNo.Text)
    J_REFDATE = N2Date2Null(txtRefDate.Text)
    If Trim(txtParticulars2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars2.Text))

    J_PAIDSTATUS = "'N'"
    J_RECEIVESTATUS = "'N'"

    If AddorEdit = "ADD" Then
        Dim rsJournal_HDDup                       As ADODB.Recordset
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
        J_JNO = N2Str2Null(txtJNo.Text)
        J_VOUCHERNO = N2Str2Null(GetVoucherNo(xJOURNALTYPE))
        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                        " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "'," & N2Date2Null(LOGDATE) & ")"
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "A", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
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
        NEW_LogAudit "E", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        CheckIfthereISCDJ txtVoucherNo
        SQL_STATEMENT = "update AMIS_Journal_Det set" & _
                        " jtype = " & J_JTYPE & "," & _
                        " jdate = " & J_JDATE & "," & _
                        " USERCODE = '" & LOGCODE & "'," & _
                        " LASTUPDATE = '" & LOGDATE & "'," & _
                        " jno = " & J_JNO & _
                        " where jtype = '" & PrevJType & "' and jno = '" & PrevJNo & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    If AddorEdit <> "ADD" Then
        rsJournal_HD.Find "jno = " & J_JNO
        cmdCancel.Value = True
        FillDetails

        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                        " debit = " & TOTDEBIT & "," & _
                        " credit = " & TOTCREDIT & "," & _
                        " tax = " & TOTTAX & "," & _
                        " outbalance = " & OUTBALANCE & _
                        " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "JOURNAL ENTRY AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
    End If
    rsRefresh
    rsJournal_HD.Find "jno = " & J_JNO
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then
        If xJOURNALTYPE = "GJ" Then cmdGJEntry_Click Else cmdAddJournal_Click
    End If
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
            '                        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                         Exit Sub
            '                    End If
            '                Else
            '                    MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                    Exit Sub
            '                End If
            '            End If
            '        End If
        End If
        '    If CheckIfBookIsOpen(xJOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
        '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
        '        Exit Sub
        '    End If

        Screen.MousePointer = 11
        ' Update Due to new log Audit : BTT 282008
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT

        'UPDATED BY: JUN --- DATE UPDATED: 11/18/2009 --- DESCRIPTION: THIS UPDATE THE AMOUNT TO ZERO AND DELETE THE AR IN AMIS_AR AND PAYMENT DETAIL IN AMIS_DETAIL
        'If COMPANY_CODE <> "HGC" Then
        If VAL_OTH_GJ_UNPOST = True Then
            Exit Sub
        End If
        Call UNPOST_GJ
        Call UNPOST_AP_GJ
        'End If
        'UPDATED BY: JUN-------------------------------------------------------------------------------------------

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        Screen.MousePointer = 0
        LogAudit "U", "JOURNAL ENTRY", txtJNo
        NEW_LogAudit "U", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        Exit Sub
    End If
ErrorCode:
    ShowVBError
End Sub

Private Sub FillGrid()
    Dim rsChartAccount2                           As ADODB.Recordset
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
        '            If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
        '                fraFindAccount.Visible = True
        '                cmdFindAccount.Visible = True
        '                cmdFindAccount.ZOrder 0
        '                fraFindAccount.ZOrder 0
        '                fraFindAccount.Enabled = True
        '                DoEvents
        '                On Error Resume Next
        '                txtSearch.SetFocus
        '            ElseIf Me.ActiveControl.Name = "cboGJAccountNo" And cboGJAccountNo.Text = "" Then
        '                fraFindAccount.Visible = True
        '                cmdFindAccount.Visible = True
        '                cmdFindAccount.ZOrder 0
        '                fraFindAccount.ZOrder 0
        '                fraFindAccount.Enabled = True
        '                DoEvents
        '                txtSearch.SetFocus
        '            ElseIf Me.ActiveControl.Name = "cboAccount" Then
        '                OkAccount
        '            ElseIf Me.ActiveControl.Name = "txtPO_No" And txtPO_No.Text = "" Then
        '                On Error Resume Next
        '                txtPO_No.SetFocus
        '            ElseIf Me.ActiveControl.Name = "txtCredit" And SetAcctType(cboAcct_Code.Text) = "C" And Val(txtCredit.Text) <= 0 And Val(txtDebit.Text) <= 0 Then
        '                On Error Resume Next
        '                txtCredit.SetFocus
        '            ElseIf Me.ActiveControl.Name = "txtDebit" And SetAcctType(cboAcct_Code.Text) = "D" And Val(txtDebit.Text) <= 0 And Val(txtCredit.Text) <= 0 Then
        '                On Error Resume Next
        '                txtDebit.SetFocus
        '            ElseIf Me.ActiveControl.Name = "txtGJCredit" And SetAcctType(cboGJAccountNo.Text) = "C" And Val(txtGJCredit.Text) <= 0 And Val(txtGJDebit.Text) <= 0 Then
        '                On Error Resume Next
        '                txtGJCredit.SetFocus
        '            ElseIf Me.ActiveControl.Name = "txtGJDebit" And SetAcctType(cboGJAccountNo.Text) = "D" And Val(txtGJDebit.Text) <= 0 And Val(txtGJCredit.Text) <= 0 Then
        '                On Error Resume Next
        '                txtGJDebit.SetFocus
        '            ElseIf Me.ActiveControl.Name = "txtGrossAmt" And NumericVal(txtGrossAmt.Text) <= 0 Then
        '                On Error Resume Next
        '                txtGrossAmt.SetFocus
        '            Else
        '                MoveKeyPress KeyCode
        '            End If
    Case vbKeyEscape
        If fraFindAccount.Visible = True Then
            If Me.ActiveControl.Name = "txtSearch" Then
                SendToBack
                SendToBackPV
                SendToBackGJ
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
                    SendToBackGJ
                    SendToBackTemplates
                    StoreMemVars
                ElseIf Me.ActiveControl.Name = "lstTemplates" Then
                    On Error Resume Next

                    txtSearchTemplates.SetFocus
                Else
                    SendToBack
                    SendToBackPV
                    SendToBackGJ
                    SendToBackTemplates
                    StoreMemVars
                End If
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
                JournalTAB.Tab = 0
                If xJOURNALTYPE = "GJ" Then
                    With frmAMISJournalEntry_GJDetails
                        .txtCode.Text = ""
                        .txtName.Text = ""
                        .txtInvoiceNo.Text = ""
                        .txtInvoiceType.Text = ""
                        .cboGJAccountNo.Text = ""
                        .txtGJAccountName.Text = ""
                        .txtGJDebit.Text = ""
                        .txtGJCredit.Text = ""
                        .cboATC2.ListIndex = -1
                        .txtRATE2.Text = ""
                        .txtTaxBase2.Text = ""
                        .cmdGJDelete = False
                        .txtGJDebit = "0.00"
                        .txtGJCredit = "0.00"
                        .cmdGJDelete.Enabled = False
                        .labClass.Caption = ""
                        .txtJtype.Text = ""
                        .txtADJ_Remarks = ""
                        .cboCDJNo.ListIndex = -1
                    End With
                    Call frmAMISJournalEntry_GJDetails.xADDorEDIT("ADD")
                    InitAccountCode
                    FormExistsShow frmAMISJournalEntry_GJDetails
                    'KeepScreenVisible frmAMIS_GJ_ENTRY
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
                    JournalTAB.Tab = 1
                    cmdPV_Entry_Click
                    JournalTAB.TabEnabled(0) = False
                    Picture1.Enabled = False
                    txtMRR_No.BackColor = &HFFFFFF
                    txtINV_No.BackColor = &HFFFFFF
                End If
            End If
        Else
            ShowInvoiceApp SetInvCode(cboInvoiceType), txtInvoiceNo.Text
        End If
    Case vbKeyF5
        cmdPost.Value = True
    Case vbKeyF6
        cmdUnPost.Value = True
    Case vbKeyF7
        cmdCancelCO.Value = True
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
        SendToBackGJ
        SendToBackTemplates
        cmdShowPostRange.Visible = True: picShowPostRange.Visible = True
        picShowPostRange.Enabled = True
        cmdShowPostRange.ZOrder 0: picShowPostRange.ZOrder 0
        On Error Resume Next
        txtFromVNo.SetFocus
    Case vbKeyF12
        If Null2String(rsJournal_HD!Status) = "C" Then

            If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub

            If MsgBox("Are you sure you want to Un-Cancel this Transaction?", vbQuestion + vbYesNo, "Un-Cancel Journal") = vbYes Then
                Screen.MousePointer = 11
                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                rsRefresh
                rsJournal_HD.Find "id = " & labID.Caption
                StoreMemVars
                Screen.MousePointer = 0
            End If
        End If
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
            Call frmALL_AuditInquiry.DisplayHistory(labID, "JOURNAL ENTRY")
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
            If xJOURNALTYPE = "APJ" Then
                If MsgBox("Set this AP Transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AP") = vbYes Then
                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtAmountToPay) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='APJ'")
                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
                End If
            End If

            ' TEMPORARY to close AR for HAI
            If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CSJ" Then
                If MsgBox("Set this AR transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AR") = vbYes Then
                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtInvoiceAmt) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='SJ'")
                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
                End If
            End If
        End If
    End If
    Exit Sub

ErrorCode:
    ShowErrMsg
End Sub

Sub InitAccountCode()
    Dim rsInitAccountCode                         As ADODB.Recordset
    Set rsInitAccountCode = New ADODB.Recordset
    rsInitAccountCode.Open "Select AcctCode from Amis_ChartAccount order by AcctCode asc", gconDMIS, adOpenKeyset
    frmAMISJournalEntry_GJDetails.cboGJAccountNo.Clear
    If Not rsInitAccountCode.EOF And Not rsInitAccountCode.BOF Then
        Do While Not rsInitAccountCode.EOF
            frmAMISJournalEntry_GJDetails.cboGJAccountNo.AddItem Null2String(rsInitAccountCode!ACCTCODE)
            rsInitAccountCode.MoveNext
        Loop
    End If
    Set rsInitAccountCode = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBack: SendToBackPV: SendToBackGJ: SendToBackTemplates
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    picPayables.Top = 1200
    picDisbursement.Top = 1200
    picReceivable.Top = 420
    'Frame1.Top = 90
    'fraATC.Visible = False: fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False
    labCheckAmt.Visible = False: txtCheckAmt.Visible = False: txtParticulars.Height = 795

    If xJOURNALTYPE = "GJ" Then
        LocalAcess = "GENERAL JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Me.Caption = "GENERAL JOURNAL DATA ENTRY"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    End If
    InitCbo
    INIT_LABEL
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
    Unload frmAMISJournalEntry_GJDetails
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

            'MsgBox "Transactions are Already Cancelled" & vbCrLf & _
             '       "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"

            'MsgBox "Journals are Already Posted" & vbCrLf & _
             '       "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            If kcnt > 0 Then
                AddorEdit = "EDIT"
                cmdJournalDelete.Visible = True
                BringToFront

                StoreJournalEntry (lstDetails.SelectedItem.SubItems(5))
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
    End If
End Sub

Private Sub lstGJ_Click()
'UPDATED BY: JUN
'DATE UPDATED: 08112009
'DESCRIPTION: DISPLAY THE INFORMATION OF CUSTOMER OR VENDOR WHO HAS AN ADJUSTMENT
    Dim RSinfo                                    As ADODB.Recordset
    Set RSinfo = New ADODB.Recordset

    INIT_LAB_LBL

    If lstGJ.ListItems.Count = 0 Then Exit Sub

    RSinfo.Open "Select ADJ_REMARKS,ENTITY,INVOICETYPE,INVOICENO,ADJ_VOUCHERNO,ADJ_JTYPE FROM AMIS_JOURNAL_DET WHERE ID = " & lstGJ.SelectedItem.SubItems(5) & "", gconDMIS, adOpenKeyset
    If Not RSinfo.EOF And Not RSinfo.BOF Then
        If Left(Null2String(RSinfo!ENTITY), 1) = "V" Then
            lblName.Caption = "Ven. Name"
            labcode.Caption = Mid(Null2String(RSinfo!ENTITY), 2, Len(Null2String(RSinfo!ENTITY)))
            labName.Caption = " " & GET_VENDOR_NAME(Mid(Null2String(RSinfo!ENTITY), 2, Len(Null2String(RSinfo!ENTITY))))
        Else
            lblName.Caption = "Cust. Name"
            labcode.Caption = Mid(Null2String(RSinfo!ENTITY), 2, Len(Null2String(RSinfo!ENTITY)))
            labName.Caption = " " & GET_CUSTNAME(Mid(Null2String(RSinfo!ENTITY), 2, Len(Null2String(RSinfo!ENTITY))))
        End If

        '        If Null2String(RSinfo!ADJ_JTYPE) <> "SJ" Then
        '           lblinvoiceno.Caption = "Journal No."
        '           labInvoiceNo.Caption = Null2String(RSinfo!INVOICENO)
        '           labJournalType.Caption = Null2String(RSinfo!ADJ_JTYPE)
        '        ElseIf Null2String(RSinfo!ADJ_JTYPE) = "SJ" Then
        '           labInvoiceNo.Caption = Null2String(RSinfo!INVOICENO)
        '           labJournalType.Caption = Null2String(RSinfo!ADJ_JTYPE)
        '        ElseIf Null2String(RSinfo!ADJ_JTYPE) = "APJ" Then
        '           labInvoiceNo.Caption = Null2String(RSinfo!INVOICENO)
        '           labJournalType.Caption = Null2String(RSinfo!ADJ_JTYPE)
        '        End If

        labInvoiceNo.Caption = Null2String(RSinfo!ADJ_VOUCHERNO)
        labJournalType.Caption = Null2String(RSinfo!ADJ_JTYPE)
        lblINVOICE_DETAIL.Caption = Null2String(RSinfo!INVOICENO)
        lblINVOICETYPE_DETAIL.Caption = Null2String(RSinfo!InvoiceType)

        If COMPANY_CODE = "HMH" Then
            Dim RS                                As ADODB.Recordset
            Set RS = New ADODB.Recordset
            RS.Open "SELECT ADJ_REMARKS FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND JTYPE = 'GJ' AND ADJ_REMARKS IS NOT NULL", gconDMIS, adOpenKeyset
            If Not RS.EOF And Not RS.BOF Then
                txtGJ_Remarks.Text = Null2String(RS!ADJ_REMARKS)
            End If
            Set RS = Nothing
        Else
            txtGJ_Remarks.Text = Null2String(RSinfo!ADJ_REMARKS)
        End If

    End If
    Set RSinfo = Nothing
End Sub

Private Sub lstGJ_DblClick()
    If lstGJ.ListItems.Count = 0 Then Exit Sub

    If Null2String(rsJournal_HD!Status) = "C" Then
        'MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
        MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"
    ElseIf Null2String(rsJournal_HD!Status) = "P" Then
        'MsgBox "Journals are Already Posted" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
        MessagePop RecLocekd, "Posted Transaction", "Journals are Already Posted and cannot be Change"
    Else
        frmAMISJournalEntry_GJDetails.xADDorEDIT ("EDIT")
        labDET.Caption = lstGJ.SelectedItem.SubItems(5)
        xJOURNALTYPE = "GJ"
        Call PASS_INFO_ENRTY(labDET.Caption)



        'ADDorEDIT = "EDIT"
        'cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        'StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub lstGJ_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call lstGJ_Click
End Sub

Private Sub lstGJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstGJ_DblClick
End Sub

Private Sub lstGJ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        AddorEdit = "EDIT"
        'cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        'StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
        cmdGJDelete_Click
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
            MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
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
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub lstTemplates_KeyPress(KeyAscii As Integer)
    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    If KeyAscii = 13 Then InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub optPrintCheck_Click()
    If optPrintCheck.Value = True Then
        picPrintCheck.Enabled = True
    Else
        picPrintCheck.Enabled = False
    End If
End Sub

Private Sub optPrintVoucher_Click()
    picPrintCheck.Enabled = False
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
            If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Or xJOURNALTYPE = "PDJ" Or xJOURNALTYPE = "CLO" Then
                Picture1.Enabled = False
                cmdGJEntry_Click
            Else
                JournalTAB.TabEnabled(1) = False
                Picture1.Enabled = False
                cmdAddJournal_Click
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
    SendToBackGJ
    SendToBackTemplates
    cmdShowPostRange.Visible = True: picShowPostRange.Visible = True
    picShowPostRange.Enabled = True
    cmdShowPostRange.ZOrder 0: picShowPostRange.ZOrder 0
    On Error Resume Next
    txtFromVNo.SetFocus
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then labPosted.Visible = False Else labPosted.Visible = True
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
    If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VDJ" Or xJOURNALTYPE = "VCJ" Then cboBankName.Text = SetBankName(txtBankCode.Text)
    If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CSJ" Or xJOURNALTYPE = "CCM" Then cboBankName2.Text = SetBankName(txtBankCode.Text)
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

'Private Sub txtCode_Change()
'cboNameofVendor.Text = SetVendorName(txtCode.Text)
'txtAddress.Caption = SetVendorAddress(txtCode.Text)
'End Sub

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

'Private Sub txtcustcode_Change()
'cboCustName.Text = SetCustomerName(txtcustcode.Text)
'End Sub

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

Private Sub txtGJAccountName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtGJAccountParticulars_GotFocus()
'If txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!" Then txtGJAccountParticulars.Text = ""
End Sub

Private Sub txtGJAccountParticulars_LostFocus()
'If txtGJAccountParticulars.Text = "" Then txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!"
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
        '        If NumericVal(txtGJDebit.Text) = 0 Then
        '            If Val(txtGJCredit.Text) = 0 Then
        '                If OUTBALANCE > 0 And TOTDEBIT > 0 Then
        '                    txtGJCredit.Text = OUTBALANCE
        '                    txtGJDebit.Text = ZERO
        '                Else
        '                    txtGJCredit.Text = ""
        '                End If
        '            Else
        '                txtGJCredit.Text = NumericVal(txtGJCredit.Text)
        '            End If
        '        Else
        '            txtGJCredit.Text = ZERO
        '        End If
    End If
End Sub

Private Sub txtGJCredit_LostFocus()
'If txtGJCredit.Text = "" Then txtGJCredit.Text = 0
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
        '        If NumericVal(txtGJCredit.Text) = 0 Then
        '            If NumericVal(txtGJDebit.Text) = 0 Then
        '                If OUTBALANCE > 0 And TOTCREDIT > 0 Then
        '                    txtGJCredit.Text = ZERO: txtGJDebit.Text = OUTBALANCE
        '                Else
        '                    txtGJDebit.Text = ""
        '                End If
        '            Else
        '                txtGJDebit.Text = NumericVal(txtGJDebit.Text)
        '            End If
        '        Else
        '            txtGJDebit.Text = ZERO
        '        End If
    End If
End Sub

Private Sub txtGJDebit_LostFocus()
'If txtGJDebit.Text = "" Then txtGJDebit.Text = 0
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
    If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CSJ" Or xJOURNALTYPE = "CCM" Then
        cboCustName.SetFocus
    Else
        On Error Resume Next
        txtParticulars2.SetFocus
    End If
End Sub

Private Sub txtMRR_No_Change()
    Dim theJtype                                  As String
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset
    If xJOURNALTYPE = "CDJ" Then
        Set rsJournal_HD2 = New ADODB.Recordset
        'Update BTT : 09012008
        If txtPO_No.Text = "VPJ" Then
            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'VPJ'")
        Else
            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'APJ'and status='P'")
        End If
        If Not rsJournal_HD2.EOF And Not rsJournal_HD2.BOF Then
            theJtype = Null2String(rsJournal_HD2!jtype)
            txtINV_No.Text = Null2String(rsJournal_HD2!JDate)
            txtProd_No.Text = Null2String(rsJournal_HD2!duedate)
            txtPVAmount.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD2!BALANCE))
            If theJtype = "VPJ" Then
                Set RS = New ADODB.Recordset
                Set RS = gconDMIS.Execute("SELECT acct_code from AMIS_journal_det where voucherno='" & txtMRR_No.Text & "' and jtype ='VPJ'")
                If Not RS.EOF And Not RS.BOF Then
                    CDJ_AP = N2Str2Null(RS!Acct_code)
                    IsVPJ = True
                End If
            Else
                CDJ_AP = ReturnAP_AccountCode("AP")
                CDJ_IS_FROM_AP = True
                IsVPJ = False
            End If
        Else
            txtINV_No.Text = ""
            txtProd_No.Text = ""
            txtPVAmount.Text = ZERO
            CDJ_IS_FROM_AP = False
        End If
    End If
End Sub

Private Sub txtMRR_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If xJOURNALTYPE = "CDJ" Or xJOURNALTYPE = "VCJ" Then
        If KeyAscii = 13 Then
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchAPJ2.Show vbModal
        End If
    End If
    If xJOURNALTYPE = "CRJ" Or xJOURNALTYPE = "CCM" Then
        If KeyAscii = 13 Then
            SEARCH_TAB = 0
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchSJ2.Show vbModal
        End If
    End If
End Sub

Private Sub txtMRR_No_LostFocus()
    If xJOURNALTYPE = "CRJ" And AddorEdit <> "EDIT" Then
        Dim rsAR_Accounts                         As New ADODB.Recordset
        Set rsAR_Accounts = New ADODB.Recordset
        Set rsAR_Accounts = gconDMIS.Execute("select Acct_Code from AMIS_Journal_Det Where (Left(Acct_Code,5) = '11-02' or Left(Acct_Code,5) = '11-03') and  VoucherNo = '" & txtVoucherNo.Text & "' AND Jtype = '" & xJOURNALTYPE & "'")
        If Not rsAR_Accounts.EOF And Not rsAR_Accounts.BOF Then
            cboARTag.Text = Setacctname(rsAR_Accounts!Acct_code)
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

Private Sub txtParticulars2_GotFocus()
    If txtParticulars2.Text = "Pls Type Your Message Here!" Then txtParticulars2.Text = ""
End Sub

Private Sub txtParticulars2_LostFocus()
    If txtParticulars2.Text = "" Then txtParticulars2.Text = "Pls Type Your Message Here!"
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

Private Sub txtTaxBase2_Change()
' Update By BTT : 09262008
'    If NumericVal(txtRATE2.Text) > 0 Then
'        txtGJCredit.Text = Round(NumericVal(txtTaxBase2.Text) * (NumericVal(txtRATE2.Text) / 100), 2)
'    End If

End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Format(txtVoucherNo, "000000")
End Sub

Sub GettheTaxBaseAmnt()
    Dim SQL                                       As String
    Dim RS                                        As New ADODB.Recordset

    If xJOURNALTYPE = "APJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '11-07002-00' and jtype = 'APJ'"
    Else
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '11-07002-00' and jtype = 'CDJ'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTaxBase.Text = N2Str2IntZero(RS!SumDebit)
    End If
    Set RS = Nothing
End Sub

Function rsCHECKINVOICENOandTYPE(xInvoiceType As String, xINVOICENO As String, XCustomerCode As String) As Boolean
'UPDATED BY: JUN
'DATE UPDATED: 10272008
'DESCRIPTION: VALIDATE INVOICENO AND INVOICETYPE
    Dim rsExist                                   As ADODB.Recordset
    Set rsExist = gconDMIS.Execute("Select * from AMIS_Journal_hd where INVOICETYPE = '" & xInvoiceType & "' AND INVOICENO = '" & xINVOICENO & "' AND JTYPE = 'SJ' and CUSTOMERCODE = '" & XCustomerCode & "'")
    If Not rsExist.EOF And Not rsExist.BOF Then
        rsCHECKINVOICENOandTYPE = True                     ' yaon
    Else
        rsCHECKINVOICENOandTYPE = False                    ' mayo
    End If
    Set rsExist = Nothing
End Function

Function GetSJVoucherNo(ByVal xINVOICENO As String, ByVal xInvoiceType As String) As Boolean
'Update BTT : 10282008
'To check if the transaction is posted
    Dim RsSJVoucher                               As New ADODB.Recordset
    Set RsSJVoucher = gconDMIS.Execute("Select Voucherno,invoicetype,invoiceno,Status from Amis_journal_hd where invoiceno=" & xINVOICENO & " and invoicetype=" & xInvoiceType & " and jtype ='SJ'")
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
    Dim RSACCT                                    As New ADODB.Recordset
    Set RSACCT = gconDMIS.Execute("SELECT * from AMIS_chartaccount where acctcode='" & XXX & "'")
    With RSACCT
        If Not .EOF And Not .BOF Then
            cboARTag.AddItem Null2String(RSACCT!Description)
        End If
    End With
    Set RSACCT = Nothing
End Function
Sub CheckIfthereISCDJ(XXX As String)
    Dim RSCDJ                                     As New ADODB.Recordset
    Set RSCDJ = gconDMIS.Execute("SELECT amount FROM AMIS_CV_DETAIL where Pv_voucherno='" & XXX & "'")
    If Not RSCDJ.EOF And Not RSCDJ.BOF Then
        gconDMIS.Execute "UPDATE AMIS_journal_hd set balance = " & TOTALPVAMOUNT - TotalAPAmountToPay & "  where voucherno='" & XXX & "' and jtype='APJ'"
    End If
    Set RSCDJ = Nothing
End Sub

Function IsAR(XXX As String) As Boolean
'LAST UPDATED BY: BERNARD IMPLEMENTED IN PAMPANGA
'UPDATED BY: JUN
'DATE UPDATE: 06/06/2009
'DESCRIPTION: CHECK IF THE ACCOUNT CODE IS SET AS AR ACCOUNT SCHEDULE

    Dim RS                                        As New ADODB.Recordset
    Dim Det                                       As New ADODB.Recordset
    Dim xACCOUNT_COUNT                            As Integer

    xACCOUNT_COUNT = 0
    IsAR = False
    Set Det = gconDMIS.Execute("select voucherno,acct_code from AMIS_journal_det where voucherno='" & XXX & "' and jtype = 'CRJ'")
    If Not (Det.EOF And Det.BOF) Then
        Do While Not Det.EOF
            Set RS = gconDMIS.Execute("select IS_SCHEDULE_ACCNT from AMIS_CHARTACCOUNT where ACCTCODE='" & Det!Acct_code & "' AND IS_SCHEDULE_ACCNT = 1")
            If Not (RS.EOF And RS.BOF) Then
                xACCOUNT_COUNT = xACCOUNT_COUNT + 1
            Else
                'ISAR = False
            End If
            Det.MoveNext
        Loop

        If xACCOUNT_COUNT >= 1 Then
            IsAR = True
            If NumericVal(TOTAL_AR_AMOUNT) <> NumericVal(TOTALPVAMOUNT) Then
                picChat.Visible = True
            Else
                picChat.Visible = False
            End If
        Else
            IsAR = False
        End If
    End If
    Set Det = Nothing
    Set RS = Nothing
End Function

Sub PASS_INFO_ENRTY(xID As Variant)
    Dim rsPASS_INFO_ENRTY                         As ADODB.Recordset
    Set rsPASS_INFO_ENRTY = New ADODB.Recordset
    rsPASS_INFO_ENRTY.Open "Select * from Amis_Journal_det where ID = '" & xID & "'", gconDMIS, adOpenKeyset
    If Not rsPASS_INFO_ENRTY.EOF And Not rsPASS_INFO_ENRTY.BOF Then
        With frmAMISJournalEntry_GJDetails
            .txtCode = Mid(Null2String(rsPASS_INFO_ENRTY!ENTITY), 2, Len(Null2String(rsPASS_INFO_ENRTY!ENTITY)))
            .txtName = GET_NAME(Null2String(rsPASS_INFO_ENRTY!ENTITY))

            If Null2String(RTrim(LTrim(rsPASS_INFO_ENRTY!ADJ_JTYPE))) = "OTH" And rsPASS_INFO_ENRTY!IS_OTHERS = True Then
                .chkOther.Value = 1
                .chkOther.Enabled = True
                .txtInvoiceNo.Enabled = False
                .txtInvoiceType.Enabled = False
                .cboCDJNo.Enabled = False
                .cboJTYPE.Enabled = False
                .txtOTH_NO.Text = Null2String(rsPASS_INFO_ENRTY!ADJ_VOUCHERNO)
                '                ElseIf Null2String(rsPASS_INFO_ENRTY!InvoiceType) <> "" Then
                '                    .txtInvoiceNo = Null2String(rsPASS_INFO_ENRTY!INVOICENO)
                '                    .txtInvoiceType = Null2String(rsPASS_INFO_ENRTY!InvoiceType)
                '                    .cboCDJNo.Enabled = False
                '                    .cboJTYPE.Enabled = False
                'ElseIf (Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE) = "CDJ" Or Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE) = "APJ" Or Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE) = "CRJ" Or Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE) = "GJ") And Null2String(RTrim(LTrim(rsPASS_INFO_ENRTY!INVOICENO))) <> "OTHERS" Then
            ElseIf rsPASS_INFO_ENRTY!IS_OTHERS = False Then    'And Null2String(rsPASS_INFO_ENRTY!INVOICENO) <> "" And Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE) <> "" Then
                .txtInvoiceNo.Enabled = False
                .txtInvoiceType.Enabled = False
                .chkOther.Enabled = False
                .cboCDJNo.Text = Null2String(rsPASS_INFO_ENRTY!ADJ_VOUCHERNO)
                .cboJTYPE.Text = Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE)
                .txtINVOICE_DETAIL.Text = Null2String(rsPASS_INFO_ENRTY!INVOICENO)
                .txtINVOICE_TYPE.Text = Null2String(rsPASS_INFO_ENRTY!InvoiceType)
            ElseIf rsPASS_INFO_ENRTY!IS_OTHERS = True And Null2String(RTrim(LTrim(rsPASS_INFO_ENRTY!ADJ_JTYPE))) = "" Then
                .chkOther.Value = 1
                .chkOther.Enabled = True
                .cboCDJNo.Enabled = False
                .cboJTYPE.Enabled = False
            End If



            .cboGJAccountNo = Null2String(rsPASS_INFO_ENRTY!Acct_code)
            .txtGJAccountName = Null2String(rsPASS_INFO_ENRTY!acct_Name)
            .txtGJDebit = ToDoubleNumber(rsPASS_INFO_ENRTY!DEBIT)
            .txtGJCredit = ToDoubleNumber(rsPASS_INFO_ENRTY!CREDIT)
            .labClass.Caption = Null2String(rsPASS_INFO_ENRTY!ENTITY)
            .txtJtype.Text = Null2String(rsPASS_INFO_ENRTY!ADJ_JTYPE)
            .txtADJ_Remarks = Null2String(rsPASS_INFO_ENRTY!ADJ_REMARKS)
            .cboATC2 = Null2String(rsPASS_INFO_ENRTY!ATC)
            .txtRATE2 = NumericVal(rsPASS_INFO_ENRTY!Rate)
            .txtTaxBase2 = NumericVal(rsPASS_INFO_ENRTY!taxbase)
            .txtJItemNo = Null2String(rsPASS_INFO_ENRTY!jitemno)
            '                If Left(RTrim(LTrim(rsPASS_INFO_ENRTY!ENTITY)), 1) = "C" Then
            '                    .cmdCustomer.Enabled = True
            '                Else
            '                    .cmdCustomer.Enabled = False
            '                End If
            '
            '                If Left(RTrim(LTrim(rsPASS_INFO_ENRTY!ENTITY)), 1) = "V" Then
            '                    .cmdVendor.Enabled = True
            '                Else
            '                    .cmdVendor.Enabled = False
            '                End If

            If Null2String(RTrim(LTrim(rsPASS_INFO_ENRTY!ADJ_JTYPE))) = "APJ" Then
                .lblCode.Caption = "Ven. Code"
                .lblName.Caption = "Ven. Name"
                .lblInvoiceNo.Caption = "MRR NO."
                .lblInvoiceType.Caption = "MRR TYPE"
            ElseIf Null2String(RTrim(LTrim(rsPASS_INFO_ENRTY!ADJ_JTYPE))) = "SJ" Then
                .lblCode.Caption = "Cust. Code"
                .lblName.Caption = "Cust. Name"
                .lblInvoiceNo.Caption = "INV. NO."
                .lblInvoiceType.Caption = "INV. TYPE"
            ElseIf Null2String(LTrim(RTrim(rsPASS_INFO_ENRTY!ADJ_JTYPE))) = "CDJ" Then
                .lblCode.Caption = "Ven. Code"
                .lblName.Caption = "Ven. Name"
                .lblInvoiceNo.Caption = "INV. NO."
                .lblInvoiceType.Caption = "INV. TYPE"
            End If
        End With
    End If
    Set rsPASS_INFO_ENRTY = Nothing
End Sub

Function GET_NAME(xCUSCODE As String) As String
    Dim rsGet_Name                                As ADODB.Recordset
    Set rsGet_Name = New ADODB.Recordset
    rsGet_Name.Open "Select AccountName from All_Entity  where Code = '" & Mid(xCUSCODE, 2, Len(xCUSCODE)) & "' and EntityCode = '" & Left(xCUSCODE, 1) & "' ", gconDMIS, adOpenKeyset
    If Not rsGet_Name.EOF And Not rsGet_Name.BOF Then
        GET_NAME = Null2String(rsGet_Name!AccountName)
    End If
    Set rsGet_Name = Nothing
End Function

Sub INIT_LABEL()
    labcode.Caption = ""
    labName.Caption = ""
    labInvoiceNo.Caption = ""
    'labInvoiceType.Caption = ""
    labJournalType.Caption = ""
End Sub

Sub INIT_LAB_LBL()
    labcode.Caption = ""
    labName.Caption = ""
    labInvoiceNo.Caption = ""
    'labInvoiceType.Caption = ""
    labJournalType.Caption = ""
    txtGJ_Remarks.Text = ""
    lblINVOICE_DETAIL.Caption = ""
    lblINVOICETYPE_DETAIL.Caption = ""
End Sub

Function GET_VENDOR_NAME(xVENDORCODE As String) As String
'DESCRIPTION: GET THE VENDOR NAME
    Dim rsGET_VENDOR_NAME                         As ADODB.Recordset
    Set rsGET_VENDOR_NAME = New ADODB.Recordset
    rsGET_VENDOR_NAME.Open "Select NAMEOFVENDOR from ALL_VENDOR_TABLE WHERE CODE = '" & xVENDORCODE & "' AND CODE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsGET_VENDOR_NAME.EOF And Not rsGET_VENDOR_NAME.BOF Then
        GET_VENDOR_NAME = Null2String(rsGET_VENDOR_NAME!nameofvendor)
    Else
        GET_VENDOR_NAME = ""
    End If
    Set rsGET_VENDOR_NAME = Nothing
End Function

Function GET_CUSTNAME(xCUSTCODE As String) As String
'DESCRIPTION: GET THE CUSTOMER NAME
    Dim rsGET_CUSTNAME                            As ADODB.Recordset
    Set rsGET_CUSTNAME = New ADODB.Recordset
    rsGET_CUSTNAME.Open "Select ACCTNAME from ALL_CUSTOMER_TABLE WHERe CUSCDE = '" & xCUSTCODE & "' AND CUSCDE IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsGET_CUSTNAME.EOF And Not rsGET_CUSTNAME.BOF Then
        GET_CUSTNAME = Null2String(rsGET_CUSTNAME!AcctName)
    Else
        GET_CUSTNAME = ""
    End If
    Set rsGET_CUSTNAME = Nothing
End Function

Sub DISPAY_INFO()
'UPDATED BY: JUN
'DATE UPDATED: 08112009
'DESCRIPTION: DISPLAY THE INFORMATION OF CUSTOMER OR VENDOR WHO HAS AN ADJUSTMENT
    Dim RSinfo                                    As ADODB.Recordset
    Set RSinfo = New ADODB.Recordset

    INIT_LAB_LBL

    If lstGJ.ListItems.Count = 0 Then Exit Sub

    RSinfo.Open "Select ADJ_REMARKS,ADJ_VOUCHERNO,ENTITY,INVOICETYPE,INVOICENO,ADJ_JTYPE FROM AMIS_JOURNAL_DET WHERE ID = '" & lstGJ.ListItems(1).SubItems(5) & "'", gconDMIS, adOpenKeyset
    If Not RSinfo.EOF And Not RSinfo.BOF Then
        If Left(Null2String(RSinfo!ENTITY), 1) = "V" Then
            lblName.Caption = "Ven. Name"
            labcode.Caption = Right(Null2String(RSinfo!ENTITY), 6)
            labName.Caption = " " & GET_VENDOR_NAME(Right(Null2String(RSinfo!ENTITY), 6))
        Else
            lblName.Caption = "Cust. Name"
            labcode.Caption = Right(Null2String(RSinfo!ENTITY), 6)
            labName.Caption = " " & GET_CUSTNAME(Right(Null2String(RSinfo!ENTITY), 6))
        End If

        labInvoiceNo.Caption = Null2String(RSinfo!ADJ_VOUCHERNO)
        labJournalType.Caption = Null2String(RSinfo!ADJ_JTYPE)
        lblINVOICE_DETAIL.Caption = Null2String(RSinfo!INVOICENO)
        lblINVOICETYPE_DETAIL.Caption = Null2String(RSinfo!InvoiceType)

        txtGJ_Remarks.Text = Null2String(RSinfo!ADJ_REMARKS)
    End If
    Set RSinfo = Nothing
End Sub

Sub GET_AR_GJ()
    Dim rsGET_AR_GJ                               As ADODB.Recordset
    Dim rsCOUNT_CODE                              As ADODB.Recordset
    Dim rsCOUNT_NAME                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xJType                                    As String
    Dim XCustomerCode                             As String
    Dim xCUST_NAME                                As String
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xInvoicedate                              As String
    Dim xAMOUNT_TO_PAY                            As Double
    Dim xAMOUNT_PAID                              As Double
    Dim xACCT_CODE                                As String
    Dim xLAST_UPDATED                             As String
    Dim xBAL                                      As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsGET_AR_GJ = New ADODB.Recordset
    rsGET_AR_GJ.Open "SELECT VOUCHERNO,JDATE,JTYPE,ENTITY,INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,IS_OTHERS,ACCT_CODE " & _
                     "FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('11-02','11-03') AND VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " and DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_AR_GJ.EOF And Not rsGET_AR_GJ.BOF Then
        Do While Not rsGET_AR_GJ.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGET_AR_GJ!jtype) & "-" & Null2String(rsGET_AR_GJ!VOUCHERNO))
            xJdate = N2Str2Null(Null2String(rsGET_AR_GJ!JDate))
            xJType = N2Str2Null(Null2String(rsGET_AR_GJ!jtype))
            XCustomerCode = N2Str2Null(Right(Null2String(rsGET_AR_GJ!ENTITY), 6))

            If Left(Null2String(rsGET_AR_GJ!ENTITY), 1) = "V" Then
                xCUST_NAME = GET_VENDOR_NAME(Right(Null2String(rsGET_AR_GJ!ENTITY), 6))
            Else
                xCUST_NAME = GET_CUSTNAME(Right(Null2String(rsGET_AR_GJ!ENTITY), 6))
            End If

            If IsNull(rsGET_AR_GJ!INVOICENO) = True And IsNull(rsGET_AR_GJ!InvoiceType) = True And rsGET_AR_GJ!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_AR_GJ!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_AR_GJ!ADJ_JTYPE))
            ElseIf IsNull(rsGET_AR_GJ!INVOICENO) = False And IsNull(rsGET_AR_GJ!InvoiceType) = False And rsGET_AR_GJ!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_AR_GJ!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_AR_GJ!InvoiceType))
            ElseIf IsNull(rsGET_AR_GJ!INVOICENO) = True And IsNull(rsGET_AR_GJ!InvoiceType) = True And rsGET_AR_GJ!IS_OTHERS = True Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_AR_GJ!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_AR_GJ!ADJ_JTYPE))
            End If

            xInvoicedate = N2Str2Null(rsGET_AR_GJ!JDate)
            xAMOUNT_TO_PAY = GET_SUM_GJ_AR(Null2String(rsGET_AR_GJ!INVOICENO), Null2String(rsGET_AR_GJ!InvoiceType), rsGET_AR_GJ!IS_OTHERS, Null2String(rsGET_AR_GJ!ADJ_VOUCHERNO), Null2String(rsGET_AR_GJ!ADJ_JTYPE), Null2String(rsGET_AR_GJ!Acct_code), XCustomerCode, Null2String(rsGET_AR_GJ!VOUCHERNO))
            xAMOUNT_PAID = 0
            xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
            xACCT_CODE = N2Str2Null(Null2String(rsGET_AR_GJ!Acct_code))
            xLAST_UPDATED = N2Date2Null(LOGDATE)


            Dim rsCheck                           As ADODB.Recordset
            Set rsCheck = New ADODB.Recordset
            rsCheck.Open "SELECT * FROM AMIS_AR WHERE SJVOUCHERNO = " & xVOUCHERNO & " AND INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            If Not rsCheck.EOF And Not rsCheck.BOF Then
            Else
                gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xINVOICENO & "," & XCustomerCode & ",'" & xCUST_NAME & "'," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & ", " & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ")"
            End If
            Set rsCheck = Nothing
            rsGET_AR_GJ.MoveNext
        Loop
    End If
    Set rsGET_AR_GJ = Nothing
End Sub

Function GET_SUM_GJ_AR(xINVOICENO As String, xInvoiceType As String, xIS_OTHERS As Boolean, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xACCT_CODE As String, xCUST_CODE As String, xVOUCHERNO As String) As Double
    Dim rsGET_SUM_GJ_AR                           As ADODB.Recordset
    Set rsGET_SUM_GJ_AR = New ADODB.Recordset

    If xINVOICENO = "" And xInvoiceType = "" And xIS_OTHERS = False Then
        rsGET_SUM_GJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_AR FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & xADJ_VOUCHERNO & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND IS_OTHERS = 0 and RIGHT(ENTITY,6) = " & xCUST_CODE & " and DEBIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "' AND VOUCHERNO='" & xVOUCHERNO & "'", gconDMIS, adOpenKeyset
    ElseIf xINVOICENO <> "" And xInvoiceType <> "" And xIS_OTHERS = False Then
        rsGET_SUM_GJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_AR FROM AMIS_JOURNAL_DET WHERE INVOICENO  = " & xINVOICENO & " AND INVOICETYPE = " & N2Str2Null(xInvoiceType) & " AND IS_OTHERS = 0 and RIGHT(ENTITY,6) = " & xCUST_CODE & " AND DEBIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'  AND VOUCHERNO='" & xVOUCHERNO & "'", gconDMIS, adOpenKeyset
    ElseIf xInvoiceType = "" And xInvoiceType = "" And xIS_OTHERS = True Then
        rsGET_SUM_GJ_AR.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_AR FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO  = " & N2Str2Null(xADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND IS_OTHERS = 1 and RIGHT(ENTITY,6) = " & xCUST_CODE & " AND  DEBIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'  AND VOUCHERNO='" & xVOUCHERNO & "'", gconDMIS, adOpenKeyset
    Else
        GET_SUM_GJ_AR = 0
        Exit Function
    End If
    If Not rsGET_SUM_GJ_AR.EOF And Not rsGET_SUM_GJ_AR.BOF Then
        GET_SUM_GJ_AR = NumericVal(rsGET_SUM_GJ_AR!SUM_GJ_AR)
    Else
        GET_SUM_GJ_AR = 0
    End If
    Set rsGET_SUM_GJ_AR = Nothing
End Function

Sub GET_GJ_PAYMENT()
    Dim rsGET_GJ_PAYMENT                          As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim XCustomerCode                             As String
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xACCT_CODE                                As String
    Dim xINVOICE_AMT                              As Double
    Dim xJType                                    As String
    Dim xSJVOUCHERNO_DETAIL                       As String

    Set rsGET_GJ_PAYMENT = New ADODB.Recordset
    rsGET_GJ_PAYMENT.Open "SELECT VOUCHERNO,CREDIT,JDATE,JTYPE,ENTITY,INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,IS_OTHERS,ACCT_CODE " & _
                          "FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('11-02','11-03') AND VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " and CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_PAYMENT.EOF And Not rsGET_GJ_PAYMENT.BOF Then
        Do While Not rsGET_GJ_PAYMENT.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            xJdate = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!JDate))
            XCustomerCode = N2Str2Null(Right(Null2String(rsGET_GJ_PAYMENT!ENTITY), 6))

            If IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "COB" Then
                    xINVOICENO = Get_COB_InvoiceNo(N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO)))
                    xInvoiceType = Get_COB_InvoiceType(N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO)))
                Else
                    If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "SJ" Then
                        xINVOICENO = N2Str2Null(Get_SJ_INVOICENO(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO), Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE)))
                        xInvoiceType = N2Str2Null(Get_SJ_INVOICETYPE(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO), Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE)))
                    Else
                        xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                        xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE))
                    End If
                End If

                If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "OTH" Then
                    xSJVOUCHERNO_DETAIL = N2Str2Null("")
                Else
                    xSJVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                End If
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = False And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = False And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!InvoiceType))
                xSJVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype) & "-" & Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            ElseIf IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = True Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE))
                xSJVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype) & "-" & Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            End If

            xINVOICE_AMT = NumericVal(rsGET_GJ_PAYMENT!CREDIT)
            xACCT_CODE = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!Acct_code))
            xJType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype))

            gconDMIS.Execute "INSERT INTO AMIS_DETAIL(INVOICENO,INVOICETYPE,INVOICEAMOUNT,CUSTOMERCODE,ACCT_CODE,JDATE,VOUCHERNO,JTYPE,SJVOUCHERNO,INVOICEDATE) " & _
                             "VALUES(" & xINVOICENO & "," & xInvoiceType & "," & xINVOICE_AMT & "," & XCustomerCode & "," & xACCT_CODE & "," & xJdate & "," & xVOUCHERNO & "," & xJType & "," & xSJVOUCHERNO_DETAIL & "," & xJdate & ")"
            
            Dim rsAR As ADODB.Recordset
            Set rsAR = New ADODB.Recordset
            rsAR.Open "SELECT * FROM AMIS_AR WHERE SJVOUCHERNO=" & xSJVOUCHERNO_DETAIL & " AND INVOICENO IS NULL", gconDMIS, adOpenForwardOnly
            If Not rsAR.EOF And Not rsAR.BOF Then
                gconDMIS.Execute ("Update AMIS_AR SET INVOICENO=" & xINVOICENO & ",INVOICETYPE = " & xInvoiceType & " WHERE SJVOUCHERNO= " & xSJVOUCHERNO_DETAIL & " ")
            End If
            Dim rsGET_AR_SUM                      As ADODB.Recordset
            Dim xSUM_AR                           As Double
            Dim xSJVOUCHERNO                      As String
            If IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                xSJVOUCHERNO = Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO)
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = False And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = False And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                xSJVOUCHERNO = Null2String(rsGET_GJ_PAYMENT!InvoiceType) & "-" & Null2String(rsGET_GJ_PAYMENT!INVOICENO)
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = True Then
                xSJVOUCHERNO = Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO)
            End If
            Set rsGET_AR_SUM = New ADODB.Recordset
            If IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "OTH" Then
                    rsGET_AR_SUM.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE INVOICENO = " & N2Str2Null(xINVOICENO) & " AND INVOICETYPE = " & N2Str2Null(xInvoiceType) & " and  CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
                Else
                    rsGET_AR_SUM.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & "  and  CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
                End If
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = False And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = False And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                rsGET_AR_SUM.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = True Then
                rsGET_AR_SUM.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & " and CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            End If

            If Not rsGET_AR_SUM.EOF And Not rsGET_AR_SUM.BOF Then
                xSUM_AR = NumericVal(rsGET_AR_SUM!AMOUNT_TOPAY)
            Else
                xSUM_AR = 0
            End If
            Set rsGET_AR_SUM = Nothing

            Dim rsGET_GJ_SUM_PAYMENT              As ADODB.Recordset
            Dim xSUM_PAYMENT                      As Double
            Dim AR_BALANCE                        As Double

            Set rsGET_GJ_SUM_PAYMENT = New ADODB.Recordset
            rsGET_GJ_SUM_PAYMENT.Open "SELECT ROUND(SUM(INVOICEAMOUNT),2) AS GJ_PAYMENT FROM AMIS_DETAIL WHERE INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            If Not rsGET_GJ_SUM_PAYMENT.EOF And Not rsGET_GJ_SUM_PAYMENT.BOF Then
                xSUM_PAYMENT = NumericVal(rsGET_GJ_SUM_PAYMENT!GJ_PAYMENT)
            Else
                xSUM_PAYMENT = 0
            End If
            Set rsGET_GJ_SUM_PAYMENT = Nothing

            AR_BALANCE = Round((xSUM_AR - xSUM_PAYMENT), 2)

'            If IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
'                If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "OTH" Then
'                    gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & AR_BALANCE & " WHERE INVOICENO = " & N2Str2Null(xINVOICENO) & " AND INVOICETYPE = " & N2Str2Null(xInvoiceType) & " and  CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
'                Else
'                    gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & AR_BALANCE & " WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & "  and  CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
'                End If
'
'            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = False And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = False And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
'                gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & AR_BALANCE & " WHERE INVOICENO = " & xINVOICENO & " AND INVOICETYPE = " & xInvoiceType & " AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
'            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = True Then
'                gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = " & xSUM_PAYMENT & ", BALANCE = " & AR_BALANCE & " WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & " and CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE = " & xACCT_CODE & ""
'            End If

            rsGET_GJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_GJ_PAYMENT = Nothing
End Sub

Function VAL_OTH_GJ_UNPOST() As Boolean
    Dim rsOTH                                     As ADODB.Recordset
    Dim rsdetail                                  As ADODB.Recordset
    Set rsOTH = New ADODB.Recordset
    rsOTH.Open "SELECT * FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND DEBIT <> 0 AND ADJ_JTYPE = 'OTH'", gconDMIS, adOpenKeyset
    If Not rsOTH.EOF And Not rsOTH.BOF Then
        Do While Not rsOTH.EOF
            Set rsdetail = New ADODB.Recordset
            rsdetail.Open "SELECT * FROM AMIS_DETAIL WHERE INVOICENO = " & N2Str2Null(rsOTH!ADJ_VOUCHERNO) & " AND INVOICETYPE = " & N2Str2Null(rsOTH!ADJ_JTYPE) & " AND CUSTOMERCODE = " & N2Str2Null(Right((rsOTH!ENTITY), 6)) & " AND ACCT_CODE = " & N2Str2Null(rsOTH!Acct_code) & "", gconDMIS, adOpenKeyset
            If Not rsdetail.EOF And Not rsdetail.BOF Then
                VAL_OTH_GJ_UNPOST = True
                MessagePop InfoFriend, "INFORMATION", "You can't unpost this transaction. It has a payment Please see voucher no: " & "" & Null2String(rsdetail!jtype) & "" & "-" & "" & Null2String(rsdetail!VOUCHERNO) & ""
                Screen.MousePointer = 0
            Else
                VAL_OTH_GJ_UNPOST = False
            End If
            Set rsdetail = Nothing
            rsOTH.MoveNext
        Loop
    End If
    Set rsOTH = Nothing
End Function

Sub UNPOST_GJ()
    Dim rsUNPOST_GJ                               As ADODB.Recordset
    Dim rsAMOUN_TOPAY                             As ADODB.Recordset
    Dim xSJVOUCHERNO                              As String
    Dim xVOUCHERNO                                As String
    Dim xBAL                                      As Double

    xBAL = 0



    xVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
    Set rsUNPOST_GJ = New ADODB.Recordset
    rsUNPOST_GJ.Open "SELECT INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,IS_OTHERS,ENTITY,ACCT_CODE " & _
                     "FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND LEFT(ACCT_CODE,5) IN('11-02','11-02') AND CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsUNPOST_GJ.EOF And Not rsUNPOST_GJ.BOF Then
        Do While Not rsUNPOST_GJ.EOF
            If IsNull(rsUNPOST_GJ!INVOICENO) = True And IsNull(rsUNPOST_GJ!InvoiceType) = True And rsUNPOST_GJ!IS_OTHERS = False Then
                xSJVOUCHERNO = Null2String(rsUNPOST_GJ!ADJ_JTYPE) & "-" & Null2String(rsUNPOST_GJ!ADJ_VOUCHERNO)
                If Null2String(rsUNPOST_GJ!ADJ_JTYPE) = "OTH" Then
                    Set rsAMOUN_TOPAY = New ADODB.Recordset
                    rsAMOUN_TOPAY.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE INVOICENO = " & N2Str2Null(rsUNPOST_GJ!ADJ_VOUCHERNO) & "  AND INVOICETYPE = " & N2Str2Null(rsUNPOST_GJ!ADJ_JTYPE) & " and  CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & "", gconDMIS, adOpenKeyset
                Else
                    Set rsAMOUN_TOPAY = New ADODB.Recordset
                    rsAMOUN_TOPAY.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & "  and  CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & "", gconDMIS, adOpenKeyset
                End If
                If Not rsAMOUN_TOPAY.EOF And Not rsAMOUN_TOPAY.BOF Then
                    xBAL = NumericVal(rsAMOUN_TOPAY!AMOUNT_TOPAY)
                End If

                If Null2String(rsUNPOST_GJ!ADJ_JTYPE) = "OTH" Then
                    gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = 0 , BALANCE = " & xBAL & " WHERE INVOICENO = " & N2Str2Null(rsUNPOST_GJ!ADJ_VOUCHERNO) & "  AND INVOICETYPE = " & N2Str2Null(rsUNPOST_GJ!ADJ_JTYPE) & " and  CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & ""
                Else
                    gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = 0 , BALANCE = " & xBAL & " WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & "  and  CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & ""
                End If

            ElseIf IsNull(rsUNPOST_GJ!INVOICENO) = False And IsNull(rsUNPOST_GJ!InvoiceType) = False And rsUNPOST_GJ!IS_OTHERS = False Then
                Set rsAMOUN_TOPAY = New ADODB.Recordset
                rsAMOUN_TOPAY.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE INVOICENO = " & N2Str2Null(rsUNPOST_GJ!INVOICENO) & " AND INVOICETYPE = " & N2Str2Null(rsUNPOST_GJ!InvoiceType) & " AND CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & "", gconDMIS, adOpenKeyset
                If Not rsAMOUN_TOPAY.EOF And Not rsAMOUN_TOPAY.BOF Then
                    xBAL = NumericVal(rsAMOUN_TOPAY!AMOUNT_TOPAY)
                End If

                'gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = 0 , BALANCE = " & xBAL & " WHERE INVOICENO = " & N2Str2Null(rsUNPOST_GJ!INVOICENO) & " AND INVOICETYPE = " & N2Str2Null(rsUNPOST_GJ!InvoiceType) & " AND CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & ""
            ElseIf IsNull(rsUNPOST_GJ!INVOICENO) = True And IsNull(rsUNPOST_GJ!InvoiceType) = True And rsUNPOST_GJ!IS_OTHERS = True Then
                xSJVOUCHERNO = Null2String(rsUNPOST_GJ!ADJ_JTYPE) & "-" & Null2String(rsUNPOST_GJ!ADJ_VOUCHERNO)
                Set rsAMOUN_TOPAY = New ADODB.Recordset
                rsAMOUN_TOPAY.Open "SELECT AMOUNT_TOPAY FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & " and CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & "", gconDMIS, adOpenKeyset
                If Not rsAMOUN_TOPAY.EOF And Not rsAMOUN_TOPAY.BOF Then
                    xBAL = NumericVal(rsAMOUN_TOPAY!AMOUNT_TOPAY)
                End If
                'gconDMIS.Execute "UPDATE AMIS_AR SET AMOUNT_PAID = 0 , BALANCE = " & xBAL & " WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & " and CUSTOMERCODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCOUNT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & ""
            End If
            rsUNPOST_GJ.MoveNext
        Loop
    End If

    gconDMIS.Execute "DELETE FROM AMIS_DETAIL WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & ""
    gconDMIS.Execute "DELETE FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xVOUCHERNO) & ""

    Set rsUNPOST_GJ = Nothing
    Set rsAMOUN_TOPAY = Nothing
End Sub

Function VALIDATE_POSTING() As Boolean
    Dim rsVAL_POSTING                             As ADODB.Recordset
    Dim rsPOSTED                                  As ADODB.Recordset
    Dim rsGET_OTH                                 As ADODB.Recordset


    Set rsVAL_POSTING = New ADODB.Recordset
    rsVAL_POSTING.Open "SELECT DEBIT,ADJ_VOUCHERNO,ADJ_JTYPE,ENTITY FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN('11-02','11-03') " & _
                       "AND JTYPE = 'GJ' AND ADJ_VOUCHERNO IS NOT NULL AND ADJ_JTYPE IS NOT NULL AND VOUCHERNO = '" & txtVoucherNo.Text & "'", gconDMIS, adOpenKeyset
    If Not rsVAL_POSTING.EOF And Not rsVAL_POSTING.BOF Then
        Do While Not rsVAL_POSTING.EOF
            Set rsPOSTED = New ADODB.Recordset
            If Null2String(rsVAL_POSTING!ADJ_JTYPE) = "OTH" Then
                rsPOSTED.Open "SELECT * FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & N2Str2Null(rsVAL_POSTING!ADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(rsVAL_POSTING!ADJ_JTYPE) & " AND STATUS = 'P' AND LEFT(ACCT_CODE,5) IN('11-02','11-03') AND JTYPE = 'GJ' AND DEBIT <> 0", gconDMIS, adOpenKeyset
            Else
                'rsPOSTED.Open "SELECT * FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & rsVAL_POSTING!ADJ_VOUCHERNO & " AND JTYPE = " & N2Str2Null(rsVAL_POSTING!ADJ_JTYPE) & " AND STATUS = 'P' AND LEFT(ACCT_CODE,5) IN('11-02','11-03')", gconDMIS, adOpenKeyset
                rsPOSTED.Open "SELECT * FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & rsVAL_POSTING!ADJ_VOUCHERNO & " AND JTYPE = " & N2Str2Null(rsVAL_POSTING!ADJ_JTYPE) & " AND STATUS = 'P'", gconDMIS, adOpenKeyset
            End If

            If rsPOSTED.EOF And rsPOSTED.BOF Then
                If Null2String(rsVAL_POSTING!ADJ_JTYPE) = "OTH" Then
                    Set rsGET_OTH = New ADODB.Recordset
                    rsGET_OTH.Open "SELECT VOUCHERNO,JTYPE FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & N2Str2Null(rsVAL_POSTING!ADJ_VOUCHERNO) & " AND ADJ_JTYPE = " & N2Str2Null(rsVAL_POSTING!ADJ_JTYPE) & " AND DEBIT <> 0 AND ENTITY = " & N2Str2Null(rsVAL_POSTING!ENTITY) & "", gconDMIS, adOpenKeyset
                    If Not rsGET_OTH.EOF And Not rsGET_OTH.BOF Then
                        If NumericVal(rsVAL_POSTING!DEBIT) <> 0 Then
                        Else
                            MessagePop InfoFriend, "INFORMATION", "You can't post this transaction.The Vouher you are adjusting is not posted. Please see " & " " & Null2String(rsGET_OTH!jtype) & " " & "-" & "" & Null2String(rsGET_OTH!VOUCHERNO) & ""
                            VALIDATE_POSTING = True
                            Exit Function
                        End If

                    End If
                    Set rsGET_OTH = Nothing
                Else
                    MessagePop InfoFriend, "INFORMATION", "You can't post this transaction.The Vouher you are adjusting is not posted. Please see " & " " & Null2String(rsVAL_POSTING!ADJ_JTYPE) & " " & "-" & "" & Null2String(rsVAL_POSTING!ADJ_VOUCHERNO) & ""
                    VALIDATE_POSTING = True
                    Exit Function
                End If
            Else
                VALIDATE_POSTING = False
            End If
            Set rsPOSTED = Nothing
            rsVAL_POSTING.MoveNext
        Loop
    End If
    Set rsVAL_POSTING = Nothing
End Function

Function CheckIfOpen(xJType As String, xAcctMonth, xAcctYear) As Boolean
    Dim rsCheckOpen                               As ADODB.Recordset
    Set rsCheckOpen = New ADODB.Recordset
    rsCheckOpen.Open "Select * from AMIS_AccountingPeriod where JType = '" & xJType & "' and Month(AcctMonth) = '" & Format(xAcctMonth, "m") & "' and Year(AcctMonth) = '" & Format(xAcctMonth, "yyyy") & "' and Status=0 and CurrPeriod = 1", gconDMIS, adOpenForwardOnly
    If Not rsCheckOpen.EOF And Not rsCheckOpen.BOF Then
        CheckIfOpen = True
    Else
        CheckIfOpen = False
    End If
    Set rsCheckOpen = Nothing
End Function

Function Get_COB_InvoiceNo(XXX As String) As String
    Dim rsGetInvoice                              As ADODB.Recordset
    Set rsGetInvoice = New ADODB.Recordset
    rsGetInvoice.Open "SELECT InvoiceNo from AMIS_JOURNAL_HD where VoucherNo = " & XXX & " and JType='COB' and Status='P'", gconDMIS, adOpenKeyset
    If Not rsGetInvoice.EOF And Not rsGetInvoice.BOF Then
        Get_COB_InvoiceNo = N2Str2Null(Null2String(rsGetInvoice!INVOICENO))
    End If
End Function

Function Get_COB_InvoiceType(XXX As String) As String
    Dim rsGetInvoice                              As ADODB.Recordset
    Set rsGetInvoice = New ADODB.Recordset
    rsGetInvoice.Open "SELECT InvoiceType from AMIS_JOURNAL_HD where VoucherNo = " & XXX & " and JType='COB' and Status='P'", gconDMIS, adOpenKeyset
    If Not rsGetInvoice.EOF And Not rsGetInvoice.BOF Then
        Get_COB_InvoiceType = N2Str2Null(Null2String(rsGetInvoice!InvoiceType))
    End If
End Function

Sub GET_AP_GJ()
    Dim rsGJ_VOUCHER                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xDUEDATE                                  As String
    Dim xJType                                    As String
    Dim XCustomerCode                             As String
    Dim xCUST_NAME                                As String
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xInvoicedate                              As String
    Dim xAMOUNT_TO_PAY                            As Double
    Dim xAMOUNT_PAID                              As Double
    Dim xACCT_CODE                                As String
    Dim xLAST_UPDATED                             As String
    Dim xBAL                                      As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsGJ_VOUCHER = New ADODB.Recordset
    rsGJ_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,DET.ENTITY,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,ACCT_CODE,IS_OTHERS,ADJ_VOUCHERNO,ADJ_JTYPE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE INNER JOIN AMIS_CHARTACCOUNT AC ON DET.ACCT_CODE=AC.ACCTCODE " & _
                      "WHERE IS_SCHEDULE_ACCNT=1  AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P' AND DET.CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGJ_VOUCHER.EOF And Not rsGJ_VOUCHER.BOF Then
        Do While Not rsGJ_VOUCHER.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGJ_VOUCHER!jtype) & "-" & Null2String(rsGJ_VOUCHER!VOUCHERNO))
            xJdate = N2Str2Null(Null2String(rsGJ_VOUCHER!JDate))
            xJType = N2Str2Null(Null2String(rsGJ_VOUCHER!jtype))
            xDUEDATE = N2Str2Null(Null2String(rsGJ_VOUCHER!duedate))
            XCustomerCode = N2Str2Null(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6))

            If Left(Null2String(rsGJ_VOUCHER!ENTITY), 1) = "V" Then
                xCUST_NAME = N2Str2Null(GET_VENDOR_NAME(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6)))
            Else
                xCUST_NAME = N2Str2Null(GET_CUSTNAME(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6)))
            End If

'            If IsNull(rsGJ_VOUCHER!INVOICENO) = True And IsNull(rsGJ_VOUCHER!InvoiceType) = True And rsGJ_VOUCHER!IS_OTHERS = False Then
'                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_VOUCHERNO))
'                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_JTYPE))
'            ElseIf IsNull(rsGJ_VOUCHER!INVOICENO) = False And IsNull(rsGJ_VOUCHER!InvoiceType) = False And rsGJ_VOUCHER!IS_OTHERS = False Then
'                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!INVOICENO))
'                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!InvoiceType))
'            ElseIf IsNull(rsGJ_VOUCHER!INVOICENO) = True And IsNull(rsGJ_VOUCHER!InvoiceType) = True And rsGJ_VOUCHER!IS_OTHERS = True Then
                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_JTYPE))
'            Else
'                xINVOICENO = N2Str2Null("")
'                xInvoiceType = N2Str2Null("")
'            End If
            xInvoicedate = N2Str2Null(Null2String(rsGJ_VOUCHER!invoicedate))
            xAMOUNT_TO_PAY = GET_SUM_GJ_AP(Null2String(rsGJ_VOUCHER!VOUCHERNO), Null2String(rsGJ_VOUCHER!jtype), Null2String(rsGJ_VOUCHER!ADJ_VOUCHERNO), Null2String(rsGJ_VOUCHER!ADJ_JTYPE), Null2String(rsGJ_VOUCHER!Acct_code), XCustomerCode)
            xAMOUNT_PAID = 0
            xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
            xACCT_CODE = N2Str2Null(Null2String(rsGJ_VOUCHER!Acct_code))
            xLAST_UPDATED = N2Str2Null(LOGDATE)

            SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                            "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xINVOICENO & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
            gconDMIS.Execute SQL_STATEMENT
            If xJOURNALTYPE = "CDJ" Then
                gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid=" & xAMOUNT_PAID & ",Balance = " & xBAL & " where JTYPE ='" & xJOURNALTYPE & "' And VOUCHERNO = " & N2Str2Null(rsGJ_VOUCHER!VOUCHERNO)
            End If
            rsGJ_VOUCHER.MoveNext
        Loop
    End If
    Set rsGJ_VOUCHER = Nothing
End Sub

Sub GET_AP_GJ2()
    Dim rsGJ_VOUCHER                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xDUEDATE                                  As String
    Dim xJType                                    As String
    Dim XCustomerCode                             As String
    Dim xCUST_NAME                                As String
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xInvoicedate                              As String
    Dim xAMOUNT_TO_PAY                            As Double
    Dim xAMOUNT_PAID                              As Double
    Dim xACCT_CODE                                As String
    Dim xLAST_UPDATED                             As String
    Dim xBAL                                      As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsGJ_VOUCHER = New ADODB.Recordset
    rsGJ_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,DET.ENTITY,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,ACCT_CODE,IS_OTHERS " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P' AND DET.DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGJ_VOUCHER.EOF And Not rsGJ_VOUCHER.BOF Then
        Do While Not rsGJ_VOUCHER.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGJ_VOUCHER!jtype) & "-" & Null2String(rsGJ_VOUCHER!VOUCHERNO))
            xJdate = N2Str2Null(Null2String(rsGJ_VOUCHER!JDate))
            xJType = N2Str2Null(Null2String(rsGJ_VOUCHER!jtype))
            xDUEDATE = N2Str2Null(Null2String(rsGJ_VOUCHER!duedate))
            XCustomerCode = N2Str2Null(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6))

            If Left(Null2String(rsGJ_VOUCHER!ENTITY), 1) = "V" Then
                xCUST_NAME = N2Str2Null(GET_VENDOR_NAME(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6)))
            Else
                xCUST_NAME = N2Str2Null(GET_CUSTNAME(Right(Null2String(rsGJ_VOUCHER!ENTITY), 6)))
            End If

            If IsNull(rsGJ_VOUCHER!INVOICENO) = True And IsNull(rsGJ_VOUCHER!InvoiceType) = True And rsGJ_VOUCHER!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_JTYPE))
            ElseIf IsNull(rsGJ_VOUCHER!INVOICENO) = False And IsNull(rsGJ_VOUCHER!InvoiceType) = False And rsGJ_VOUCHER!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!InvoiceType))
            ElseIf IsNull(rsGJ_VOUCHER!INVOICENO) = True And IsNull(rsGJ_VOUCHER!InvoiceType) = True And rsGJ_VOUCHER!IS_OTHERS = True Then
                xINVOICENO = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGJ_VOUCHER!ADJ_JTYPE))
            Else
                xINVOICENO = N2Str2Null("")
                xInvoiceType = N2Str2Null("")
            End If
            xInvoicedate = N2Str2Null(Null2String(rsGJ_VOUCHER!invoicedate))
            xAMOUNT_TO_PAY = 0
            xAMOUNT_PAID = GET_SUM_GJ_AP2(Null2String(rsGJ_VOUCHER!IS_OTHERS), Null2String(rsGJ_VOUCHER!VOUCHERNO), Null2String(rsGJ_VOUCHER!jtype), Null2String(rsGJ_VOUCHER!Acct_code), XCustomerCode)
            xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
            xACCT_CODE = N2Str2Null(Null2String(rsGJ_VOUCHER!Acct_code))
            xLAST_UPDATED = N2Str2Null(LOGDATE)

            SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                            "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xINVOICENO & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
            gconDMIS.Execute SQL_STATEMENT
            If xJOURNALTYPE = "CDJ" Then
                gconDMIS.Execute "Update AMIS_JOURNAL_HD Set AmountPaid=" & xAMOUNT_PAID & ",Balance = " & xBAL & " where JTYPE ='" & xJOURNALTYPE & "' And VOUCHERNO = " & N2Str2Null(rsGJ_VOUCHER!VOUCHERNO)
            End If
            rsGJ_VOUCHER.MoveNext
        Loop
    End If
    Set rsGJ_VOUCHER = Nothing
End Sub

Function GET_SUM_GJ_AP(xVOUCHERNO, xJType, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xACCT_CODE As String, xCUST_CODE As String) As Double
    Dim rsGET_SUM_GJ_AP                           As ADODB.Recordset
    Set rsGET_SUM_GJ_AP = New ADODB.Recordset

    'If xIS_OTHERS = False Then
        rsGET_SUM_GJ_AP.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_GJ_AP FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & xVOUCHERNO & " AND JTYPE = " & N2Str2Null(xJType) & " AND ADJ_VOUCHERNO = " & xADJ_VOUCHERNO & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " and RIGHT(ENTITY,6) = " & xCUST_CODE & " and CREDIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    'ElseIf xIS_OTHERS = True Then
    '    rsGET_SUM_GJ_AP.Open "SELECT ROUND(SUM(CREDIT),2) AS SUM_GJ_AP FROM AMIS_JOURNAL_DET WHERE VOUCHERNO  = " & xADJ_VOUCHERNO & " AND JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND IS_OTHERS = 1 and RIGHT(ENTITY,6) = " & xCUST_CODE & " AND  CREDIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
   ' Else
    
    If Not rsGET_SUM_GJ_AP.EOF And Not rsGET_SUM_GJ_AP.BOF Then
        GET_SUM_GJ_AP = NumericVal(rsGET_SUM_GJ_AP!SUM_GJ_AP)
    Else
        GET_SUM_GJ_AP = 0
    End If
    Set rsGET_SUM_GJ_AP = Nothing
End Function

Function GET_SUM_GJ_AP2(xIS_OTHERS As Boolean, xADJ_VOUCHERNO As String, xADJ_JTYPE As String, xACCT_CODE As String, xCUST_CODE As String) As Double
    Dim rsGET_SUM_GJ_AP                           As ADODB.Recordset
    Set rsGET_SUM_GJ_AP = New ADODB.Recordset

    If xIS_OTHERS = False Then
        rsGET_SUM_GJ_AP.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_AP FROM AMIS_JOURNAL_DET WHERE ADJ_VOUCHERNO = " & xADJ_VOUCHERNO & " AND ADJ_JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND IS_OTHERS = 0 and RIGHT(ENTITY,6) = " & xCUST_CODE & " and CREDIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    ElseIf xIS_OTHERS = True Then
        rsGET_SUM_GJ_AP.Open "SELECT ROUND(SUM(DEBIT),2) AS SUM_GJ_AP FROM AMIS_JOURNAL_DET WHERE VOUCHERNO  = " & xADJ_VOUCHERNO & " AND JTYPE = " & N2Str2Null(xADJ_JTYPE) & " AND IS_OTHERS = 1 and RIGHT(ENTITY,6) = " & xCUST_CODE & " AND  DEBIT <> 0 AND ACCT_CODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    Else
        GET_SUM_GJ_AP2 = 0
        Exit Function
    End If
    If Not rsGET_SUM_GJ_AP.EOF And Not rsGET_SUM_GJ_AP.BOF Then
        GET_SUM_GJ_AP2 = NumericVal(rsGET_SUM_GJ_AP!SUM_GJ_AP)
    Else
        GET_SUM_GJ_AP2 = 0
    End If
    Set rsGET_SUM_GJ_AP = Nothing
End Function

Sub GET_AP_GJ_PAYMENT()
    Dim rsGET_GJ_PAYMENT                          As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xAP_VOUCHERnO As String
    Dim xVENDORCODE                               As String
    Dim xVENDORNAME
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xLAST_UPDATED As String
    Dim xACCT_CODE                                As String
    Dim xAMOUNT                                   As Double
    Dim xAMOUNT_TO_PAY As Double
    Dim xBAL As Double
    Dim xJType                                    As String
    Dim xAPVOUCHERNO_DETAIL                       As String

    Set rsGET_GJ_PAYMENT = New ADODB.Recordset
    rsGET_GJ_PAYMENT.Open "SELECT VOUCHERNO,DEBIT,JDATE,JTYPE,ENTITY,INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,IS_OTHERS,ACCT_CODE " & _
                          "FROM AMIS_JOURNAL_DET WHERE LEFT(ACCT_CODE,5) IN ('21-01','21-02','21-07') AND VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " and DEBIT <> 0", gconDMIS, adOpenKeyset
    If Not rsGET_GJ_PAYMENT.EOF And Not rsGET_GJ_PAYMENT.BOF Then
        Do While Not rsGET_GJ_PAYMENT.EOF
            xVOUCHERNO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            xJdate = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!JDate))
            xVENDORCODE = N2Str2Null(Right(Null2String(rsGET_GJ_PAYMENT!ENTITY), 6))
            If Left(Null2String(rsGET_GJ_PAYMENT!ENTITY), 1) = "V" Then
                xVENDORNAME = N2Str2Null(GET_VENDOR_NAME(Right(Null2String(rsGET_GJ_PAYMENT!ENTITY), 6)))
            Else
                xVENDORNAME = N2Str2Null(GET_CUSTNAME(Right(Null2String(rsGET_GJ_PAYMENT!ENTITY), 6)))
            End If
            If IsNull(rsGET_GJ_PAYMENT!INVOICENO) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE))

                If Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) = "OTH" Then
                    xAPVOUCHERNO_DETAIL = N2Str2Null("")
                Else
                    xAPVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                End If
            ElseIf IsNull(rsGET_GJ_PAYMENT!INVOICENO) = False And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = False And rsGET_GJ_PAYMENT!IS_OTHERS = False Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!InvoiceType))
                xAPVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype) & "-" & Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            ElseIf IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And IsNull(rsGET_GJ_PAYMENT!InvoiceType) = True And rsGET_GJ_PAYMENT!IS_OTHERS = True Then
                xINVOICENO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
                xInvoiceType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE))
                xAPVOUCHERNO_DETAIL = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype) & "-" & Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            End If
            xAMOUNT_TO_PAY = 0
            xAMOUNT = NumericVal(rsGET_GJ_PAYMENT!DEBIT)
            xBAL = xAMOUNT_TO_PAY - xAMOUNT
            xACCT_CODE = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!Acct_code))
            xJType = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype))
            xLAST_UPDATED = LOGDATE
            xAP_VOUCHERnO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!jtype) + "-" + Null2String(rsGET_GJ_PAYMENT!VOUCHERNO))
            If xInvoiceType = "'OTH'" Then
                SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,LASTUPDATED,JDATE) " & _
                                "VALUES(" & xAP_VOUCHERnO & "," & xInvoiceType & "," & xINVOICENO & "," & xVENDORCODE & "," & xVENDORNAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT & "," & xBAL & "," & xACCT_CODE & "," & xLAST_UPDATED & "," & xJdate & ")"
                gconDMIS.Execute SQL_STATEMENT
            Else
                gconDMIS.Execute "INSERT INTO AMIS_DETAILS(AMOUNTPAID,VENDORCODE,ACCT_CODE,JDATE,VOUCHERNO,JTYPE,PV_VOUCHERNO) " & _
                                 "VALUES(" & xAMOUNT & "," & xVENDORCODE & "," & xACCT_CODE & "," & xJdate & "," & xVOUCHERNO & "," & xJType & "," & xAPVOUCHERNO_DETAIL & ")"
            End If

            Dim rsGET_AP_SUM                      As ADODB.Recordset
            Dim xSUM_AP                           As Double
            Dim xAPVOUCHERNO                      As String

            xAPVOUCHERNO = N2Str2Null(Null2String(rsGET_GJ_PAYMENT!ADJ_JTYPE) & "-" & Null2String(rsGET_GJ_PAYMENT!ADJ_VOUCHERNO))
            Set rsGET_AP_SUM = New ADODB.Recordset
            rsGET_AP_SUM.Open "SELECT AMOUNT2PAY FROM AMIS_AP WHERE VOUCHERNO = " & xAPVOUCHERNO & " and VENDOR_CODE = " & xVENDORCODE & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            If Not rsGET_AP_SUM.EOF And Not rsGET_AP_SUM.BOF Then
                xSUM_AP = NumericVal(rsGET_AP_SUM!AMOUNT2PAY)
            Else
                xSUM_AP = 0
            End If
            Set rsGET_AP_SUM = Nothing

            Dim rsGET_GJ_SUM_PAYMENT              As ADODB.Recordset
            Dim xSUM_PAYMENT                      As Double
            Dim AP_BALANCE                        As Double

            Set rsGET_GJ_SUM_PAYMENT = New ADODB.Recordset
            rsGET_GJ_SUM_PAYMENT.Open "SELECT ROUND(SUM(AMOUNTPAID),2) AS GJ_PAYMENT FROM AMIS_DETAILS WHERE PV_VOUCHERNO = " & xAPVOUCHERNO & " AND VENDORCODE = " & xVENDORCODE & " AND ACCT_CODE = " & xACCT_CODE & "", gconDMIS, adOpenKeyset
            If Not rsGET_GJ_SUM_PAYMENT.EOF And Not rsGET_GJ_SUM_PAYMENT.BOF Then
                xSUM_PAYMENT = NumericVal(rsGET_GJ_SUM_PAYMENT!GJ_PAYMENT)
            Else
                xSUM_PAYMENT = 0
            End If
            Set rsGET_GJ_SUM_PAYMENT = Nothing

            AP_BALANCE = Round((xSUM_AP - xSUM_PAYMENT), 2)

            'gconDMIS.Execute "UPDATE AMIS_AP SET AMOUNTPAID = " & xSUM_PAYMENT & ", BALANCE = " & AP_BALANCE & " WHERE VOUCHERNO = " & xAPVOUCHERNO & "  and  VENDOR_CODE = " & xVENDORCODE & " AND ACCT_CODE = " & xACCT_CODE & ""
            rsGET_GJ_PAYMENT.MoveNext
        Loop
    End If
    Set rsGET_GJ_PAYMENT = Nothing
End Sub

Sub UNPOST_AP_GJ()
    Dim rsUNPOST_GJ                               As ADODB.Recordset
    Dim rsAMOUNT_TOPAY                            As ADODB.Recordset
    Dim xAPVOUCHERNO                              As String
    Dim xVOUCHERNO                                As String
    Dim xBAL                                      As Double
    Dim xAMOUNT_PAID                              As Double
    xBAL = 0

    Set rsAMOUNT_TOPAY = New ADODB.Recordset


    Set rsUNPOST_GJ = New ADODB.Recordset
    rsUNPOST_GJ.Open "SELECT INVOICENO,INVOICETYPE,ADJ_VOUCHERNO,ADJ_JTYPE,IS_OTHERS,ENTITY,ACCT_CODE,DEBIT " & _
                     "FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND LEFT(ACCT_CODE,5) IN('21-01','21-02','21-07') AND CREDIT <> 0", gconDMIS, adOpenKeyset
    If Not rsUNPOST_GJ.EOF And Not rsUNPOST_GJ.BOF Then
        
'        Do While Not rsUNPOST_GJ.EOF
'            rsAMOUNT_TOPAY.Open "SELECT AMOUNT2PAY,AMOUNTPAID,BALANCE FROM AMIS_AP WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & "  and  VENDOR_CODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & "", gconDMIS, adOpenKeyset
'            If Not rsAMOUNT_TOPAY.EOF And Not rsAMOUNT_TOPAY.BOF Then
'                xAMOUNT_PAID = NumericVal(rsAMOUNT_TOPAY!AMOUNTPAID) - NumericVal(rsUNPOST_GJ!DEBIT)
'                xBAL = NumericVal(rsAMOUNT_TOPAY!BALANCE) + NumericVal(rsUNPOST_GJ!DEBIT)
'            End If
'
'            gconDMIS.Execute "UPDATE AMIS_AP SET AMOUNTPAID = " & xAMOUNT_PAID & " , BALANCE = " & xBAL & " WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & "  and VENDOR_CODE = " & N2Str2Null(Right(rsUNPOST_GJ!ENTITY, 6)) & " AND ACCT_CODE = " & N2Str2Null(rsUNPOST_GJ!Acct_code) & ""
'
'            rsUNPOST_GJ.MoveNext
'        Loop
    End If
    xVOUCHERNO = Null2String(xJOURNALTYPE) & "-" & Null2String(txtVoucherNo.Text)
    gconDMIS.Execute "DELETE FROM AMIS_DETAILS WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & ""
    gconDMIS.Execute "DELETE FROM AMIS_AP WHERE VOUCHERNO = " & N2Str2Null(xVOUCHERNO) & ""

    Set rsUNPOST_GJ = Nothing
    Set rsAMOUNT_TOPAY = Nothing
End Sub

Function Get_SJ_INVOICENO(xVOUCHERNO As String, xJType As String) As String
    Dim rsSJ_Invoice                              As ADODB.Recordset
    Set rsSJ_Invoice = New ADODB.Recordset
    rsSJ_Invoice.Open "SELECT INVOICENO FROM AMIS_JOURNAL_HD WHERE VOUCHERNO ='" & xVOUCHERNO & "' AND JTYPE='" & xJType & "'", gconDMIS, adOpenKeyset
    If Not rsSJ_Invoice.EOF And Not rsSJ_Invoice.BOF Then
        Get_SJ_INVOICENO = Null2String(rsSJ_Invoice!INVOICENO)
    End If
End Function

Function Get_SJ_INVOICETYPE(xVOUCHERNO As String, xJType As String) As String
    Dim rsSJ_Invoice                              As ADODB.Recordset
    Set rsSJ_Invoice = New ADODB.Recordset
    rsSJ_Invoice.Open "SELECT INVOICETYPE FROM AMIS_JOURNAL_HD WHERE VOUCHERNO ='" & xVOUCHERNO & "' AND JTYPE='" & xJType & "'", gconDMIS, adOpenKeyset
    If Not rsSJ_Invoice.EOF And Not rsSJ_Invoice.BOF Then
        Get_SJ_INVOICETYPE = Null2String(rsSJ_Invoice!InvoiceType)
    End If
End Function

Function GJ_REMARKS_XXX() As String
    Dim GJ_REMARKS                                As ADODB.Recordset
    Dim xENTITY2                                  As String
    Set GJ_REMARKS = New ADODB.Recordset
    GJ_REMARKS.Open "SELECT DISTINCT ADJ_REMARKS,ENTITY FROM AMIS_JOURNAL_DET WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND JTYPE = 'GJ'", gconDMIS, adOpenKeyset
    If Not GJ_REMARKS.EOF And Not GJ_REMARKS.BOF Then
        Do While Not GJ_REMARKS.EOF
            If IsNull(GJ_REMARKS!ADJ_REMARKS) <> True Then
                GJ_REMARKS_XXX = GJ_REMARKS_XXX + Null2String(GJ_REMARKS!ADJ_REMARKS)
                GJ_REMARKS_XXX = GJ_REMARKS_XXX + Chr(13)
            End If
            GJ_REMARKS.MoveNext
        Loop
    End If
    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET REMARKS=NULL WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET REMARKS='" & GJ_REMARKS_XXX & "' WHERE VOUCHERNO = '" & txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    '            If xENTITY = "C" Then
    '                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET CUSTOMERCODE='" & xENTITY2 & "' WHERE VOUCHERNO = '" & frmAMIS_GJ_JOURNAL_ENTRY.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    '            Else
    '                gconDMIS.Execute ("UPDATE AMIS_JOURNAL_HD SET VENDORCODE='" & xENTITY2 & "' WHERE VOUCHERNO = '" & frmAMIS_GJ_JOURNAL_ENTRY.txtVoucherNo.Text & "' AND JTYPE = 'GJ'")
    '            End If
    Set GJ_REMARKS = Nothing
End Function
