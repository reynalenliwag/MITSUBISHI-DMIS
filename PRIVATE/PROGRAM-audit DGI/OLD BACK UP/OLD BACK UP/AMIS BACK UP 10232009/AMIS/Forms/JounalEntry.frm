VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMISJournalEntry 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   7755
   ClientLeft      =   11040
   ClientTop       =   4800
   ClientWidth     =   9885
   ForeColor       =   &H00FFFFFF&
   Icon            =   "JounalEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   9885
   Begin TabDlg.SSTab JournalTAB 
      Height          =   4215
      Left            =   180
      TabIndex        =   131
      Top             =   2490
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   7435
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
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
      TabPicture(0)   =   "JounalEntry.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label53"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAddJournal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraAddJournal"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraDetails"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Command5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "[<F4> Add &Details]   [<Ctrl> + <D> View &Details]   "
      TabPicture(1)   =   "JounalEntry.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picPV_Detail"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdPV_Entry"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "picPV_Entry"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command5 
         Caption         =   "::"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74670
         TabIndex        =   247
         Top             =   3630
         Width           =   435
      End
      Begin VB.PictureBox fraDetails 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   -74910
         ScaleHeight     =   3435
         ScaleWidth      =   9405
         TabIndex        =   152
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
            TabIndex        =   153
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
            MouseIcon       =   "JounalEntry.frx":0902
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
            TabIndex        =   154
            Top             =   2940
            Width           =   9345
            Begin VB.PictureBox picChat 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   60
               ScaleHeight     =   345
               ScaleWidth      =   6195
               TabIndex        =   155
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
                  TabIndex        =   156
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
               TabIndex        =   158
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
               TabIndex        =   160
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
               TabIndex        =   159
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
               TabIndex        =   157
               Top             =   90
               Width           =   1275
            End
         End
      End
      Begin VB.PictureBox fraAddJournal 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   -74790
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   162
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
            MouseIcon       =   "JounalEntry.frx":0A64
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":0BB6
            Style           =   1  'Graphical
            TabIndex        =   195
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
            MouseIcon       =   "JounalEntry.frx":0EF4
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":1046
            Style           =   1  'Graphical
            TabIndex        =   178
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
            TabIndex        =   174
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
            TabIndex        =   172
            Top             =   330
            Width           =   1100
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   2310
            TabIndex        =   169
            Top             =   -30
            Width           =   4425
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   315
               Left            =   30
               TabIndex        =   171
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393217
               BackColor       =   16777215
               MultiLine       =   0   'False
               TextRTF         =   $"JounalEntry.frx":1371
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
               TabIndex        =   170
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
            TabIndex        =   167
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
            TabIndex        =   168
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
            TabIndex        =   166
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin VB.Frame fraATC 
            Height          =   915
            Left            =   2340
            TabIndex        =   179
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
               TabIndex        =   183
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
               TabIndex        =   184
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
               TabIndex        =   185
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
               TabIndex        =   186
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
               TabIndex        =   181
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
               TabIndex        =   180
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
               TabIndex        =   182
               Top             =   240
               Width           =   1725
            End
         End
         Begin VB.Frame fraComp 
            Height          =   915
            Left            =   2340
            TabIndex        =   187
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
               TabIndex        =   193
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
               TabIndex        =   192
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
               TabIndex        =   191
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
               TabIndex        =   190
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
               TabIndex        =   189
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
               TabIndex        =   188
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
            MouseIcon       =   "JounalEntry.frx":1404
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":1556
            Style           =   1  'Graphical
            TabIndex        =   194
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
            TabIndex        =   175
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
            TabIndex        =   163
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
            TabIndex        =   164
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
            TabIndex        =   165
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
            TabIndex        =   173
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
            TabIndex        =   177
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
            TabIndex        =   176
            Top             =   420
            Width           =   2685
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1845
         Left            =   -74880
         TabIndex        =   161
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
         MICON           =   "JounalEntry.frx":18A6
      End
      Begin VB.PictureBox picPV_Entry 
         BackColor       =   &H00FF8080&
         Height          =   1575
         Left            =   210
         ScaleHeight     =   1515
         ScaleWidth      =   9105
         TabIndex        =   137
         Top             =   750
         Width           =   9165
         Begin VB.CommandButton Command4 
            Caption         =   ".."
            Height          =   345
            Left            =   5010
            TabIndex        =   243
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
            TabIndex        =   221
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
            MouseIcon       =   "JounalEntry.frx":18C2
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":1A14
            Style           =   1  'Graphical
            TabIndex        =   151
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
            MouseIcon       =   "JounalEntry.frx":1D52
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":1EA4
            Style           =   1  'Graphical
            TabIndex        =   149
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
            TabIndex        =   148
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
            TabIndex        =   146
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
            TabIndex        =   143
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
            TabIndex        =   147
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
            TabIndex        =   144
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
            MouseIcon       =   "JounalEntry.frx":21CF
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":2321
            Style           =   1  'Graphical
            TabIndex        =   150
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
            TabIndex        =   220
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
            TabIndex        =   219
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
            TabIndex        =   142
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
            TabIndex        =   138
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
            TabIndex        =   139
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
            TabIndex        =   140
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
            TabIndex        =   141
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
            TabIndex        =   145
            Top             =   420
            Width           =   1305
         End
      End
      Begin wizButton.cmd cmdPV_Entry 
         Height          =   1635
         Left            =   180
         TabIndex        =   136
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
         MICON           =   "JounalEntry.frx":2671
      End
      Begin VB.PictureBox picPV_Detail 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3795
         Left            =   60
         ScaleHeight     =   3795
         ScaleWidth      =   9405
         TabIndex        =   132
         Top             =   90
         Width           =   9405
         Begin MSComctlLib.ListView lstPV_Detail 
            Height          =   3285
            Left            =   60
            TabIndex        =   133
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
            MouseIcon       =   "JounalEntry.frx":268D
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
            TabIndex        =   134
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
            TabIndex        =   135
            Top             =   3450
            Width           =   1275
         End
      End
      Begin VB.Label Label53 
         Caption         =   "[    ] View Unapplied Payments"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74670
         TabIndex        =   246
         Top             =   3660
         Width           =   4365
      End
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   10020
      ScaleHeight     =   795
      ScaleWidth      =   2715
      TabIndex        =   236
      Top             =   6840
      Width           =   2775
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption56 
         Height          =   825
         Left            =   -30
         TabIndex        =   244
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
      TabIndex        =   222
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
         TabIndex        =   230
         Top             =   30
         Width           =   2775
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   30
            TabIndex        =   233
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
            TabIndex        =   235
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
            TabIndex        =   234
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
            TabIndex        =   232
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
            TabIndex        =   231
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
         TabIndex        =   223
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
            MouseIcon       =   "JounalEntry.frx":27EF
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   227
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
            MouseIcon       =   "JounalEntry.frx":2941
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   224
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
            MouseIcon       =   "JounalEntry.frx":2A93
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   242
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
            MouseIcon       =   "JounalEntry.frx":2BE5
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   228
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
            MouseIcon       =   "JounalEntry.frx":2D37
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   226
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
            MouseIcon       =   "JounalEntry.frx":2E89
            MousePointer    =   99  'Custom
            Style           =   1  'Graphical
            TabIndex        =   225
            Top             =   2610
            Width           =   2625
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   345
            Index           =   1
            Left            =   0
            TabIndex        =   229
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
      Left            =   180
      ScaleHeight     =   2610
      ScaleWidth      =   9690
      TabIndex        =   0
      Top             =   0
      Width           =   9690
      Begin VB.PictureBox picDisbursement 
         BorderStyle     =   0  'None
         Height          =   1245
         Left            =   90
         ScaleHeight     =   1245
         ScaleWidth      =   9525
         TabIndex        =   25
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   30
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
            TabIndex        =   28
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
            TabIndex        =   26
            Text            =   "000226"
            Top             =   30
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtParticulars 
            Height          =   795
            Left            =   4380
            TabIndex        =   33
            Top             =   420
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1402
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            TextRTF         =   $"JounalEntry.frx":2FDB
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
            TabIndex        =   218
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
            TabIndex        =   29
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
            TabIndex        =   34
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   27
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
         TabIndex        =   37
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
            TabIndex        =   53
            Text            =   "Invoice Type"
            Top             =   900
            Width           =   4950
         End
         Begin VB.CheckBox chkNonVat 
            Caption         =   "Non-Vat"
            Height          =   285
            Left            =   1140
            TabIndex        =   49
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
            TabIndex        =   56
            Top             =   930
            Width           =   1755
         End
         Begin RichTextLib.RichTextBox txtRemarks2 
            Height          =   705
            Left            =   4560
            TabIndex        =   62
            Top             =   1350
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   1244
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            TextRTF         =   $"JounalEntry.frx":306F
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   42
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
            TabIndex        =   39
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
            TabIndex        =   63
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
            TabIndex        =   58
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
            TabIndex        =   40
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   48
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
            TabIndex        =   51
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
            TabIndex        =   57
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
            TabIndex        =   41
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
            TabIndex        =   52
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
            TabIndex        =   47
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
            TabIndex        =   44
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
            TabIndex        =   43
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
            TabIndex        =   38
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
            TabIndex        =   60
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
            TabIndex        =   61
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
            TabIndex        =   59
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
            TabIndex        =   50
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   13
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
         TabIndex        =   15
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
            TabIndex        =   16
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
            TabIndex        =   24
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
            TabIndex        =   21
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
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   60
            Width           =   2325
         End
         Begin RichTextLib.RichTextBox txtRemarks 
            Height          =   765
            Left            =   4380
            TabIndex        =   23
            Top             =   420
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   1349
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            TextRTF         =   $"JounalEntry.frx":3106
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
            TabIndex        =   19
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
            TabIndex        =   22
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
            TabIndex        =   20
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
            TabIndex        =   17
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
         TabIndex        =   245
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
         TabIndex        =   11
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
         TabIndex        =   8
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
         TabIndex        =   5
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
         TabIndex        =   14
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
         TabIndex        =   12
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
      TabIndex        =   212
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
         MouseIcon       =   "JounalEntry.frx":319D
         MousePointer    =   99  'Custom
         TabIndex        =   216
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
         MouseIcon       =   "JounalEntry.frx":34A7
         MousePointer    =   99  'Custom
         TabIndex        =   215
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
         MouseIcon       =   "JounalEntry.frx":37B1
         MousePointer    =   99  'Custom
         TabIndex        =   214
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
         MouseIcon       =   "JounalEntry.frx":3ABB
         MousePointer    =   99  'Custom
         TabIndex        =   213
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
      TabIndex        =   107
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
         TabIndex        =   108
         Top             =   0
         Width           =   2775
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1170
      TabIndex        =   106
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
      MICON           =   "JounalEntry.frx":3DC5
   End
   Begin VB.PictureBox picTemplates 
      Height          =   4125
      Left            =   1260
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   109
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
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   111
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
         MouseIcon       =   "JounalEntry.frx":3DE1
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
         TabIndex        =   112
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
      TabIndex        =   65
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
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   270
         Width           =   9195
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   4515
         Left            =   90
         TabIndex        =   68
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
         MouseIcon       =   "JounalEntry.frx":3F43
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
         TabIndex        =   69
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
         TabIndex        =   67
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
         TabIndex        =   70
         Top             =   5310
         Width           =   9225
      End
   End
   Begin VB.CommandButton cmdPrinting 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Command1"
      Height          =   2445
      Left            =   3450
      TabIndex        =   122
      Top             =   1830
      Width           =   2775
   End
   Begin wizButton.cmd cmdShowPostRange 
      Height          =   2385
      Left            =   3540
      TabIndex        =   113
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
      MICON           =   "JounalEntry.frx":40A5
   End
   Begin VB.PictureBox picShowPostRange 
      Height          =   2235
      Left            =   3600
      ScaleHeight     =   2175
      ScaleWidth      =   2535
      TabIndex        =   114
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
         TabIndex        =   120
         Top             =   1350
         Width           =   2295
      End
      Begin wizProgBar.Prg prgPostRange 
         Height          =   285
         Left            =   90
         TabIndex        =   121
         Top             =   1800
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         Picture         =   "JounalEntry.frx":40C1
         ForeColor       =   0
         BarPicture      =   "JounalEntry.frx":40DD
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
         TabIndex        =   119
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
         TabIndex        =   117
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
         TabIndex        =   115
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
         TabIndex        =   118
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
         TabIndex        =   116
         Top             =   480
         Width           =   735
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   5775
      Left            =   210
      TabIndex        =   64
      Top             =   240
      Width           =   9555
      _ExtentX        =   16854
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
      MICON           =   "JounalEntry.frx":40F9
   End
   Begin VB.PictureBox picPrinting 
      Height          =   2265
      Left            =   3600
      ScaleHeight     =   2205
      ScaleWidth      =   2535
      TabIndex        =   123
      Top             =   1920
      Width           =   2595
      Begin VB.PictureBox picPrintCheck 
         Enabled         =   0   'False
         Height          =   885
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   2355
         TabIndex        =   125
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
            TabIndex        =   126
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
            TabIndex        =   127
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
            TabIndex        =   128
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
         TabIndex        =   130
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
         TabIndex        =   129
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
         TabIndex        =   124
         Top             =   60
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   8130
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   196
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
         MouseIcon       =   "JounalEntry.frx":4115
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":4267
         Style           =   1  'Graphical
         TabIndex        =   198
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
         MouseIcon       =   "JounalEntry.frx":45A5
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":46F7
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   60
      ScaleHeight     =   900
      ScaleWidth      =   9735
      TabIndex        =   199
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
         MouseIcon       =   "JounalEntry.frx":4A47
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":4B99
         Style           =   1  'Graphical
         TabIndex        =   211
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
         MouseIcon       =   "JounalEntry.frx":4EFF
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":5051
         Style           =   1  'Graphical
         TabIndex        =   210
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
         MouseIcon       =   "JounalEntry.frx":53B7
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":5509
         Style           =   1  'Graphical
         TabIndex        =   209
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
         MouseIcon       =   "JounalEntry.frx":5843
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":5995
         Style           =   1  'Graphical
         TabIndex        =   208
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
         MouseIcon       =   "JounalEntry.frx":5CDA
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":5E2C
         Style           =   1  'Graphical
         TabIndex        =   207
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
         MouseIcon       =   "JounalEntry.frx":6151
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":62A3
         Style           =   1  'Graphical
         TabIndex        =   206
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
         MouseIcon       =   "JounalEntry.frx":65FF
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":6751
         Style           =   1  'Graphical
         TabIndex        =   205
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
         MouseIcon       =   "JounalEntry.frx":6A64
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":6BB6
         Style           =   1  'Graphical
         TabIndex        =   204
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
         MouseIcon       =   "JounalEntry.frx":6F06
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":7058
         Style           =   1  'Graphical
         TabIndex        =   203
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
         MouseIcon       =   "JounalEntry.frx":73B6
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":7508
         Style           =   1  'Graphical
         TabIndex        =   202
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
         MouseIcon       =   "JounalEntry.frx":7802
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":7954
         Style           =   1  'Graphical
         TabIndex        =   201
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
         MouseIcon       =   "JounalEntry.frx":7CAC
         MousePointer    =   99  'Custom
         Picture         =   "JounalEntry.frx":7DFE
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox picGJ 
      BorderStyle     =   0  'None
      Height          =   6315
      Left            =   150
      ScaleHeight     =   6315
      ScaleWidth      =   9555
      TabIndex        =   71
      Top             =   420
      Width           =   9555
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   555
         Left            =   0
         TabIndex        =   237
         Top             =   5730
         Width           =   9405
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
            TabIndex        =   240
            Text            =   "Text1"
            Top             =   210
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
            Left            =   6150
            MaxLength       =   14
            TabIndex        =   239
            Text            =   "Text1"
            Top             =   210
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
            Left            =   7710
            MaxLength       =   14
            TabIndex        =   238
            Text            =   "Text1"
            Top             =   210
            Width           =   1485
         End
         Begin VB.Label Label19 
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
            TabIndex        =   241
            Top             =   210
            Width           =   1275
         End
      End
      Begin VB.PictureBox picGJEntry 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         ForeColor       =   &H80000008&
         Height          =   1995
         Left            =   120
         ScaleHeight     =   1965
         ScaleWidth      =   9225
         TabIndex        =   76
         Top             =   2880
         Width           =   9255
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
            TabIndex        =   90
            Text            =   "Combo1"
            Top             =   690
            Width           =   4305
         End
         Begin VB.Frame fraATC2 
            BackColor       =   &H00F5F5F5&
            Height          =   915
            Left            =   2310
            TabIndex        =   92
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
               TabIndex        =   98
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
               TabIndex        =   97
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
               TabIndex        =   96
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
               TabIndex        =   95
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
               TabIndex        =   94
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
               TabIndex        =   93
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
               TabIndex        =   99
               Top             =   540
               Width           =   855
            End
         End
         Begin VB.CommandButton cmdGJCancel 
            BackColor       =   &H00F2EFE9&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8160
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "JounalEntry.frx":815D
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":82AF
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   1050
            Width           =   1005
         End
         Begin VB.CommandButton cmdGJSave 
            BackColor       =   &H00F2EFE9&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   7140
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "JounalEntry.frx":85C1
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":8713
            Style           =   1  'Graphical
            TabIndex        =   101
            Top             =   1050
            Width           =   975
         End
         Begin VB.CommandButton cmdGJDelete 
            BackColor       =   &H00F2EFE9&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   90
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "JounalEntry.frx":8B55
            MousePointer    =   99  'Custom
            Picture         =   "JounalEntry.frx":8CA7
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   1050
            Width           =   1005
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
            TabIndex        =   81
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin RichTextLib.RichTextBox txtGJAccountName 
            Height          =   315
            Left            =   2340
            TabIndex        =   83
            Top             =   330
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   15920873
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"JounalEntry.frx":8FB1
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
            TabIndex        =   84
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
            TabIndex        =   86
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
            Left            =   7140
            TabIndex        =   102
            Top             =   1110
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
            TabIndex        =   82
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin RichTextLib.RichTextBox txtGJAccountParticulars 
            Height          =   885
            Left            =   2340
            TabIndex        =   104
            Top             =   2250
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1561
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   0   'False
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"JounalEntry.frx":9044
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
            TabIndex        =   91
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
            TabIndex        =   105
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
            TabIndex        =   88
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
            TabIndex        =   89
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
            TabIndex        =   85
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
            TabIndex        =   80
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
            TabIndex        =   79
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
            TabIndex        =   77
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
            TabIndex        =   87
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
            TabIndex        =   78
            Top             =   60
            Width           =   2205
         End
      End
      Begin wizButton.cmd cmdGJEntry 
         Height          =   2115
         Left            =   60
         TabIndex        =   75
         Top             =   2820
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3731
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
         MICON           =   "JounalEntry.frx":90DB
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   90
         Top             =   3870
      End
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   60
         TabIndex        =   73
         Top             =   330
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"JounalEntry.frx":90F7
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
      Begin MSComctlLib.ListView lstGJ 
         Height          =   4335
         Left            =   60
         TabIndex        =   74
         Top             =   1110
         Width           =   9405
         _ExtentX        =   16589
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "JounalEntry.frx":918E
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
         TabIndex        =   72
         Top             =   60
         Width           =   1695
      End
   End
   Begin VB.Label lblVPJAcctCode 
      Caption         =   "dont delete this "
      Height          =   165
      Left            =   11220
      TabIndex        =   217
      Top             =   1080
      Width           =   1845
   End
End
Attribute VB_Name = "frmAMISJournalEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_hd                                       As New ADODB.Recordset
Dim rsJournal_Det                                      As New ADODB.Recordset
Dim RSPV_DETAIL                                        As New ADODB.Recordset
Dim rsCV_Detail                                        As New ADODB.Recordset
Dim rsCRJ_Detail                                       As New ADODB.Recordset
Dim RSJV_DETAIL                                        As New ADODB.Recordset
Dim rsChartAccount                                     As New ADODB.Recordset
Dim rsJournal_HD2                                      As New ADODB.Recordset
Dim rsProfile                                          As New ADODB.Recordset
Dim rsCheckJournal_HD                                  As New ADODB.Recordset
Dim rsVENDOR                                           As New ADODB.Recordset
Dim rsPayTerm                                          As New ADODB.Recordset
Dim rsBanks                                            As New ADODB.Recordset
Dim rsCUSTOMER                                         As New ADODB.Recordset
Dim rsInvoiceType                                      As New ADODB.Recordset
Dim RSATC                                              As New ADODB.Recordset
Dim kcnt, JCNT                                         As Integer
Attribute JCNT.VB_VarUserMemId = 1073938448
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938450
Dim SearchBy                                           As String
Attribute SearchBy.VB_VarUserMemId = 1073938451
Public CDJ_CIB                                         As String
Attribute CDJ_CIB.VB_VarUserMemId = 1073938452
Public CDJ_AP                                          As String
Attribute CDJ_AP.VB_VarUserMemId = 1073938453
Dim LocalAcess                                         As String
Attribute LocalAcess.VB_VarUserMemId = 1073938454
Dim TOTDEBIT                                           As Double
Attribute TOTDEBIT.VB_VarUserMemId = 1073938455
Dim TOTCREDIT                                          As Double
Attribute TOTCREDIT.VB_VarUserMemId = 1073938456
Dim TOTTAX                                             As Double
Attribute TOTTAX.VB_VarUserMemId = 1073938457
Dim OUTBALANCE                                         As Double
Attribute OUTBALANCE.VB_VarUserMemId = 1073938458
Dim TOTAL_AR_AMOUNT                                    As Double
Attribute TOTAL_AR_AMOUNT.VB_VarUserMemId = 1073938459
Dim TOTALPVAMOUNT                                      As Double
Attribute TOTALPVAMOUNT.VB_VarUserMemId = 1073938460
Dim COMP_SJ_OUTPUT_TAX                                 As Double
Attribute COMP_SJ_OUTPUT_TAX.VB_VarUserMemId = 1073938461
Dim PrevJType                                          As String
Attribute PrevJType.VB_VarUserMemId = 1073938462
Dim PrevJNo                                            As String
Attribute PrevJNo.VB_VarUserMemId = 1073938463
Dim PREVINVOICETYPE                                    As String
Attribute PREVINVOICETYPE.VB_VarUserMemId = 1073938464
Dim PREVINVOICENO                                      As String
Attribute PREVINVOICENO.VB_VarUserMemId = 1073938465
Dim PREVPV_VOUCHERNO                                   As String
Attribute PREVPV_VOUCHERNO.VB_VarUserMemId = 1073938466
Dim PREVPV_AMOUNT                                      As Double
Attribute PREVPV_AMOUNT.VB_VarUserMemId = 1073938467
Dim DIRECTDISBURSEMENTVOUCHERNO                        As String
Attribute DIRECTDISBURSEMENTVOUCHERNO.VB_VarUserMemId = 1073938468
Dim CDJ_IS_FROM_AP                                     As Boolean
Attribute CDJ_IS_FROM_AP.VB_VarUserMemId = 1073938469
Dim ISVPJ                                              As Boolean
Attribute ISVPJ.VB_VarUserMemId = 1073938470
Dim TotalARAmountToPay                                 As Double
Attribute TotalARAmountToPay.VB_VarUserMemId = 1073938471
Dim TOTAL_AP_AMOUNT                                    As Double
Attribute TOTAL_AP_AMOUNT.VB_VarUserMemId = 1073938472
Dim TOTALAPAMOUNTTOPAY                                 As Double
Attribute TOTALAPAMOUNTTOPAY.VB_VarUserMemId = 1073938473
Dim SJVOUCHERNO                                        As String
Attribute SJVOUCHERNO.VB_VarUserMemId = 1073938474
Dim APJINVOICENO                                       As String
Attribute APJINVOICENO.VB_VarUserMemId = 1073938475
Dim APJINVOICETYPE                                     As String
Attribute APJINVOICETYPE.VB_VarUserMemId = 1073938476

Function GetVoucherNo(XXX As String) As String
    Dim rsJournal_hd                                   As ADODB.Recordset
    Set rsJournal_hd = New ADODB.Recordset
    Set rsJournal_hd = gconDMIS.Execute("Select * from AMIS_Journal_HD Where Jtype = '" & XXX & "' Order by VoucherNo desc")
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        GetVoucherNo = Format(NumericVal(rsJournal_hd!VOUCHERNO) + 1, "000000")
    Else
        GetVoucherNo = "000001"
    End If
End Function

Function Setacctcode(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where Description = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctcode = UCase(Null2String(rsChartAccount2!acctcode))
    Else
        Setacctcode = ""
    End If
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function SetAcctType(VVV As Variant) As String
    Dim rsChartAccount2                                As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,AcctType from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctType = SetDebitCredit(Null2String(rsChartAccount2!AcctType))
    Else
        SetAcctType = ""
    End If
End Function

Function SetBankCode(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankname = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankCode = Null2String(rsBanks!bankcode)
        CDJ_CIB = N2Str2Null(rsBanks!acctcode)
    Else
        SetBankCode = ""
        CDJ_CIB = "NULL"
    End If
End Function

Function SetBankName(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        rsBanks.Open "Select bankcode,bankname,acctcode from CMIS_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            SetBankName = Null2String(rsBanks!BankName)
            CDJ_CIB = N2Str2Null(rsBanks!acctcode)
        Else
            SetBankName = ""
            CDJ_CIB = "NULL"
        End If
    Else
        rsBanks.Open "Select bankcode,bankname,acctcode from ALL_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            SetBankName = Null2String(rsBanks!BankName)
            CDJ_CIB = N2Str2Null(rsBanks!acctcode)
        Else
            SetBankName = ""
            CDJ_CIB = "NULL"
        End If
    End If
End Function

Function SetCustomerCode(CCC As Variant)
    Set rsCUSTOMER = New ADODB.Recordset
    '    rsCustomer.Open "Select cuscde,acctname from ALL_CUSTMASTER_AMIS where acctname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsCUSTOMER.Open "Select custcode,custname from ALL_CUSTMASTER_AMIS where custname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        SetCustomerCode = Null2String(rsCUSTOMER!CUSTCODE)
    Else
        SetCustomerCode = ""
    End If
End Function

Function SetCustomerName(CCC As Variant)
    Set rsCUSTOMER = New ADODB.Recordset
    Set rsCUSTOMER = gconDMIS.Execute("Select custname from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(CCC))
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        SetCustomerName = Null2String(rsCUSTOMER!CUSTNAME)
    Else
        SetCustomerName = ""
    End If
    Set rsCUSTOMER = Nothing
End Function

Function SetCustomerCreditTerm(CCC As Variant)
    Set rsCUSTOMER = New ADODB.Recordset
    Set rsCUSTOMER = gconDMIS.Execute("Select CREDITDAYS from ALL_Customer_Table where cuscde = " & N2Str2Null(CCC))
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        SetCustomerCreditTerm = Null2String(rsCUSTOMER!CREDITDAYS)
    Else
        SetCustomerCreditTerm = 0
    End If
    Set rsCUSTOMER = Nothing
End Function

Function SetDebitCredit(VVV As Variant) As String
    Dim rsAccountType                                  As ADODB.Recordset
    Set rsAccountType = New ADODB.Recordset
    rsAccountType.Open "Select Code,DebitCredit from AMIS_Acctype where Code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAccountType.EOF And Not rsAccountType.BOF Then
        If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
            If txtAcct_Name.Text = "ACCOUNTS PAYABLE - TRADE" Then SetDebitCredit = "D"
        ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
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
        SetVendorCode = Null2String(rsVENDOR!code)
    Else
        SetVendorCode = ""
    End If
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!Nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function StoreGJEntry(ByVal ID As Variant)
    On Error GoTo Errorcode
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,JNo,acct_code,acct_name,debit,jitemno,credit,tax,atc,rate,taxbase from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labGJID.Caption = rsJournal_Det!ID
        txtGJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboGJAccountNo.Text = Null2String(rsJournal_Det!ACCT_CODE)
        txtGJAccountName.Text = Null2String(rsJournal_Det!acct_Name)
        txtGJDebit.Text = N2Str2Zero(rsJournal_Det!DEBIT)
        txtGJCredit.Text = N2Str2Zero(rsJournal_Det!CREDIT)
        If fraATC2.Visible = True Then
            cboJVSupCust.Text = SetVendorName(Null2String(rsJournal_hd!VendorCode))
            If Null2String(rsJournal_Det!ATC) <> "" Then
                cboATC2.Text = Null2String(rsJournal_Det!ATC)
            Else
                cboATC2.ListIndex = 0
            End If
            txtRATE2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Rate))
            txtTaxBase2.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!taxbase))
        End If
        StoreGJParticulars Null2String(rsJournal_Det!JNo), Null2String(rsJournal_Det!jitemno)
    End If
    Exit Function

Errorcode:
    Resume Next
End Function

Function StoreGJParticulars(ByVal JNo As Variant, ByVal ItemNo As Variant)
    Set RSJV_DETAIL = New ADODB.Recordset
    RSJV_DETAIL.Open "select JNo,ItemNo,Particulars from AMIS_JV_Detail where JNo = " & N2Str2Null(JNo) & " and ItemNo = " & N2Str2Null(ItemNo), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSJV_DETAIL.EOF And Not RSJV_DETAIL.BOF Then
        txtGJAccountParticulars.Text = Null2String(RSJV_DETAIL!Particulars)
    End If
End Function

Function StoreJournalEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,acct_code,acct_name,debit,jitemno,credit,tax,grossamt,netamt,ATC,RATE,TAXBASE from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labDetID.Caption = rsJournal_Det!ID
        labPartNo.Caption = Null2String(rsJournal_Det!ACCT_CODE)
        txtJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboAcct_Code.Text = Null2String(rsJournal_Det!ACCT_CODE)
        txtAcct_Name.Text = Null2String(rsJournal_Det!acct_Name)
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!tax))
        txtGrossAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!grossamt))
        txtNetAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!netamt))
        If JOURNALTYPE = "APJ" And fraATC.Visible = True Then
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
    Dim rsREPORWITHDealer                              As ADODB.Recordset
    Set rsREPORWITHDealer = New ADODB.Recordset
    Set rsREPORWITHDealer = gconDMIS.Execute("SELECT  dbo.CSMS_SellingDealer.DealerCode AS VEHICLE_DEALER_CODE, dbo.CSMS_Repor.INVOICE FROM dbo.CSMS_Repor INNER JOIN dbo.CSMS_CusVeh ON dbo.CSMS_Repor.PLATE_NO = dbo.CSMS_CusVeh.VCOND_NO INNER JOIN dbo.CSMS_SellingDealer ON dbo.CSMS_CusVeh.SELLING_DEALER = dbo.CSMS_SellingDealer.DealerCode Where dbo.CSMS_Repor.INVOICE = '" & XXX & "'")
    If Not rsREPORWITHDealer.EOF And Not rsREPORWITHDealer.BOF Then
        StoreDealerCode = Null2String(rsREPORWITHDealer!VEHICLE_DEALER_CODE)
    End If
    Set rsREPORWITHDealer = Nothing
End Function

Function StorePVEntry(ByVal ID As Variant)
    If JOURNALTYPE = "APJ" Then
        Set RSPV_DETAIL = New ADODB.Recordset
        RSPV_DETAIL.Open "select * from AMIS_PV_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPV_DETAIL.EOF And Not RSPV_DETAIL.BOF Then
            labPVID.Caption = RSPV_DETAIL!ID
            txtPVItemNo.Text = Null2String(RSPV_DETAIL!ItemNo)
            txtPO_No.Text = Null2String(RSPV_DETAIL!po_no)
            txtMRR_No.Text = Null2String(RSPV_DETAIL!MRR_No)
            txtINV_No.Text = Null2String(RSPV_DETAIL!INV_NO)
            txtProd_No.Text = Null2String(RSPV_DETAIL!Prod_No)
            txtPVAmount.Text = N2Str2Zero(RSPV_DETAIL!amount)
        End If
    ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        Set rsCRJ_Detail = New ADODB.Recordset
        rsCRJ_Detail.Open "select * from AMIS_CRJ_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            labPVID.Caption = rsCRJ_Detail!ID
            txtPVItemNo.Text = Null2String(rsCRJ_Detail!ItemNo)
            txtPO_No.Text = txtVoucherNo.Text
            ReturnAccountDescription (Null2String(rsCRJ_Detail!j_class))
            txtPO_No.Enabled = False
            cboARTag.Text = Setacctname(Null2String(rsCRJ_Detail!j_class))
            txtMRR_No.Text = Null2String(rsCRJ_Detail!InvoiceType)
            txtINV_No.Text = Null2String(rsCRJ_Detail!INVOICENO)
            txtProd_No.Text = Null2String(rsCRJ_Detail!invoicedate)
            txtPVAmount.Text = N2Str2Zero(rsCRJ_Detail!invoiceamount)
            PREVINVOICETYPE = Null2String(rsCRJ_Detail!InvoiceType)
            PREVINVOICENO = Null2String(rsCRJ_Detail!INVOICENO)
            PREVPV_AMOUNT = N2Str2Zero(rsCRJ_Detail!invoiceamount)
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
            PREVPV_VOUCHERNO = Null2String(rsCV_Detail!pv_voucherno)
            txtINV_No.Text = Null2String(rsCV_Detail!docdate)
            txtProd_No.Text = Null2String(rsCV_Detail!duedate)
            txtPVAmount.Text = N2Str2Zero(rsCV_Detail!amount)
            txtMRR_No.Enabled = False
            txtProd_No.Enabled = True
            txtINV_No.Enabled = True
            PREVPV_AMOUNT = N2Str2Zero(rsCV_Detail!amount)
        End If
    End If
End Function

Function ReturnAP_AccountCode(XXX As String) As String
    Dim rsChartAccount                                 As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE2 = 'AP' AND TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnAP_AccountCode = Null2String(rsChartAccount!acctcode)
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
    cmdGJEntry.ZOrder 0
    cmdGJEntry.Visible = True
    picGJEntry.ZOrder 0
    picGJEntry.Visible = True
    picGJEntry.Enabled = True
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
If JOURNALTYPE <> "SJ" Then
   JournalTAB.Tab = 1
   txtMRR_No.BackColor = &HFFFFFF
   txtINV_No.BackColor = &HFFFFFF
 Else
    ShowInvoiceApp SetInvCode(cboInvoiceType), txtInvoiceNo.Text
 End If
End Sub

Private Sub cmdNoCharge_Click()
If JOURNALTYPE = "SJ" Then
    ReturnInvoiceNo txtVoucherNo, JOURNALTYPE
    With frmAMIS_Payment
        frmAMIS_Payment.FillPaymentdetail AMIS_Invoiceno, AMIS_Invoicetype
        frmAMIS_Payment.Show
    End With
End If
If JOURNALTYPE = "APJ" Then
    With frmAMIS_Payment
        frmAMIS_Payment.FillPaymentdetail txtVoucherNo, ""
        frmAMIS_Payment.Show
    End With
End If
End Sub

Sub cmdPVSave_Click()
    On Error GoTo Errorcode
    Dim Ans                                            As String
    txtMRR_No.BackColor = &HFFFFFF
    txtINV_No.BackColor = &HFFFFFF
    If AddorEdit = "ADD" Then
        Dim rsPV_DetailClone                           As ADODB.Recordset
        Set rsPV_DetailClone = New ADODB.Recordset
        rsPV_DetailClone.Open "select * from AMIS_PV_Detail where PO_NO = " & N2Str2Null(txtPO_No.Text) & " and MRR_NO = " & N2Str2Null(txtMRR_No.Text) & " and INV_NO = " & N2Str2Null(txtINV_No.Text), gconDMIS
        If Not rsPV_DetailClone.EOF And Not rsPV_DetailClone.BOF Then
            MsgBox "PO Number : " & txtPO_No.Text & " with MRR Number : " & txtMRR_No.Text & " and Invoice Number : " & txtINV_No.Text & " already used in this transaction", vbInformation, "Error in PO Number, MRR Number, Invoice Number Validation"
            Exit Sub
        End If
    End If
    'UPDATE BY BTT : 11/27/2008
    If JOURNALTYPE <> "APJ" Then
            If Len(txtMRR_No.Text) = 0 Then
                If JOURNALTYPE = "CDJ" Then
                    MsgBox "Invalid Link.APJ/VPJ is missing", vbInformation, "WARNING"
                ElseIf JOURNALTYPE = "CRJ" Then
                    MsgBox "Invalid Link.Invoice Type is missing", vbInformation, "WARNING"
                Else
                    MsgBox "Invalid Link", vbInformation, "WARNING"
                End If
                txtMRR_No.BackColor = &HFFFF80
                Exit Sub
            End If
    End If
    If JOURNALTYPE <> "APJ" Then
        If Len(txtINV_No.Text) = 0 Then
            If JOURNALTYPE = "CRJ" Then
                MsgBox "Missing invoice No", vbExclamation, "WARNING"
                txtINV_No.BackColor = &HFFFF80
            End If
            Exit Sub
        End If
    End If
    
'    If JOURNALTYPE = "CRJ" Then
'            If cboARTag.Text = "" Then
'            MsgBox "Tagging of AR is required..", vbExclamation, "WARNING"
'            cboARTag.BackColor = &HFFFF80
'            Exit Sub
'        End If
'    End If
     If JOURNALTYPE = "CRJ" Then
            'UPDATED BY: JUN
            'DATE UPDATE: 07312009
            'DESCRIPTION: CHECK IF THE IS AN AR SCHEDULE
            If CHECK_IF_SCHED_ACCNT(txtVoucherNo.Text) = True Then
                If cboARTag.Text = "" Then
                MsgBox "Tagging of AR is required..", vbExclamation, "WARNING"
                cboARTag.BackColor = &HFFFF80
                Exit Sub
            Else
                'ALLOW SAVING OF ENTRY IT IS NOT AN AR SCHEDULE
            End If
     End If
    
    
    If JOURNALTYPE = "CRJ" Then
        If Not (txtMRR_No = "AI" Or txtMRR_No = "VI" Or txtMRR_No = "SI" Or txtMRR_No = "PI" Or txtMRR_No = "MI" Or txtMRR_No = "CI") Then
            MsgBox "Invalid Invoice Type.", vbExclamation, "WARNING"
            txtMRR_No.BackColor = &HFFFF80
            Exit Sub
        End If

        
        If AddorEdit = "ADD" Or AddorEdit = "EDIT" Then
            If getJTYPE(RTrim(LTrim(txtINV_No)), RTrim(LTrim(txtMRR_No))) = "SJ" Then
                If rsCHECKINVOICENOandTYPE(RTrim(LTrim(txtMRR_No)), RTrim(LTrim(txtINV_No)), RTrim(LTrim(txtCustCode))) = False Then
                    MsgBox "Your invoice is Not existing in the Sales journal or your trying to link a wrong customer Code", vbInformation, "Please verify your Invioce or Customer Code"
                    Exit Sub
                End If
            ElseIf getJTYPE(RTrim(LTrim(txtINV_No)), RTrim(LTrim(txtMRR_No))) = "COB" Then
                Ans = MsgBox("You are adding detail from Customer Opening balance.Are you sure this is correct?", vbQuestion + vbYesNo, "Information")
                If Ans = vbYes Then
                    ' Go save the data
                Else
                    Exit Sub
                End If
            ElseIf txtINV_No = "INTRO" Then
                Ans = MsgBox("You are adding detail with INT-RO.Are you sure this is correct?", vbQuestion + vbYesNo, "Information")
                If Ans = vbYes Then
                    ' Go save the data
                Else
                    Exit Sub
                End If
            Else
                MsgBox "Please verify your Invioce/Customer Code..Not exist/Wrong in the Sales Journal", vbInformation, "Information"
                txtINV_No.BackColor = &HFFFF80
                Exit Sub
            End If
        End If
    End If



    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO         As String
    Dim J_JVOUCHERNO, J_JDATE                          As String
    Dim PV_AMOUNT                                      As Double
    Dim PV_STATUS, PV_ITEMNO                           As String

    J_JVOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JDATE = N2Str2Null(txtJDate.Text)
    PV_ITEMNO = N2Str2Null(Format(txtPVItemNo.Text, "0000"))
    PV_PONO = N2Str2Null(txtPO_No.Text)
    PV_MRRNO = N2Str2Null(txtMRR_No.Text)             ' TYPE
    PV_INVNO = N2Str2Null(txtINV_No.Text)             ' NO
    PV_PRODNO = N2Str2Null(txtProd_No.Text)           ' DATE
    PV_AMOUNT = NumericVal(txtPVAmount.Text)          ' AMOUNT
    PV_STATUS = "'N'"


    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        Dim rsJournal_HD_APJ                           As ADODB.Recordset
        Dim rsPV_Detail_APJ                            As ADODB.Recordset

        Set rsJournal_HD_APJ = New ADODB.Recordset
        Set rsJournal_HD_APJ = gconDMIS.Execute("Select VoucherNo,VendorCode from AMIS_Journal_HD Where Jtype = 'APJ' and VendorCode = '" & txtCode.Text & "' Order By VoucherNo Asc")
        If Not rsJournal_HD_APJ.EOF And Not rsJournal_HD_APJ.BOF Then
            Do While Not rsJournal_HD_APJ.EOF
                Set rsPV_Detail_APJ = New ADODB.Recordset
                Set rsPV_Detail_APJ = gconDMIS.Execute("Select * from AMIS_PV_Detail Where (Inv_No = " & PV_INVNO & " OR Prod_No = " & PV_PRODNO & ") AND VoucherNo = " & N2Str2Null(rsJournal_HD_APJ!VOUCHERNO))
                If Not rsPV_Detail_APJ.EOF And Not rsPV_Detail_APJ.BOF Then
                    Screen.MousePointer = 0
                    MsgBox "Invoice No. or Prod No. Already Used in PV Number - " & Null2String(rsPV_Detail_APJ!VOUCHERNO)
                    Exit Sub
                Else
                    rsJournal_HD_APJ.MoveNext
                End If
            Loop
        End If
        'FOR APJ
        If JOURNALTYPE = "APJ" Then
            SQL_STATEMENT = "insert into AMIS_PV_Detail " & _
                            "(JTYPE,VoucherNo,JDATE,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                          " values ('" & JOURNALTYPE & "'," & J_JVOUCHERNO & "," & J_JDATE & ", " & PV_ITEMNO & ", " & PV_PONO & _
                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                            ", " & PV_STATUS & ")"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
        'FOR CDJ
        If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Or JOURNALTYPE = "VDJ" Then
            SQL_STATEMENT = "insert into AMIS_CV_Detail " & _
                            "(CV_JTYPE,JTYPE,VoucherNo,itemno,PV_VoucherNo,DocDate,DueDate,Amount,status)" & _
                          " values ('" & JOURNALTYPE & "'," & PV_PONO & "," & J_JVOUCHERNO & ", " & PV_ITEMNO & _
                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                            ", " & PV_STATUS & ")"
            gconDMIS.Execute SQL_STATEMENT
            SQL_STATEMENT = "update AMIS_Journal_HD set PaidStatus = 'N' where VoucherNo = " & PV_MRRNO & " and (Jtype =  " & PV_PONO & ")"    'update by BTT : 09-24-2008
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            ' Start of Cumputing Balance
            'Set rsCheckJournal_HD = New ADODB.Recordset
            'If JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then
            '    Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & J_JVOUCHERNO & " and jtype='" & JOURNALTYPE & "'")    'update by BTT : 09-24-2008
            'Else                                      ' for cdj
            '    Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and jtype=" & PV_PONO & "")    'update by BTT : 09-24-2008
            'End If
            'If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            '    If Null2String(rsCheckJournal_HD!jtype) = "APJ" Then
            '        If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'Y'," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = Balance - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        Else
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'N'," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        End If
            '    End If
            '    If Null2String(rsCheckJournal_HD!jtype) = "VPJ" Then
            '        If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'Y'," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VPJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        Else
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'N'," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VPJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        End If
             '   End If
             '   If Null2String(rsCheckJournal_HD!jtype) = "VDJ" Then
             '       If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
             '           SQL_STATEMENT = "update AMIS_Journal_HD set" & _
             '                         " PaidStatus = 'Y'," & _
             '                         " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
             '                         " [Balance] = [Balance] - " & PV_AMOUNT & _
             '                         " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VDJ'"
             '           gconDMIS.Execute SQL_STATEMENT
             '           NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
             '       Else
             '           SQL_STATEMENT = "update AMIS_Journal_HD set" & _
             '                         " PaidStatus = 'N'," & _
             '                         " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
             '                         " [Balance] = [Balance] - " & PV_AMOUNT & _
             '                         " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VDJ'"
             '           gconDMIS.Execute SQL_STATEMENT
             '           NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
             '       End If
             '   End If
            'End If
        End If
        
        'FOR CRJ
        If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
            SQL_STATEMENT = "insert into AMIS_CRJ_Detail " & _
                            "(J_CLASS,CR_TYPE,VoucherNo,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
                          " values ('" & Setacctcode(cboARTag) & "','" & JOURNALTYPE & "'," & J_JVOUCHERNO & ", " & PV_ITEMNO & _
                            ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                            ", " & PV_STATUS & ")"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            'Update By BTT : 10282008
            'To Classify the type of invoice
            'If GetSJVoucherNo(PV_INVNO, PV_MRRNO) = True Then
                'SQL_STATEMENT = "UPDATE AMIS_CRJ_detail set SJ_voucherno =" & Null2String(SJVoucherno) & ", J_CLASS = '" & Setacctcode(cboARTag) & "' where ID='" & labPVID.Caption & "'"
                'gconDMIS.Execute SQL_STATEMENT
            'End If
            'Commented by BTT - Not needed code because AR/AP have its own process
            'Set rsCheckJournal_HD = New ADODB.Recordset
            'Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
            'If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            '   If N2Str2Zero(rsCheckJournal_HD!INVOICEAMT) <= PV_AMOUNT Then
            '        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                      " ReceiveStatus = 'Y' " & "," & _
            '                      " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                      " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                      " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
            '        gconDMIS.Execute SQL_STATEMENT
            '        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '    Else
            '        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                      " ReceiveStatus = 'N' " & "," & _
            '                      " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                      " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                      " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
            '        gconDMIS.Execute SQL_STATEMENT
            '        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
             '   End If
            
            'Else
            '    Set rsCheckJournal_HD = New ADODB.Recordset
            '    Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
            '    If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            '        If N2Str2Zero(rsCheckJournal_HD!INVOICEAMT) <= PV_AMOUNT Then
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " ReceiveStatus = 'Y' " & "," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        Else
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " ReceiveStatus = 'N' " & "," & _
            '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
            '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
            '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        End If
            '    End If
            'End If
        End If
    Else
    ' End of Balance computation
        'EDIT MODE
        If JOURNALTYPE = "APJ" Then
            SQL_STATEMENT = "update AMIS_PV_Detail set" & _
                          " VoucherNo = " & J_JVOUCHERNO & "," & _
                          " itemno = " & PV_ITEMNO & "," & _
                          " PO_No = " & PV_PONO & "," & _
                          " MRR_No = " & PV_MRRNO & "," & _
                          " INV_No = " & PV_INVNO & "," & _
                          " PROD_No = " & PV_PRODNO & "," & _
                          " Amount = " & PV_AMOUNT & "," & _
                          " status = " & PV_STATUS & _
                          " where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
        If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
            SQL_STATEMENT = "update AMIS_CV_Detail set" & _
                          " VoucherNo = " & J_JVOUCHERNO & "," & _
                          " itemno = " & PV_ITEMNO & "," & _
                          " PV_VoucherNo = " & PV_MRRNO & "," & _
                          " DocDate = " & PV_INVNO & "," & _
                          " DueDate = " & PV_PRODNO & "," & _
                          " Amount = " & PV_AMOUNT & "," & _
                          " status = " & PV_STATUS & _
                          " where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
           ' Set rsCheckJournal_HD = New ADODB.Recordset
           ' Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and (Jtype = " & N2Str2Null(txtPO_No.Text) & ")")
           ' If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
           '     If Null2String(rsCheckJournal_HD!jtype) = "APJ" Then
           '         If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
           '             SQL_STATEMENT = "update AMIS_Journal_HD set" & _
           '                           " PaidStatus = 'Y'," & _
           '                           " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
           '                           " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
           '                           " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
           '             gconDMIS.Execute SQL_STATEMENT
           '             NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        Else
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'N'," & _
            '                          " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
            '                          " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
            '            gconDMIS.Execute SQL_STATEMENT
             '           NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        End If
            '    End If
            '    If Null2String(rsCheckJournal_HD!jtype) = "VPJ" Then
            '        If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'Y'," & _
            '                          " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
            '                          " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VPJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '        Else
            '            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                          " PaidStatus = 'N'," & _
            '                          " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
            '                          " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
            '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VPJ'"
            '            gconDMIS.Execute SQL_STATEMENT
            '        End If
              '  End If
              '  If Null2String(rsCheckJournal_HD!jtype) = "VDJ" Then
              '      If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
              '          SQL_STATEMENT = "update AMIS_Journal_HD set" & _
              '                        " PaidStatus = 'Y'," & _
              '                        " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
              '                        " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
              '                        " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VDJ'"
              '          gconDMIS.Execute SQL_STATEMENT
              '          NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
              '      Else
              '          SQL_STATEMENT = "update AMIS_Journal_HD set" & _
              '                        " PaidStatus = 'N'," & _
              '                        " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
              '                        " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
              '                        " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VDJ'"
              '          gconDMIS.Execute SQL_STATEMENT
              '      End If
              '  End If
            'End If
        End If
        If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
            SQL_STATEMENT = "update AMIS_CRJ_Detail set" & _
                          " VoucherNo = " & J_JVOUCHERNO & "," & _
                          " itemno = " & PV_ITEMNO & "," & _
                          " INVOICETYPE = " & PV_MRRNO & "," & _
                          " INVOICENO = " & PV_INVNO & "," & _
                          " INVOICEDATE = " & PV_PRODNO & "," & _
                          " INVOICEAMOUNT = " & PV_AMOUNT & "," & _
                          " J_CLASS = '" & Setacctcode(cboARTag) & "'," & _
                          " status = " & PV_STATUS & _
                          " where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            'Update By BTT - 10282008
            If GetSJVoucherNo(PV_INVNO, PV_MRRNO) = True Then
                SQL_STATEMENT = "UPDATE AMIS_CRJ_detail set SJ_voucherno =" & N2Str2Null(SJVOUCHERNO) & ",J_CLASS='" & Setacctcode(cboARTag) & "' where ID='" & labPVID & "'"
                gconDMIS.Execute SQL_STATEMENT
            End If
            NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            'SQL_STATEMENT = "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype='SJ'"
            'gconDMIS.Execute SQL_STATEMENT
            'NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            'SQL_STATEMENT = "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype='CSJ'"
            'gconDMIS.Execute SQL_STATEMENT
            'NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            'Set rsCheckJournal_HD = New ADODB.Recordset
            'Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
            'If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            '    If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
            '        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                      " ReceiveStatus = 'Y' " & "," & _
            '                      " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
            '                      " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
            '                      " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
            '        gconDMIS.Execute SQL_STATEMENT
            '        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '    Else
            '        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
            '                      " ReceiveStatus = 'N' " & "," & _
            '                      " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
            '                      " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
            '                      " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
            '        gconDMIS.Execute SQL_STATEMENT
            '        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            '    End If
            'Else
             '   Set rsCheckJournal_HD = New ADODB.Recordset
             '   Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
             '   If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
             '       If N2Str2Zero(rsCheckJournal_HD!BALANCE) <= PV_AMOUNT Then
             '           SQL_STATEMENT = "update AMIS_Journal_HD set" & _
             '                         " ReceiveStatus = 'Y' " & "," & _
             '                         " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
             '                         " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
             '                         " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
             '           gconDMIS.Execute SQL_STATEMENT
             '           NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
             '       Else
             '           SQL_STATEMENT = "update AMIS_Journal_HD set" & _
             '                         " ReceiveStatus = 'N' " & "," & _
             '                         " AmountPaid = (AmountPaid - " & PrevPV_Amount & ") + " & PV_AMOUNT & "," & _
             '                         " [Balance] = (Balance + " & PrevPV_Amount & ") - " & PV_AMOUNT & _
             '                         " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
             '           gconDMIS.Execute SQL_STATEMENT
             '           NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
             '       End If
             '   End If
            'End If
        End If
    End If ' end of ADD



    FillDetails
    rsRefresh
    'On Error Resume Next
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
    If JOURNALTYPE <> "CDJ" Then
        If AddorEdit = "ADD" Then
            cmdPV_Entry_Click
        Else
            cmdPVCancel.Value = True
        End If
    Else
        SendToBackPV
        Dim J_VOUCHERNO, J_ACCT_CODE, J_ACCT_NAME, J_JTYPE, J_JNO As String
        Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET   As Double
        Dim J_STATUS, J_JITEMNO                        As String
        Dim rsDet As New ADODB.Recordset
        'Dim TOTAL_DEBIT, TOTAL_CREDIT As Double
        'TOTAL_DEBIT = 0: TOTAL_CREDIT = 0
        Set rsDet = gconDMIS.Execute("Select Voucherno from AMIS_journal_det where voucherno ='" & txtVoucherNo.Text & "' and jtype = 'CDJ'") ' to check if there is already account entry
        If rsDet.EOF And rsDet.BOF Then
            If CDJ_IS_FROM_AP = True Then
                gconDMIS.Execute ("Delete from AMIS_Journal_Det where jtype = 'CDJ' and voucherno = " & N2Str2Null(txtVoucherNo.Text))

                J_JDATE = N2Str2Null(txtJDate.Text)
                J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
                J_JTYPE = "'CDJ'"
                J_JNO = N2Str2Null(txtJNo.Text)
    
                J_JITEMNO = "'0001'"
                'Case to case Company_code
                If COMPANY_CODE = "HAI" Or COMPANY_CODE = "HBK" Then
                        'Update by BTT - 06242008
                        If ISVPJ = True Then
                            J_ACCT_CODE = CDJ_AP
                            J_ACCT_NAME = N2Str2Null(Setacctname(J_ACCT_CODE))
                        Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("AP"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("AP")))
                        End If
                    ElseIf COMPANY_CODE = "HGC" Then
                        ' Update By BTT : 1/27/2009 : To get the acctual Liabilities account in the AP transaction
                        Set rsDet = gconDMIS.Execute("Select Acct_code,acct_name from AMIS_journal_det where voucherno =" & PV_MRRNO & " and jtype = 'APJ' and (left(acct_code,5)='21-01' or left(acct_code,5)='21-02') ")
                        If Not rsDet.EOF And Not rsDet.BOF Then
                            J_ACCT_CODE = N2Str2Null(rsDet!ACCT_CODE)
                            J_ACCT_NAME = N2Str2Null(rsDet!acct_Name)
                            Else
                            J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                            J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                        End If
                    Else
                        J_ACCT_CODE = N2Str2Null(ReturnAP_AccountCode("GENERAL"))
                        J_ACCT_NAME = N2Str2Null(Setacctname(ReturnAP_AccountCode("GENERAL")))
                    End If
                    J_DEBIT = NumericVal(txtTotalPV_Amount.Text)
                    J_CREDIT = 0
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                'TOTAL_DEBIT = TOTAL_DEBIT + J_DEBIT
                    If J_ACCT_CODE = "" Then
                        MsgBox "Invalid Entry..Please verifry the entry", vbExclamation, "WARNING"
                        Exit Sub
                    End If
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & "," & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
    
                    J_JITEMNO = "'0002'"
                    J_ACCT_CODE = CDJ_CIB
                    J_ACCT_NAME = N2Str2Null(Setacctname(CDJ_CIB))
                    J_DEBIT = 0
                    J_CREDIT = NumericVal(txtTotalPV_Amount.Text)
                    J_TAX = 0
                    J_GROSS = 0
                    J_NET = 0
                    J_STATUS = "'N'"
                    'TOTAL_CREDIT = TOTAL_CREDIT + J_CREDIT
                    gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                                 "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                               " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                                 ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                                 ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
                    FillDetails
                    If JOURNALTYPE = "VDJ" Then
                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                  " credit = " & TOTCREDIT & "," & _
                                  " tax = " & TOTTAX & "," & _
                                  " outbalance = " & OUTBALANCE & _
                                  " where id = " & labID.Caption
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
                    ElseIf JOURNALTYPE = "VCJ" Then
                    SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                  " debit = " & TOTDEBIT & "," & _
                                  " tax = " & TOTTAX & "," & _
                                  " outbalance = " & OUTBALANCE & _
                                  " where id = " & labID.Caption
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
                    Else
                        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                                  " debit = " & TOTDEBIT & "," & _
                                  " credit = " & TOTCREDIT & "," & _
                                  " tax = " & TOTTAX & "," & _
                                  " outbalance = " & OUTBALANCE & _
                                  " where id = " & labID.Caption
                        gconDMIS.Execute SQL_STATEMENT
                        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
                    End If
                    StoreMemvars
            End If
        End If ' end of cheking the accounting entry
    End If


    '   AXP 1062008 FOR CHECKING: FOR KIM CHECKING
    '   SHOULD ALSO BE INCLUDED IN DELETE BUTTON TOO
    '   UPDATE AMIS JOURNAL HEADER BALANCE AFTER PAYMENT HAS BEEN DONE IN DISBURSMENT

       Dim RS_TOTALPAY As ADODB.Recordset
       If JOURNALTYPE = "CDJ" And txtPO_No = "APJ" Then
       Set RS_TOTALPAY = gconDMIS.Execute("SELECT isnull(SUM(AMOUNT),0)  AS TOTALPAYMENTS  FROM AMIS_CV_Detail WHERE CV_JTYPE='CDJ' AND JTYPE='VPJ' AND PV_VOUCHERNO=" & N2Str2Null(txtMRR_No))
         If Not RS_TOTALPAY.EOF Or Not RS_TOTALPAY.BOF Then
            gconDMIS.Execute ("update amis_journal_hd  set  BALANCE=AMOUNTTOPAY-" & RS_TOTALPAY!TOTALPAYMENTS & " WHERE JTYPE='APJ' AND VOUCHERNO= " & N2Str2Null(txtMRR_No))
         End If
        ElseIf JOURNALTYPE = "CDJ" And LTrim(RTrim(txtPO_No)) = "VPJ" Then
        Set RS_TOTALPAY = gconDMIS.Execute("SELECT isnull(SUM(AMOUNT),0)  AS TOTALPAYMENTS  FROM AMIS_CV_Detail WHERE CV_JTYPE='CDJ' AND JTYPE='VPJ' AND PV_VOUCHERNO=" & N2Str2Null(txtMRR_No))
             If Not RS_TOTALPAY.EOF Or Not RS_TOTALPAY.BOF Then
            gconDMIS.Execute ("update amis_journal_hd  set  BALANCE=AMOUNTTOPAY-" & RS_TOTALPAY!TOTALPAYMENTS & " WHERE JTYPE='VPJ' AND VOUCHERNO= " & N2Str2Null(txtMRR_No))
         End If
        End If
    
    JournalTAB.TabEnabled(0) = True
    Picture1.Enabled = True
    cboARTag.BackColor = &H80000005
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub
Sub FillDetails()
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0: TotalARAmountToPay = 0
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE: TOTAL_AP_AMOUNT = 0: TOTALAPAMOUNTTOPAY = 0: PREVPV_AMOUNT = 0
    Dim J_ITemNo, PV_ITEMNO                            As Integer
    If JOURNALTYPE <> "GJ" And JOURNALTYPE <> "OPB" And JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "PDJ" And JOURNALTYPE <> "CLO" Then
        lstDetails.Sorted = False: lstDetails.ListItems.Clear
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " and jtype = '" & JOURNALTYPE & "' order by jitemno asc")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            Screen.MousePointer = 11
            rsJournal_Det.MoveFirst: TOTAL_AR_AMOUNT = 0
            Do While Not rsJournal_Det.EOF
                kcnt = kcnt + 1
                If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
                lstDetails.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
                lstDetails.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!ACCT_CODE)
                lstDetails.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
                lstDetails.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
                If Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "11-02" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "11-03" Then
                    TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + N2Str2Zero(rsJournal_Det!CREDIT)
                    TotalARAmountToPay = TotalARAmountToPay + N2Str2Zero(rsJournal_Det!DEBIT)
                End If
                If Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-01" Or Left(Null2String(rsJournal_Det!ACCT_CODE), 5) = "21-02" Then
                    TOTAL_AP_AMOUNT = TOTAL_AP_AMOUNT + N2Str2Zero(rsJournal_Det!CREDIT)
                    TOTALAPAMOUNTTOPAY = TOTALAPAMOUNTTOPAY + N2Str2Zero(rsJournal_Det!DEBIT)
                End If
                lstDetails.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
                lstDetails.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
                If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then COMP_SJ_OUTPUT_TAX = 0
                TOTDEBIT = TOTDEBIT + Round(NumericVal(N2Str2Zero(rsJournal_Det!DEBIT)), 2)
                TOTCREDIT = TOTCREDIT + Round(NumericVal(N2Str2Zero(rsJournal_Det!CREDIT)), 2)
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
            If JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Or JOURNALTYPE = "CSJ" Or JOURNALTYPE = "CCM" Then
            Else
                cmdPost.Enabled = False
            End If
        End If

        'DISPLAY JOURNAL DETAILS
        JCNT = 0
        TOTALPVAMOUNT = 0
        txtTotalPV_Amount.Text = ZERO
        If JOURNALTYPE = "APJ" Then
            lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
            Set RSPV_DETAIL = New ADODB.Recordset
            Set RSPV_DETAIL = gconDMIS.Execute("select * from AMIS_PV_Detail where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
            If Not RSPV_DETAIL.EOF And Not RSPV_DETAIL.BOF Then
                Screen.MousePointer = 11
                RSPV_DETAIL.MoveFirst: TOTALPVAMOUNT = 0
                Do While Not RSPV_DETAIL.EOF
                    JCNT = JCNT + 1
                    If Null2String(RSPV_DETAIL!ItemNo) = "" Then PV_ITEMNO = JCNT Else PV_ITEMNO = Null2String(RSPV_DETAIL!ItemNo)
                    lstPV_Detail.ListItems.Add JCNT, , Format(PV_ITEMNO, "0000")
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 1, , Null2String(RSPV_DETAIL!po_no)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 2, , Null2String(RSPV_DETAIL!MRR_No)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 3, , Null2String(RSPV_DETAIL!INV_NO)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 4, , Null2String(RSPV_DETAIL!Prod_No)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(RSPV_DETAIL!amount))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 6, , RSPV_DETAIL!ID
                    TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(RSPV_DETAIL!amount))
                    RSPV_DETAIL.MoveNext
                Loop
                lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
                txtTotalPV_Amount.Text = TOTALPVAMOUNT
                Screen.MousePointer = 0
            End If
        End If
        ' DISPLAY DETAIL (F4)
        If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Or JOURNALTYPE = "VDJ" Then
            lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
            Set rsCV_Detail = New ADODB.Recordset
            Set rsCV_Detail = gconDMIS.Execute("select * from AMIS_CV_Detail where CV_JTYPE = '" & JOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
            If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
                Screen.MousePointer = 11
                If JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then    ' Update by BTT 11252008
                    lstPV_Detail.ColumnHeaders(3).Text = "Doc Date"
                    lstPV_Detail.ColumnHeaders(4).Text = "Due Date"
                End If
                rsCV_Detail.MoveFirst: TOTALPVAMOUNT = 0
                Do While Not rsCV_Detail.EOF
                    JCNT = JCNT + 1
                    If Null2String(rsCV_Detail!ItemNo) = "" Then PV_ITEMNO = JCNT Else PV_ITEMNO = Null2String(rsCV_Detail!ItemNo)
                    lstPV_Detail.ListItems.Add JCNT, , Format(PV_ITEMNO, "0000")
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 1, , Null2String(rsCV_Detail!pv_voucherno)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 2, , Null2String(rsCV_Detail!docdate)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 3, , Null2String(rsCV_Detail!duedate)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsCV_Detail!amount))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsCV_Detail!amount))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 6, , rsCV_Detail!ID
                    TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsCV_Detail!amount))
                    rsCV_Detail.MoveNext
                Loop
                lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
                txtTotalPV_Amount.Text = TOTALPVAMOUNT
                Screen.MousePointer = 0
            End If
        End If
        If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
            lstPV_Detail.ColumnHeaders(2).Width = lstPV_Detail.ColumnHeaders(2).Width + lstPV_Detail.ColumnHeaders(5).Width
            lstPV_Detail.ColumnHeaders(5).Width = 1
            lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("select * from AMIS_CRJ_Detail where CR_TYPE = '" & JOURNALTYPE & "' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                Screen.MousePointer = 11
                rsCRJ_Detail.MoveFirst: TOTALPVAMOUNT = 0
                Do While Not rsCRJ_Detail.EOF
                    JCNT = JCNT + 1
                    If Null2String(rsCRJ_Detail!ItemNo) = "" Then PV_ITEMNO = JCNT Else PV_ITEMNO = Null2String(rsCRJ_Detail!ItemNo)
                    lstPV_Detail.ListItems.Add JCNT, , Format(PV_ITEMNO, "0000")
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 1, , SetInvType(Null2String(rsCRJ_Detail!InvoiceType))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 2, , Null2String(rsCRJ_Detail!INVOICENO)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 3, , Null2String(rsCRJ_Detail!invoicedate)
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsCRJ_Detail!invoiceamount))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsCRJ_Detail!invoiceamount))
                    lstPV_Detail.ListItems(JCNT).ListSubItems.Add 6, , rsCRJ_Detail!ID
                    TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsCRJ_Detail!invoiceamount))
                    rsCRJ_Detail.MoveNext
                Loop
                lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
                txtTotalPV_Amount.Text = TOTALPVAMOUNT
                Screen.MousePointer = 0
            End If
            'CHECK IF JOURNAL USES AR ACCOUNT
            If TOTAL_AR_AMOUNT > 0 Then
                If NumericVal(TOTAL_AR_AMOUNT) <> NumericVal(TOTALPVAMOUNT) Then
                    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CRJ" Then
                        picChat.Visible = True
                    End If
                Else
                    picChat.Visible = False
                End If
            Else
                picChat.Visible = False
            End If
        Else
            picChat.Visible = False
        End If
    Else
        txtGJTotDebit.Text = ZERO: txtGJTotCredit.Text = ZERO: txtGJOutBalance.Text = ZERO
        lstGJ.Sorted = False: lstGJ.ListItems.Clear
        Set rsJournal_Det = New ADODB.Recordset
        Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & JOURNALTYPE & "' order by jitemno asc")
        If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
            Screen.MousePointer = 11
            rsJournal_Det.MoveFirst
            Do While Not rsJournal_Det.EOF
                kcnt = kcnt + 1
                If Null2String(rsJournal_Det!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJournal_Det!jitemno)
                lstGJ.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
                lstGJ.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJournal_Det!ACCT_CODE)
                lstGJ.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJournal_Det!acct_Name)
                lstGJ.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
                lstGJ.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
                lstGJ.ListItems(kcnt).ListSubItems.Add 5, , rsJournal_Det!ID
                TOTDEBIT = TOTDEBIT + NumericVal(N2Str2Zero(rsJournal_Det!DEBIT))
                TOTCREDIT = TOTCREDIT + NumericVal(N2Str2Zero(rsJournal_Det!CREDIT))
                TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJournal_Det!tax))
                rsJournal_Det.MoveNext
            Loop
            lstGJ.Sorted = True: lstGJ.Refresh
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

            Screen.MousePointer = 0
        End If
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccount2                                As ADODB.Recordset
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
    Dim rsTemplate_Header                              As ADODB.Recordset
    lstTemplates.Enabled = False
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & JOURNALTYPE & "' AND description like '" & XXX & "%' order by description asc")
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
    Dim rsTemplate_Header                              As ADODB.Recordset
    lstTemplates.Enabled = False
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & JOURNALTYPE & "' order by description asc")
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
    rsJournal_hd.Bookmark = rsFind(rsJournal_hd.Clone, "jno", Format(DDD, "000000")).Bookmark
    StoreMemvars
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

    If JOURNALTYPE = "APJ" Or JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then
        Set RSATC = New ADODB.Recordset
        Set RSATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
        If Not RSATC.EOF And Not RSATC.BOF Then
            'Combo_Loadval cboATC, rsATC
            RSATC.MoveFirst: cboATC.AddItem ""
            Do While Not RSATC.EOF
                cboATC.AddItem Null2String(RSATC!ATC)
                RSATC.MoveNext
            Loop
        End If
        Set RSATC = Nothing

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

    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CSJ" Or JOURNALTYPE = "CCM" Then
        Set rsBanks = New ADODB.Recordset
        Set rsBanks = gconDMIS.Execute("select bankname from CMIS_Banks order by bankname asc")
        If Not rsBanks.EOF And Not rsBanks.BOF Then
            Combo_Loadval cboBankName2, rsBanks
        End If
        Set rsBanks = Nothing

        Set rsPayTerm = New ADODB.Recordset
        Set rsPayTerm = gconDMIS.Execute("select pay_Code from ALL_PayTerm order by pay_desc asc")
        If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
            Combo_Loadval cboPayTerm2, rsPayTerm
        End If
        Set rsPayTerm = Nothing
        InitCustomer
    End If
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        Set rsInvoiceType = New ADODB.Recordset
        rsInvoiceType.Open "select InvType from ALL_InvoiceType order by InvType asc", gconDMIS
        If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
            rsInvoiceType.MoveFirst
            cboInvoiceType.Clear
            Do While Not rsInvoiceType.EOF
                cboInvoiceType.AddItem Null2String(rsInvoiceType!INVTYPE)
                rsInvoiceType.MoveNext
            Loop
        End If
    End If
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        cboInvoiceType.Clear
        cboInvoiceType.AddItem "CASH"
        cboInvoiceType.AddItem "CARD"
        cboInvoiceType.AddItem "CHECK"
    End If
    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
        Set RSATC = New ADODB.Recordset
        Set RSATC = gconDMIS.Execute("Select ATC from AMIS_ATC order by ATC asc")
        If Not RSATC.EOF And Not RSATC.BOF Then
            'Combo_Loadval cboATC, rsATC
            RSATC.MoveFirst: cboATC2.Clear: cboATC2.AddItem ""
            Do While Not RSATC.EOF
                cboATC2.AddItem Null2String(RSATC!ATC)
                RSATC.MoveNext
            Loop
        End If
        Set RSATC = Nothing

        Set rsVENDOR = New ADODB.Recordset
        Set rsVENDOR = gconDMIS.Execute("select nameofvendor from ALL_Vendor order by nameofvendor asc")
        If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
            Combo_Loadval cboJVSupCust, rsVENDOR
        End If
        Set rsVENDOR = Nothing
    End If
    Dim rsAR_Accounts                                  As ADODB.Recordset


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
    Set rsCUSTOMER = New ADODB.Recordset
    'Set rsCustomer = gconDMIS.Execute("select acctname from ALL_CUSTMASTER_AMIS where (acctname <> '' and acctname is not null) order by acctname asc")
    Set rsCUSTOMER = gconDMIS.Execute("select custname from ALL_CUSTMASTER_AMIS where (custname <> '' and custname is not null) order by custname asc")
    If Not rsCUSTOMER.EOF And Not rsCUSTOMER.BOF Then
        Combo_Loadval cboCustName, rsCUSTOMER
    End If
    Set rsCUSTOMER = Nothing
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

Sub InitGrid()
    If JOURNALTYPE = "CDJ" Then
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
    If JOURNALTYPE = "APJ" Then
        cboATC.ListIndex = 0
        txtRATE.Text = "0"
        txtTaxBase.Text = ZERO
    End If
End Sub

Sub InitMemVars()
    Dim rsJournal_HDDup                                As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & JOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
    txtJDate.Text = LOGDATE:

    CDJ_CIB = ""
    CDJ_AP = ""

    'Accounts Payable Module'
    txtCode.Text = ""
    txtAddress.Caption = "":
    txtInvoiceDate.Text = LOGDATE
    txtDueDate.Text = LOGDATE:
    txtBankCode.Text = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    '---------------------------'
    'Cash Disbursement Module'
    txtCheckNo.Text = "": txtCheckDate.Text = LOGDATE: txtPayCode.Text = ""
    cboNameofVendor.Text = ""
    txtTotDebit.Text = ZERO: txtTotCredit.Text = ZERO
    txtAmountToPay.Text = ZERO: txtOutBalance.Text = ZERO
    txtCheckAmt.Text = ZERO
    txtParticulars2.Locked = False
    txtParticulars.Text = "Pls Type Your Message Here!"
    txtParticulars2.Text = "Pls Type Your Message Here!"
    '---------------------------'
    'Accounts Receivable Module'
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
    InitGrid
    SendToBack
End Sub

Sub InitPV_Detail()
    txtPVItemNo.Text = Format(JCNT + 1, "0000")
    txtMRR_No.Text = ""
    If JOURNALTYPE = "APJ" Then
        txtPO_No.Text = "": txtINV_No.Text = "": txtProd_No.Text = ""
        txtPVAmount.Text = ZERO
    ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
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
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                  As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
    Dim J_STATUS, J_JITEMNO                            As String
    Dim rsTemplate_Details                             As ADODB.Recordset
    Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("Select * from AMIS_Template_Details Where TemplateCode = " & XXX)
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        rsTemplate_Details.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsTemplate_Details.EOF
            kcnt = kcnt + 1
            J_JDATE = N2Date2Null(txtJDate.Text)
            J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
            J_JTYPE = N2Str2Null(JOURNALTYPE)
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
        StoreMemvars
        FillDetails
        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
            lstDetails.SetFocus
        End If

        Screen.MousePointer = 0
    End If
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
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
    Else
        cboAcct_Code.Text = labAccountCode.Caption
        'If JOURNALTYPE = "APJ" Or JOURNALTYPE = "SJ" Then
        '   txtGrossAmt.SetFocus
        If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
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
    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
        If cboGJAccountNo.Text <> "" Then
            If SetAcctType(cboGJAccountNo.Text) = "C" Then
                On Error Resume Next
                txtGJCredit.SetFocus
            Else
                On Error Resume Next
                txtGJDebit.SetFocus
            End If
        End If
    Else
        If JOURNALTYPE = "APJ" Or JOURNALTYPE = "SJ" Then
            txtGrossAmt.SetFocus
        Else
            If cboAcct_Code.Text <> "" Then
                If SetAcctType(cboAcct_Code.Text) = "C" Then
                    txtCredit.SetFocus
                Else
                    txtDebit.SetFocus
                End If
            End If
        End If
    End If
End Sub

Sub rsRefresh()
    If JOURNALTYPE = "APJ" Then Me.Caption = "ACCOUNTS PAYABLE JOURNAL ENTRY"
    If JOURNALTYPE = "CDJ" Then Me.Caption = "CASH DISBURSEMENT JOURNAL ENTRY"
    If JOURNALTYPE = "SJ" Then Me.Caption = "SALES JOURNAL ENTRY"
    If JOURNALTYPE = "CRJ" Then Me.Caption = "CASH RECEIPTS JOURNAL ENTRY"
    If JOURNALTYPE = "GJ" Then Me.Caption = "GENERAL JOURNAL DATA ENTRY"
    If JOURNALTYPE = "ADJ" Then Me.Caption = "CLIENT ADJUSTING JOURNAL ENTRIES"
    If JOURNALTYPE = "PDJ" Then Me.Caption = "PROPOSED ADJUSTING JOURNAL ENTRIES"
    If JOURNALTYPE = "CLO" Then Me.Caption = "CLOSING ENTRIES"
    Set rsJournal_hd = New ADODB.Recordset
    rsJournal_hd.Open "select * from AMIS_Journal_HD where jtype = '" & JOURNALTYPE & "' order by ID asc", gconDMIS, adOpenKeyset
End Sub

Sub SearchVoucherNo(XXX As String)
    If XXX <> "" Then
        On Error GoTo Errorcode
        rsJournal_hd.Bookmark = rsFind(rsJournal_hd.Clone, "voucherno", XXX).Bookmark
    End If
    StoreMemvars
    Exit Sub

Errorcode:
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
    cmdGJEntry.ZOrder 1
    cmdGJEntry.Visible = False
    picGJEntry.ZOrder 1
    picGJEntry.Visible = False
    picGJEntry.Enabled = False
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

Sub StoreMemvars()
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        labID.Caption = rsJournal_hd!ID
        txtJNo.Text = Null2String(rsJournal_hd!JNo)
        txtVoucherNo.Text = Null2String(rsJournal_hd!VOUCHERNO)
        txtJDate.Text = Format(Null2String(rsJournal_hd!jdate), "DD-MMM-YY")
        txtInvoiceDate.Text = Format(Null2String(rsJournal_hd!invoicedate), "DD-MMM-YY")
        txtDueDate.Text = Format(Null2String(rsJournal_hd!duedate), "DD-MMM-YY")
        txtPayCode.Text = Null2String(rsJournal_hd!paytype)
        txtTerms.Text = Null2String(rsJournal_hd!TERMS)
        If SetPayDesc(Null2String(rsJournal_hd!paytype)) = "" Then
            cboPayType.ListIndex = -1
        Else
            cboPayType.Text = SetPayDesc(Null2String(rsJournal_hd!paytype))
        End If
        If JOURNALTYPE = "APJ" Or JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then
            txtCode.Text = Null2String(rsJournal_hd!VendorCode)
            cboNameofVendor.Text = SetVendorName(txtCode.Text)
            CURRENT_VENDORCODE = Null2String(rsJournal_hd!VendorCode)
            txtAddress.Caption = SetVendorAddress(txtCode.Text)
            cboBankName.Text = SetBankName(Null2String(rsJournal_hd!bankcode))
            txtCheckAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!AmountPaid))
            APJINVOICENO = Null2String(rsJournal_hd!INVOICENO)
            APJINVOICETYPE = Null2String(rsJournal_hd!InvoiceType)
        End If
        If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CSJ" Or JOURNALTYPE = "CCM" Then
            CURRENT_CUSCODE = Null2String(rsJournal_hd!CustomerCode)
            txtCustCode.Text = Null2String(rsJournal_hd!CustomerCode)
            cboCustName.Text = SetCustomerName(Null2String(rsJournal_hd!CustomerCode))
            If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
                cboInvoiceType.Text = SetInvType(Null2String(rsJournal_hd!InvoiceType))
                If txtTerms.Text = "CHG" Then
                    'SHOULD APPEAR CUSTOMER CREDIT TERMS
                    cboPayTerm2.Text = SetCustomerCreditTerm(Null2String(rsJournal_hd!CustomerCode))
                Else
                    cboPayTerm2.Text = ""
                End If
            Else
                cboInvoiceType.Text = Null2String(rsJournal_hd!paytype)
            End If
            'SHOW DEALER FOR SERVICE INVOICE TRANSACTIONS
            If cboInvoiceType.Text = "SI" Then
                txtDealer.Text = StoreDealerCode(Null2String(rsJournal_hd!INVOICENO))
            Else
                txtDealer.Text = ""
            End If
            If Left(Null2String(rsJournal_hd!INVOICENO), 2) = "NV" Then
                chkNonVat.Value = 1
                txtInvoiceNo.Text = Right(Null2String(rsJournal_hd!INVOICENO), 6)
            Else
                chkNonVat.Value = 0
                txtInvoiceNo.Text = Null2String(rsJournal_hd!INVOICENO)
            End If
            txtInvoiceDate2.Text = Null2String(rsJournal_hd!invoicedate)
            txtInvoiceAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!InvoiceAmt))
            cboBankName2.Text = SetBankName(Null2String(rsJournal_hd!bankcode))
            txtRefNo.Text = Null2String(rsJournal_hd!refno)
            txtRefDate.Text = Null2String(rsJournal_hd!RefDate)
        End If
        If JOURNALTYPE = "APJ" Then
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where PV_VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                picRefCDJ.ZOrder 0: picRefCDJ.Visible = True
                RefCDJ.Caption = "Ref CDJ# " & Null2String(rsCRJ_Detail!VOUCHERNO)
            Else
                picRefCDJ.ZOrder 1: picRefCDJ.Visible = False
                RefCDJ.Caption = ""
            End If
        End If
        If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
            Set rsCRJ_Detail = New ADODB.Recordset
            Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail Where InvoiceType = " & N2Str2Null(rsJournal_hd!InvoiceType) & " And InvoiceNo = " & N2Str2Null(txtInvoiceNo.Text) & " And InvoiceAmount = " & NumericVal(txtInvoiceAmt.Text))
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                RefCRJ.BorderStyle = 1: RefCRJ.Caption = "Ref CRJ# " & Null2String(rsCRJ_Detail!VOUCHERNO)
            Else
                RefCRJ.BorderStyle = 0: RefCRJ.Caption = ""
            End If
        End If
        If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then txtParticulars2.Locked = True
        txtBankCode.Text = Null2String(rsJournal_hd!bankcode)
        txtCheckNo.Text = Null2String(rsJournal_hd!CheckNo)
        txtCheckDate.Text = Null2String(rsJournal_hd!CheckDate)
        txtParticulars.Text = Null2String(rsJournal_hd!Remarks)
        txtParticulars2.Text = Null2String(rsJournal_hd!Remarks)
        txtTotDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!DEBIT))
        txtTotCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!CREDIT))
        txtOutBalance.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!OUTBALANCE))
        txtAmountToPay.Text = ToDoubleNumber(N2Str2Zero(rsJournal_hd!amounttopay))
        txtRemarks.Text = Null2String(rsJournal_hd!Remarks)
        txtRemarks2.Text = Null2String(rsJournal_hd!Remarks)
        If Null2String(rsJournal_hd!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "*** CANCELLED ***"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
    rsJournal_hd.Find "VoucherNo = " & N2Str2Null(XXX)
    StoreMemvars
End Sub

Private Sub cboAcct_Code_Change()
    Dim DEALER_ITW_COMPENSATION                        As String
    Dim DEALER_ITW_EXPANDED                            As String
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
        If JOURNALTYPE = "CLO" Then
            Dim rsJournal_HDDet                        As ADODB.Recordset
            Set rsJournal_HDDet = New ADODB.Recordset
            rsJournal_HDDet.Open "select SUM(DEBIT) AS TOTAL_DEBIT,SUM(CREDIT) AS TOTAL_CREDIT from AMIS_vw_vLEDGER where Jdate <= '" & txtJDate.Text & "' and Acct_Code = '" & cboAcct_Code.Text & "'", gconDMIS
            If Not rsJournal_HDDet.EOF And Not rsJournal_HDDet.BOF Then
                If N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT) > 0 Then
                    txtGJDebit.Text = ZERO
                    txtGJCredit.Text = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                Else
                    txtGJDebit.Text = Abs(N2Str2Zero(rsJournal_HDDet!TOTAL_DEBIT) - N2Str2Zero(rsJournal_HDDet!TOTAL_CREDIT))
                    txtGJCredit.Text = ZERO
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
    Set RSATC = New ADODB.Recordset
    Set RSATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC.Text))
    If Not RSATC.EOF And Not RSATC.BOF Then
        txtRATE.Text = N2Str2Zero(RSATC!Rate)
        If NumericVal(txtRATE.Text) > 0 Then
            txtCredit.Text = Round(NumericVal(txtTaxBase.Text) * (NumericVal(txtRATE.Text) / 100), 2)
        End If
    End If
    Set RSATC = Nothing
End Sub

Private Sub cboATC2_Click()
    Set RSATC = New ADODB.Recordset
    Set RSATC = gconDMIS.Execute("Select * from AMIS_ATC WHERE ATC = " & N2Str2Null(cboATC2.Text))
    If Not RSATC.EOF And Not RSATC.BOF Then
        txtRATE2.Text = N2Str2Zero(RSATC!Rate)
    End If
    Set RSATC = Nothing
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
    Dim DEALER_ITW_COMPENSATION                        As String
    Dim DEALER_ITW_EXPANDED                            As String
    txtGJAccountName.Text = Setacctname(cboGJAccountNo.Text)
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
    If cboGJAccountNo.Text = DEALER_ITW_COMPENSATION Or cboGJAccountNo.Text = DEALER_ITW_EXPANDED Then
        fraATC2.Visible = True: labATC.Visible = True: cboJVSupCust.Visible = True
        On Error Resume Next
        cboATC2.SetFocus
    Else
        fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False
    End If
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
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    Dim rsProfile                                      As ADODB.Recordset
    Dim AccountingMonth, AccountingYear                As Integer
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        AccountingMonth = rsProfile!PERIODMONTH
        AccountingYear = rsProfile!PERIODYEAR
    End If
    'TEMPORARY DISABLED BY FML - 2/15/2008 2:45AM
    '    Dim rsDetails                                      As ADODB.Recordset
    '    Set rsDetails = New ADODB.Recordset
    '    Set rsDetails = gconDMIS.Execute("Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit, VoucherNo from AMIS_Journal_Det Where jtype = '" & JOURNALTYPE & "' and Month(Jdate) = " & AccountingMonth & " and Year(Jdate) = " & AccountingYear & " and Status <> 'C' group by VoucherNo order by VoucherNo asc")
    '    Dim SQL
    '    SQL = "Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit, VoucherNo from AMIS_Journal_Det Where jtype = '" & JOURNALTYPE & "' and Month(Jdate) = " & AccountingMonth & " and Year(Jdate) = " & AccountingYear & " and Status <> 'C' group by VoucherNo order by VoucherNo asc"
    '    If Not rsDetails.EOF And Not rsDetails.EOF Then
    '        Screen.MousePointer = 11
    '        Do While Not rsDetails.EOF
    '            If Round(rsDetails!TotalDebit, 2) <> Round(rsDetails!Totalcredit, 2) Then
    '                Screen.MousePointer = 0
    '                MsgBox "TOTAL DEBIT: [" & Round(rsDetails!TotalDebit, 2) & "] TOTAL CREDIT: [" & Round(rsDetails!Totalcredit, 2) & "]" & vbCrLf & _
                     '                       "Warning: " & JOURNALTYPE & "-" & rsDetails!vOUCHERNO & " is still not balance or has zero details" & vbCrLf & _
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
    SendToBack
    InitMemVars
    FillDetails

    lstDetails.Enabled = False
    On Error Resume Next
    'txtVoucherNo.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdAddAccount_Click()
    Screen.MousePointer = 11
    REFRESH_ACCOUNT = True
    frmAMISFILESChartOfAccount.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdAddJournal_Click()
    If JOURNALTYPE = "CDJ" Then
        If DIRECTDISBURSEMENTVOUCHERNO <> txtVoucherNo.Text Then
            If MsgBox("Add from Accounts Payable?", vbQuestion + vbYesNo, "Disbursement for Purchases") = vbYes Then
                If MsgBox("Warning: This Disbursement will have the default entry of AP and Cash In Bank, Continue?", vbQuestion + vbYesNo) = vbYes Then
                    SendToBackPV
                    BringToFrontPV
                    AddorEdit = "ADD"
                    cmdPVDelete.Visible = False
                    InitPV_Detail
                    CDJ_IS_FROM_AP = True
                    frmAMISSearchAPJ2.Show vbModal
                    JournalTAB.Tab = 1
                    'cmdPVSave_Click
                Else
                    GoTo CDJ_ISDirectDisbursement
                End If
            Else
                GoTo CDJ_ISDirectDisbursement
            End If
        Else
            SendToBack
            cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
            fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
            fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
            AddorEdit = "ADD"
            InitJournal
            On Error Resume Next
            cboAcct_Code.SetFocus
        End If
    Else
        SendToBack
        cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
        fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
        fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
        AddorEdit = "ADD"
        InitJournal
        On Error Resume Next
        cboAcct_Code.SetFocus
    End If
    Exit Sub

CDJ_ISDirectDisbursement:
    CDJ_IS_FROM_AP = False
    DIRECTDISBURSEMENTVOUCHERNO = txtVoucherNo.Text
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
    StoreMemvars
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdCancelCO_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_CancelEntry", LocalAcess) = False Then Exit Sub
    '    If CheckIfBookIsOpen(JOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If
    If MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Cancel Journal") = vbYes Then
        Screen.MousePointer = 11
        If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
            With FrmCancelTransaction
                .lblTransaction_type = JOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
        End If
        If JOURNALTYPE = "GJ" Then
            With FrmCancelTransaction
                .lblTransaction_type = JOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
        End If
        Screen.MousePointer = 0
        ' UPDATE DUE TO NEW AUDIT : BTT 08282008
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'C' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        SQL_STATEMENT = "update AMIS_Journal_Det set status = 'C' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        If JOURNALTYPE = "APJ" Then
            'Update By BTT 06282008
            With FrmCancelTransaction
                .lblTransaction_type = JOURNALTYPE
                .LblTransactionNo = txtVoucherNo.Text
                FrmCancelTransaction.Show
                If CANCEL_ANS = "NO" Then Exit Sub
                Screen.MousePointer = 0
            End With
            SQL_STATEMENT = "update AMIS_PV_Detail set status = 'C' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
        If JOURNALTYPE = "CDJ" Then
            'Update By BTT 06282008

            With FrmCancelTransaction
                .lblTransaction_type = JOURNALTYPE
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
                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            End If
        End If
        If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
            'Update By BTT 06282008
            With FrmCancelTransaction
                .lblTransaction_type = JOURNALTYPE
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
                NEW_LogAudit "C", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            End If

        End If
        rsRefresh
        rsJournal_hd.Find "id = " & labID.Caption
        StoreMemvars
        Screen.MousePointer = 0
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub

    AddorEdit = "EDIT"
    PrevJType = UCase(JOURNALTYPE)
    PrevJNo = Format(txtJNo.Text, "000000")
    lstDetails.Enabled = False
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_hd!ID
    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then txtParticulars2.Locked = False
    On Error Resume Next
    'txtVoucherNo.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFind_Click()
    On Error GoTo Errorcode:

    If JOURNALTYPE = "APJ" Then
        frmAMISSearchAPJ.Show vbModal
    ElseIf JOURNALTYPE = "VDJ" Then
        frmAMISSearchVDJ.Show vbModal
    ElseIf JOURNALTYPE = "CDJ" Then
        frmAMISSearchCDJ.Show vbModal
    ElseIf JOURNALTYPE = "VCJ" Then
        frmAMISSearchVCJ.Show vbModal
    ElseIf JOURNALTYPE = "SJ" Then
        frmAMISSearchSJ.Show vbModal
    ElseIf JOURNALTYPE = "CSJ" Then
        frmAMISSearchCDM.Show vbModal
    ElseIf JOURNALTYPE = "CRJ" Then
        frmAMISSearchCRJ.Show vbModal
    ElseIf JOURNALTYPE = "CCM" Then
        frmAMISSearchCCM.Show vbModal
    Else
        frmAMISSearchGJ.Show vbModal
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdFirst_Click()
    On Error GoTo Errorcode:

    rsJournal_hd.MoveFirst
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdGJCancel_Click()
    SendToBackGJ
    StoreMemvars
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

    Dim cnt                                            As Integer
    Dim rsJournalDup                                   As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemNo,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(JOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_hd!VOUCHERNO) & " order by ID asc", gconDMIS
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
    rsJournal_hd.Find "id = " & labID.Caption
    cmdGJCancel.Value = True

    If lstGJ.ListItems.Count > 0 And lstGJ.Enabled = True Then
        lstGJ.SetFocus
    End If
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

    On Error GoTo Errorcode
    If cboGJAccountNo.Text = "" Then
        MsgBox "Account Code must have a value", vbInformation, "Error Encountered!"
        Exit Sub
    End If

    If NumericVal(txtGJDebit.Text) > 0 And NumericVal(txtGJCredit.Text) > 0 Then
        MsgBox "Invalid Journal Entry! Debit and Credit Amount can not have both Amount!", vbCritical, "Invalid Entry!"
        Exit Sub
    End If

    'If AddorEdit = "ADD" Then
    '    Dim rsJournal_DetClone                         As ADODB.Recordset
    '    Set rsJournal_DetClone = New ADODB.Recordset
    '    rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(JOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
    '    If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
    '        MsgBox "Account Code already used in this transaction", vbInformation, "Error in Part Number Validation"
    '        Exit Sub
    '    End If
    'End If


    If JOURNALTYPE = "GJ" Then
        If Left(cboGJAccountNo.Text, 5) = "11-02" Or Left(cboGJAccountNo.Text, 5) = "11-03" Then
            If MsgBox("A/R Codes must have a DM/CM to update the Customer Subsidiary" & vbCrLf & " or use DM/CM for A/R Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/R schedule") = vbYes Then
                'save in audit trail
                MsgBox "Reminder: You must use DM/CM to update A/R Schedule", vbInformation, "Confirmation Logged in Audit Trail."
            Else
                Exit Sub
            End If
        End If
        If Left(cboGJAccountNo.Text, 5) = "21-01" Then
            If MsgBox("A/P Codes must have a DM/CM to update the Vendors Subsidiary" & vbCrLf & " or use DM/CM for A/P Entries. Would you like to continue?", vbQuestion + vbYesNo, "Warning: Possible Update that will not update the A/P schedule") = vbYes Then
                'save in audit trail
                MsgBox "Reminder: You must use DM/CM to update A/P Schedule", vbInformation, "Confirmation Logged in Audit Trail."
            Else
                Exit Sub
            End If
        End If
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                  As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                As String
    Dim J_DEBIT, J_CREDIT, J_TAX                       As Double
    Dim J_STATUS, J_JITEMNO                            As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(JOURNALTYPE)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_JITEMNO = N2Str2Null(txtGJItemNo.Text)
    J_ACCT_CODE = N2Str2Null(cboGJAccountNo.Text)
    J_ACCT_NAME = N2Str2Null(txtGJAccountName.Text)
    J_DEBIT = Round(NumericVal(txtGJDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtGJCredit.Text), 2)
    J_TAX = Round(NumericVal(txtTax.Text), 2)
    J_STATUS = "'N'"

    Dim J_SUPCODE, J_ATC                               As String
    Dim J_RATE, J_TAXBASE                              As Double
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
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
    If AddorEdit = "ADD" Then cmdGJEntry_Click Else cmdGJCancel_Click
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub cmdJournalCancel_Click()
    SendToBack
    StoreMemvars
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
    Dim cnt                                            As Integer
    Dim rsJournalDup                                   As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemno,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(JOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_hd!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            'UPDATE DUE TO NEW AUDIT : BTT 08292008
            SQL_STATEMENT = "update AMIS_Journal_Det set JItemno = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            gconDMIS.Execute SQL_STATEMENT
            rsJournalDup.MoveNext
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        Loop
    End If
    FillDetails
    If JOURNALTYPE = "VDJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " credit = " & TOTCREDIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    ElseIf JOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " debit = " & TOTDEBIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " debit = " & TOTDEBIT & "," & _
                      " credit = " & TOTCREDIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    rsRefresh
    On Error Resume Next
    rsJournal_hd.Find "id = " & labID.Caption
    cmdJournalCancel.Value = True
    JournalTAB.TabEnabled(1) = True
    Picture1.Enabled = True
    If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
        lstDetails.SetFocus
    End If
End Sub

Private Sub cmdJournalSave_Click()
    Dim ValidateAccount As New ADODB.Recordset
    On Error GoTo Errorcode
    If cboAcct_Code.Text = "" Or Setacctname(cboAcct_Code.Text) = "" Then
        MsgBox "Account Code and Description must have a value", vbInformation, "Error Encountered!"
        Exit Sub
    End If


    'NOT TO ALLOW INPUT OF SAME ACCOUNT CODE
    '    If AddorEdit = "ADD" Then
    '        Dim rsJournal_DetClone                         As ADODB.Recordset
    '        Set rsJournal_DetClone = New ADODB.Recordset
    '        rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(JOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
    '        If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
    '            MsgBox "Account Code already used in this transaction", vbInformation, "Error in Account Code Validation"
    '            Exit Sub
    '        End If
    '    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                  As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME                As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET       As Double
    Dim J_STATUS, J_JITEMNO                            As String
    Dim J_ATC                                          As String
    Dim J_RATE, J_TAXBASE                              As Double

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(Format(txtVoucherNo.Text, "000000"))
    J_JTYPE = N2Str2Null(JOURNALTYPE)
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
        NEW_LogAudit "AA", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
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
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    FillDetails
    If JOURNALTYPE = "VDJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " credit = " & TOTCREDIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    ElseIf JOURNALTYPE = "VCJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " debit = " & TOTDEBIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                      " debit = " & TOTDEBIT & "," & _
                      " credit = " & TOTCREDIT & "," & _
                      " tax = " & TOTTAX & "," & _
                      " outbalance = " & OUTBALANCE & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    rsRefresh
    On Error Resume Next
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
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

Errorcode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdLast_Click()
    On Error GoTo Errorcode:

    rsJournal_hd.MoveLast
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdNext_Click()
    On Error GoTo Errorcode:

    rsJournal_hd.MoveNext
    If rsJournal_hd.EOF Then
        rsJournal_hd.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
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
    On Error GoTo Errorcode:
    
    Dim KimyDKid                                       As Integer

    For KimyDKid = 1 To lstDetails.ListItems.Count
        If lstDetails.ListItems(KimyDKid).ListSubItems(2).Text = "" Then
            MsgBox "Warning: Invalid Account Description Encountered!", vbCritical, "Can not Post!"
            Exit Sub
        End If
    Next

    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then Exit Sub
    If JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "PDJ" And JOURNALTYPE <> "OPB" Then
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
        If Not rsProfile.EOF And Not rsProfile.BOF Then
            If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
                If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
                    MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                    Exit Sub
                End If
            Else
                MsgBox "Warning: Action not authorized!" & vbCrLf & "Invalid Accouting Period", vbExclamation, "Error!"
                Exit Sub
            End If
        End If
    End If
    
    If NumericVal(txtTotDebit) <> NumericVal(txtTotCredit) Then
        MsgBox "Entry is not balanced. Posting of Entry Not Allowed.", vbInformation
        Exit Sub
    End If
    '    If CheckIfBookIsOpen(JOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If
    'Update by BTT : to update the balance of the SJ
    'If JOURNALTYPE = "CRJ" Then
    '    UpdateBalanceSJ txtVoucherNo.Text, True
    'End If
    
    Dim rsCheckDetails                                 As ADODB.Recordset
    Dim rsCheckCRJDetails                              As ADODB.Recordset
    Dim TotalCRJ_Credit                                As Double

    Screen.MousePointer = 11
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        If Trim(cboCustName.Text) = "" Then
            MsgBox "Warning: Posting is Not Allowed! Customer Name is Required!", vbInformation, "Missing Fields"
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        If Trim(cboCustName.Text) = "" Then
            MsgBox "Warning: Posting is Not Allowed! Customer Name is Required!", vbInformation, "Missing Fields"
            Screen.MousePointer = 0
            Exit Sub
        End If
        If picChat.Visible = True Then
            Screen.MousePointer = 0
            MsgBox "Warning: A/R Credit is not equal to details", vbCritical, "Error!"
            If COMPANY_CODE = "HAI" Then
                MsgBox "HAI is temporarily exempted in AR restrictions", vbInformation, "You can proceed"
                GoTo PostJournal
            End If
        Else
            GoTo PostJournal
        End If
    ElseIf JOURNALTYPE = "APJ" Then
        If Trim(cboNameofVendor.Text) = "" Then
            MsgBox "Warning: Posting is Not Allowed! Vendor Name is Required!", vbInformation, "Missing Fields"
            Screen.MousePointer = 0
            Exit Sub
        End If
        If Trim(txtPayCode.Text) = "" Or Trim(cboPayType.Text) = "" Then
            MsgBox "Warning: Payment Term/Type is Required!", vbInformation, "Missing Fields"
            Screen.MousePointer = 0
            Exit Sub
        End If
        If COMPANY_CODE <> "HAI" Then
            Set rsCheckDetails = New ADODB.Recordset
            Set rsCheckDetails = gconDMIS.Execute("Select Acct_Code, Credit from AMIS_Journal_Det Where (left(Acct_Code,5) = '21-01' OR left(Acct_Code,5) = '21-02'  OR left(Acct_Code,5) = '21-07') and Jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCheckDetails.EOF And Not rsCheckDetails.BOF Then
                rsCheckDetails.MoveFirst
                TotalCRJ_Credit = 0
                Do While Not rsCheckDetails.EOF
                    TotalCRJ_Credit = TotalCRJ_Credit + N2Str2Zero(rsCheckDetails!CREDIT)
                    rsCheckDetails.MoveNext
                Loop
                If TotalCRJ_Credit > 0 Then
                    If Round(NumericVal(txtAmountToPay.Text), 2) = Round(TotalCRJ_Credit, 2) Then
                        GoTo PostJournal
                    Else
                        Screen.MousePointer = 0
                        MsgBox "Warning: A/P Credit is not equal to Amount to Pay", vbCritical, "Error!"
                        Exit Sub
                    End If
                Else
                    GoTo PostJournal
                End If
            Else
                GoTo PostJournal
            End If
        Else
            GoTo PostJournal
        End If
    Else
        If JOURNALTYPE = "CDJ" Then
            If Trim(cboNameofVendor.Text) = "" Then
                MsgBox "Warning: Posting is Not Allowed! Vendor Name is Required!", vbInformation, "Missing Fields"
                Screen.MousePointer = 0
                Exit Sub
            End If
            Set rsCheckDetails = New ADODB.Recordset
            Set rsCheckDetails = gconDMIS.Execute("Select Acct_Code, debit from AMIS_Journal_Det Where (left(Acct_Code,5) = '21-01' or left(Acct_Code,5) = '21-02' or left(Acct_Code,5) = '21-07') and Jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
            If Not rsCheckDetails.EOF And Not rsCheckDetails.BOF Then
                rsCheckDetails.MoveFirst
                TotalCRJ_Credit = 0
                Do While Not rsCheckDetails.EOF
                    TotalCRJ_Credit = TotalCRJ_Credit + N2Str2Zero(rsCheckDetails!DEBIT)
                    rsCheckDetails.MoveNext
                Loop
                Set rsCheckCRJDetails = New ADODB.Recordset
                Set rsCheckCRJDetails = gconDMIS.Execute("Select SUM(AMOUNT) as TUTALINVAMT from AMIS_CV_Detail Where VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
                If Not rsCheckCRJDetails.EOF And Not rsCheckCRJDetails.BOF Then
                    If N2Str2Zero(rsCheckCRJDetails!TUTALINVAMT) = Round(TotalCRJ_Credit, 2) Or N2Str2Zero(rsCheckCRJDetails!TUTALINVAMT) = 0 Then
                        GoTo PostJournal
                    Else
                        Screen.MousePointer = 0
                        MsgBox "Warning: A/P Debit is not equal to details", vbCritical, "Error!"
                        Exit Sub
                    End If
                End If
            Else
                GoTo PostJournal
            End If
        Else
            GoTo PostJournal
        End If
    End If
    Screen.MousePointer = 0
    LogAudit "P", "JOURNAL ENTRY", txtJNo
    Exit Sub

PostJournal:
    If NumericVal(txtTotDebit.Text) <> NumericVal(txtTotCredit.Text) Then
        MsgBox "Warning: Total Debit is not equal to Total Credit", vbCritical, "Cannot be Posted!"
        Exit Sub
    End If
    If JOURNALTYPE = "SJ" Then
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',AmountToPay = " & TotalARAmountToPay & ",Balance = " & TotalARAmountToPay & "- AmountPaid where jtype = 'SJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    ElseIf JOURNALTYPE = "APJ" Then
        If TOTALAPAMOUNTTOPAY = 0 Then TOTALAPAMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',Balance = " & TOTALAPAMOUNTTOPAY & "- AmountPaid where jtype = 'APJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    ElseIf JOURNALTYPE = "CDJ" Then
        If TOTALAPAMOUNTTOPAY > 0 Then
            SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',Balance = " & TOTALPVAMOUNT - TOTALAPAMOUNTTOPAY & " where jtype = 'CDJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        Else
            SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',Balance = 0 where Jtype = 'CDJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        End If
    ElseIf JOURNALTYPE = "VDJ" Then
        If TOTALAPAMOUNTTOPAY = 0 Then TOTALAPAMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',Balance = " & TOTALAPAMOUNTTOPAY & "- AmountPaid, Debit = " & TOTALAPAMOUNTTOPAY & " where jtype = 'VDJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    ElseIf JOURNALTYPE = "CSJ" Then
        If TotalARAmountToPay = 0 Then TotalARAmountToPay = NumericVal(txtInvoiceAmt.Text)
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P',AmountToPay = " & TotalARAmountToPay & ",Balance = " & TotalARAmountToPay & "- AmountPaid, Debit = " & TotalARAmountToPay & " where jtype = 'CSJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    ElseIf JOURNALTYPE = "CCM" Then
        If TOTAL_AR_AMOUNT = 0 Then TOTAL_AR_AMOUNT = NumericVal(txtInvoiceAmt.Text)
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P', Credit = " & TOTAL_AR_AMOUNT & " where jtype = 'CCM' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    ElseIf JOURNALTYPE = "VCJ" Then
        If TOTAL_AP_AMOUNT = 0 Then TOTAL_AP_AMOUNT = NumericVal(txtCheckAmt.Text)
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P', Credit = " & TOTAL_AP_AMOUNT & " where jtype = 'VCJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    Else
        SQL_STATEMENT = "update AMIS_Journal_HD set status = 'P' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    End If
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo

    SQL_STATEMENT = "update AMIS_Journal_Det set status = 'P' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo

    'UPDATED BY: JUN
    'DATE UPDATED: 05-30-2009
    'DESCRIPTION: VALIDATE IF ALL ENTRY IN AMIS_JOURNAL_DET WAS TAG AS POSTED IF NOT UPDATE THE STATUS INTO POSTED
    Dim rsCHECK_POSTED As ADODB.Recordset
        Set rsCHECK_POSTED = gconDMIS.Execute("SELECT STATUS FROM AMIS_JOURNAL_DET where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " AND STATUS <> 'P'")
        If Not rsCHECK_POSTED.EOF And Not rsCHECK_POSTED.BOF Then
            gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        Else
            'ALL DETAILS ARE POSTED
        End If
    Set rsCHECK_POSTED = Nothing
    
    If JOURNALTYPE = "APJ" Then
        SQL_STATEMENT = "update AMIS_PV_Detail set status = 'P' where jtype = 'APJ' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    If JOURNALTYPE = "CDJ" Then
        SQL_STATEMENT = "update AMIS_CV_Detail set status = 'P' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    If JOURNALTYPE = "CRJ" Then
        SQL_STATEMENT = "update AMIS_CRJ_Detail set status = 'P' where VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    rsRefresh
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
    Exit Sub
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPostRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsgBox "Temporarily Disabled to ensure proper posting of Journals (AR/AP Issue)", vbInformation, "Netspeed Advice"
    Exit Sub
    If txtToVNo.Text < txtFromVNo.Text Then
        MsgBox "Error: Invalid Voucher No. Range", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    txtFromVNo.Text = Format(txtFromVNo.Text, "000000")
    txtToVNo.Text = Format(txtToVNo.Text, "000000")
    Dim rsCheckVouchers, rsCheckUnBalancedVouchers     As ADODB.Recordset
    Set rsCheckVouchers = New ADODB.Recordset
    Set rsCheckVouchers = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD where Jtype = '" & JOURNALTYPE & "' AND VoucherNo = '" & txtToVNo.Text & "'")
    If rsCheckVouchers.EOF And rsCheckVouchers.BOF Then
        MsgBox "Error: Voucher No. Range Exceeds Current Records Available.", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    Dim KIM, JOY, YZA                                  As Integer
    Screen.MousePointer = 11
    JOY = 0
    YZA = NumericVal(txtToVNo.Text) - NumericVal(txtFromVNo.Text)
    picShowPostRange.Enabled = False
    For KIM = txtFromVNo.Text To txtToVNo.Text
        Set rsCheckVouchers = New ADODB.Recordset
        Set rsCheckVouchers = gconDMIS.Execute("Select JType,VoucherNo,Debit,Credit,Status from AMIS_Journal_HD Where JType = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000")))
        If Not rsCheckVouchers.EOF And Not rsCheckVouchers.BOF Then
            Set rsCheckUnBalancedVouchers = New ADODB.Recordset
            Set rsCheckUnBalancedVouchers = gconDMIS.Execute("Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit from AMIS_Journal_Det Where jtype = '" & JOURNALTYPE & "' and Status <> 'C' and VoucherNo = " & N2Str2Null(Format(KIM, "000000")))
            If Round(rsCheckUnBalancedVouchers!TotalDebit, 2) <> Round(rsCheckUnBalancedVouchers!Totalcredit, 2) Then
                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                If JOURNALTYPE = "APJ" Then gconDMIS.Execute "update AMIS_PV_Detail set status = 'N' where jtype = 'APJ' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                If JOURNALTYPE = "CDJ" Then gconDMIS.Execute "update AMIS_CV_Detail set status = 'N' where VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                If JOURNALTYPE = "CRJ" Then gconDMIS.Execute "update AMIS_CRJ_Detail set status = 'N' where VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
            Else
                If Null2String(rsCheckVouchers!Status) = "N" Then
                    If N2Str2Zero(rsCheckVouchers!DEBIT) = N2Str2Zero(rsCheckVouchers!CREDIT) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "APJ" Then gconDMIS.Execute "update AMIS_PV_Detail set status = 'P' where status = 'N' and jtype = 'APJ' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CDJ" Then gconDMIS.Execute "update AMIS_CV_Detail set status = 'P' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CRJ" Then gconDMIS.Execute "update AMIS_CRJ_Detail set status = 'P' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!jtype) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
                    End If
                ElseIf Null2String(rsCheckVouchers!Status) = "C" Then
                    gconDMIS.Execute "update AMIS_Journal_HD set status = 'C' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    gconDMIS.Execute "update AMIS_Journal_Det set status = 'C' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    If JOURNALTYPE = "APJ" Then gconDMIS.Execute "update AMIS_PV_Detail set status = 'C' where status = 'N' and jtype = 'APJ' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    If JOURNALTYPE = "CDJ" Then gconDMIS.Execute "update AMIS_CV_Detail set status = 'C' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    If JOURNALTYPE = "CRJ" Then gconDMIS.Execute "update AMIS_CRJ_Detail set status = 'C' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                Else
                    If N2Str2Zero(rsCheckVouchers!DEBIT) = N2Str2Zero(rsCheckVouchers!CREDIT) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "APJ" Then gconDMIS.Execute "update AMIS_PV_Detail set status = 'P' where status = 'N' and jtype = 'APJ' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CDJ" Then gconDMIS.Execute "update AMIS_CV_Detail set status = 'P' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CRJ" Then gconDMIS.Execute "update AMIS_CRJ_Detail set status = 'P' where status = 'N' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "APJ" Then gconDMIS.Execute "update AMIS_PV_Detail set status = 'N' where jtype = 'APJ' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CDJ" Then gconDMIS.Execute "update AMIS_CV_Detail set status = 'N' where VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        If JOURNALTYPE = "CRJ" Then gconDMIS.Execute "update AMIS_CRJ_Detail set status = 'N' where VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!jtype) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
                    End If
                End If
            End If
        End If
        JOY = JOY + 1
        If YZA <> 0 Then prgPostRange.Value = (JOY / YZA) * 100
        DoEvents
    Next
    cmdShowPostRange.Visible = False: picShowPostRange.Visible = False
    rsRefresh
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
    Screen.MousePointer = 0
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdPrevious_Click()
    On Error GoTo Errorcode:

    rsJournal_hd.MovePrevious
    If rsJournal_hd.BOF Then
        rsJournal_hd.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdPrint_Click()
    Dim Ans                                            As String
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", LocalAcess) = False Then Exit Sub

    Ans = MsgBox("Are you sure do you want to print this Transaction?", vbQuestion + vbYesNo, "Print Transaction")
    If Ans = vbYes Then

        'For Reprint Routin : Update by BTT
        If JOURNALTYPE = "GJ" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If JOURNALTYPE = "APJ" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If JOURNALTYPE = "CDJ" Then
            If optPrintVoucher.Value = True Then
                If JOURNALTYPE = "CDJ" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
            End If
        End If

        If JOURNALTYPE = "SJ" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If JOURNALTYPE = "CRJ" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If JOURNALTYPE = "CLO" Then SaveReprintInformation JOURNALTYPE, MODULENAME, txtVoucherNo.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
        If JOURNALTYPE = "GJ" Then ShowReport "GeneralJournal", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "GENERAL JOURNAL PRINTOUT", LOGDATE, False
        If JOURNALTYPE = "APJ" Then ShowReport "AccountsPayable", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "ACCOUNTS PAYABLE JOURNAL PRINTOUT", LOGDATE, False
        If JOURNALTYPE = "CDJ" Then cmdPrinting.ZOrder 0: picPrinting.ZOrder 0: If COMPANY_CODE = "HSB" Then optCHINBANK.Caption = "BPI Bank"
        If JOURNALTYPE = "SJ" Then ShowReport "SalesJournal", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "SALES JOURNAL PRINTOUT", LOGDATE, False
        If JOURNALTYPE = "CRJ" Then ShowReport "CashReceipts", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CASH RECEIPTS JOURNAL PRINTOUT", LOGDATE, False
        If JOURNALTYPE = "CLO" Then ShowReport "ClosingEntries", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CLOSING ENTRIES", LOGDATE, False
        If JOURNALTYPE = "CSJ" Then ShowReport "CustomerCreditAdjustment", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CUSTOMER DEBIT MEMO", LOGDATE, False
        If JOURNALTYPE = "CCM" Then ShowReport "CustomerDebitAdjustment", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CUSTOMER CREDIT MEMO", LOGDATE, False
        If JOURNALTYPE = "VDJ" Then ShowReport "AdjusmentPayable", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "VENDOR DEBIT MEMO", LOGDATE, False
        If JOURNALTYPE = "VCJ" Then ShowReport "AdjusmentCPayable", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "VENDOR CREDIT MEMO", LOGDATE, False
        NEW_LogAudit "PX", "JOURNAL ENTRY", "", "", "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If


    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPV_Entry_Click()
    SendToBackPV
    BringToFrontPV
    AddorEdit = "ADD"
    cmdPVDelete.Visible = False
    InitPV_Detail
    On Error Resume Next
    If JOURNALTYPE = "APJ" Then
        On Error Resume Next
        txtPO_No.SetFocus
    Else
        On Error Resume Next
        txtMRR_No.SetFocus
    End If
End Sub

Private Sub cmdPVCancel_Click()
    SendToBackPV
    StoreMemvars
    JournalTAB.TabEnabled(0) = True
    Picture1.Enabled = True
End Sub

Private Sub cmdPVDelete_Click()
    If labPVID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If JOURNALTYPE = "APJ" Then
        If MsgBox("Delete This PV Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            SQL_STATEMENT = "delete from AMIS_PV_Detail where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
    ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        If MsgBox("Delete This CRJ Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            SQL_STATEMENT = "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PREVINVOICETYPE & "' and InvoiceNo = '" & PREVINVOICENO & "' and Jtype = 'SJ'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            SQL_STATEMENT = "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PREVINVOICETYPE & "' and InvoiceNo = '" & PREVINVOICENO & "' and Jtype = 'CSJ'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            SQL_STATEMENT = "delete from AMIS_CRJ_Detail where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
    Else
        If MsgBox("Delete This CV Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            SQL_STATEMENT = "delete from AMIS_CV_Detail where id = " & labPVID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
    End If
    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO         As String
    Dim J_JVOUCHERNO                                   As String
    Dim PV_AMOUNT                                      As Double
    Dim PV_STATUS, PV_ITEMNO                           As String

    J_JVOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    PV_ITEMNO = N2Str2Null(Format(txtPVItemNo.Text, "0000"))
    PV_PONO = N2Str2Null(txtPO_No.Text)
    PV_MRRNO = N2Str2Null(txtMRR_No.Text)             ' TYPE
    PV_INVNO = N2Str2Null(txtINV_No.Text)             'NO
    PV_PRODNO = N2Str2Null(txtProd_No.Text)           ' DATE
    PV_AMOUNT = NumericVal(txtPVAmount.Text)          'AMT
    PV_STATUS = "'N'"

    If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and Jtype = " & PV_PONO & "")    'Update By BTT : 09-24-2008
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            If Null2String(rsCheckJournal_HD!jtype) = "APJ" Then
                SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                              " PaidStatus = 'N' " & "," & _
                              " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                              " Balance = (Balance + " & PV_AMOUNT & ")" & _
                              " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            End If
            If Null2String(rsCheckJournal_HD!jtype) = "VPJ" Then
                SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                              " PaidStatus = 'N' " & "," & _
                              " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                              " Balance = (Balance + " & PV_AMOUNT & ")" & _
                              " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VPJ'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            End If
            If Null2String(rsCheckJournal_HD!jtype) = "VDJ" Then
                SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                              " PaidStatus = 'N' " & "," & _
                              " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                              " Balance = (Balance + " & PV_AMOUNT & ")" & _
                              " where VoucherNo = " & PV_MRRNO & " and Jtype = 'VDJ'"
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
            End If
        End If
    End If
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                          " ReceiveStatus = 'N' " & "," & _
                          " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                          " Balance = Balance + " & PV_AMOUNT & _
                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                          " ReceiveStatus = 'N' " & "," & _
                          " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                          " Balance = Balance + " & PV_AMOUNT & _
                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "XX", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
    End If
    FillDetails
    rsRefresh
    On Error Resume Next
    JournalTAB.TabEnabled(0) = True
    rsJournal_hd.Find "id = " & labID.Caption
    cmdPVCancel.Value = True
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim rsfindDup, rsProfile                           As ADODB.Recordset


    If IsNull(txtJNo.Text) = True Then
        MessagePop RecSaveError, "Error!", "Journal No. must not be empty"
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where jtype = '" & JOURNALTYPE & "' and jno = '" & txtJNo.Text & "' order by jtype,jno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MessagePop RecSaveError, "Error!", "Journal No. already exist!"
                Exit Sub
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
        Exit Sub
    End If
    If JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "OPB" And JOURNALTYPE <> "PDJ" Then
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
    '    If CheckIfBookIsOpen(JOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                  As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_CUSTOMERNAME                                 As String
    Dim J_DEBIT, J_CREDIT, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_CHECKNO                            As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE            As String
    Dim J_INVOICETYPE, J_INVOICENO                     As String
    Dim J_CHECKDATE, J_BANKCODE                        As String
    Dim J_REFNO, J_REFDATE                             As String
    Dim J_TERMS, J_DEALER                              As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                  As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(Format(txtVoucherNo.Text, "000000"))
    J_JTYPE = N2Str2Null(JOURNALTYPE)

    If JOURNALTYPE = "APJ" Or JOURNALTYPE = "VDJ" Then
        ' ' update by BTT
        J_INVOICEDATE = N2Str2Null(txtInvoiceDate.Text)
        J_BALANCE = NumericVal(txtAmountToPay.Text)
        J_INVOICETYPE = N2Str2Null(APJINVOICETYPE)
        J_INVOICENO = N2Str2Null(APJINVOICENO)
        J_AMOUNTPAID = 0
    ElseIf JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        J_INVOICEDATE = N2Str2Null(txtInvoiceDate2.Text)
        J_BALANCE = NumericVal(txtInvoiceAmt.Text)
        J_AMOUNTPAID = 0
    ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        J_INVOICEDATE = N2Str2Null(txtInvoiceDate2.Text)
        J_BALANCE = 0
        J_AMOUNTPAID = 0
    ElseIf JOURNALTYPE = "CDJ" Then
        If Trim(txtCheckNo.Text) = "" Then
            ShowIsRequiredMsg "Check Number"
            Exit Sub
        End If
        If IsDate(txtCheckDate.Text) = False Then
            ShowIsRequiredMsg "Check Date"
            Exit Sub
        End If
        If Trim(txtCheckNo.Text) = "" Then
            ShowIsRequiredMsg "Check Number"
            Exit Sub
        End If
        J_INVOICEDATE = "NULL"
        J_BALANCE = 0
        J_AMOUNTPAID = NumericVal(txtAmountToPay.Text)
    Else
        J_INVOICEDATE = "NULL"
        J_BALANCE = 0
        J_AMOUNTPAID = 0
    End If
    J_DUEDATE = N2Str2Null(txtDueDate.Text)
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        J_PAYTYPE = N2Str2Null(cboInvoiceType.Text)
    Else
        J_PAYTYPE = N2Str2Null(txtPayCode.Text)
    End If
    J_JNO = N2Str2Null(txtJNo.Text)
    If JOURNALTYPE = "VDJ" Or JOURNALTYPE = "CSJ" Then
        J_DEBIT = NumericVal(txtAmountToPay.Text)
        J_CREDIT = NumericVal(txtTotCredit.Text)
    ElseIf JOURNALTYPE = "VCJ" Then
        J_DEBIT = NumericVal(txtTotDebit.Text)
        J_CREDIT = NumericVal(txtCheckAmt.Text)
        J_AMOUNTPAID = NumericVal(txtCheckAmt.Text)
    ElseIf JOURNALTYPE = "CCM" Then
        J_DEBIT = NumericVal(txtTotDebit.Text)
        J_CREDIT = NumericVal(txtInvoiceAmt.Text)
        J_AMOUNTPAID = NumericVal(txtInvoiceAmt.Text)
    Else
        J_DEBIT = NumericVal(txtTotDebit.Text)
        J_CREDIT = NumericVal(txtTotCredit.Text)
    End If
    J_OUTBALANCE = NumericVal(txtOutBalance.Text)
    J_AMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
    J_STATUS = "'N'"

    J_CHECKNO = N2Str2Null(txtCheckNo.Text)
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        J_TERMS = N2Str2Null(txtTerms.Text)
        J_DEALER = N2Str2Null(txtDealer.Text)
    Else
        J_TERMS = "NULL"
        J_DEALER = "NULL"
    End If
    If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "CRJ" Or JOURNALTYPE = "VCJ" Or JOURNALTYPE = "CCM" Then
        J_CHECKDATE = N2Str2Null(txtCheckDate.Text)
    Else
        J_CHECKDATE = "NULL"
    End If
    J_BANKCODE = N2Str2Null(txtBankCode.Text)

    J_CUSTOMERNAME = "NULL"
    If JOURNALTYPE = "APJ" Or JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then
        J_CUSTOMERCODE = "'999999'"
        If Trim(txtCode.Text) = "" Then
            MsgBox "Invalid Supplier!", vbCritical, "Error"
            Exit Sub
        End If
        J_VENDORCODE = N2Str2Null(txtCode.Text)
    Else
        J_VENDORCODE = "'999999'"
        If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
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
    End If
    If JOURNALTYPE <> "APJ" Then ' update by BTT
    J_INVOICETYPE = N2Str2Null(SetInvCode(cboInvoiceType.Text))
    End If
    
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        If chkNonVat.Value = 1 Then
            J_INVOICENO = N2Str2Null("NV" & Format(txtInvoiceNo.Text, "000000"))
        Else
            J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
        End If
    Else
        If JOURNALTYPE <> "APJ" Then ' update by BTT
        J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
        End If
    End If
    J_INVOICEAMT = NumericVal(txtInvoiceAmt.Text)
    J_REFNO = N2Str2Null(txtRefNo.Text)
    J_REFDATE = N2Date2Null(txtRefDate.Text)
    If JOURNALTYPE = "APJ" Or JOURNALTYPE = "VDJ" Then
        If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks.Text))
    ElseIf JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
        If Trim(txtParticulars.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars.Text))
    ElseIf JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        If Trim(txtRemarks2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks2.Text))
    ElseIf JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        If Trim(txtRemarks2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks2.Text))
    Else
        If Trim(txtParticulars2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars2.Text))
    End If
    J_PAIDSTATUS = "'N'"
    J_RECEIVESTATUS = "'N'"

    If AddorEdit = "ADD" Then
        Dim rsJournal_HDDup                            As ADODB.Recordset
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
        J_JNO = N2Str2Null(txtJNo.Text)
        J_VOUCHERNO = N2Str2Null(GetVoucherNo(JOURNALTYPE))
        SQL_STATEMENT = "Insert into AMIS_Journal_HD" & _
                      " (jdate,voucherno,jtype,vendorcode,customercode,customername,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE)" & _
                      " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & J_CUSTOMERNAME & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                        ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "')"
        gconDMIS.Execute SQL_STATEMENT

        NEW_LogAudit "A", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
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
        NEW_LogAudit "E", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        CheckIfthereISCDJ txtVoucherNo
        SQL_STATEMENT = "update AMIS_Journal_Det set" & _
                      " jtype = " & J_JTYPE & "," & _
                      " jdate = " & J_JDATE & "," & _
                      " USERCODE = '" & LOGCODE & "'," & _
                      " LASTUPDATE = '" & LOGDATE & "'," & _
                      " jno = " & J_JNO & _
                      " where jtype = '" & PrevJType & "' and jno = '" & PrevJNo & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    End If
    If AddorEdit <> "ADD" Then
        rsJournal_hd.Find "jno = " & J_JNO
        cmdCancel.Value = True
        FillDetails
        If JOURNALTYPE = "VDJ" Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                          " debit = " & J_DEBIT & "," & _
                          " credit = " & TOTCREDIT & "," & _
                          " outbalance = " & OUTBALANCE & _
                          " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "JOURNAL ENTRY AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        ElseIf JOURNALTYPE = "VCJ" Then
            SQL_STATEMENT = "update AMIS_Journal_HD set" _
                          & " debit = " & TOTDEBIT & "," _
                          & " credit = " & J_CREDIT & "," _
                          & " outbalance = " & OUTBALANCE _
                          & " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "JOURNAL ENTRY AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        Else
            SQL_STATEMENT = "update AMIS_Journal_HD set" & _
                          " debit = " & TOTDEBIT & "," & _
                          " credit = " & TOTCREDIT & "," & _
                          " tax = " & TOTTAX & "," & _
                          " outbalance = " & OUTBALANCE & _
                          " where id = " & labID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "E", "JOURNAL ENTRY AMOUNT", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
        End If
    End If
    rsRefresh
    rsJournal_hd.Find "jno = " & J_JNO
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then
        If JOURNALTYPE = "GJ" Then cmdGJEntry_Click Else cmdAddJournal_Click
    End If
    Exit Sub
Errorcode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200713:18
Private Sub cmdUnPost_Click()
    On Error GoTo Errorcode:

    If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub

    If JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "PDJ" And JOURNALTYPE <> "OPB" Then
        Set rsProfile = New ADODB.Recordset
        Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
        If Not rsProfile.EOF And Not rsProfile.BOF Then
            If Year(txtJDate.Text) = rsProfile!PERIODYEAR Then
                If Month(txtJDate.Text) <> rsProfile!PERIODMONTH Then
                    MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                     Exit Sub
                End If
            Else
                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
                Exit Sub
            End If
        End If
    End If
    '    If CheckIfBookIsOpen(JOURNALTYPE, Month(txtJDate.Text), Year(txtJDate.Text)) = False Then
    '        MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
    '        Exit Sub
    '    End If
    
    
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
        Dim rsCRJ_Detail                               As ADODB.Recordset
        Set rsCRJ_Detail = New ADODB.Recordset
        Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail where INVOICETYPE = '" & SetInvCode(cboInvoiceType.Text) & "' AND INVOICENO = '" & txtInvoiceNo.Text & "' AND INVOICEDATE = '" & txtInvoiceDate.Text & "' and status <> 'C'")
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            MsgBox "Warning: This Sales Journal is already link to Cash Receipts Voucher No. " & Null2String(rsCRJ_Detail!VOUCHERNO) & vbCrLf & _
                 "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
            Exit Sub
        End If
    End If
    If JOURNALTYPE = "APJ" Then
        Dim rsCV_Detail                                As ADODB.Recordset
        Set rsCV_Detail = New ADODB.Recordset
        Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail where PV_VoucherNo = '" & txtVoucherNo.Text & "' and status <> 'C'")
        If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
            MsgBox "Warning: This AP Journal is already link to Cash Disbursement Voucher No. " & Null2String(rsCV_Detail!VOUCHERNO) & vbCrLf & _
                 "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    ' Update Due to new log Audit : BTT 282008
    SQL_STATEMENT = "update AMIS_Journal_HD set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute SQL_STATEMENT
    SQL_STATEMENT = "update AMIS_Journal_Det set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute SQL_STATEMENT
    rsRefresh
    rsJournal_hd.Find "id = " & labID.Caption
    StoreMemvars
    Screen.MousePointer = 0
    LogAudit "U", "JOURNAL ENTRY", txtJNo
    NEW_LogAudit "U", "JOURNAL ENTRY", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, JOURNALTYPE, txtJNo
    Exit Sub
Errorcode:
    ShowVBError
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

End Sub

Private Sub Command1_Click()
  If Module_Access(LOGID, "SYSTEM SETUP", "SYSTEM") = False Then Exit Sub
  frmAMISProfile.Show
End Sub

Private Sub Command3_Click()
' update by BTT 2/3/2009
If JOURNALTYPE = "SJ" Then
   
    ReturnInvoiceNo txtVoucherNo, JOURNALTYPE
    With frmAMIS_Payment
        frmAMIS_Payment.FillPaymentdetail AMIS_Invoiceno, AMIS_Invoicetype
        frmAMIS_Payment.Show
    End With
End If
If JOURNALTYPE = "APJ" Then
    With frmAMIS_Payment
        frmAMIS_Payment.FillPaymentdetail txtVoucherNo, ""
        frmAMIS_Payment.Show
    End With
End If
If JOURNALTYPE = "CRJ" Then
   
    OR_NUMBER_GLOBAL = txtInvoiceNo.Text
    frmORPaymentDetail.Show vbModal
End If
End Sub

Private Sub Command4_Click()
If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
        If Trim(txtMRR_No.Text) = "" Then frmAMISSearchAPJ2.Show vbModal
        
        End If
        If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        SEARCH_TAB = 0
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchSJ2.Show vbModal
       End If
End Sub

Private Sub Command5_Click()
    frmAMIS_UNAPPLIED_PAYMENT.Combo1.Text = "Customer Name"
    frmAMIS_UNAPPLIED_PAYMENT.txtSearch = RTrim(LTrim(cboCustName.Text))
    frmAMIS_UNAPPLIED_PAYMENT.Show
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Errorcode

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = Me.Caption
            Call frmALL_AuditInquiry.DisplayHistory(labID, JOURNALTYPE)

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
            ElseIf Me.ActiveControl.Name = "cboGJAccountNo" And cboGJAccountNo.Text = "" Then
                fraFindAccount.Visible = True
                cmdFindAccount.Visible = True
                cmdFindAccount.ZOrder 0
                fraFindAccount.ZOrder 0
                fraFindAccount.Enabled = True
                DoEvents
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
            ElseIf Me.ActiveControl.Name = "txtGJCredit" And SetAcctType(cboGJAccountNo.Text) = "C" And Val(txtGJCredit.Text) <= 0 And Val(txtGJDebit.Text) <= 0 Then
                On Error Resume Next
                txtGJCredit.SetFocus
            ElseIf Me.ActiveControl.Name = "txtGJDebit" And SetAcctType(cboGJAccountNo.Text) = "D" And Val(txtGJDebit.Text) <= 0 And Val(txtGJCredit.Text) <= 0 Then
                On Error Resume Next
                txtGJDebit.SetFocus
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
                    SendToBackGJ
                    SendToBackTemplates
                    StoreMemvars
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
                        StoreMemvars
                    ElseIf Me.ActiveControl.Name = "lstTemplates" Then
                        On Error Resume Next

                        txtSearchTemplates.SetFocus
                    Else
                        SendToBack
                        SendToBackPV
                        SendToBackGJ
                        SendToBackTemplates
                        StoreMemvars
                    End If
                End If
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsJournal_hd!Status) = "C" Then
                    MsgBox "Journals are Already Cancelled" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                ElseIf Null2String(rsJournal_hd!Status) = "P" Then
                    MsgBox "Journals are Already Posted" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                Else
                    JournalTAB.Tab = 0
                    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
                        Picture1.Enabled = False
                        cmdGJEntry_Click
                    Else
                        JournalTAB.TabEnabled(1) = False
                        Picture1.Enabled = False
                        cmdAddJournal_Click
                    End If
                End If
            End If
        Case vbKeyF4
            If JOURNALTYPE <> "SJ" Then
                If Picture1.Visible = True Then
                    If Null2String(rsJournal_hd!Status) = "C" Then
                        MsgBox "Journals are Already Cancelled" & vbCrLf & _
                               "and cannot be Change", vbInformation, "Edit Not Allowed!"
                    ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
                If Null2String(rsJournal_hd!Status) = "C" Then
                    MsgBox "Journals are Already Cancelled" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
            If Null2String(rsJournal_hd!Status) = "C" Then

                If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub

                If MsgBox("Are you sure you want to Un-Cancel this Transaction?", vbQuestion + vbYesNo, "Un-Cancel Journal") = vbYes Then
                    Screen.MousePointer = 11
                    gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                    gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & JOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                    rsRefresh
                    rsJournal_hd.Find "id = " & labID.Caption
                    StoreMemvars
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
            If JOURNALTYPE = "APJ" Then
                If MsgBox("Set this AP Transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AP") = vbYes Then
                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtAmountToPay) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='APJ'")
                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
                End If
            End If

            ' TEMPORARY to close AR for HAI
            If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
                If MsgBox("Set this AR transaction as Already Paid?", vbQuestion + vbYesNo, "Manual Close AR") = vbYes Then
                    gconDMIS.Execute ("Update AMIS_journal_hd set balance = 0, amountpaid='" & NumericVal(txtInvoiceAmt) & "',paidstatus='Y' where voucherno='" & txtVoucherNo & "' and jtype='SJ'")
                    MsgBox "Setting of transaction as Paid Successfully Done.", vbInformation, "Confirmed"
                End If
            End If
        End If
    End If
    Exit Sub

Errorcode:
    ShowErrMsg
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
    fraATC.Visible = False: fraATC2.Visible = False: labATC.Visible = False: cboJVSupCust.Visible = False
    labCheckAmt.Visible = False: txtCheckAmt.Visible = False: txtParticulars.Height = 795

    If JOURNALTYPE <> "CRJ" Then
        Label53.Visible = False
        Command5.Visible = False
    End If
    
    
    If JOURNALTYPE = "APJ" Then
        LocalAcess = "ACCOUNTS PAYABLE JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False
        Me.Caption = "ACCOUNTS PAYABLE JOURNAL DATA ENTRY"
        labSupplierPayTo = "Supplier Code"
        picGJ.Visible = False: picPayables.Visible = True: picPayables.ZOrder 0: picPayables.Enabled = True
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labPV1.Caption = "PO Number": labPV2.Caption = "MRR Number"
        labPV3.Caption = "Invoice Number": labPV4.Caption = "Product Number"
        labTax.Caption = "Input Tax": RefCRJ.Visible = False
    ElseIf JOURNALTYPE = "CDJ" Then
        CDJ_IS_FROM_AP = False
        LocalAcess = "CASH DISBURSEMENT JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False
        Me.Caption = "CASH DISBURSEMENT JOURNAL DATA ENTRY"
        labSupplierPayTo = "Pay To": RefCRJ.Visible = False
        picGJ.Visible = False: labDueDate.Visible = False: txtDueDate.Visible = False
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picDisbursement.Visible = True: picDisbursement.ZOrder 0: picDisbursement.Enabled = True
    ElseIf JOURNALTYPE = "SJ" Then
        LocalAcess = "SALES JOURNAL"
        chkNonVat.Visible = False: SJ_SHOW = True
        JournalTAB.TabEnabled(1) = False: labBankName.Visible = False: cboBankName2.Visible = False
        'labParticulars.Top = 960: 'txtRemarks2.Top = 930: txtRemarks2.Height = 1125
        Me.Caption = "SALES JOURNAL DATA ENTRY"
        labSupplierPayTo = "Supplier Code"
        labType.Caption = "Invoice Type": LabNo.Caption = "Invoice No."
        labDate.Caption = "Invoice Date": labAmt.Caption = "Invoice Amt."
        picGJ.Visible = False: RefCRJ.Visible = True
        picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labTax.Caption = "Output Tax"
    ElseIf JOURNALTYPE = "CRJ" Then
        LocalAcess = "CASH RECEIPTS JOURNAL"
        chkNonVat.Visible = True
        txtInvoiceNo.Left = 2040
        txtInvoiceNo.Width = 975
        fraComp.Visible = False
         Command3.Caption = "View OR Detail"
        Me.Caption = "CASH RECEIPTS JOURNAL DATA ENTRY"
        picGJ.Visible = False: RefCRJ.Visible = False
        labType.Caption = "Payment Type": LabNo.Caption = "O.R. No."
        labDate.Caption = "O.R. Date": labAmt.Caption = "O.R. Amount": labTerms.Visible = False
        picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labPV1.Caption = "Voucher No": txtPO_No.Enabled = False
        labPV2.Caption = "Invoice Type": labPV3.Caption = "Invoice No.": labPV4.Caption = "Invoice Date"
        lstPV_Detail.ColumnHeaders(2).Text = "Invoice Type"
        lstPV_Detail.ColumnHeaders(3).Text = "Invoice No."
        lstPV_Detail.ColumnHeaders(4).Text = "Invoice Date"
        lstPV_Detail.ColumnHeaders(5).Text = "Invoice Amt."
        Label51.Visible = False:
        Label52.Visible = True: cboARTag.Visible = True
    ElseIf JOURNALTYPE = "GJ" Then
        LocalAcess = "GENERAL JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Me.Caption = "GENERAL JOURNAL DATA ENTRY"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    ElseIf JOURNALTYPE = "ADJ" Then
        LocalAcess = "CLIENT ADJUSTING JOURNAL ENTRIES"
        Label3.Caption = "ADJ No."
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Me.Caption = "CLIENT ADJUSTING JOURNAL ENTRIES"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    ElseIf JOURNALTYPE = "PDJ" Then
        LocalAcess = "PROPOSED ADJUSTING JOURNAL ENTRIES"
        Label3.Caption = "ADJ No."
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Me.Caption = "PROPOSED ADJUSTING JOURNAL ENTRIES"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    ElseIf JOURNALTYPE = "CLO" Then
        LocalAcess = "CLOSING ENTRIES"
        Label3.Caption = "CLO No."
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Me.Caption = "CLOSING ENTRIES"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    ElseIf JOURNALTYPE = "OPB" Then
        LocalAcess = "ACCOUNT OPENING BALANCE"
        chkNonVat.Visible = False
        fraComp.Visible = False: RefCRJ.Visible = False
        Label3.Caption = "Ref. No.": Label5.Caption = "Ref. Date"
        Me.Caption = "OPENING BALANCES"
        picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
        labOutBalance.Visible = False: txtOutBalance.Visible = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picPayables.Enabled = False: picDisbursement.Enabled = False
        txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    ElseIf JOURNALTYPE = "VDJ" Then
        'FOR DM/CM WINDOW
        Label3.Caption = "DM No."
        Me.BackColor = &HC0E0FF
        Frame1.BackColor = &HC0E0FF
        picPayables.BackColor = &HC0E0FF
        Picture1.BackColor = &HC0E0FF
        Picture2.BackColor = &HC0E0FF
        LocalAcess = "ACCOUNTS PAYABLE JOURNAL"       '"GENERAL JOURNAL - VENDOR DEBIT MEMO"
        chkNonVat.Visible = False
        fraComp.Visible = False
        Me.Caption = "GENERAL JOURNAL - VENDOR DEBIT MEMO"
        labSupplierPayTo = "Supplier Code"
        picGJ.Visible = False: picPayables.Visible = True: picPayables.ZOrder 0: picPayables.Enabled = True
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labPV1.Caption = "PO Number": labPV2.Caption = "MRR Number"
        labPV3.Caption = "Invoice Number": labPV4.Caption = "Product Number"
        labTax.Caption = "Input Tax": RefCRJ.Visible = False
    ElseIf JOURNALTYPE = "VCJ" Then
        'FOR DM/CM WINDOW
        Label3.Caption = "CM No."
        Me.BackColor = &HC0E0FF
        Frame1.BackColor = &HC0E0FF
        picDisbursement.BackColor = &HC0E0FF
        Picture1.BackColor = &HC0E0FF
        Picture2.BackColor = &HC0E0FF
        txtParticulars.Height = 375
        labCheckAmt.Visible = True
        txtCheckAmt.Visible = True
        CDJ_IS_FROM_AP = False
        LocalAcess = "CASH DISBURSEMENT JOURNAL"
        chkNonVat.Visible = False
        fraComp.Visible = False
        Me.Caption = "GENERAL JOURNAL - VENDOR CREDIT MEMO"
        labSupplierPayTo = "Pay To": RefCRJ.Visible = False
        picGJ.Visible = False: labDueDate.Visible = False: txtDueDate.Visible = False
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
        picDisbursement.Visible = True: picDisbursement.ZOrder 0: picDisbursement.Enabled = True
    ElseIf JOURNALTYPE = "CSJ" Then
        'FOR DM/CM WINDOW
        Label3.Caption = "DM No."
        Me.BackColor = &HC0E0FF
        Frame1.BackColor = &HC0E0FF
        picReceivable.BackColor = &HC0E0FF
        Picture1.BackColor = &HC0E0FF
        Picture2.BackColor = &HC0E0FF
        LocalAcess = "SALES JOURNAL"
        chkNonVat.Visible = False: SJ_SHOW = True
        JournalTAB.TabEnabled(1) = False: labBankName.Visible = False: cboBankName2.Visible = False
        'labParticulars.Top = 960: 'txtRemarks2.Top = 930: txtRemarks2.Height = 1125
        Me.Caption = "GENERAL JOURNAL - CUSTOMER DEBIT MEMO"
        labSupplierPayTo = "Supplier Code"
        labType.Caption = "Invoice Type": LabNo.Caption = "Invoice No."
        labDate.Caption = "Invoice Date": labAmt.Caption = "Invoice Amt."
        picGJ.Visible = False: RefCRJ.Visible = True
        picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labTax.Caption = "Output Tax"
    ElseIf JOURNALTYPE = "CCM" Then
        Label3.Caption = "CM No."
        Me.BackColor = &HC0E0FF
        Frame1.BackColor = &HC0E0FF
        picReceivable.BackColor = &HC0E0FF
        Picture1.BackColor = &HC0E0FF
        Picture2.BackColor = &HC0E0FF
        chkNonVat.BackColor = &HC0E0FF
        LocalAcess = "CASH RECEIPTS JOURNAL"
        chkNonVat.Visible = True
        txtInvoiceNo.Left = 2040
        txtInvoiceNo.Width = 975
        fraComp.Visible = False
        Me.Caption = "GENERAL JOURNAL - CUSTOMER CREDIT MEMO"
        picGJ.Visible = False: RefCRJ.Visible = False
        labType.Caption = "Payment Type": LabNo.Caption = "O.R. No."
        labDate.Caption = "O.R. Date": labAmt.Caption = "O.R. Amount": labTerms.Visible = False
        picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
        picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
        picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
        labPV1.Caption = "Voucher No": txtPO_No.Enabled = False
        labPV2.Caption = "Invoice Type": labPV3.Caption = "Invoice No.": labPV4.Caption = "Invoice Date"
        lstPV_Detail.ColumnHeaders(2).Text = "Invoice Type"
        lstPV_Detail.ColumnHeaders(3).Text = "Invoice No."
        lstPV_Detail.ColumnHeaders(4).Text = "Invoice Date"
        lstPV_Detail.ColumnHeaders(5).Text = "Invoice Amt."
    End If
    InitGrid
    InitCbo
    InitMemVars
    txtSearch.Text = "": txtSearchTemplates.Text = ""
    rsRefresh
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        rsJournal_hd.MoveLast
    End If
    'If JOURNALTYPE = "SJ" Then picInvoiceDet.Visible = True Else picInvoiceDet.Visible = False
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
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
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        If Null2String(rsJournal_hd!Status) = "C" Then
            MessagePop RecLocekd, "Editing Not Allowed", "Transactions are Already Cancelled && cannot be Change"

            'MsgBox "Transactions are Already Cancelled" & vbCrLf & _
             '       "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
            If Null2String(rsJournal_hd!Status) = "C" Then
                MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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

Private Sub lstGJ_DblClick()
    If Null2String(rsJournal_hd!Status) = "C" Then
        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
    If Not rsJournal_hd.EOF And Not rsJournal_hd.BOF Then
        If Null2String(rsJournal_hd!Status) = "C" Then
            MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_hd!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            If JCNT > 0 Then
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
                If Null2String(rsJournal_hd!Status) = "C" Then
                    MsgBox "Journals are Already Cancelled" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                ElseIf Null2String(rsJournal_hd!Status) = "P" Then
                    MsgBox "Journals are Already Posted" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                Else
                    JournalTAB.Tab = 0
                    If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Or JOURNALTYPE = "PDJ" Or JOURNALTYPE = "CLO" Then
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
                If Null2String(rsJournal_hd!Status) = "C" Then
                    MsgBox "Journals are Already Cancelled" & vbCrLf & _
                           "and cannot be Change", vbInformation, "Edit Not Allowed!"
                ElseIf Null2String(rsJournal_hd!Status) = "P" Then
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
    If JOURNALTYPE = "APJ" Or JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VDJ" Or JOURNALTYPE = "VCJ" Then cboBankName.Text = SetBankName(txtBankCode.Text)
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CSJ" Or JOURNALTYPE = "CCM" Then cboBankName2.Text = SetBankName(txtBankCode.Text)
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
                If txtAcct_Name.Text = "OUTPUT TAX" And JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" Then
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
    If JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "PDJ" And JOURNALTYPE <> "CLO" Then
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
    If JOURNALTYPE <> "ADJ" And JOURNALTYPE <> "PDJ" And JOURNALTYPE <> "CLO" Then
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
    If JOURNALTYPE = "CDJ" Then
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
    If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CSJ" Or JOURNALTYPE = "CCM" Then
        cboCustName.SetFocus
    Else
        On Error Resume Next
        txtParticulars2.SetFocus
    End If
End Sub

Private Sub txtMRR_No_Change()
    Dim theJtype                                       As String
    Dim SQL                                            As String
    Dim rs                                             As New ADODB.Recordset
    If JOURNALTYPE = "CDJ" Then
        Set rsJournal_HD2 = New ADODB.Recordset
        'Update BTT : 09012008
        If txtPO_No.Text = "VPJ" Then
            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'VPJ'")
        Else
            Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay,Balance from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'APJ'and status='P'")
        End If
        If Not rsJournal_HD2.EOF And Not rsJournal_HD2.BOF Then
            theJtype = Null2String(rsJournal_HD2!jtype)
            txtINV_No.Text = Null2String(rsJournal_HD2!jdate)
            txtProd_No.Text = Null2String(rsJournal_HD2!duedate)
            txtPVAmount.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD2!BALANCE))
            If theJtype = "VPJ" Then
                Set rs = New ADODB.Recordset
                Set rs = gconDMIS.Execute("SELECT acct_code from AMIS_journal_det where voucherno='" & txtMRR_No.Text & "' and jtype ='VPJ'")
                If Not rs.EOF And Not rs.BOF Then
                    CDJ_AP = N2Str2Null(rs!ACCT_CODE)
                    ISVPJ = True
                End If
            Else
                CDJ_AP = ReturnAP_AccountCode("AP")
                CDJ_IS_FROM_AP = True
                ISVPJ = False
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
    If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "VCJ" Then
        If KeyAscii = 13 Then
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchAPJ2.Show vbModal
        End If
    End If
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
        If KeyAscii = 13 Then
            SEARCH_TAB = 0
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchSJ2.Show vbModal
        End If
    End If
End Sub

Private Sub txtMRR_No_LostFocus()
    If JOURNALTYPE = "CRJ" And AddorEdit <> "EDIT" Then
        Dim rsAR_Accounts                              As New ADODB.Recordset
        Set rsAR_Accounts = New ADODB.Recordset
        Set rsAR_Accounts = gconDMIS.Execute("select Acct_Code from AMIS_Journal_Det Where (Left(Acct_Code,5) = '11-02' or Left(Acct_Code,5) = '11-03') and  VoucherNo = '" & txtVoucherNo.Text & "' AND Jtype = '" & JOURNALTYPE & "'")
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
    If JOURNALTYPE = "CRJ" Then
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
    If JOURNALTYPE = "CRJ" Or JOURNALTYPE = "CCM" Then
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

Private Sub txtsearch_Change()
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
        If JOURNALTYPE = "SJ" Or JOURNALTYPE = "CSJ" And txtAcct_Name.Text = "OUTPUT TAX" Then
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
    If NumericVal(txtRATE2.Text) > 0 Then
        txtGJCredit.Text = Round(NumericVal(txtTaxBase2.Text) * (NumericVal(txtRATE2.Text) / 100), 2)
    End If

End Sub

Private Sub txtVoucherNo_LostFocus()
    txtVoucherNo.Text = Format(txtVoucherNo, "000000")
End Sub
Sub GettheTaxBaseAmnt()
    Dim SQL                                            As String
    Dim rs                                             As New ADODB.Recordset

    If JOURNALTYPE = "APJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '11-07002-00' and jtype = 'APJ'"
    Else
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & txtVoucherNo & "' and Acct_code <> '11-07002-00' and jtype = 'CDJ'"
    End If
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)

    If Not rs.EOF And Not rs.BOF Then
        txtTaxBase.Text = N2Str2IntZero(rs!SumDebit)
    End If
    Set rs = Nothing
End Sub

Function rsCHECKINVOICENOandTYPE(xinvoiceType As String, xInvoiceNo As String, XCustomerCode As String) As Boolean
    'UPDATED BY: JUN
    'DATE UPDATED: 10272008
    'DESCRIPTION: VALIDATE INVOICENO AND INVOICETYPE
    Dim rsExist                                        As ADODB.Recordset
    Set rsExist = gconDMIS.Execute("Select * from AMIS_Journal_hd where INVOICETYPE = '" & xinvoiceType & "' AND INVOICENO = '" & xInvoiceNo & "' AND JTYPE = 'SJ' and CUSTOMERCODE = '" & XCustomerCode & "'")
    If Not rsExist.EOF And Not rsExist.BOF Then
        rsCHECKINVOICENOandTYPE = True                ' yaon
    Else
        rsCHECKINVOICENOandTYPE = False               ' mayo
    End If
    Set rsExist = Nothing
End Function
Function GetSJVoucherNo(ByVal xInvoiceNo As String, ByVal xinvoiceType As String) As Boolean
    'Update BTT : 10282008
    'To check if the transaction is posted
    Dim RsSJVoucher                                    As New ADODB.Recordset
    Set RsSJVoucher = gconDMIS.Execute("Select Voucherno,invoicetype,invoiceno,Status from Amis_journal_hd where invoiceno=" & xInvoiceNo & " and invoicetype=" & xinvoiceType & " and jtype ='SJ'")
    If Not RsSJVoucher.EOF And Not RsSJVoucher.BOF Then
        If (RsSJVoucher!Status) = "P" Then
            GetSJVoucherNo = True
            SJVOUCHERNO = Null2String(RsSJVoucher!VOUCHERNO)
        Else
            GetSJVoucherNo = False
            MsgBox "Transaction is not posted to sales journal..Please verify", vbExclamation, "WARNING"
        End If
    End If
    Set RsSJVoucher = Nothing
End Function
Function ReturnAccountDescription(XXX As String)
    Dim RSACCT                                         As New ADODB.Recordset
    Set RSACCT = gconDMIS.Execute("SELECT * from AMIS_chartaccount where acctcode='" & XXX & "'")
    With RSACCT
        If Not .EOF And Not .BOF Then
            cboARTag.AddItem Null2String(RSACCT!Description)
        End If
    End With
    Set RSACCT = Nothing
End Function
Sub CheckIfthereISCDJ(XXX As String)
    Dim RSCDJ                                          As New ADODB.Recordset
    Set RSCDJ = gconDMIS.Execute("SELECT amount FROM AMIS_CV_DETAIL where Pv_voucherno='" & XXX & "'")
    If Not RSCDJ.EOF And Not RSCDJ.BOF Then
        gconDMIS.Execute "UPDATE AMIS_journal_hd set balance = " & TOTALPVAMOUNT - TOTALAPAMOUNTTOPAY & "  where voucherno='" & XXX & "' and jtype='APJ'"
    End If
    Set RSCDJ = Nothing
End Sub

Function CHECK_IF_SCHED_ACCNT(xVoucherNo As String) As Boolean
    'UPDATED BY: JUN
    'DATE UPDATED: 7-12-2009
    'DESCRIPTION: CHECK IF THERE IS AN AR SCHEDULE ACCOUNT IF NOT ALLOW SAVE WITHOUT TAGGING AS AR ACCOUNT
    Dim rsCHECK_IF_SCHED_ACCNT As ADODB.Recordset
    Dim SHED As Integer
    Dim NOT_SCHED As Integer
        SHED = 0
        NOT_SCHED = 0
        Set rsCHECK_IF_SCHED_ACCNT = New ADODB.Recordset
        rsCHECK_IF_SCHED_ACCNT.Open "Select Acct_Code From Amis_Journal_det where VoucherNo = '" & xVoucherNo & "' and Jtype = '" & JOURNALTYPE & "' and DEBIT <> 0 " & _
                                    "AND Acct_Code IN(SELECT AcctCode FROM Amis_ChartAccount where IS_SCHEDULE_ACCNT = 1)", gconDMIS, adOpenKeyset
        If Not rsCHECK_IF_SCHED_ACCNT.EOF And Not rsCHECK_IF_SCHED_ACCNT.BOF Then
            CHECK_IF_SCHED_ACCNT = True
        Else
            CHECK_IF_SCHED_ACCNT = False
        End If
    Set rsCHECK_IF_SCHED_ACCNT = Nothing
End Function

