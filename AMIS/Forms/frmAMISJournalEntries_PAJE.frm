VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmAMISJournalEntries_PAJE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PAJE Journal Entry"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
   Icon            =   "frmAMISJournalEntries_PAJE.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   30
      ScaleHeight     =   915
      ScaleWidth      =   9135
      TabIndex        =   43
      Top             =   1800
      Width           =   9135
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
         Height          =   675
         Left            =   0
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   240
         Width           =   9135
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   1005
      End
   End
   Begin VB.TextBox txtOutBalance 
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
      Left            =   2130
      MaxLength       =   14
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   1890
      Width           =   1515
   End
   Begin VB.TextBox txtTotDebit 
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
      Left            =   3690
      MaxLength       =   14
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1890
      Width           =   1515
   End
   Begin VB.TextBox txtTotCredit 
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
      Left            =   5250
      MaxLength       =   14
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   1890
      Width           =   1485
   End
   Begin VB.PictureBox fraAddJournal 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   30
      ScaleHeight     =   1635
      ScaleWidth      =   9105
      TabIndex        =   14
      Top             =   30
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
         MouseIcon       =   "frmAMISJournalEntries_PAJE.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries_PAJE.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   765
         Width           =   705
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
         MouseIcon       =   "frmAMISJournalEntries_PAJE.frx":1512
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries_PAJE.frx":1664
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   765
         Width           =   705
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   2310
         TabIndex        =   15
         Top             =   0
         Width           =   4425
         Begin VB.TextBox txtAcct_Name 
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
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   330
            Width           =   4335
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
            TabIndex        =   16
            Top             =   90
            Width           =   2205
         End
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
         TabIndex        =   3
         Top             =   330
         Width           =   1100
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
         TabIndex        =   4
         Top             =   330
         Width           =   1100
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
         MouseIcon       =   "frmAMISJournalEntries_PAJE.frx":19B4
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntries_PAJE.frx":1B06
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   765
         Visible         =   0   'False
         Width           =   705
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
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   330
         Width           =   2295
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
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   330
         Width           =   855
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
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   330
         Width           =   585
      End
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
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   330
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CommandButton cmdEntity 
         Caption         =   "Entity"
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
         Left            =   6930
         Picture         =   "frmAMISJournalEntries_PAJE.frx":1E31
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   765
         Width           =   705
      End
      Begin VB.CommandButton cmdParticulars 
         Caption         =   "Particular"
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
         Left            =   750
         Picture         =   "frmAMISJournalEntries_PAJE.frx":2EB3
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1725
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Frame FrameNoteDetail 
         Height          =   735
         Left            =   2340
         TabIndex        =   46
         Top             =   810
         Width           =   4365
         Begin VB.Label txtnotedetail 
            Caption         =   " Delete shcedule before changing/editing the account. "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   330
            TabIndex        =   48
            Top             =   180
            Width           =   3885
         End
         Begin VB.Label Label22 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Index           =   1
            Left            =   150
            TabIndex        =   47
            Top             =   210
            Width           =   165
         End
      End
      Begin VB.Frame fraATC 
         Height          =   915
         Left            =   2340
         TabIndex        =   19
         Top             =   690
         Width           =   4365
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
            MaxLength       =   50
            TabIndex        =   13
            Top             =   510
            Width           =   1725
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
            TabIndex        =   12
            Top             =   510
            Width           =   615
         End
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
            TabIndex        =   11
            Top             =   510
            Width           =   1425
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
            TabIndex        =   23
            Top             =   240
            Width           =   1725
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
            TabIndex        =   22
            Top             =   240
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
            TabIndex        =   21
            Top             =   240
            Width           =   1365
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
            TabIndex        =   20
            Top             =   540
            Width           =   855
         End
      End
      Begin VB.Frame fraComp 
         Height          =   915
         Left            =   2340
         TabIndex        =   24
         Top             =   660
         Width           =   4365
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
            Left            =   120
            MaxLength       =   10
            TabIndex        =   8
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
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   9
            Top             =   510
            Width           =   1300
         End
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
            TabIndex        =   10
            Top             =   510
            Width           =   1300
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
            TabIndex        =   27
            Top             =   240
            Width           =   1365
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
            TabIndex        =   26
            Top             =   240
            Width           =   1215
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
            TabIndex        =   25
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Label lblCode 
         Height          =   255
         Left            =   930
         TabIndex        =   40
         Top             =   1740
         Width           =   1035
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
         TabIndex        =   34
         Top             =   420
         Width           =   2685
      End
      Begin VB.Label labDetID 
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
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
         TabIndex        =   33
         Top             =   390
         Width           =   915
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
         TabIndex        =   32
         Top             =   360
         Width           =   855
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
         TabIndex        =   31
         Top             =   60
         Width           =   795
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
         TabIndex        =   30
         Top             =   60
         Width           =   885
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Code"
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
         TabIndex        =   28
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Label lblClass 
      Height          =   225
      Left            =   6810
      TabIndex        =   41
      Top             =   1890
      Width           =   975
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   1725
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   9255
      _Version        =   655364
      _ExtentX        =   16325
      _ExtentY        =   3043
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
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmAMISJournalEntries_PAJE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_Det                                           As New ADODB.Recordset
Dim rsChartAccount                                          As New ADODB.Recordset
Dim rsATC                                                   As New ADODB.Recordset
Dim xJOURNALTYPE                                            As String
Dim AddorEdit                                               As String
Event NewJournalEntries(AcctCode As String, DESCRIPTION As String, Debit As Double, Credit As Double, DetID As Long)
Dim WithEvents frmNewAMISJournalEntry_Chart                 As frmAMISJournalEntry_Chart
Attribute frmNewAMISJournalEntry_Chart.VB_VarHelpID = -1
Dim kcnt                                                    As Integer
Dim LocalAcess                                              As String
Dim WithEvents frmNewEntity                                 As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim JOURNAL_DETID                                           As Long
Dim xEntityClass                                            As String
''CHECKING FOR DETAILS Updated by NORMAN
Dim ACCT_HEADER                                             As New ADODB.Recordset
Dim HEADER_ACCT                                             As String
Dim ACCOUNT_OPENING                                         As New ADODB.Recordset
Dim ACCOUNT_CLOSING                                         As New ADODB.Recordset
Dim OPENING_ACCOUNT                                         As String
Dim CLOSING_ACCOUNT                                         As String
Public xD_JType                                             As String
Public xD_Voucherno                                         As String
Private Sub cboAcct_Code_Change()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    'GettheTaxBaseAmnt
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
'        fraATC.Visible = True
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
    End If
    
    Dim rsIT As ADODB.Recordset
    Set rsIT = New ADODB.Recordset
    Set rsIT = gconDMIS.Execute("select * from AMIS_CHARTACCOUNT where AcctCode = '" & cboAcct_Code.Text & "'  and TranType1 = 'INPUT TAX'")
    If Not rsIT.EOF And Not rsIT.BOF Then
        fraComp.Visible = True
        txtGrossAmt.Text = "0.00"
        txtTax.Text = "0.00"
        txtNetAmt.Text = "0.00"
    Else
        fraComp.Visible = False
    End If
    
    
End Sub

Private Sub cboAcct_Code_Click()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String

    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")

    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

    GettheTaxBaseAmnt
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
    End If

    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_GotFocus()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String

    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")

    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

    GettheTaxBaseAmnt
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
        On Error Resume Next
        cboATC.SetFocus
    Else
        fraATC.Visible = False
    End If
End Sub

Private Sub cboAcct_Code_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        If cboAcct_Code.Text = "" Then
            cmdSelect_Click
        End If
    Case vbKeyEscape
        If cboAcct_Code.Text = "" Then
            cmdJournalCancel_Click
        End If
    Case Else
        'MoveKeyPress KeyCode
    End Select
    
End Sub

Private Sub cboAcct_Code_LostFocus()
    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String
    'DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        fraATC.Visible = True
    Else
        fraATC.Visible = False
    End If
End Sub

Private Sub cboATC_Click()
'UPDATED: ACL 06252010
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

Private Sub cmdEntity_Click()

''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "EXIT TRANSACTION!"
        cmdJournalCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "EXIT TRANSACTION!"
        cmdJournalCancel.Value = True
        Exit Sub
    End If
''END COUNTER CHECKING BY NRE

    Set frmNewEntity = New frmEntity
    Call frmNewEntity.LOADJOURNAL("PAJ")
    frmNewEntity.Show 1
End Sub

Private Sub cmdJournalCancel_Click()
'frmAMISJournalEntry_APJ.JournalTAB.TabEnabled(0) = True
'frmAMISJournalEntry_APJ.Picture1.Enabled = True
    Call frmAMISJouirnalEntry_PAJE.load_voucher_chk
    LOAD_NEWJOURNAL = False
    Unload Me
End Sub

Private Sub cmdJournalDelete_Click()

    ''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "Warning!"
        cmdJournalCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "Warning!"
        cmdJournalCancel.Value = True
        Exit Sub
    End If
    ''END OF COUNTER CHECKING BY NRE
    
'If Function_Access(LOGID, "Acess_Delete", LocalAcess) = False Then Exit Sub
    Dim xACCT_CODE                                          As String
    If frmAMISJouirnalEntry_PAJE.labDetID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "System Message"
        Exit Sub
    End If
    If MsgBox("Delete this Journal entry, are you sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        If CheckARDetails(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(5)) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the AR details"
            Exit Sub
        End If

        If CheckAPDetails(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(5)) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the AP details"
            Exit Sub
        End If

        If CheckARPaymentDetails(xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(5)) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the Payment details"
            Exit Sub
        End If

        If CheckAPPaymentDetails(xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(5)) = True Then
            MessagePop InfoWarning, "System Message", "Action not allowed. Check for the Payment details"
            Exit Sub
        End If

        SQL_STATEMENT = "DELETE FROM AMIS_JOURNAL_DET WHERE id = " & frmAMISJouirnalEntry_PAJE.labDetID.Caption & ""
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "GENERAL JOURNAL", SQL_STATEMENT, frmAMISJouirnalEntry_PAJE.labID.Caption, "DT", frmAMISJouirnalEntry_PAJE.txtVoucherNo, xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.labDetID.Caption
    End If
    LOAD_NEWJOURNAL = False
    Unload Me
    Call frmAMISJouirnalEntry_PAJE.StoreSearch(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text)
End Sub

Private Sub cmdJournalSave_Click()
    On Error GoTo ErrorCode
    Dim str_MSG                                             As String
    str_MSG = "Error in saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf
    
''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "EXIT TRANSACTION!"
        cmdJournalCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "EXIT TRANSACTION!"
        cmdJournalCancel.Value = True
        Exit Sub
    End If
''END COUNTER CHECKING BY NRE
    
    
    If CHECK_IF_SCHED(cboAcct_Code.Text) = True Then
        If lblCode.Caption = "" Then
            MsgBox "Please select an entity", vbInformation, "System Message"
            Exit Sub
        ElseIf txtGJ_Remarks.Text = "" Then
            MessagePop RecSaveError, "System Message", "Field is empty!"
            txtGJ_Remarks.SetFocus
            Exit Sub
        End If
    End If
    
    If fraATC.Visible = True Then
        If cboATC.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Please select ATC code"
            cboATC.SetFocus
            Exit Sub
        End If
        If txtTaxBase.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Tax base amount must have a value"
            txtTaxBase.SetFocus
            Exit Sub
        End If
        If txtRATE.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Tax rate must have a value"
            txtRATE.SetFocus
            Exit Sub
        End If
    End If
    
    If txtDebit.Text = 0 And txtCredit.Text = 0 Then
            MessagePop InfoFriend, "INFORMATION", "Debit or Credit must have a value"
            Exit Sub
    End If
    


    gconDMIS.BeginTrans
    If JournalEntriesNew = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "General Journal")
        MsgBox str_MSG, vbCritical, "Journal Entry Error "
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans

    If AddorEdit = "ADD" Then frmAMISJouirnalEntry_PAJE.LOAD_JOURNALENTRY Else cmdJournalCancel_Click
    '    If AddorEdit = "EDIT" Then
    '        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then
    '            lstDetails.SetFocus
    '        End If
    '    End If

    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Function JournalEntriesNew() As Boolean
    On Error GoTo ErrorCode
    Dim J_JDATE                                             As String
    Dim J_VOUCHERNO                                         As String
    Dim J_JTYPE                                             As String
    Dim J_JNO                                               As String
    Dim J_ACCT_CODE                                         As String
    Dim J_ACCT_NAME                                         As String
    Dim J_STATUS                                            As String
    Dim J_JITEMNO                                           As String
    Dim J_INVOICENO                                         As String
    Dim J_INVOICETYPE                                       As String
    Dim xADJ_TYPE                                           As String
    Dim xADJ_VOUCHERNO                                      As String
    Dim xIS_OTHERS                                          As Integer
    Dim xADJ_REMARKS                                        As String
    Dim J_DEBIT                                             As Double
    Dim J_CREDIT                                            As Double
    Dim J_TAX                                               As Double
    Dim J_SUPCODE, J_ATC                                    As String
    Dim J_RATE, J_TAXBASE, J_NETAMT, J_GROSSAMT             As Double
    Dim J_CUSCDE                                            As String
    Dim xCUSNAME                                            As String

    Dim DEALER_ITW_COMPENSATION                             As String
    Dim DEALER_ITW_EXPANDED                                 As String

    DEALER_ITW_COMPENSATION = ReturnWithholdingTax("COMPENSATION")
    DEALER_ITW_EXPANDED = ReturnWithholdingTax("EXPANDED")

  
    If cboAcct_Code.Text = DEALER_ITW_EXPANDED Then
        J_ATC = N2Str2Null(cboATC.Text)
        J_RATE = NumericVal(txtRATE.Text)
        J_TAXBASE = NumericVal(txtTaxBase.Text)
    Else
        J_ATC = "NULL"
        J_RATE = 0
        J_TAXBASE = 0
    End If
    
    If fraComp.Visible = True Then
        J_TAX = Round(NumericVal(txtTax.Text), 2)
        J_GROSSAMT = Round(NumericVal(txtGrossAmt.Text), 2)
        J_NETAMT = Round(NumericVal(txtNetAmt.Text), 2)
    Else
        J_TAX = "0.00"
        J_GROSSAMT = "0.00"
        J_NETAMT = "0.00"
    End If

    J_JDATE = N2Date2Null(frmAMISJouirnalEntry_PAJE.txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(frmAMISJouirnalEntry_PAJE.txtJNo.Text)
    J_ACCT_CODE = N2Str2Null(cboAcct_Code.Text)
    J_ACCT_NAME = N2Str2Null(txtAcct_Name.Text)
    J_DEBIT = Round(NumericVal(txtDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtCredit.Text), 2)
    J_STATUS = "'N'"
    J_JITEMNO = "NULL"
    J_INVOICENO = "NULL"
    J_INVOICETYPE = "NULL"
    J_CUSCDE = N2Str2Null(lblClass.Caption + lblCode.Caption)
    xADJ_TYPE = "NULL"
    xADJ_VOUCHERNO = "NULL"
    xADJ_REMARKS = N2Str2Null(txtGJ_Remarks.Text)
    xIS_OTHERS = 0

    If cboAcct_Code.Text = "" Then
        MsgBox "Please select account code.", vbInformation, "System Message"
        cboAcct_Code.SetFocus
        JournalEntriesNew = True
        Exit Function
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into AMIS_Journal_Det " & _
                        "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status,ATC,RATE,TAXBASE,INVOICENO,INVOICETYPE,ENTITY,ADJ_VOUCHERNO,ADJ_JTYPE,Adj_Remarks,IS_OTHERS,JOURNAL_HD_ID,NETAMT,GROSSAMT)" & _
                        " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                        ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                        ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & "," & J_ATC & "," & J_RATE & "," & J_TAXBASE & "," & J_INVOICENO & "," & J_INVOICETYPE & "," & J_CUSCDE & "," & xADJ_VOUCHERNO & "," & xADJ_TYPE & ", " & xADJ_REMARKS & "," & xIS_OTHERS & "," & frmAMISJouirnalEntry_PAJE.labID.Caption & "," & J_NETAMT & "," & J_GROSSAMT & ")"
        gconDMIS.Execute SQL_STATEMENT
        JOURNAL_DETID = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_DET", J_JTYPE, "JTYPE")

        NEW_LogAudit "AA", "PROPOSED ADJUSTING JOURNAL ENTRIES", SQL_STATEMENT, frmAMISJouirnalEntry_PAJE.labID, "", frmAMISJouirnalEntry_PAJE.txtVoucherNo, xJOURNALTYPE, CStr(JOURNAL_DETID)

        MessagePop RecSave, "INFORMATION", "Record successfully added"
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
                         " status = " & J_STATUS & "," & _
                         " invoiceno = " & J_INVOICENO & "," & _
                         " invoicetype= " & J_INVOICETYPE & "," & _
                         " entity = " & J_CUSCDE & "," & _
                         " adj_voucherno = " & xADJ_VOUCHERNO & "," & _
                         " adj_jtype = " & xADJ_TYPE & "," & _
                         " Adj_Remarks = " & xADJ_REMARKS & "," & _
                         " IS_OTHERS = " & xIS_OTHERS & "," & _
                         " NETAMT = " & J_NETAMT & "," & _
                         " GROSSAMT = " & J_GROSSAMT & _
                         " where id = " & frmAMISJouirnalEntry_PAJE.labDetID.Caption
        NEW_LogAudit "EE", "CLIENT ADJUSTING JOURNAL ENTRIES", SQL_STATEMENT, frmAMISJouirnalEntry_PAJE.labID, "", frmAMISJouirnalEntry_PAJE.txtVoucherNo, xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.labDetID.Caption
        MessagePop RecSave, "INFORMATION", "Record Succesfully updated"
    End If
    CHECK_DETAILS
    frmAMISJouirnalEntry_PAJE.labDetID.Caption = ""
    LOAD_NEWJOURNAL = False
    JournalEntriesNew = True
    Unload Me
    Call frmAMISJouirnalEntry_PAJE.StoreSearch(J_VOUCHERNO)
    Exit Function
ErrorCode:
    Err_handler = "Error Number : " & Err.Number & vbCrLf & "Error Description :" & Err.DESCRIPTION
    JournalEntriesNew = False
End Function

Private Sub cmdSelect_Click()
'frmAMISJournalEntry_Chart.Caption = frmAMISJournalEntry_APJ.Caption
    Set frmNewAMISJournalEntry_Chart = frmAMISJournalEntry_Chart
    frmNewAMISJournalEntry_Chart.Show 1
    DoEvents
    'On Error Resume Next
    'frmNewAMISJournalEntry_Chart.txtSearch.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        '            If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
        '                cmdSelect_Click
        '                DoEvents
        '                On Error Resume Next
        '                'frmAMISJournalEntry_Chart.txtSearch.SetFocus
        '            End If
        If Me.ActiveControl.Name = "cboAcct_Code" Then
            OkAccount
        ElseIf Me.ActiveControl.Name = "txtCredit" And SetAcctType(cboAcct_Code.Text) = "C" And Val(txtCredit.Text) <= 0 And Val(txtDebit.Text) <= 0 Then
            On Error Resume Next
            txtCredit.SetFocus
        ElseIf Me.ActiveControl.Name = "txtDebit" And SetAcctType(cboAcct_Code.Text) = "D" And Val(txtDebit.Text) <= 0 And Val(txtCredit.Text) <= 0 Then
            On Error Resume Next
            txtDebit.SetFocus
        Else
            'MoveKeyPress KeyCode
        End If
    Case vbKeyEscape
        If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
            Unload Me
        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitJournal
    InitCbo
    fraATC.Visible = False
    fraComp.Visible = False
End Sub

Sub InitJournal()
'frmAMISJournalEntry_APJ.txtJItemNo.Text = Format(kcnt + 1, "0000")
    cboAcct_Code.Text = ""
    txtAcct_Name.Text = ""
    txtDebit.Text = ZERO
    txtCredit.Text = ZERO
    txtTax.Text = ZERO
    txtGrossAmt.Text = ZERO
    txtNetAmt.Text = ZERO
    '    txtSearch.Text = ""
    cboATC.Text = ""
    txtRATE.Text = "0"
    txtTaxBase.Text = ZERO
End Sub

Sub LOADJOURNAL(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Sub xADDorEDIT(XXX As String)
    AddorEdit = XXX
End Sub

Sub InitCbo()
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("select acctcode from AMIS_ChartAccount order by acctcode asc")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        Combo_Loadval cboAcct_Code, rsChartAccount
    End If
    Set rsChartAccount = Nothing

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
End Sub
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

Function ReturnWithholdingTax(XXX As String)
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = '" & XXX & "'")
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnWithholdingTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Sub GettheTaxBaseAmnt()
    Dim SQL                                                 As String
    Dim RS                                                  As New ADODB.Recordset

    If xJOURNALTYPE = "PAJ" Then
        SQL = "select sum(debit) as SumDebit from AMIS_journal_det where voucherno = '" & frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text & "' and Acct_code <> '" & ReturnInPutTax & "' and jtype = 'PAJ'"
    End If
    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    If Not RS.EOF And Not RS.BOF Then
        txtTaxBase.Text = ToDoubleNumber(N2Str2Zero(RS!SumDebit))
    End If
    Set RS = Nothing
End Sub

Function ReturnInPutTax()
    Dim rsChartAccount                                      As ADODB.Recordset
    Set rsChartAccount = New ADODB.Recordset
    If COMPANY_CODE = "HCC" Then
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where DESCRIPTION = 'INPUT TAX'")
    Else
        Set rsChartAccount = gconDMIS.Execute("Select AcctCode from AMIS_ChartAccount where TRANTYPE1 = 'INPUT TAX'")
    End If
    If Not rsChartAccount.EOF And Not rsChartAccount.BOF Then
        ReturnInPutTax = Null2String(rsChartAccount!AcctCode)
    End If
    Set rsChartAccount = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
End Sub

Public Sub frmNewAMISJournalEntry_Chart_ChartAccount(ACCT_CODE As String, DESCRIPTION As String)
    cboAcct_Code.Text = ACCT_CODE
    txtAcct_Name.Text = DESCRIPTION
End Sub

Private Sub txtCredit_GotFocus()
    If NumericVal(txtDebit.Text) = 0 Then
        If Val(txtCredit.Text) = 0 Then
            If NumericVal(txtNetAmt.Text) > 0 Then
                txtDebit.Text = ZERO
                txtCredit.Text = NumericVal(txtNetAmt.Text)
            Else
                'If OUTBALANCE > 0 And TOTDEBIT > 0 Then

                If NumericVal(txtOutBalance.Text) > 0 And NumericVal(txtTotDebit.Text) > 0 Then
                    txtCredit.Text = NumericVal(txtOutBalance.Text)
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

Private Sub txtCredit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
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
                    txtDebit.Text = ZERO: txtCredit.Text = NumericVal(txtOutBalance.Text)
                Else
                    If NumericVal(txtOutBalance.Text) > 0 And NumericVal(txtTotCredit.Text) > 0 Then
                        txtCredit.Text = ZERO: txtDebit.Text = NumericVal(txtOutBalance.Text)
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

Private Sub txtDebit_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtDebit_LostFocus()
    If txtDebit.Text = "" Then txtDebit.Text = 0
End Sub

Function StoreJournalEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,acct_code,acct_name,debit,jitemno,credit,tax,grossamt,netamt,ATC,RATE,TAXBASE,ENTITY,ADJ_REMARKS from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labDetID.Caption = rsJournal_Det!ID
        'labPartNo.Caption = Null2String(rsJournal_Det!ACCT_CODE)
        'txtJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboAcct_Code.Text = Null2String(rsJournal_Det!ACCT_CODE)
        txtAcct_Name.Text = Null2String(rsJournal_Det!acct_Name)
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Debit))
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!Credit))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!tax))
        txtGrossAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!grossamt))
        txtNetAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!netamt))
        lblClass.Caption = Null2String(Left(rsJournal_Det!ENTITY, 1))
        lblCode.Caption = Null2String(Right(rsJournal_Det!ENTITY, 6))
        txtGJ_Remarks.Text = Null2String(rsJournal_Det!ADJ_REMARKS)
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
  ''UPDATED BY NORMAN FOR CHECKING DETAILS
Call DETAIL_CHECKING
If OPENING_ACCOUNT <> "" Or CLOSING_ACCOUNT <> "" Then
    cboAcct_Code.Enabled = False
    
Else
     cboAcct_Code.Enabled = True
End If
If (HEADER_ACCT = "21" Or HEADER_ACCT = "11") And OPENING_ACCOUNT = "" And CLOSING_ACCOUNT = "" Then
    FrameNoteDetail.Visible = False
    FrameNoteDetail.ZOrder 1
ElseIf HEADER_ACCT = "" Then
    FrameNoteDetail.Visible = False
    FrameNoteDetail.ZOrder 1
Else
    FrameNoteDetail.Visible = True
    FrameNoteDetail.ZOrder 0
End If

If HEADER_ACCT = "11" Then
    If OPENING_ACCOUNT <> "" Then
        txtDebit.Enabled = True
        txtCredit.Enabled = False
        cmdEntity.Enabled = False
    ElseIf CLOSING_ACCOUNT <> "" Then
        txtDebit.Enabled = False
        txtCredit.Enabled = True
        cmdEntity.Enabled = False
    Else
        txtDebit.Enabled = True
        txtCredit.Enabled = True
        cmdEntity.Enabled = True
    End If
ElseIf HEADER_ACCT = "21" Then
     If OPENING_ACCOUNT <> "" Then
        txtDebit.Enabled = False
        txtCredit.Enabled = True
        cmdEntity.Enabled = False
    ElseIf CLOSING_ACCOUNT <> "" Then
        txtDebit.Enabled = True
        txtCredit.Enabled = False
        cmdEntity.Enabled = False
    Else
        txtDebit.Enabled = True
        txtCredit.Enabled = True
        cmdEntity.Enabled = True
    End If
Else
    txtDebit.Enabled = True
    txtCredit.Enabled = True
    cmdEntity.Enabled = True
End If

End Function

Function GetJNo(xVOUCHERNO As String) As String
'DESCRIPTION: GET THE THE HIGHEST JNO
    Dim rsgetJNO                                            As ADODB.Recordset
    Set rsgetJNO = gconDMIS.Execute("Select JNO From Amis_Journal_hd where Voucherno = '" & xVOUCHERNO & "' and Jtype = 'GJ'")
    If Not rsgetJNO.EOF And Not rsgetJNO.BOF Then
        GetJNo = Null2String(rsgetJNO!JNo)
    Else
        GetJNo = "000001"
    End If
    Set rsgetJNO = Nothing
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

Sub OkAccount()
    If cboAcct_Code.Text <> "" Then
        If SetAcctType(cboAcct_Code.Text) = "C" Then
            On Error Resume Next
            txtCredit.SetFocus
        Else
            On Error Resume Next
            txtDebit.SetFocus
        End If
    End If
End Sub

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    lblCode.Caption = strCode
    lblClass.Caption = strEntityClass
    frmAMISJouirnalEntry_PAJE.labcode.Caption = strCode
    frmAMISJouirnalEntry_PAJE.labName.Caption = strAccountName
    xEntityClass = strEntityClass
End Sub

Function CHECK_IF_SCHED(xACCT_CODE As String) As Boolean
'DESCRIPITON:CHECK AR ACCOUNT SCHEDULE THEN IF AR ACCOUNT RETURN TRUE
    Dim rsCHECK_IF_SCHED                                    As ADODB.Recordset
    Set rsCHECK_IF_SCHED = New ADODB.Recordset
    rsCHECK_IF_SCHED.Open "Select AcctCode from Amis_ChartAccount where AcctCode = '" & xACCT_CODE & "' and Is_Schedule_Accnt = 1", gconDMIS, adOpenKeyset
    If Not rsCHECK_IF_SCHED.EOF And Not rsCHECK_IF_SCHED.BOF Then
        CHECK_IF_SCHED = True
    Else
        CHECK_IF_SCHED = False
    End If
    Set rsCHECK_IF_SCHED = Nothing
End Function

Sub CHECK_DETAILS()
    If AddorEdit = "ADD" Then
        If CheckIfARDebitNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtDebit.Text) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, cboAcct_Code.Text, frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtDebit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_Details.xD_JType = xJOURNALTYPE
            frmAMISJournalEntry_Details.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
            frmAMISJournalEntry_Details.Show 1
        ElseIf CheckIfAPDebitNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtDebit.Text) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, xJOURNALTYPE, cboAcct_Code.Text, frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtDebit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_DetailPayment.xD_JType = xJOURNALTYPE
            frmAMISJournalEntry_DetailPayment.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfARCreditNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtCredit.Text) = True Then
            Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, xJOURNALTYPE, cboAcct_Code.Text, frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtCredit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, JOURNAL_DETID)
            frmAMISJournalEntry_DetailPayment.xD_JType = xJOURNALTYPE
            frmAMISJournalEntry_DetailPayment.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
            frmAMISJournalEntry_DetailPayment.Show 1
        ElseIf CheckIfAPCreditNotZero(cboAcct_Code.Text, CheckIfARAccount(N2Str2Null(cboAcct_Code.Text)), txtCredit.Text) = True Then
            Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, cboAcct_Code.Text, frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtCredit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, JOURNAL_DETID)
             frmAMISJournalEntry_Details.xD_JType = xJOURNALTYPE
            frmAMISJournalEntry_Details.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
            frmAMISJournalEntry_Details.Show 1
        End If
    Else
        If CheckIfSameAccount(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), cboAcct_Code.Text, xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, txtDebit.Text, txtCredit.Text) = False Then
            MessagePop InfoWarning, "System Message", "Action not allowed. " & Chr(13) & "See details for account code " & frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1)
            Exit Sub
        Else
            If CheckIfARDebitNotZero(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1))), txtDebit.Text) = True Then
                If CheckIfBalanceAR(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), NumericVal(txtDebit.Text), xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtDebit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, frmAMISJouirnalEntry_PAJE.labDetID.Caption)
                    frmAMISJournalEntry_Details.xD_JType = xJOURNALTYPE
                    frmAMISJournalEntry_Details.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
                    frmAMISJournalEntry_Details.Show 1
                End If
            ElseIf CheckIfAPDebitNotZero(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1))), txtDebit.Text) = True Then
                If CheckIfBalanceAPDetails(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), NumericVal(txtDebit.Text), xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtDebit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, frmAMISJouirnalEntry_PAJE.labDetID.Caption)
                    frmAMISJournalEntry_DetailPayment.xD_JType = xJOURNALTYPE
                    frmAMISJournalEntry_DetailPayment.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
                    frmAMISJournalEntry_DetailPayment.Show 1
                End If
            ElseIf CheckIfARCreditNotZero(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1))), txtCredit.Text) = True Then
                If CheckIfBalanceARDetails(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), NumericVal(txtCredit.Text), xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_DetailPayment.LOAD_DATA(frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, xJOURNALTYPE, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtCredit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, frmAMISJouirnalEntry_PAJE.labDetID.Caption)
                    frmAMISJournalEntry_DetailPayment.xD_JType = xJOURNALTYPE
                    frmAMISJournalEntry_DetailPayment.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
                    frmAMISJournalEntry_DetailPayment.Show 1
                End If
            ElseIf CheckIfAPCreditNotZero(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), CheckIfARAccount(N2Str2Null(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1))), txtCredit.Text) = True Then
                If CheckIfBalanceAP(frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), NumericVal(txtCredit.Text), xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text) = False Then
                    Call frmAMISJournalEntry_Details.LOAD_DATA(xJOURNALTYPE + "-" + frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text, frmAMISJouirnalEntry_PAJE.lstGJ.SelectedItem.SubItems(1), frmAMISJouirnalEntry_PAJE.txtJDate.Text, txtCredit.Text, lblClass.Caption, lblCode.Caption, txtDebit.Text, frmAMISJouirnalEntry_PAJE.labDetID.Caption)
                    frmAMISJournalEntry_Details.xD_JType = xJOURNALTYPE
                    frmAMISJournalEntry_Details.xD_Voucherno = frmAMISJouirnalEntry_PAJE.txtVoucherNo.Text
                    frmAMISJournalEntry_Details.Show 1
                End If
            End If
        End If
    End If
End Sub

Sub DETAIL_CHECKING()
OPENING_ACCOUNT = ""
CLOSING_ACCOUNT = ""
HEADER_ACCT = ""
Set ACCT_HEADER = New ADODB.Recordset
ACCT_HEADER.Open "SELECT HEADERS FROM AMIS_CHARTACCOUNT WHERE ACCTCODE ='" & cboAcct_Code.Text & "' AND IS_SCHEDULE_ACCNT = 1", gconDMIS, adOpenForwardOnly
If Not ACCT_HEADER.EOF And Not ACCT_HEADER.BOF Then
    HEADER_ACCT = Null2String(ACCT_HEADER!HEADERS)
End If
If HEADER_ACCT = "11" Then
     If txtDebit.Text > 0 Then
        Set ACCOUNT_OPENING = New ADODB.Recordset
        ACCOUNT_OPENING.Open "SELECT * FROM AMIS_AR WHERE JOURNAL_DET_ID ='" & labDetID.Caption & "'", gconDMIS, adOpenForwardOnly
        If Not ACCOUNT_OPENING.EOF And Not ACCOUNT_OPENING.BOF Then
        OPENING_ACCOUNT = Null2String(ACCOUNT_OPENING!Journal_Det_ID)
        End If
    Else
        Set ACCOUNT_CLOSING = New ADODB.Recordset
        ACCOUNT_CLOSING.Open "SELECT * FROM AMIS_DETAIL WHERE JOURNAL_DET_ID ='" & labDetID.Caption & "'", gconDMIS, adOpenForwardOnly
        If Not ACCOUNT_CLOSING.EOF And Not ACCOUNT_CLOSING.BOF Then
        CLOSING_ACCOUNT = Null2String(ACCOUNT_CLOSING!Journal_Det_ID)
        End If
    End If
ElseIf HEADER_ACCT = "21" Then
    If txtDebit.Text > 0 Then
        Set ACCOUNT_CLOSING = New ADODB.Recordset
        ACCOUNT_CLOSING.Open "SELECT * FROM AMIS_DETAILS WHERE JOURNAL_DET_ID ='" & labDetID.Caption & "'", gconDMIS, adOpenForwardOnly
        If Not ACCOUNT_CLOSING.EOF And Not ACCOUNT_CLOSING.BOF Then
        CLOSING_ACCOUNT = Null2String(ACCOUNT_CLOSING!Journal_Det_ID)
        End If
    Else
        Set ACCOUNT_OPENING = New ADODB.Recordset
        ACCOUNT_OPENING.Open "SELECT * FROM AMIS_AP WHERE JOURNAL_DET_ID ='" & labDetID.Caption & "'", gconDMIS, adOpenForwardOnly
        If Not ACCOUNT_OPENING.EOF And Not ACCOUNT_OPENING.BOF Then
        OPENING_ACCOUNT = Null2String(ACCOUNT_OPENING!Journal_Det_ID)
        End If
    End If
Else
End If
End Sub

Private Sub txtGrossAmt_Change()
txtTax.Text = Round(NumericVal(NumericVal(txtGrossAmt.Text) / 1.12 * 0.12), 2)
txtNetAmt.Text = Round(NumericVal(NumericVal(txtGrossAmt.Text) / 1.12), 2)
txtDebit.Text = txtTax.Text
End Sub

'SJR 6/13/14
Private Sub txtTaxBase_Change()
txtCredit.Text = Round(NumericVal(txtTaxBase.Text) * (NumericVal(txtRATE.Text) / 100), 2)
End Sub
'SJR 6/13/14

