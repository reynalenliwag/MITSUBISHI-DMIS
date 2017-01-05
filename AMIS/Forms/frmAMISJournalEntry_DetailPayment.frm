VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAMISJournalEntry_DetailPayment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schedules"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8100
   Icon            =   "frmAMISJournalEntry_DetailPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   870
      TabIndex        =   37
      Text            =   "Total:"
      Top             =   4500
      Width           =   1185
   End
   Begin VB.TextBox lblAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
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
      Left            =   2100
      MaxLength       =   15
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   4500
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   4560
      ScaleHeight     =   825
      ScaleWidth      =   3480
      TabIndex        =   22
      Top             =   5520
      Width           =   3480
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
         Left            =   2760
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   2070
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Print this Record"
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
         Left            =   690
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   0
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":1EA0
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":1FF2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add Record"
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
         Left            =   690
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":2305
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":2457
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move to Next Record"
         Top             =   960
         Visible         =   0   'False
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
         Left            =   0
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":27AF
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":2901
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Previous Record"
         Top             =   960
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   30
      ScaleHeight     =   1695
      ScaleWidth      =   8025
      TabIndex        =   15
      Top             =   450
      Width           =   8055
      Begin VB.TextBox txtCusCde 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "000226"
         Top             =   60
         Width           =   1095
      End
      Begin VB.TextBox txtCustomerName 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "000226"
         Top             =   60
         Width           =   4905
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "::"
         Enabled         =   0   'False
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
         Left            =   2580
         TabIndex        =   1
         Top             =   60
         Width           =   345
      End
      Begin VB.TextBox txtInvoiceDate 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "88/88/8888"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox txtInvoiceAmt 
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   870
         Width           =   1485
      End
      Begin VB.TextBox txtInvoiceNo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "000000"
         Top             =   450
         Width           =   1485
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   795
         Left            =   3120
         TabIndex        =   6
         Top             =   870
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1402
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   2
         MaxLength       =   100
         Appearance      =   0
         TextRTF         =   $"frmAMISJournalEntry_DetailPayment.frx":2C60
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   180
         Width           =   570
      End
      Begin VB.Label labDate 
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
         Left            =   60
         TabIndex        =   19
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label labParticulars 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   3120
         TabIndex        =   18
         Top             =   570
         Width           =   990
      End
      Begin VB.Label LabNo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Left            =   60
         TabIndex        =   17
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label labAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amt."
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
         Left            =   60
         TabIndex        =   16
         Top             =   990
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lstDetails 
      Height          =   2205
      Left            =   0
      TabIndex        =   7
      Top             =   2190
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3889
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
      Appearance      =   0
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
      MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":2CF7
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item #"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Inv. No."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Inv. Amt."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Inv. Date"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Particulars"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   8040
      TabIndex        =   23
      Top             =   4470
      Width           =   8040
      Begin VB.Timer tmReal 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   1320
         Top             =   360
      End
      Begin VB.Timer tmrSarah 
         Interval        =   300
         Left            =   840
         Top             =   360
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
         Left            =   0
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":2E59
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":2FAB
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
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
         Left            =   7290
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":32D6
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":3428
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Cancel"
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
         Left            =   6600
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":3766
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":38B8
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Save Entry"
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
         Left            =   5910
         MouseIcon       =   "frmAMISJournalEntry_DetailPayment.frx":3C08
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_DetailPayment.frx":3D5A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.TextBox txtDueDate 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   26
      Text            =   "88/88/8888"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblRealSJVoucherNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   7920
      TabIndex        =   38
      Top             =   120
      Width           =   60
   End
   Begin VB.Label lblJType 
      Height          =   255
      Left            =   1440
      TabIndex        =   35
      Top             =   7470
      Width           =   1575
   End
   Begin VB.Label lblVoucherNo 
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   7110
      Width           =   1575
   End
   Begin VB.Label lblSJVoucherNo 
      Height          =   255
      Left            =   1440
      TabIndex        =   33
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblBalance 
      Height          =   255
      Left            =   1440
      TabIndex        =   32
      Top             =   6330
      Width           =   1575
   End
   Begin VB.Label lblEntityClass 
      Height          =   285
      Left            =   1440
      TabIndex        =   31
      Top             =   5910
      Width           =   1575
   End
   Begin VB.Label lblInvoiceType 
      Height          =   345
      Left            =   1440
      TabIndex        =   30
      Top             =   5460
      Width           =   1575
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
      Left            =   90
      TabIndex        =   27
      Top             =   4740
      Visible         =   0   'False
      Width           =   885
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   405
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   8145
      _Version        =   655364
      _ExtentX        =   14367
      _ExtentY        =   714
      _StockProps     =   14
      Caption         =   "Add/Edit Schedules"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   4005
      Left            =   0
      TabIndex        =   14
      Top             =   420
      Width           =   8145
      _Version        =   655364
      _ExtentX        =   14367
      _ExtentY        =   7064
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
Attribute VB_Name = "frmAMISJournalEntry_DetailPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents frmNewEntity                                 As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim xEntityClass                                            As String
Dim xVOUCHERNO                                              As String
Dim xJType                                                  As String
Dim xInvoiceType                                            As String
Dim xInvoiceNo                                              As String
Dim XCustomerCode                                           As String
Dim xCUSTOMERNAME                                           As String
Dim xINVOICEAMOUNT                                          As Double
Dim xACCT_CODE                                              As String
Dim xInvoicedate                                            As String
Dim xLAST_UPDATED                                           As String
Dim xJdate                                                  As String
Dim xREMARKS                                                As String
Dim cntDetails                                              As Integer
Dim AddorEdit                                               As String
Dim SQL_STATEMENT                                           As String
Dim xREFCODE                                                As String
Dim xSJVOUCHERNO                                            As String
Dim xAMOUNT                                                 As Double
Dim xdebit                                                  As Double
Dim xDUEDATE                                                As String
Dim PAYCODE                                                 As String
Dim xJOURNAL_DET_ID                                         As Long
Dim rsAR                                                    As ADODB.Recordset
Public xD_JType                                             As String
Public xD_Voucherno                                         As String
Public xJDATEVAL                                            As String


'SJR DECLARATIONS
'--Showing validation for ahead details (F2)--'
Public ssDetailInvDate                                      As String

'--Showing SJVoucher number--'
Dim rsReal                                                  As ADODB.Recordset
Dim LST                                                     As ListItem
'SJR DECLARATIONS

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    picDetails.Enabled = True
    lstDetails.Enabled = False
    initMemvars
    'Picture1.Visible = False
    'Picture2.Visible = True
End Sub

Private Sub cmdCancel_Click()

If xD_JType = "APJ" Then
Call frmAMISJournalEntry_APJ.load_voucher_chk
ElseIf xD_JType = "CDJ" Then
Call frmAMISJournalEntry_CDJ.load_voucher_chk
ElseIf xD_JType = "SJ" Then
Call frmAMISJournalEntry_SJ.load_voucher_chk
ElseIf xD_JType = "CRJ" Then
Call frmAMISJournalEntry_CRJ.load_voucher_chk
ElseIf xD_JType = "DRJ" Then
Call frmAMISJournalEntry_DRJ.load_voucher_chk
ElseIf xD_JType = "GJ" Then
Call frmAMISJournalEntry_GJ.load_voucher_chk
ElseIf xD_JType = "CSJ" Then
Call frmAMISJournalEntry_CDM.load_voucher_chk
End If



'    picDetails.Enabled = False
'    Picture1.Visible = True
'    Picture2.Visible = False
'    lstDetails.Enabled = True
'    If lstDetails.ListItems.Count > 0 Then
'        Call StoreMemVars(lstDetails.ListItems.Item(1).SubItems(6))
'    End If

    If NumericVal(lblAmount.Text) <> NumericVal(xAMOUNT) Then
        If MsgBox("GL Amount is not equal to SL, proceed?", vbYesNo + vbQuestion, "System Message") = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdDelete_Click()

    ''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    End If
    ''END OF COUNTER CHECKING BY NRE
    
    If lstDetails.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to Delete this Detail?", vbQuestion + vbYesNo, "Delete AR Detail") = vbYes Then
            If CheckIfARAccount(xACCT_CODE) = True Then
                Dim rsCRJID As New ADODB.Recordset
                Dim CRJ_ID As String
                Set rsCRJID = New ADODB.Recordset
                Set rsCRJID = gconDMIS.Execute("SELECT CRJ_ID FROM AMIS_DETAIL WHERE ID = " & lstDetails.SelectedItem.SubItems(6))
                    If Not rsCRJID.EOF And Not rsCRJID.BOF Then
                        CRJ_ID = N2Str2Null(rsCRJID!CRJ_ID)
                        SQL_STATEMENT = "DELETE AMIS_CRJ_DETAIL where id = " & CRJ_ID & ""
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                SQL_STATEMENT = "DELETE from AMIS_DETAIL where id = " & lstDetails.SelectedItem.SubItems(6)
            Else
                Dim rsCV_ID As New ADODB.Recordset
                Dim CV_ID As String
                Set rsCV_ID = New ADODB.Recordset
                Set rsCV_ID = gconDMIS.Execute("SELECT CV_ID FROM AMIS_DETAILS WHERE ID = " & lstDetails.SelectedItem.SubItems(6))
                    If Not rsCV_ID.EOF And Not rsCV_ID.BOF Then
                        CV_ID = N2Str2Null(rsCV_ID!CV_ID)
                        SQL_STATEMENT = "DELETE AMIS_CV_DETAIL where id = " & CV_ID & ""
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                SQL_STATEMENT = "DELETE from AMIS_DETAILS where id = " & lstDetails.SelectedItem.SubItems(6)
            End If
            gconDMIS.Execute SQL_STATEMENT
        End If
        Call rsRefresh(xVOUCHERNO, xJType, xACCT_CODE, CURRENT_CUSCODE, xJOURNAL_DET_ID)
        Call FillGrid(xVOUCHERNO)
    End If
    initMemvars
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    picDetails.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstDetails.Enabled = False
End Sub

Private Sub cmdExit_Click()
    If NumericVal(lblAmount.Text) <> NumericVal(xAMOUNT) Then
        If MsgBox("GL Amount is not equal to SL, proceed?", vbYesNo + vbQuestion, "System Message") = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click()
    
    ''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    End If
    ''END OF COUNTER CHECKING BY NRE
    
    AddorEdit = "ADD"
    picDetails.Enabled = True
    'lstDetails.Enabled = False
    initMemvars
    Call frmAMISSearchAR.CHECKACCOUNT(xACCT_CODE)
    If xD_JType = "GJ" Then
        frmAMISSearchAR.chkCurrentCust = True
        frmAMISSearchAR.chkCurrentCust.Locked = True
    Else
        frmAMISSearchAR.chkCurrentCust = True
        frmAMISSearchAR.chkCurrentCust.Locked = False
    End If
    frmAMISSearchAR.Show 1
End Sub

Private Sub cmdSave_Click()
        
    ''COUNTER CHECKING BY NRE
    If CIP(xD_JType, xD_Voucherno) = "P" Then
        MsgBox "Warning: Transaction Already Posted", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    ElseIf CIP(xD_JType, xD_Voucherno) = "C" Then
        MsgBox "Warning: Transaction Already Cancelled", vbCritical, "Warning!"
        cmdCancel.Value = True
        Exit Sub
    End If
    ''END OF COUNTER CHECKING BY NRE
    
    If txtCusCde.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtCusCde.SetFocus
        Exit Sub
    ElseIf txtInvoiceNo.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtInvoiceNo.SetFocus
        Exit Sub
    ElseIf txtInvoiceAmt.Text = "" Or txtInvoiceAmt.Text = "0.00" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtInvoiceAmt.SetFocus
        Exit Sub
        '    ElseIf txtInvoiceDate.Text = "" Or IsDate(txtInvoiceDate.Text) = False Then
        '        MessagePop RecSaveError, "System Message", "Invalid is date!"
        '        txtInvoiceDate.SetFocus
        '        Exit Sub
    ElseIf txtRemarks.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtRemarks.SetFocus
        Exit Sub
    ElseIf txtRemarks.Text = "Pls Type Your Message Here!" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtRemarks.SetFocus
        Exit Sub
    End If
    If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then xREMARKS = "NULL" Else xREMARKS = N2Str2Null(Trim(txtRemarks.Text))

    xVOUCHERNO = N2Str2Null(xVOUCHERNO)
    xJType = N2Str2Null(xJType)
    xInvoiceType = N2Str2Null(lblInvoiceType.Caption)
    xInvoiceNo = N2String(Trim(txtInvoiceNo.Text))
    XCustomerCode = N2Str2Null(txtCusCde.Text)
    xCUSTOMERNAME = N2Str2Null(txtCustomerName.Text)
    xINVOICEAMOUNT = NumericVal(txtInvoiceAmt.Text)
    xACCT_CODE = N2Str2Null(xACCT_CODE)
    xInvoicedate = (txtInvoiceDate)
    xLAST_UPDATED = N2Str2Null(LOGDATE)
    xJdate = (xJdate)
    xDUEDATE = N2Str2Null(txtDueDate.Text)
    xEntityClass = N2Str2Null(xEntityClass)
    

    
    If Null2String(xEntityClass) = "" Then
        Dim rsEntityCode As ADODB.Recordset
        Set rsEntityCode = New ADODB.Recordset
        Set rsEntityCode = gconDMIS.Execute("SELECT ENTITYCODE FROM ALL_ENTITY WHERE CODE = " & XCustomerCode & " AND ACCOUNTNAME = " & xCUSTOMERNAME & "")
        If Not rsEntityCode.EOF And Not rsEntityCode.BOF Then
            xEntityClass = N2Str2Null(rsEntityCode!ENTITYCODE)
        End If
    End If
    xSJVOUCHERNO = N2Str2Null(lblSJVoucherNo.Caption)

    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(xACCT_CODE) = True Then
    'ar here
    
        rsAR.Open "SELECT AMOUNT_TOPAY-ISNULL((SELECT SUM(INVOICEAMOUNT) FROM AMIS_DETAIL WHERE INVOICENO=AR.INVOICENO AND INVOICETYPE=AR.INVOICETYPE AND ACCT_CODE=AR.ACCOUNT_CODE AND CUSTOMERCODE=AR.CUSTOMERCODE),0) AS BALANCE FROM AMIS_AR AR WHERE INVOICETYPE=" & xInvoiceType & " AND LTRIM(INVOICENO)='" & xInvoiceNo & "' AND CUSTOMERCODE = " & XCustomerCode & " AND ACCOUNT_CODE=" & xACCT_CODE & "", gconDMIS, adOpenForwardOnly
            If xInvoicedate < xJdate Then
                If MsgBox("Invoice date is ahead from journal date. Would you like to proceed?", vbInformation + vbYesNo, "System Message") = vbYes Then
                Else
                    Exit Sub
                End If
            End If
    
    Else
    'ap here
        rsAR.Open "SELECT AMOUNT2PAY-ISNULL((SELECT SUM(AMOUNTPAID) FROM AMIS_DETAILS WHERE INVOICENO=AP.INVOICENO AND ACCT_CODE=AP.ACCT_CODE AND VENDORCODE=AP.VENDOR_CODE AND PV_VOUCHERNO=AP.VOUCHERNO),0) AS BALANCE FROM AMIS_AP AP WHERE isnull(INVOICETYPE,'NULL') in(" & xInvoiceType & ",'NULL') and INVOICENO='" & xInvoiceNo & "' AND ACCT_CODE=" & xACCT_CODE & " AND VENDOR_CODE = " & XCustomerCode & " ", gconDMIS, adOpenForwardOnly
            If xInvoicedate < xJdate Then
                If MsgBox("Invoice date is ahead from journal date. Would you like to proceed?", vbInformation + vbYesNo, "System Message") = vbYes Then
                Else
                Exit Sub
                End If
            End If
            
    End If

    If AddorEdit = "ADD" Then
        If Not rsAR.EOF And Not rsAR.BOF Then
            If CheckIfARAccount(xACCT_CODE) = True And xdebit = 0 Then
                If rsAR!BALANCE < xINVOICEAMOUNT Then
                    MsgBox "Please check!" & Chr(13) & "Payment is greater than invoice amount!" & Chr(13) & "Current Balance: " & N2String(ToDoubleNumber(rsAR!BALANCE)), vbInformation, "System Message"
                    Exit Sub
                End If
                SQL_STATEMENT = "INSERT INTO AMIS_DETAIL(VOUCHERNO,JTYPE,SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,INVOICEAMOUNT,ACCT_CODE,INVOICEDATE,MODIFIED_DATE,JDATE,REMARKS,ENTITYCODE,REFCODE,JOURNAL_DET_ID) " & _
                                "VALUES(" & xVOUCHERNO & "," & xJType & ", " & xSJVOUCHERNO & "," & xInvoiceType & ",'" & xInvoiceNo & "'," & XCustomerCode & "," & xINVOICEAMOUNT & "," & xACCT_CODE & "," & N2Str2Null(xInvoicedate) & "," & xLAST_UPDATED & "," & N2Str2Null(xJdate) & "," & xREMARKS & "," & xEntityClass & "," & xREFCODE & "," & xJOURNAL_DET_ID & ")"
            ElseIf CheckIfARAccount(xACCT_CODE) = False And xdebit > 0 Then
                If rsAR!BALANCE < xINVOICEAMOUNT Then
                    MsgBox "Please check!" & Chr(13) & "Payment is greater than invoice amount!", vbInformation, "System Message"
                    Exit Sub
                End If
                SQL_STATEMENT = "INSERT INTO AMIS_DETAILS(VOUCHERNO,JTYPE,PV_VOUCHERNO,INVOICETYPE,INVOICENO,VENDORCODE,AMOUNTPAID,ACCT_CODE,INVOICEDATE,JDATE,REMARKS,ENTITYCODE,REFCODE,JOURNAL_DET_ID) " & _
                                "VALUES(" & xVOUCHERNO & "," & xJType & ", " & xSJVOUCHERNO & "," & xInvoiceType & ",'" & xInvoiceNo & "'," & XCustomerCode & "," & xINVOICEAMOUNT & "," & xACCT_CODE & "," & N2Str2Null(xInvoicedate) & "," & N2Str2Null(xJdate) & "," & xREMARKS & "," & xEntityClass & "," & xREFCODE & "," & xJOURNAL_DET_ID & ")"

            End If
            gconDMIS.Execute SQL_STATEMENT
            Call rsRefresh(xVOUCHERNO, xJType, xACCT_CODE, CURRENT_CUSCODE, xJOURNAL_DET_ID)
            Call FillGrid(xVOUCHERNO)
        Else
            MessagePop RecSaveError, "System Message", "No such record."
        End If
    ElseIf AddorEdit = "EDIT" Then
        If Not rsAR.EOF And Not rsAR.BOF Then
            If lstDetails.ListItems.Count > 0 Then
                If CheckIfARAccount(xACCT_CODE) = True Then
                    SQL_STATEMENT = "UPDATE AMIS_DETAIL SET CUSTOMERCODE=" & XCustomerCode & ",INVOICENO='" & xInvoiceNo & "',INVOICETYPE=" & xInvoiceType & ",INVOICEAMOUNT=" & xINVOICEAMOUNT & ",INVOICEDATE=" & N2Str2Null(xInvoicedate) & ",REMARKS=" & xREMARKS & ",ENTITYCODE=" & xEntityClass & ",REFCODE=" & xREFCODE & ", JOURNAL_DET_ID = " & xJOURNAL_DET_ID & " WHERE ID=" & lstDetails.SelectedItem.SubItems(6) & ""
                Else
                    SQL_STATEMENT = "UPDATE AMIS_DETAILS SET VENDORCODE=" & XCustomerCode & ",INVOICENO='" & xInvoiceNo & "',INVOICETYPE=" & xInvoiceType & ",AMOUNTPAID=" & xINVOICEAMOUNT & ",INVOICEDATE=" & N2Str2Null(xInvoicedate) & ",REMARKS=" & xREMARKS & ",ENTITYCODE=" & xEntityClass & ",REFCODE=" & xREFCODE & ", JOURNAL_DET_ID = " & xJOURNAL_DET_ID & "  WHERE ID=" & lstDetails.SelectedItem.SubItems(6) & ""
                End If
                gconDMIS.Execute SQL_STATEMENT
                Call rsRefresh(xVOUCHERNO, xJType, xACCT_CODE, CURRENT_CUSCODE, xJOURNAL_DET_ID)
                Call FillGrid(xVOUCHERNO)
            End If
        Else
            MessagePop RecSaveError, "System Message", "No such record."
        End If
    End If
    SQL_STATEMENT = "UPDATE AMIS_JOURNAL_DET SET ADJ_REMARKS = " & xREMARKS & " WHERE ID = " & xJOURNAL_DET_ID & ""
    'gconDMIS.Execute SQL_STATEMENT
    initMemvars
End Sub

Private Sub cmdSelect_Click()
    Set frmNewEntity = New frmEntity
    Call frmNewEntity.LOADJOURNAL("SCHED")
    frmNewEntity.Show 1
End Sub

Sub initMemvars()
    PAYCODE = ""
    txtCusCde.Text = ""
    txtCustomerName.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate.Text = xJdate
    txtDueDate.Text = LOGDATE
    txtInvoiceAmt.Text = ZERO
    lblInvoiceType.Caption = ""
    lblBalance.Caption = ZERO
    lblEntityClass.Caption = ""
    lblSJVoucherNo.Caption = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initMemvars
    picDetails.Enabled = True
    Call rsRefresh(xVOUCHERNO, xJType, xACCT_CODE, CURRENT_CUSCODE, xJOURNAL_DET_ID)
    Call FillGrid(xVOUCHERNO)
        If lstDetails.ListItems.Count > 0 Then
        Call StoreMemVars(lstDetails.ListItems.Item(1).SubItems(6))
    End If
    txtInvoiceDate.Enabled = False
  End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If NumericVal(lblAmount.text) <> NumericVal(frmAMISJournalEntry_APJ.txtDebit.Text) Then
'        If MsgBox("GL Amount is not equal to SL, proceed?", vbYesNo + vbQuestion, "System Message") = vbNo Then
'            Cancel = 1
'        Else
'            Unload Me
'        End If
'    Else
'        Unload Me
'    End If
End Sub

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    txtCusCde.Text = strCode
    txtCustomerName.Text = strAccountName
    lblEntityClass.Caption = strEntityClass
    If lblEntityClass.Caption = "V" Then
        PAYCODE = Vendor_PayCode(txtCusCde.Text)
        If IsDate(txtInvoiceDate.Text) = True Then
            txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
        End If
    End If
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    AddorEdit = "EDIT"
    StoreMemVars (lstDetails.SelectedItem.SubItems(6))

End Sub

Private Sub tmReal_Timer()
lblRealSJVoucherNo.ForeColor = &H80000002
tmrSarah.Enabled = True
tmReal.Enabled = False
End Sub

Private Sub tmrSarah_Timer()
lblRealSJVoucherNo.ForeColor = &H80&
tmrSarah.Enabled = False
tmReal.Enabled = True
End Sub

Private Sub txtInvoiceAmt_GotFocus()
    If txtInvoiceAmt.Text = "0.00" Or txtInvoiceAmt.Text = "" Then
        txtInvoiceAmt.Text = ""
    Else
        txtInvoiceAmt.Text = NumericVal(txtInvoiceAmt.Text)
    End If
End Sub

Private Sub txtInvoiceAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtInvoiceAmt_LostFocus()
    txtInvoiceAmt.Text = ToDoubleNumber(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceDate_Change()
    On Error Resume Next
    If IsDate(txtInvoiceDate.Text) = True Then
        txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
    End If
End Sub

Private Sub txtInvoiceDate_GotFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_LostFocus()
    If txtRemarks.Text = "" Then txtRemarks.Text = "Pls Type Your Message Here!"
End Sub
Sub LOAD_DATA(VOUCHERNO As String, JTYPE As String, AcctCode As String, JDATE As String, amount As Double, ENTITYCLASS As String, CUSCDE As String, Debit As Double, Journal_Det_ID As Long)
    xVOUCHERNO = N2Str2Null(VOUCHERNO)
    xJType = N2Str2Null(JTYPE)
    xACCT_CODE = N2Str2Null(AcctCode)
    xJdate = Format(JDATE, "mm/dd/yyyy")
    xAMOUNT = amount
    xdebit = Debit
    xREFCODE = N2Str2Null(ENTITYCLASS + CUSCDE)
    xEntityClass = ENTITYCLASS
    CURRENT_CUSCODE = CUSCDE
    xJOURNAL_DET_ID = Journal_Det_ID
End Sub


Function FillGrid(XXX)
    cntDetails = 0
    lstDetails.ListItems.Clear
    lblAmount.Text = "0.00"
    If Not rsAR.EOF And Not rsAR.BOF Then
        cntDetails = 1
        Do While Not rsAR.EOF
            Set LST = lstDetails.ListItems.Add(, , Format(cntDetails, "0000"))
            LST.SubItems(1) = Null2String(rsAR!CustomerCode)
            LST.SubItems(2) = Null2String(rsAR!INVOICENO)
            LST.SubItems(3) = ToDoubleNumber(rsAR!invoiceamount)
            LST.SubItems(4) = Null2String(rsAR!invoicedate)
            LST.SubItems(5) = Null2String(rsAR!remarks)
            LST.SubItems(6) = Null2String(rsAR!ID)
            lblAmount = lblAmount + rsAR!invoiceamount
            rsAR.MoveNext
            cntDetails = cntDetails + 1
        Loop
                
       cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    lblAmount = ToDoubleNumber(lblAmount)
End Function

Private Sub StoreMemVars(ID As String)
    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(xACCT_CODE) = True Then
        rsAR.Open "SELECT CUSTOMERCODE,INVOICENO,INVOICETYPE,INVOICEDATE,INVOICEAMOUNT,REMARKS,ENTITYCODE,NULL AS DUEDATE,ID FROM AMIS_DETAIL WHERE ID=" & ID & "", gconDMIS, adOpenForwardOnly
    Else
        rsAR.Open "SELECT VENDORCODE AS CUSTOMERCODE,INVOICENO,INVOICETYPE,INVOICEDATE,AMOUNTPAID AS INVOICEAMOUNT,REMARKS,ENTITYCODE,NULL AS DUEDATE,ID FROM AMIS_DETAILS WHERE ID=" & ID & "", gconDMIS, adOpenForwardOnly
    End If
    If Not rsAR.EOF And Not rsAR.BOF Then
        txtCusCde.Text = Null2String(rsAR!CustomerCode)
        txtCustomerName.Text = ACCOUNTNAME(Null2String(rsAR!ENTITYCODE), Null2String(rsAR!CustomerCode))
        txtInvoiceNo.Text = Null2String(rsAR!INVOICENO)
        lblInvoiceType.Caption = Null2String(rsAR!INVOICETYPE)
        txtInvoiceDate.Text = Null2String(rsAR!invoicedate)
        txtInvoiceAmt.Text = ToDoubleNumber(rsAR!invoiceamount)
        txtRemarks.Text = Null2String(rsAR!remarks)
        lblEntityClass.Caption = Null2String(rsAR!ENTITYCODE)
        txtDueDate.Text = Null2String(rsAR!DUEDATE)
    End If
    
        'SJR 07/1/14
      
        Set rsReal = New ADODB.Recordset
        If Left(xACCT_CODE, 3) = "'21" Then 'if AP
                Set rsReal = gconDMIS.Execute("Select * from amis_details where vendorcode='" & Null2String(rsAR!CustomerCode) & "' and invoiceno='" & txtInvoiceNo.Text & "' and acct_code=" & xACCT_CODE & " and voucherno=" & xVOUCHERNO & " and jtype =" & xJType & " and entitycode='" & Null2String(rsAR!ENTITYCODE) & "'")
                If Not rsReal.EOF And Not rsReal.BOF Then
                lblRealSJVoucherNo.Caption = Null2String(rsReal!PV_VoucherNo)
                End If
        ElseIf Left(xACCT_CODE, 3) = "'11" Then 'if AR
                Set rsReal = gconDMIS.Execute("SELECT top 1 ar.sjvoucherno as voucher FROM AMIS_AR AR INNER JOIN AMIS_DETAIL DT ON AR.CUSTOMERCODE=DT.CUSTOMERCODE AND AR.INVOICENO=DT.INVOICENO AND AR.INVOICETYPE=DT.INVOICETYPE AND AR.ACCOUNT_CODE= DT.ACCT_CODE WHERE dt.customercode='" & Null2String(rsAR!CustomerCode) & "' and dt.invoiceno='" & txtInvoiceNo.Text & "' and dt.acct_code=" & xACCT_CODE & " and dt.voucherno=" & xVOUCHERNO & " and dt.jtype =" & xJType & " and dt.entitycode='" & Null2String(rsAR!ENTITYCODE) & "'")
                If Not rsReal.EOF And Not rsReal.BOF Then
                lblRealSJVoucherNo.Caption = Null2String(rsReal!Voucher)
                End If
        Else
                lblRealSJVoucherNo.Caption = ""
        End If
        
        'SJR 07/1/14

    End Sub

Private Sub rsRefresh(VOUCHERNO As String, JTYPE As String, AcctCode As String, CUSCDE As String, Journal_Det_ID As Long)
    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(AcctCode) = True And xdebit = 0 Then
        If Journal_Det_ID > 0 Then
            rsAR.Open "SELECT CUSTOMERCODE,INVOICENO,INVOICEAMOUNT,INVOICEDATE,REMARKS,ID FROM AMIS_DETAIL WHERE VOUCHERNO=" & VOUCHERNO & " AND JTYPE= " & JTYPE & " AND ACCT_CODE =" & AcctCode & " AND JOURNAL_DET_ID = " & Journal_Det_ID & " ORDER BY ID", gconDMIS, adOpenForwardOnly
        Else
            rsAR.Open "SELECT CUSTOMERCODE,INVOICENO,INVOICEAMOUNT,INVOICEDATE,REMARKS,ID FROM AMIS_DETAIL WHERE VOUCHERNO=" & VOUCHERNO & " AND JTYPE= " & JTYPE & " AND ACCT_CODE =" & AcctCode & "  ORDER BY ID", gconDMIS, adOpenForwardOnly
        End If
    ElseIf CheckIfARAccount(AcctCode) = False And xdebit > 0 Then
        If Journal_Det_ID > 0 Then
            rsAR.Open "SELECT VENDORCODE AS CUSTOMERCODE,INVOICENO,AMOUNTPAID AS INVOICEAMOUNT,INVOICEDATE,REMARKS,ID FROM AMIS_DETAILS WHERE VOUCHERNO=" & VOUCHERNO & " AND JTYPE= " & JTYPE & " AND ACCT_CODE =" & AcctCode & " AND JOURNAL_DET_ID = " & Journal_Det_ID & "  ORDER BY ID", gconDMIS, adOpenForwardOnly
        Else
            rsAR.Open "SELECT VENDORCODE AS CUSTOMERCODE,INVOICENO,AMOUNTPAID AS INVOICEAMOUNT,INVOICEDATE,REMARKS,ID FROM AMIS_DETAILS WHERE VOUCHERNO=" & VOUCHERNO & " AND JTYPE= " & JTYPE & " AND ACCT_CODE =" & AcctCode & "  ORDER BY ID", gconDMIS, adOpenForwardOnly
        End If
    End If
End Sub

Function Vendor_PayCode(XXX As String) As String
    Dim rsVENDOR                                            As ADODB.Recordset
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

Function SetPayNoDays(VVV As Variant) As Integer
    Dim rsPayTerm                                           As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_Desc,no_days from ALL_PayTerm where PAY_CODE = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayNoDays = Null2String(rsPayTerm!no_Days)
    Else
        SetPayNoDays = 0
    End If
End Function

Function ACCOUNTNAME(ENTITYCODE As String, Code As String) As String
    Dim rsAccountName                                       As ADODB.Recordset
    Set rsAccountName = New ADODB.Recordset
    rsAccountName.Open "SELECT ACCOUNTNAME FROM ALL_ENTITY WHERE ENTITYCODE ='" & ENTITYCODE & "' AND CODE = '" & Code & "'", gconDMIS, adOpenForwardOnly
    If Not rsAccountName.EOF And Not rsAccountName.BOF Then
        ACCOUNTNAME = rsAccountName!ACCOUNTNAME
    End If
    Set rsAccountName = Nothing
End Function




