VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSTables_Tax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Withholding Tax Table"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMSTablesTax.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   8610
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   6240
      ScaleHeight     =   855
      ScaleWidth      =   2415
      TabIndex        =   33
      Top             =   7170
      Width           =   2415
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
         Left            =   1560
         MouseIcon       =   "frmHRMSTablesTax.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSTablesTax.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Exit Window"
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
         Left            =   870
         MouseIcon       =   "frmHRMSTablesTax.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSTablesTax.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picTable 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   30
      ScaleHeight     =   4965
      ScaleWidth      =   8535
      TabIndex        =   5
      Top             =   2100
      Width           =   8565
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1785
         Left            =   60
         ScaleHeight     =   1755
         ScaleWidth      =   8415
         TabIndex        =   39
         Top             =   1770
         Visible         =   0   'False
         Width           =   8445
         Begin VB.CommandButton cmdCEdit 
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
            Left            =   7650
            MouseIcon       =   "frmHRMSTablesTax.frx":0EF0
            MousePointer    =   99  'Custom
            Picture         =   "frmHRMSTablesTax.frx":1042
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Cancel"
            Top             =   900
            Width           =   705
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   1110
            TabIndex        =   40
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   2010
            TabIndex        =   42
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   2910
            TabIndex        =   44
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   3810
            TabIndex        =   46
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   4
            Left            =   4680
            TabIndex        =   48
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   5580
            TabIndex        =   50
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   6450
            TabIndex        =   52
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   7350
            TabIndex        =   54
            Top             =   420
            Width           =   855
         End
         Begin VB.CommandButton cmdSEdit 
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
            Left            =   6960
            MouseIcon       =   "frmHRMSTablesTax.frx":1380
            MousePointer    =   99  'Custom
            Picture         =   "frmHRMSTablesTax.frx":14D2
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Save Entry"
            Top             =   900
            Width           =   705
         End
         Begin VB.Label lblCode 
            BackColor       =   &H00C0FFFF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1380
            Left            =   60
            TabIndex        =   58
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   6
            Left            =   5040
            TabIndex        =   57
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   6
            Left            =   4110
            TabIndex        =   53
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   5910
            TabIndex        =   51
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   5
            Left            =   2310
            TabIndex        =   49
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   5
            Left            =   1440
            TabIndex        =   47
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   3270
            TabIndex        =   45
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   7680
            TabIndex        =   43
            Top             =   120
            Width           =   120
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   6780
            TabIndex        =   41
            Top             =   120
            Width           =   120
         End
      End
      Begin MSComctlLib.ListView lsvSingle 
         Height          =   1305
         Left            =   60
         TabIndex        =   6
         Top             =   360
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmHRMSTablesTax.frx":1822
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "1"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "2"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "3"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "4"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "5"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "6"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "7"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "8"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lsvMarried 
         Height          =   1305
         Left            =   60
         TabIndex        =   60
         Top             =   3600
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmHRMSTablesTax.frx":1984
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "1"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "2"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "3"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "4"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "5"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "6"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "7"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "8"
            Object.Width           =   1587
         EndProperty
      End
      Begin MSComctlLib.ListView lsvHead 
         Height          =   1305
         Left            =   60
         TabIndex        =   61
         Top             =   1980
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   2302
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmHRMSTablesTax.frx":1AE6
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "1"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "2"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "3"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "4"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "5"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "6"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "7"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "8"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Table For Married employee with Qualified dependent child(ren)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   60
         TabIndex        =   63
         Top             =   3330
         Width           =   6945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Table For Head of theFamily Employee with dependent child(ren)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   60
         TabIndex        =   62
         Top             =   1710
         Width           =   7095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Table For employee without Dependent Children "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   59
         Top             =   90
         Width           =   5295
      End
   End
   Begin VB.PictureBox picSSSTable 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   30
      ScaleHeight     =   1995
      ScaleWidth      =   8520
      TabIndex        =   0
      Top             =   0
      Width           =   8550
      Begin VB.PictureBox picCol 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1005
         Left            =   1230
         ScaleHeight     =   975
         ScaleWidth      =   7215
         TabIndex        =   16
         Top             =   930
         Width           =   7245
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   6330
            TabIndex        =   32
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   5430
            TabIndex        =   31
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   4560
            TabIndex        =   30
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   4
            Left            =   3660
            TabIndex        =   29
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   2790
            TabIndex        =   28
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   1890
            TabIndex        =   27
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   990
            TabIndex        =   26
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtPer 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   540
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   6330
            TabIndex        =   24
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   6
            Left            =   5430
            TabIndex        =   23
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   5
            Left            =   4560
            TabIndex        =   22
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   4
            Left            =   3660
            TabIndex        =   21
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   3
            Left            =   2790
            TabIndex        =   20
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   1890
            TabIndex        =   19
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   1
            Left            =   990
            TabIndex        =   18
            Top             =   90
            Width           =   855
         End
         Begin VB.TextBox txtExem 
            Alignment       =   1  'Right Justify
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
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   90
            Width           =   855
         End
      End
      Begin VB.ComboBox cboBasis 
         Height          =   360
         ItemData        =   "frmHRMSTablesTax.frx":1C48
         Left            =   1290
         List            =   "frmHRMSTablesTax.frx":1C58
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   150
         Width           =   5985
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   7080
         TabIndex        =   15
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   7980
         TabIndex        =   14
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   3510
         TabIndex        =   13
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   1650
         TabIndex        =   12
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   2610
         TabIndex        =   11
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   6240
         TabIndex        =   10
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   4380
         TabIndex        =   9
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   5280
         TabIndex        =   8
         Top             =   660
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Exemptions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   4
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Columns"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   330
         TabIndex        =   3
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Basis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   630
         TabIndex        =   1
         Top             =   270
         Width           =   495
      End
   End
   Begin Crystal.CrystalReport rptSalaryGrade 
      Left            =   180
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7020
      ScaleHeight     =   885
      ScaleWidth      =   1695
      TabIndex        =   36
      Top             =   7170
      Visible         =   0   'False
      Width           =   1695
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
         MouseIcon       =   "frmHRMSTablesTax.frx":1C96
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSTablesTax.frx":1DE8
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   75
         MouseIcon       =   "frmHRMSTablesTax.frx":2126
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMSTablesTax.frx":2278
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSTables_Tax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub DisplayBasisInformation()
    Dim RSTMP                                                         As New ADODB.Recordset
    Dim ITEM                                                          As ListItem

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_TaxTableDetails Where TaxBasis = '" & cboBasis.Text & "' And Status = '" & "S" & "' Order By TaxCode Asc")
    lsvSingle.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvSingle.ListItems.Add(, , RSTMP!TAXCODE)
            ITEM.SubItems(1) = Format(RSTMP!Col1, "#,###,##0.00")
            ITEM.SubItems(2) = Format(RSTMP!Col2, "#,###,##0.00")
            ITEM.SubItems(3) = Format(RSTMP!Col3, "#,###,##0.00")
            ITEM.SubItems(4) = Format(RSTMP!Col4, "#,###,##0.00")
            ITEM.SubItems(5) = Format(RSTMP!Col5, "#,###,##0.00")
            ITEM.SubItems(6) = Format(RSTMP!Col6, "#,###,##0.00")
            ITEM.SubItems(7) = Format(RSTMP!Col7, "#,###,##0.00")
            ITEM.SubItems(8) = Format(RSTMP!Col8, "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_TaxTableDetails Where TaxBasis = '" & cboBasis.Text & "' And Status = '" & "H" & "' Order By TaxCode Asc")
    lsvHead.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvHead.ListItems.Add(, , RSTMP!TAXCODE)
            ITEM.SubItems(1) = Format(RSTMP!Col1, "#,###,##0.00")
            ITEM.SubItems(2) = Format(RSTMP!Col2, "#,###,##0.00")
            ITEM.SubItems(3) = Format(RSTMP!Col3, "#,###,##0.00")
            ITEM.SubItems(4) = Format(RSTMP!Col4, "#,###,##0.00")
            ITEM.SubItems(5) = Format(RSTMP!Col5, "#,###,##0.00")
            ITEM.SubItems(6) = Format(RSTMP!Col6, "#,###,##0.00")
            ITEM.SubItems(7) = Format(RSTMP!Col7, "#,###,##0.00")
            ITEM.SubItems(8) = Format(RSTMP!Col8, "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If

    Set RSTMP = gconDMIS.Execute("Select * From HRMS_TaxTableDetails Where TaxBasis = '" & cboBasis.Text & "' And Status = '" & "M" & "' Order By TaxCode Asc")
    lsvMarried.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set ITEM = lsvMarried.ListItems.Add(, , RSTMP!TAXCODE)
            ITEM.SubItems(1) = Format(RSTMP!Col1, "#,###,##0.00")
            ITEM.SubItems(2) = Format(RSTMP!Col2, "#,###,##0.00")
            ITEM.SubItems(3) = Format(RSTMP!Col3, "#,###,##0.00")
            ITEM.SubItems(4) = Format(RSTMP!Col4, "#,###,##0.00")
            ITEM.SubItems(5) = Format(RSTMP!Col5, "#,###,##0.00")
            ITEM.SubItems(6) = Format(RSTMP!Col6, "#,###,##0.00")
            ITEM.SubItems(7) = Format(RSTMP!Col7, "#,###,##0.00")
            ITEM.SubItems(8) = Format(RSTMP!Col8, "#,###,##0.00")

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Sub FillExemAndPerc()
    Dim RSTMP                                                         As New ADODB.Recordset

    Set RSTMP = gconDMIS.Execute("Select * from HRMS_TaxTable Where TaxBasis = '" & cboBasis.Text & "'")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        txtExem(0).Text = Format(RSTMP!EXP1, "#,###,##0.00")
        txtExem(1).Text = Format(RSTMP!EXp2, "#,###,##0.00")
        txtExem(2).Text = Format(RSTMP!EXp3, "#,###,##0.00")
        txtExem(3).Text = Format(RSTMP!EXp4, "#,###,##0.00")
        txtExem(4).Text = Format(RSTMP!EXp5, "#,###,##0.00")
        txtExem(5).Text = Format(RSTMP!EXp6, "#,###,##0.00")
        txtExem(6).Text = Format(RSTMP!EXp7, "#,###,##0.00")
        txtExem(7).Text = Format(RSTMP!EXp8, "#,###,##0.00")
        txtPer(0).Text = Format(RSTMP!Per1, "#,###,##0.00")
        txtPer(1).Text = Format(RSTMP!Per2, "#,###,##0.00")
        txtPer(2).Text = Format(RSTMP!Per3, "#,###,##0.00")
        txtPer(3).Text = Format(RSTMP!Per4, "#,###,##0.00")
        txtPer(4).Text = Format(RSTMP!Per5, "#,###,##0.00")
        txtPer(5).Text = Format(RSTMP!Per6, "#,###,##0.00")
        txtPer(6).Text = Format(RSTMP!Per7, "#,###,##0.00")
        txtPer(7).Text = Format(RSTMP!Per8, "#,###,##0.00")
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cboBasis_Change()
    Call DisplayBasisInformation
    Call FillExemAndPerc
End Sub

Private Sub cboBasis_Click()
    Call DisplayBasisInformation
    Call FillExemAndPerc
End Sub

Private Sub cboBasis_LostFocus()
    Call DisplayBasisInformation
    Call FillExemAndPerc
End Sub

Private Sub cmdCancel_Click()
    Picture1.Visible = True
    Picture2.Visible = False
    picTable.Enabled = False
    cboBasis.Enabled = True

    Call cboBasis_Click
End Sub

Private Sub cmdCEdit_Click()
    picEdit.Visible = False
    lsvSingle.Enabled = True
    lsvMarried.Enabled = True
    lsvHead.Enabled = True

    Picture2.Visible = True
    picCol.Enabled = True
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "ACESS_EDIT", "TABLE BIR") = False Then Exit Sub
    picTable.Enabled = True
    picCol.Enabled = True
    cboBasis.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Save Exemptions And Percentage", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        gconDMIS.Execute ("Update HRMS_TaxTable Set EXP1 = " & CDbl(txtExem(0).Text) & _
                          ",EXP2 = " & CDbl(txtExem(1).Text) & _
                          ",EXP3 = " & CDbl(txtExem(2).Text) & _
                          ",EXP4 = " & CDbl(txtExem(3).Text) & _
                          ",EXP5 = " & CDbl(txtExem(4).Text) & _
                          ",EXP6 = " & CDbl(txtExem(5).Text) & _
                          ",EXP7 = " & CDbl(txtExem(6).Text) & _
                          ",EXP8 = " & CDbl(txtExem(7).Text) & _
                          ",PER1 = " & CDbl(txtPer(0).Text) & _
                          ",PER2 = " & CDbl(txtPer(1).Text) & _
                          ",PER3 = " & CDbl(txtPer(2).Text) & _
                          ",PER4 = " & CDbl(txtPer(3).Text) & _
                          ",PER5 = " & CDbl(txtPer(4).Text) & _
                          ",PER6 = " & CDbl(txtPer(5).Text) & _
                          ",PER7 = " & CDbl(txtPer(6).Text) & _
                          ",PER8 = " & CDbl(txtPer(7).Text) & _
                        " Where TaxBasis = '" & cboBasis.Text & "'")

        Call LogAudit("E", "UPDATE TAX TABLE", cboBasis.Text)
        Call ShowSuccessFullyUpdated

        Picture1.Visible = True
        Picture2.Visible = False
        picTable.Enabled = False
        picCol.Enabled = False
        cboBasis.Enabled = True

        Call cboBasis_Click
    End If
End Sub

Private Sub cmdSEdit_Click()
    If MsgBox("Update Tax Table", vbQuestion + vbYesNo, "Confirm Update") = vbYes Then
        gconDMIS.Execute ("Update HRMS_TaxTableDetails Set Col1 = " & txtEdit(0).Text & _
                          ",Col2 = " & NumericVal(txtEdit(1).Text) & _
                          ",Col3 = " & NumericVal(txtEdit(2).Text) & _
                          ",Col4 = " & NumericVal(txtEdit(3).Text) & _
                          ",Col5 = " & NumericVal(txtEdit(4).Text) & _
                          ",Col6 = " & NumericVal(txtEdit(5).Text) & _
                          ",Col7 = " & NumericVal(txtEdit(6).Text) & _
                          ",Col8 = " & NumericVal(txtEdit(7).Text) & _
                        " Where TaxBasis = '" & cboBasis.Text & "' And TaxCode = '" & lblCode.Caption & "'")

        picEdit.Visible = False
        lsvSingle.Enabled = True
        lsvHead.Enabled = True
        lsvMarried.Enabled = True
        Picture2.Visible = True
        picCol.Enabled = True

        Call ShowSuccessFullyUpdated
        Call DisplayBasisInformation

        Call LogAudit("E", "UPDATE TAX TABLE DETAILS", cboBasis.Text)
    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    cboBasis.ListIndex = 0
End Sub

Private Sub lsvHead_DblClick()
    Dim INDEX                                                         As Double

    If Not lsvHead.ListItems.count = 0 Then
        With lsvHead
            Picture2.Visible = False
            INDEX = .SelectedItem.INDEX
            .Enabled = False
            picCol.Enabled = False
            lsvSingle.Enabled = False
            lsvMarried.Enabled = False

            picEdit.Visible = True

            lblCode.Caption = .ListItems(INDEX).Text
            txtEdit(0).Text = .ListItems(INDEX).ListSubItems(1)
            txtEdit(1).Text = .ListItems(INDEX).ListSubItems(2)
            txtEdit(2).Text = .ListItems(INDEX).ListSubItems(3)
            txtEdit(3).Text = .ListItems(INDEX).ListSubItems(4)
            txtEdit(4).Text = .ListItems(INDEX).ListSubItems(5)
            txtEdit(5).Text = .ListItems(INDEX).ListSubItems(6)
            txtEdit(6).Text = .ListItems(INDEX).ListSubItems(7)
            txtEdit(7).Text = .ListItems(INDEX).ListSubItems(8)
        End With
    End If
End Sub

Private Sub lsvMarried_DblClick()
    Dim INDEX                                                         As Double

    If Not lsvMarried.ListItems.count = 0 Then
        With lsvMarried
            Picture2.Visible = False
            INDEX = .SelectedItem.INDEX
            .Enabled = False
            picCol.Enabled = False
            lsvSingle.Enabled = False
            lsvMarried.Enabled = False

            picEdit.Visible = True

            lblCode.Caption = .ListItems(INDEX).Text
            txtEdit(0).Text = .ListItems(INDEX).ListSubItems(1)
            txtEdit(1).Text = .ListItems(INDEX).ListSubItems(2)
            txtEdit(2).Text = .ListItems(INDEX).ListSubItems(3)
            txtEdit(3).Text = .ListItems(INDEX).ListSubItems(4)
            txtEdit(4).Text = .ListItems(INDEX).ListSubItems(5)
            txtEdit(5).Text = .ListItems(INDEX).ListSubItems(6)
            txtEdit(6).Text = .ListItems(INDEX).ListSubItems(7)
            txtEdit(7).Text = .ListItems(INDEX).ListSubItems(8)
        End With
    End If
End Sub

Private Sub lsvSingle_DblClick()
    Dim INDEX                                                         As Double

    If Not lsvSingle.ListItems.count = 0 Then
        With lsvSingle
            Picture2.Visible = False
            INDEX = .SelectedItem.INDEX
            .Enabled = False
            picCol.Enabled = False
            lsvHead.Enabled = False
            lsvMarried.Enabled = False

            picEdit.Visible = True

            lblCode.Caption = .ListItems(INDEX).Text
            txtEdit(0).Text = .ListItems(INDEX).ListSubItems(1)
            txtEdit(1).Text = .ListItems(INDEX).ListSubItems(2)
            txtEdit(2).Text = .ListItems(INDEX).ListSubItems(3)
            txtEdit(3).Text = .ListItems(INDEX).ListSubItems(4)
            txtEdit(4).Text = .ListItems(INDEX).ListSubItems(5)
            txtEdit(5).Text = .ListItems(INDEX).ListSubItems(6)
            txtEdit(6).Text = .ListItems(INDEX).ListSubItems(7)
            txtEdit(7).Text = .ListItems(INDEX).ListSubItems(8)
        End With
    End If
End Sub

Private Sub txtEdit_KeyUp(INDEX As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then txtEdit(INDEX).Text = ""
End Sub

