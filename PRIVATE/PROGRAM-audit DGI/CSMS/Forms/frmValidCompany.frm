VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form FrmValid_Company 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "test"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   10200
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   9210
      Top             =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   9690
      Top             =   690
   End
   Begin VB.PictureBox picAddframe 
      BackColor       =   &H00FF8080&
      Height          =   5355
      Left            =   1590
      ScaleHeight     =   5295
      ScaleWidth      =   7035
      TabIndex        =   8
      Top             =   1020
      Width           =   7095
      Begin VB.Frame frmLISt 
         BackColor       =   &H00FF8080&
         Height          =   2685
         Left            =   30
         TabIndex        =   15
         Top             =   1740
         Width           =   6975
         Begin MSComctlLib.ListView lstlist 
            Height          =   2505
            Left            =   30
            TabIndex        =   30
            Top             =   120
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   4419
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
            Appearance      =   1
            MousePointer    =   99
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmValidCompany.frx":0000
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   9703
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Object.Width           =   0
            EndProperty
         End
      End
      Begin VB.Frame frmmiddle 
         BackColor       =   &H00FF8080&
         Height          =   705
         Left            =   30
         TabIndex        =   13
         Top             =   1020
         Width           =   6975
         Begin VB.TextBox txtfind 
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
            Height          =   435
            Left            =   90
            TabIndex        =   14
            Top             =   180
            Width           =   6795
         End
      End
      Begin VB.Frame frmTop 
         BackColor       =   &H00FF8080&
         Enabled         =   0   'False
         Height          =   1035
         Left            =   30
         TabIndex        =   9
         Top             =   -30
         Width           =   6975
         Begin VB.TextBox txtModuletype 
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
            Left            =   1560
            TabIndex        =   31
            Top             =   150
            Width           =   1785
         End
         Begin VB.TextBox txtDescription 
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
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   570
            Width           =   5325
         End
         Begin VB.ComboBox cmbModuletype_add 
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
            Left            =   1560
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   150
            Width           =   1785
         End
         Begin VB.Label labID 
            Height          =   285
            Left            =   6300
            TabIndex        =   29
            Top             =   180
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
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
            Left            =   180
            TabIndex        =   12
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Module Type"
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
            Left            =   150
            TabIndex        =   10
            Top             =   210
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   60
         ScaleHeight     =   855
         ScaleWidth      =   7035
         TabIndex        =   17
         Top             =   4440
         Width           =   7035
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
            Left            =   1410
            MouseIcon       =   "frmValidCompany.frx":0162
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":02B4
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Move to Previous Record"
            Top             =   30
            Visible         =   0   'False
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
            Left            =   2100
            MouseIcon       =   "frmValidCompany.frx":0613
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":0765
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Move to Next Record"
            Top             =   30
            Visible         =   0   'False
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
            Left            =   2790
            MouseIcon       =   "frmValidCompany.frx":0ABD
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":0C0F
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Find a Record"
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
            Left            =   3480
            MouseIcon       =   "frmValidCompany.frx":0F09
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":105B
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Add Record"
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
            Left            =   4170
            MouseIcon       =   "frmValidCompany.frx":136E
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":14C0
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Edit Selected Record"
            Top             =   30
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
            Left            =   4860
            MouseIcon       =   "frmValidCompany.frx":181C
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":196E
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Delete Selected Record"
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
            Left            =   5550
            MouseIcon       =   "frmValidCompany.frx":1C99
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":1DEB
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Print this Record"
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdEscape 
            Caption         =   "E&scape"
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
            Left            =   6240
            MouseIcon       =   "frmValidCompany.frx":2151
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":22A3
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Exit Window"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   5550
         ScaleHeight     =   885
         ScaleWidth      =   2940
         TabIndex        =   26
         Top             =   4440
         Width           =   2940
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
            Left            =   0
            MouseIcon       =   "frmValidCompany.frx":2609
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":275B
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Save Entry"
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
            Left            =   720
            MouseIcon       =   "frmValidCompany.frx":2AAB
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":2BFD
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Cancel"
            Top             =   30
            Width           =   705
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   7365
      Left            =   30
      ScaleHeight     =   7305
      ScaleWidth      =   10065
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Height          =   825
         Left            =   30
         TabIndex        =   5
         Top             =   6480
         Width           =   9975
         Begin VB.CommandButton Command1 
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
            Height          =   615
            Left            =   8790
            MouseIcon       =   "frmValidCompany.frx":2F3B
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":308D
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Exit Window"
            Top             =   150
            Width           =   1155
         End
         Begin VB.CommandButton CmdGrantAll 
            Caption         =   "&Install All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6540
            MouseIcon       =   "frmValidCompany.frx":33F3
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":3545
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Move to Previous Record"
            Top             =   150
            Width           =   1155
         End
         Begin VB.CommandButton cmddenieall 
            Caption         =   "&Uninstall All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   7680
            MouseIcon       =   "frmValidCompany.frx":38A4
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":39F6
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Move to Next Record"
            Top             =   150
            Width           =   1125
         End
         Begin VB.CommandButton lbladd 
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
            Height          =   615
            Left            =   5280
            MouseIcon       =   "frmValidCompany.frx":3D4E
            MousePointer    =   99  'Custom
            Picture         =   "frmValidCompany.frx":3EA0
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Add Record"
            Top             =   150
            Width           =   1245
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00C0FFFF&
            Height          =   555
            Left            =   1770
            ScaleHeight     =   495
            ScaleWidth      =   3405
            TabIndex        =   47
            Top             =   180
            Width           =   3465
            Begin VB.Label lblmodule 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   30
               TabIndex        =   48
               Top             =   90
               Width           =   3375
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFC0C0&
            Height          =   555
            Left            =   60
            ScaleHeight     =   495
            ScaleWidth      =   1635
            TabIndex        =   44
            Top             =   180
            Width           =   1695
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Module Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   60
               TabIndex        =   46
               Top             =   120
               Width           =   1545
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Module Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   90
               TabIndex        =   45
               Top             =   120
               Width           =   1545
            End
         End
         Begin VB.Label Label6 
            Height          =   465
            Left            =   6270
            TabIndex        =   38
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   -30
         Width           =   10005
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   345
            Left            =   3060
            TabIndex        =   43
            Top             =   180
            Width           =   345
         End
         Begin VB.TextBox txtsearch 
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
            Left            =   5370
            TabIndex        =   42
            Top             =   180
            Width           =   4575
         End
         Begin VB.ComboBox cboModuletype 
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
            Left            =   1260
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   180
            Width           =   1785
         End
         Begin VB.Label Label23 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Module Type"
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
            Left            =   90
            TabIndex        =   4
            Top             =   210
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Search Text Area"
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
            Left            =   3810
            TabIndex        =   3
            Top             =   210
            Width           =   1515
         End
      End
      Begin VB.PictureBox Pic1 
         BackColor       =   &H00FF8080&
         Height          =   5925
         Left            =   30
         ScaleHeight     =   5865
         ScaleWidth      =   9945
         TabIndex        =   6
         Top             =   600
         Width           =   10005
         Begin VB.Frame Frame2 
            Height          =   6015
            Left            =   -30
            TabIndex        =   7
            Top             =   -120
            Width           =   10005
            Begin VB.Frame Frame6 
               Height          =   5535
               Left            =   60
               TabIndex        =   39
               Top             =   450
               Width           =   9885
               Begin MSComctlLib.ListView lstwithAccess 
                  Height          =   5295
                  Left            =   30
                  TabIndex        =   40
                  ToolTipText     =   "Double Click To Uninstall"
                  Top             =   180
                  Width           =   4845
                  _ExtentX        =   8546
                  _ExtentY        =   9340
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   16711680
                  BackColor       =   12632256
                  BorderStyle     =   1
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "frmValidCompany.frx":41B3
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Object.Width           =   8645
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   0
                  EndProperty
               End
               Begin MSComctlLib.ListView lstWithNoAccess 
                  Height          =   5295
                  Left            =   4890
                  TabIndex        =   41
                  ToolTipText     =   "Double Click To Install"
                  Top             =   180
                  Width           =   4935
                  _ExtentX        =   8705
                  _ExtentY        =   9340
                  View            =   3
                  LabelEdit       =   1
                  LabelWrap       =   -1  'True
                  HideSelection   =   -1  'True
                  AllowReorder    =   -1  'True
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   255
                  BackColor       =   12632256
                  BorderStyle     =   1
                  Appearance      =   1
                  MousePointer    =   99
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  MouseIcon       =   "frmValidCompany.frx":4315
                  NumItems        =   2
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Object.Width           =   8644
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Object.Width           =   0
                  EndProperty
               End
            End
            Begin VB.Frame Frame5 
               Height          =   5475
               Left            =   4890
               TabIndex        =   36
               Top             =   480
               Width           =   5055
               Begin XtremeReportControl.ReportControl rptDenied 
                  Height          =   5295
                  Left            =   60
                  TabIndex        =   37
                  ToolTipText     =   "Double Click to Grant Access"
                  Top             =   120
                  Width           =   4965
                  _Version        =   655364
                  _ExtentX        =   8758
                  _ExtentY        =   9340
                  _StockProps     =   64
               End
            End
            Begin VB.Frame Frame3 
               Height          =   5475
               Left            =   90
               TabIndex        =   34
               Top             =   480
               Width           =   4785
               Begin XtremeReportControl.ReportControl rptAccess 
                  Height          =   5295
                  Left            =   30
                  TabIndex        =   35
                  ToolTipText     =   "Double Click To Denied Access"
                  Top             =   120
                  Width           =   4695
                  _Version        =   655364
                  _ExtentX        =   8281
                  _ExtentY        =   9340
                  _StockProps     =   64
               End
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Not Installed"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   345
               Left            =   5010
               TabIndex        =   33
               Top             =   210
               Width           =   4935
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Installed"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   30
               TabIndex        =   32
               Top             =   180
               Width           =   4905
            End
         End
      End
   End
End
Attribute VB_Name = "FrmValid_Company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsValidation                            As ADODB.Recordset
Dim rsDescription                           As ADODB.Recordset
Dim rsShowCompanyaccess                     As ADODB.Recordset
Dim AddorEdit                               As String
Dim SQL                                     As String
Dim cnt                                     As Integer
Dim cnt2                                    As Integer
Private Sub cboModuleType_Click()
    lstWithNoAccess.ListItems.Clear
    lstwithAccess.ListItems.Clear
    cnt = 0
    cnt2 = 0
End Sub



Private Sub cmdAdd_Click()
    ShowCommand
    AddorEdit = "ADD"
    frmTop.Enabled = True
    cmbModuletype_add.ListIndex = 0
    frmLISt.Enabled = False
    txtDescription.Text = ""
    txtModuletype.Visible = False
    cmbModuletype_add.Visible = True
    frmmiddle.Enabled = False
End Sub
Private Sub cmdCancel_Click()
    Picture4.Visible = True
    Picture5.Visible = False
    Picture5.ZOrder (1)
    Picture4.ZOrder (0)
    frmTop.Enabled = False
    Call rsRefresh
    Call FillGrid
    txtModuletype.Visible = False
    cmbModuletype_add.Visible = True
    frmLISt.Enabled = True
    frmmiddle.Enabled = True
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Are you sure you want to delete this data?", vbCritical + vbYesNo, "Warning") = vbNo Then Exit Sub
AddorEdit = "DELETE"
    Call SaveRecords
End Sub

Private Sub cmdEdit_Click()
    ShowCommand
    AddorEdit = "EDIT"
    frmTop.Enabled = True
    frmLISt.Enabled = False
    txtModuletype.Visible = False
    cmbModuletype_add.Visible = True
    frmmiddle.Enabled = False
    On Error Resume Next
    cmbModuletype_add.SetFocus
End Sub

Private Sub cmdEscape_Click()
    picAddframe.ZOrder (1)
    Pic1.ZOrder (0)
    Frame1.Enabled = True
    Pic1.Enabled = True
    Frame4.Enabled = True
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    txtfind.Text = ""
    txtfind.SetFocus
End Sub

Private Sub CmdGrantAll_Click()
    Dim rsgrantaccessall                       As New ADODB.Recordset
    Set rsgrantaccessall = gconDMIS.Execute("Update Valid_Company set valid = '1'where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "'")
    Set rsgrantaccessall = Nothing
'    Call fillNoAccess
'    Call fillWithAccess
    Call listWithAccess
    Call listWithNOAccess
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
        rsValidation.MoveNext
    If rsValidation.EOF Then
        rsValidation.MoveLast
        ShowLastRecordMsg
    End If
        Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    
        rsValidation.MovePrevious
    If rsValidation.BOF Then
        rsValidation.MoveFirst
        ShowLastRecordMsg
    End If
        Call StoreMemVars
End Sub

Private Sub cmdSave_Click()
    Call SaveRecords
End Sub

Private Sub cmddenieall_Click()
    Dim rsdeniedaccessall                       As New ADODB.Recordset
    Set rsdeniedaccessall = gconDMIS.Execute("Update Valid_Company set valid = '0' where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "'")
    Set rsdeniedaccessall = Nothing
'    Call fillNoAccess
'    Call fillWithAccess
    Call listWithAccess
    Call listWithNOAccess
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Call listWithAccess
    Call listWithNOAccess
End Sub

Private Sub Form_Load()
    Call addeditdelete
    Call InitializeReportControl
 End Sub
Private Sub lbladd_Click()
    'Call cmdCancel_Click
    picAddframe.ZOrder (0)
    Pic1.ZOrder (1)
    Frame1.Enabled = False
    Pic1.Enabled = False
    Frame4.Enabled = False
End Sub
Function ShowCommand()
    Picture4.Visible = False
    Picture5.Visible = True
    Picture5.ZOrder (0)
End Function

Private Sub SaveRecords()
    Dim CMD                                            As ADODB.Command
    Dim RET_ID                                         As Integer
    
    Set CMD = New ADODB.Command
    CMD.ActiveConnection = gconDMIS
    CMD.CommandType = adCmdStoredProc
    CMD.CommandText = "ValidCompany"
    
    With CMD.Parameters
        .Append CMD.CreateParameter("@ACTION", adVarChar, adParamInput, 10, AddorEdit)
        .Append CMD.CreateParameter("@MODULE_TYPE", adVarChar, adParamInput, 10, UCase(cmbModuletype_add))
        .Append CMD.CreateParameter("@DESCRIPTION", adVarChar, adParamInput, 100, UCase(txtDescription))
        .Append CMD.CreateParameter("@ID", adInteger, adParamInput, , labID)
        .Append CMD.CreateParameter("@RET_ID", adInteger, adParamOutput)
    End With
    
    On Error GoTo SQL_ERROR

    CMD.Execute
    RET_ID = CMD("@RET_ID")
    If AddorEdit = "ADD" Then
        MessagePop InfoFriend, "Save", "Record Successfully saved"
        Call rsRefresh
        Call StoreMemVars
        Call FillGrid
    ElseIf AddorEdit = "EDIT" Then
        MessagePop InfoFriend, "Save", "Record Successfully Updated"
        Call rsRefresh
        Call StoreMemVars
        Call FillGrid
        If AddorEdit = "EDIT" Then
             On Error Resume Next
             rsValidation.Find ("ID=" & RET_ID)
        End If
    Else
        MessagePop InfoFriend, "Delete", "Record Successfully Delete"
        Call rsRefresh
        Call StoreMemVars
        Call FillGrid
    End If
    cmdCancel.Value = True
    Exit Sub
    
SQL_ERROR:
    MessagePop InfoStop, "Error", "" & Err.Description
    Err.Clear
End Sub

Sub rsRefresh()
    Set rsValidation = New ADODB.Recordset
    SQL = "select MAINMODULENAME,DESCRIPTION,ID from VALID_COMPANY order by MAINMODULENAME,DESCRIPTION DESC"
    Call rsValidation.Open(SQL, gconDMIS, adOpenForwardOnly, adLockReadOnly)
    'Set rsValidation = gconDMIS.Execute("select MAINMODULENAME,DESCRIPTION,ID from VALID_COMPANY order by MAINMODULENAME,DESCRIPTION asc")
End Sub

Sub StoreMemVars()
    If Not rsValidation.EOF And Not rsValidation.BOF Then
        txtModuletype.Text = Null2String(rsValidation!MAINMODULENAME)
        txtDescription.Text = Null2String(rsValidation!Description)
        labID.Caption = Null2String(rsValidation!ID)
    Else
        Call ShowNoRecord
    End If
End Sub
 Sub Initvars()
    cboModuleType.Clear
    cmbModuletype_add.Clear
    txtDescription.Text = ""
    lstlist.ListItems.Clear
 End Sub
Private Sub FillGrid()
    If Not (rsValidation.EOF And rsValidation.BOF) Then
        Listview_Loadval Me.lstlist.ListItems, rsValidation
        lstlist.Refresh
        
        lstlist.ColumnHeaders(1).Text = "Module Type"
        lstlist.ColumnHeaders(2).Text = "Description"
        lstlist.ColumnHeaders(3).Text = "ID"
    Else
        lstlist.ListItems.Clear
    End If
 End Sub
 
Sub InitCombo()
    Dim rsModuleType                           As ADODB.Recordset
    Set rsModuleType = gconDMIS.Execute("Select Distinct MAINMODULENAME from all_rams_modules order by MAINMODULENAME asc")
    If Not rsModuleType.EOF And Not rsModuleType.BOF Then
        cmbModuletype_add.Clear: rsModuleType.MoveFirst
        Do While Not rsModuleType.EOF
            cmbModuletype_add.AddItem Null2String(rsModuleType!MAINMODULENAME)
            cboModuleType.AddItem Null2String(rsModuleType!MAINMODULENAME)
            
            rsModuleType.MoveNext
        Loop
            cmbModuletype_add.AddItem " ", 0
            cboModuleType.AddItem " ", 0
      End If
      Set rsModuleType = Nothing
 End Sub

Private Sub lstlist_DblClick()
    Call cmdEdit_Click
End Sub

Private Sub lstlist_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    txtModuletype.Visible = True
    Me.cmbModuletype_add.Visible = False
    txtModuletype.Text = lstlist.SelectedItem.Text
    Me.txtDescription.Text = lstlist.SelectedItem.SubItems(1)
    Me.labID.Caption = Me.lstlist.SelectedItem.SubItems(2)
End Sub
Sub FillSearchGrid(XXX As String)
    Dim rsCompanyValidation2                                       As ADODB.Recordset
    lstlist.Sorted = False
    lstlist.ListItems.Clear
    lstlist.Enabled = False
    Set rsCompanyValidation2 = New ADODB.Recordset
    Set rsCompanyValidation2 = gconDMIS.Execute("select top 50 MAINMODULENAME , DESCRIPTION, ID from VALID_COMPANY where MAINMODULENAME like'" & ReplaceQuote(XXX) & "%' or DESCRIPTION like'" & ReplaceQuote(XXX) & "%' order by ID asc")
    If Not (rsCompanyValidation2.EOF And rsCompanyValidation2.BOF) Then
        Listview_Loadval Me.lstlist.ListItems, rsCompanyValidation2
        lstlist.Refresh
        lstlist.Enabled = True
    End If
End Sub

Private Sub lstwithAccess_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstwithAccess
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

Private Sub lstwithAccess_DblClick()
    If cnt = 0 Then Exit Sub
        Dim rsdenieaccess                       As New ADODB.Recordset
        Set rsdenieaccess = New ADODB.Recordset
        Set rsdenieaccess = gconDMIS.Execute("Update Valid_Company set valid = '0' where ID = '" & Null2String(lstwithAccess.SelectedItem.SubItems(1)) & "'")
        Call listWithAccess
        Call listWithNOAccess
End Sub

Private Sub lstWithNoAccess_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstWithNoAccess
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

Private Sub lstWithNoAccess_DblClick()
If cnt2 = 0 Then Exit Sub
    Dim rsgrantaccess                       As New ADODB.Recordset
    Set rsgrantaccess = New ADODB.Recordset
    Set rsgrantaccess = gconDMIS.Execute("Update Valid_Company set valid = '1' where ID = '" & Null2String(lstWithNoAccess.SelectedItem.SubItems(1)) & "'")
    Call listWithAccess
    Call listWithNOAccess
End Sub

Private Sub rptAccess_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    Metrics.ForeColor = Label4.ForeColor
End Sub

Private Sub rptAccess_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    Dim rsgdenieaccess                       As New ADODB.Recordset
    Set rsgdenieaccess = gconDMIS.Execute("Update Valid_Company set valid = '0' where ID = '" & Row.Record(1).Value & "'")
    Set rsgdenieaccess = Nothing
    Call fillWithAccess
    Call fillNoAccess
End Sub

Private Sub rptDenied_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    Metrics.ForeColor = Label5.ForeColor
End Sub

Private Sub rptDenied_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal ITEM As XtremeReportControl.IReportRecordItem)
    Dim rsgrantaccess                       As New ADODB.Recordset
    Set rsgrantaccess = gconDMIS.Execute("Update Valid_Company set valid = '1' where ID = '" & Row.Record(1).Value & "'")
    Set rsgrantaccess = Nothing
    Call fillWithAccess
    Call fillNoAccess
End Sub

Private Sub Timer1_Timer()

If cboModuleType.Text <> "" Then
    If cboModuleType.Text = "CSMS" Then
        lblmodule.Caption = "Car Service"
    ElseIf cboModuleType.Text = "PMIS" Then
        lblmodule.Caption = "Parts"
    ElseIf cboModuleType.Text = "AMIS" Then
        lblmodule.Caption = "Accounting"
    ElseIf cboModuleType.Text = "CMIS" Then
        lblmodule.Caption = "Casher Monitoring"
    ElseIf cboModuleType.Text = "HRMS" Then
        lblmodule.Caption = "Human Resource"
    ElseIf cboModuleType.Text = "SMIS" Then
        lblmodule.Caption = "Sales Monitoring"
    ElseIf cboModuleType.Text = "CRIS" Then
        lblmodule.Caption = "Customer Relation"
    Else
        lblmodule.Caption = ""
    End If
    If lblmodule.Visible = True Then
       lblmodule.Visible = False
    Else
       lblmodule.Visible = True
    End If
End If

End Sub



Private Sub Timer2_Timer()
If Label7.Visible = True Then
    Label8.Visible = True
    Label7.Visible = False
ElseIf Label8.Visible = True Then
    Label8.Visible = False
    Label7.Visible = True
End If
End Sub

Private Sub txtfind_Change()
    Call FillSearchGrid(txtfind.Text)
End Sub

Function addeditdelete()
    picAddframe.ZOrder (1)
    Pic1.ZOrder (0)
    Call Initvars
    Call InitCombo
    Call rsRefresh
    Call StoreMemVars
    Call FillGrid
End Function

Sub InitializeReportControl()
    Screen.MousePointer = 11
    
    With rptAccess
        .Columns.DeleteAll
        .Columns.Add 0, "Description", 400, True::              .Columns(0).Resizable = False:                 .Columns(0).AllowRemove = False
        .Columns.Add 1, "ID", 50, True:                         .Columns(1).AllowRemove = False
                
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.TextFont.Weight = 540
        .PaintManager.CaptionFont.Bold = True
    End With
    
    With rptDenied
        .Columns.DeleteAll
        .Columns.Add 0, "Description", 400, True::              .Columns(0).Resizable = False:                 .Columns(0).AllowRemove = False
        .Columns.Add 1, "ID", 50, True:                       .Columns(1).AllowRemove = False
                
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots
        .PaintManager.VerticalGridStyle = xtpGridSmallDots
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.TextFont.Weight = 540
        .PaintManager.CaptionFont.Bold = True
    End With

 
    Screen.MousePointer = 0
End Sub
Private Sub fillNoAccess()
Dim rsNoAccess                                       As ADODB.Recordset
Dim RECNoAccess                                      As XtremeReportControl.ReportRecord

   Set rsNoAccess = New ADODB.Recordset
   Set rsNoAccess = gconDMIS.Execute("select Description,ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and isnull(valid,0) <> '1'order by DESCRIPTION asc ")
   
   If Not rsNoAccess.EOF And Not rsNoAccess.BOF Then
        rptDenied.Records.DeleteAll
        While Not rsNoAccess.EOF
            Set RECNoAccess = rptDenied.Records.Add
            RECNoAccess.AddItem (Trim(rsNoAccess!Description))
            RECNoAccess.AddItem (Trim(rsNoAccess!ID))

    
            rsNoAccess.MoveNext
            Set RECNoAccess = Nothing
        Wend
            rptDenied.Populate
            Screen.MousePointer = 0
            Set rsNoAccess = Nothing
            rptDenied.Enabled = True
    Else
        rptDenied.Records.DeleteAll
        rptDenied.Enabled = False
        
    End If
End Sub

Private Sub fillWithAccess()
Dim rswithaccess                                       As ADODB.Recordset
Dim RECWithAccess                                      As XtremeReportControl.ReportRecord

   Set rswithaccess = New ADODB.Recordset
   Set rswithaccess = gconDMIS.Execute("select Description,ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and isnull(valid,0) = '1'order by DESCRIPTION asc ")
   
   If Not rswithaccess.EOF And Not rswithaccess.BOF Then
        rptAccess.Records.DeleteAll
        While Not rswithaccess.EOF
            Set RECWithAccess = rptAccess.Records.Add
            RECWithAccess.AddItem (Trim(rswithaccess!Description))
            RECWithAccess.AddItem (Trim(rswithaccess!ID))

    
            rswithaccess.MoveNext
            Set RECWithAccess = Nothing
        Wend
            rptAccess.Populate
            Screen.MousePointer = 0
            Set RECWithAccess = Nothing
            rptAccess.Enabled = True
    Else
        rptAccess.Records.DeleteAll
        rptAccess.Enabled = False
        ShowNoRecord
    End If
 
End Sub

Private Sub listWithAccess()
Dim rswithaccess                                       As ADODB.Recordset

   Set rswithaccess = New ADODB.Recordset
   Set rswithaccess = gconDMIS.Execute("select Description,ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and isnull(valid,0) = '1'order by DESCRIPTION asc ")
   
    If Not (rswithaccess.EOF And rswithaccess.BOF) Then
        Listview_Loadval Me.lstwithAccess.ListItems, rswithaccess
        lstwithAccess.Refresh
        
        lstwithAccess.ColumnHeaders(1).Text = "Description"
        lstwithAccess.ColumnHeaders(2).Text = "ID"
        cnt = cnt + 1
    Else
        lstwithAccess.ListItems.Clear
        cnt = 0
    End If

End Sub

Private Sub listWithNOAccess()
Dim rswithnoaccess                                       As ADODB.Recordset

   Set rswithnoaccess = New ADODB.Recordset
   Set rswithnoaccess = gconDMIS.Execute("select Description,ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and isnull(valid,0) <> '1'order by DESCRIPTION asc ")
   
    If Not (rswithnoaccess.EOF And rswithnoaccess.BOF) Then
        Listview_Loadval Me.lstWithNoAccess.ListItems, rswithnoaccess
        lstWithNoAccess.Refresh
        
        lstWithNoAccess.ColumnHeaders(1).Text = "Description"
        lstWithNoAccess.ColumnHeaders(2).Text = "ID"
        cnt2 = cnt2 + 1
    Else
        lstWithNoAccess.ListItems.Clear
        cnt2 = 0
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rswithaccess                                     As ADODB.Recordset
    Dim rswithnoaccess                                   As ADODB.Recordset

    lstwithAccess.Sorted = False
    lstwithAccess.ListItems.Clear
    lstwithAccess.Enabled = False
    Set rswithaccess = New ADODB.Recordset
    Set rswithaccess = gconDMIS.Execute("select top 50  DESCRIPTION, ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and DESCRIPTION like'" & ReplaceQuote(XXX) & "%' and isnull(valid,0) = '1' order by ID asc")
    If Not (rswithaccess.EOF And rswithaccess.BOF) Then
        Listview_Loadval Me.lstwithAccess.ListItems, rswithaccess
        lstwithAccess.Refresh
        lstwithAccess.Enabled = True
        
    Else
        lstwithAccess.ListItems.Clear
    End If
    
    lstWithNoAccess.Sorted = False
    lstWithNoAccess.ListItems.Clear
    lstWithNoAccess.Enabled = False
    Set rswithnoaccess = New ADODB.Recordset
    Set rswithnoaccess = gconDMIS.Execute("select top 50  DESCRIPTION, ID from VALID_COMPANY where MAINMODULENAME = '" & Null2String(cboModuleType.Text) & "' and DESCRIPTION like'" & ReplaceQuote(XXX) & "%' and isnull(valid,0) <> '1' order by ID asc")
    If Not (rswithnoaccess.EOF And rswithnoaccess.BOF) Then
        Listview_Loadval Me.lstWithNoAccess.ListItems, rswithnoaccess
        lstWithNoAccess.Refresh
        lstWithNoAccess.Enabled = True
        
    Else
        lstWithNoAccess.ListItems.Clear
    End If

End Sub

Private Sub txtSearch_Change()
    FillSearchGrid2 Null2String(txtSearch.Text)
End Sub
