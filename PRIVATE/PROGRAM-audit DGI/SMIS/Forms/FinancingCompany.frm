VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmSMIS_Files_FinancingCo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financing Company"
   ClientHeight    =   7230
   ClientLeft      =   315
   ClientTop       =   600
   ClientWidth     =   5820
   ForeColor       =   &H00FCFCFC&
   Icon            =   "FinancingCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   5820
   Begin VB.PictureBox picFinRateDetail 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   6810
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   2490
      ScaleWidth      =   3945
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdAddDetails 
         Caption         =   "Add Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   2130
         Width           =   1095
      End
      Begin VB.CommandButton cmdDeleteDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2130
         MouseIcon       =   "FinancingCompany.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Delete Detail"
         Top             =   1920
         Width           =   555
      End
      Begin VB.TextBox txtRuralRate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1665
         MaxLength       =   5
         TabIndex        =   21
         Top             =   795
         Width           =   2115
      End
      Begin VB.CommandButton cmdSaveDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2670
         MouseIcon       =   "FinancingCompany.frx":0D47
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":0E99
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Save Detail"
         Top             =   1920
         Width           =   555
      End
      Begin VB.TextBox txtUrbanRate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1665
         MaxLength       =   5
         TabIndex        =   19
         ToolTipText     =   "Type status of tag number (e.g. U for unposted)"
         Top             =   1155
         Width           =   2115
      End
      Begin VB.TextBox txtDownpayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1665
         TabIndex        =   18
         ToolTipText     =   "Type status of tag number (e.g. U for unposted)"
         Top             =   1515
         Width           =   2115
      End
      Begin VB.TextBox txtTerms 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   17
         ToolTipText     =   "Type status of tag number (e.g. U for unposted)"
         Top             =   420
         Width           =   2115
      End
      Begin VB.CommandButton cmdExitDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3210
         MouseIcon       =   "FinancingCompany.frx":11E9
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":133B
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Exit Entry"
         Top             =   1920
         Width           =   555
      End
      Begin XtremeShortcutBar.ShortcutCaption cap2 
         Height          =   330
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   3945
         _Version        =   655364
         _ExtentX        =   6959
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "::: Add Details:::"
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
      Begin VB.Label lblQuotationParticular 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TERMS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Left            =   960
         TabIndex        =   26
         Top             =   480
         Width           =   600
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "RURAL RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   3
         Left            =   510
         TabIndex        =   25
         Top             =   855
         Width           =   1005
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "URBAN RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   0
         Left            =   510
         TabIndex        =   24
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label zlblC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DOWNPAYMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   210
         Index           =   2
         Left            =   300
         TabIndex        =   23
         Top             =   1575
         Width           =   1260
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   6420
      Left            =   0
      ScaleHeight     =   6420
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   -45
      Width           =   5835
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   5250
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   30
         TabIndex        =   35
         Top             =   -30
         Width           =   5670
         Begin VB.CommandButton cmdSelect 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1530
            TabIndex        =   46
            Top             =   420
            Width           =   375
         End
         Begin VB.TextBox txtCompany 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            MaxLength       =   80
            TabIndex        =   40
            Top             =   1050
            Width           =   5535
         End
         Begin VB.TextBox txtCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   39
            Text            =   "1234567890"
            Top             =   390
            Width           =   1425
         End
         Begin VB.TextBox txtLocation 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            MaxLength       =   80
            TabIndex        =   38
            Top             =   1695
            Width           =   5550
         End
         Begin VB.TextBox txtContactPerson 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            MaxLength       =   80
            TabIndex        =   37
            Top             =   2325
            Width           =   5550
         End
         Begin VB.TextBox txtContactNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            MaxLength       =   80
            TabIndex        =   36
            Top             =   2925
            Width           =   5550
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
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
            Height          =   225
            Left            =   90
            TabIndex        =   45
            Top             =   810
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Height          =   225
            Left            =   90
            TabIndex        =   44
            Top             =   150
            Width           =   435
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
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
            Height          =   225
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   1455
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person(s)"
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
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   42
            Top             =   2085
            Width           =   1545
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact No(s)"
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
            Height          =   225
            Index           =   2
            Left            =   90
            TabIndex        =   41
            Top             =   2715
            Width           =   1155
         End
      End
      Begin VB.Frame fraSearch 
         Height          =   3000
         Left            =   60
         TabIndex        =   29
         Top             =   3300
         Width           =   5685
         Begin VB.TextBox txtSearch 
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   90
            MaxLength       =   35
            TabIndex        =   32
            Top             =   420
            Width           =   5520
         End
         Begin VB.OptionButton optCompany 
            Caption         =   "Company &Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton optCode 
            Caption         =   "&Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1770
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   150
            Width           =   1245
         End
         Begin MSComctlLib.ListView lstCompany 
            Height          =   2070
            Left            =   60
            TabIndex        =   33
            Top             =   840
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   3651
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   0   'False
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
            MouseIcon       =   "FinancingCompany.frx":16A1
            NumItems        =   0
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTags 
      Height          =   2700
      Left            =   6810
      TabIndex        =   28
      Top             =   600
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   4763
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      ForeColorFixed  =   0
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483628
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      FormatString    =   "Terms         | Rural Rate  |  Urban Rate  |     Downpayment            |ID"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FinancingCompany.frx":1803
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   -180
      ScaleHeight     =   945
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   6315
      Width           =   6675
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
         Left            =   5190
         MouseIcon       =   "FinancingCompany.frx":1B1D
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":1C6F
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   4500
         MouseIcon       =   "FinancingCompany.frx":1FD5
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":2127
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   3810
         MouseIcon       =   "FinancingCompany.frx":248D
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":25DF
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   3120
         MouseIcon       =   "FinancingCompany.frx":293B
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":2A8D
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Delete Selected Record"
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
         Left            =   2430
         MouseIcon       =   "FinancingCompany.frx":2DB8
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":2F0A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Add Record"
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
         Left            =   1740
         MouseIcon       =   "FinancingCompany.frx":321D
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":336F
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   1050
         MouseIcon       =   "FinancingCompany.frx":3669
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":37BB
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   360
         MouseIcon       =   "FinancingCompany.frx":3B13
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":3C65
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4290
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   9
      Top             =   6300
      Width           =   1800
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
         MouseIcon       =   "FinancingCompany.frx":3FC4
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":4116
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancel"
         Top             =   45
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
         Left            =   30
         MouseIcon       =   "FinancingCompany.frx":4454
         MousePointer    =   99  'Custom
         Picture         =   "FinancingCompany.frx":45A6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.Label labid 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LABID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   6750
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label labFincomRateID 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "labFincomID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   45
      TabIndex        =   1
      Top             =   6780
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmSMIS_Files_FinancingCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFinCom                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub FillSearchGrid(XXX As String)
    Dim rsFinCom                                                      As ADODB.Recordset
    lstCompany.Sorted = False
    lstCompany.ListItems.Clear
    lstCompany.Enabled = False
    Set rsFinCom = New ADODB.Recordset
    If optCode.Value = True Then
        Set rsFinCom = gconDMIS.Execute("select  code , Company, ID from SMIS_FinCom where Code like'" & ReplaceQuote(XXX) & "%' order by Company asc")
    Else
        Set rsFinCom = gconDMIS.Execute("select  code , Company, ID from SMIS_FinCom where Company like'" & ReplaceQuote(XXX) & "%' order by Company asc")
    End If

    If Not (rsFinCom.EOF And rsFinCom.BOF) Then
        Listview_Loadval Me.lstCompany.ListItems, rsFinCom
        lstCompany.Refresh
        lstCompany.Enabled = True
    End If

End Sub

Sub initMemvars()
    txtCode.Text = ""
    txtCompany.Text = ""
    txtLocation.Text = ""
    txtContactNo.Text = ""
    txtContactPerson.Text = ""
    grdTags.ColWidth(4) = 0


End Sub

Sub rsRefresh()
    Set rsFinCom = New ADODB.Recordset
    rsFinCom.Open "select * from SMIS_FinCom order by id DESC", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    Dim recRs                                                         As ADODB.Recordset
    If Not rsFinCom.EOF And Not rsFinCom.BOF Then
        labID.Caption = rsFinCom!ID
        txtCode.Text = Null2String(rsFinCom!Code)
        txtCompany.Text = Null2String(rsFinCom!company)
        txtLocation.Text = Null2String(rsFinCom!Location)
        txtContactNo.Text = Null2String(rsFinCom!CONTACTNO)
        txtContactPerson.Text = Null2String(rsFinCom!ContactPerson)
        '    Set recRs = gconDMIS.Execute("Select * from SMIS_FINCOM_RATE WHERE FINCOMID=" & labid)
        '   grdTags.Rows = 1

        'While Not recRs.EOF
        '    grdTags.AddItem _
             '            recRs!Term & Chr(9) & _
             '                       recRs!RPerct & Chr(9) & _
             '                       recRs!UPerct & Chr(9) & _
             '                       FormatNumber(recRs!DownPayment) & Chr(9) _
             '                     & recRs!ID
        '   recRs.MoveNext
        'Wend
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "FINANCING COMPANY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "ADD"
    initMemvars
    lstCompany.Enabled = False
    txtSearch.Enabled = False
    CAP2.Caption = ":::ADD:::"
    Frame1.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    cmdAddDetails.Enabled = False
    grdTags.Rows = 1
    On Error Resume Next
    txtCode.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdAddDetails_Click()
    txtTerms.Enabled = True
    labFincomRateID = 0
    txtTerms.Text = 0
    txtRuralRate.Text = 0
    txtUrbanRate.Text = 0
    txtDownpayment.Text = 0
    cmdDeleteDetail.Enabled = False
    ShowHidePictureBox2 picFinRateDetail, True
    CAP2.Caption = "::ADD::"
    On Error Resume Next
    txtTerms.SetFocus
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    Frame1.Enabled = False: cmdAddDetails.Enabled = False: picAdds.Visible = True: picSaves.Visible = False:: fraSearch.Enabled = True
    lstCompany.Enabled = True
    txtSearch.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "FINANCING COMPANY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If Not rsFinCom.BOF Or Not rsFinCom.EOF Then
        If ShowConfirmDelete = True Then
            SQL_STATEMENT = "delete from SMIS_FinCom where id = " & labID.Caption

            gconDMIS.Execute (SQL_STATEMENT)
            '**************NEW AUDIT TRAIL***************
            NEW_LogAudit "X", "FINANCING COMPANY", SQL_STATEMENT, N2Str2Null(labID), "", "CODE: " & txtCode, "", ""
            '**************NEW AUDIT TRAIL***************

            '**************RESET THE VARIABLE***************
            SQL_STATEMENT = ""
            '**************RESET THE VARIABLE***************

            SQL_STATEMENT = "delete from SMIS_FinCom_Rate where fincomid = " & labID.Caption
            gconDMIS.Execute (SQL_STATEMENT)
            '**************NEW AUDIT TRAIL***************
            NEW_LogAudit "XX", "FINANCING COMPANY", SQL_STATEMENT, N2Str2Null(labID), "", "CODE: " & txtCode, "", ""
            '**************NEW AUDIT TRAIL***************
            LogAudit "X", "FINANCIAL COMPANY", txtCompany
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If

    rsRefresh
    FillSearchGrid ""
    cmdCancel_Click





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdDeleteDetail_Click()
    If Function_Access(LOGID, "Acess_DELETE", "FINANCING COMPANY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("Delete from  SMIS_FINCOM_Rate Where ID=" & labFincomRateID)
        StoreMemVars
        ShowHidePictureBox2 picFinRateDetail, False
        ShowDeletedMsg
    End If





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "FINANCING COMPANY") = False Then Exit Sub
    On Error GoTo ErrorCode:

    AddorEdit = "EDIT"
    Frame1.Enabled = True
    cmdAddDetails.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    fraSearch.Enabled = False
    On Error Resume Next
    txtCode.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExitDetail_Click()
    ShowHidePictureBox2 picFinRateDetail, False
    On Error Resume Next
    grdTags.SetFocus
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsFinCom.MoveNext
    If rsFinCom.EOF Then
        rsFinCom.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsFinCom.MovePrevious
    If rsFinCom.BOF Then
        rsFinCom.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "FINANCING COMPANY") = False Then Exit Sub

    Screen.MousePointer = 11
    CrystalReport1.Formulas(0) = "COMPANYNAME='" & COMPANY_NAME & "'"
    CrystalReport1.Formulas(1) = "COMPANYADDRESS='" & COMPANY_ADDRESS & "'"
    PrintSQLReport CrystalReport1, SMIS_REPORT_PATH & "Listing\FinancingCompany.rpt", "", DMIS_REPORT_Connection, 1

    NEW_LogAudit "V", "FINANCING COMPANY", "", N2Str2Null(labID), "", "", "", ""

    LogAudit "V", "FINANCIAL COMPANY", txtCompany
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()

    Dim vtxtCode                                                      As String
    Dim vtxtCompany                                                   As String
    Dim vtxtContactNo                                                 As String
    Dim vtxtContactPerson                                             As String
    Dim VTXTLocation                                                  As String
    Dim lng                                                           As Integer

    On Error GoTo ErrorCode:

    If txtCode.Text = "" Or txtCompany.Text = "" Then
        ShowIsRequiredMsg "Code and Company"
        On Error Resume Next
        txtCode.SetFocus
        Exit Sub
    End If

    vtxtCode = N2Str2Null(txtCode.Text)
    vtxtCompany = N2Str2Null(UCase(txtCompany.Text))
    VTXTLocation = N2Str2Null(UCase(txtLocation.Text))
    vtxtContactNo = N2Str2Null(txtContactNo.Text)
    vtxtContactPerson = N2Str2Null(UCase(txtContactPerson.Text))


    lng = gconDMIS.Execute("select Count(*) from SMIS_FINCOM WHERE Code=" & N2Str2Null(txtCode)).Fields(0).Value
    If AddorEdit = "ADD" Then
        If lng >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    Else
        If lng >= 1 And UCase(Null2String(rsFinCom!Code)) <> UCase(txtCode) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Code Already Exist"
            Exit Sub
        End If
    End If



    cmdAddDetails.Visible = True
    If AddorEdit = "ADD" Then
        If Not rsFinCom.EOF And Not rsFinCom.BOF Then
            rsFinCom.MoveLast
            labID.Caption = NumericVal(rsFinCom!ID) + 1
        End If
        SQL_STATEMENT = "Insert into SMIS_FinCom" & _
                      " (CODE,COMPANY,LOCATION,CONTACTNO, CONTACTPERSON)" & _
                      " values (" & vtxtCode & ", " & vtxtCompany & ", " & VTXTLocation & ", " & vtxtContactNo & ", " & vtxtContactPerson & ")"

        gconDMIS.Execute (SQL_STATEMENT)
        '********NEW LOG AUDIT***********
        NEW_LogAudit "A", "FINANCING COMPANY", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtCode), "CODE", "SMIS_FINCOM"), "", "CODE :" & N2Str2Null(txtCode), "", ""
        '********NEW LOG AUDIT***********

        cmdAddDetails_Click

        LogAudit "A", "FINANCIAL COMPANY", txtCompany

    Else
        SQL_STATEMENT = "update SMIS_FinCom set" & _
                      " Code = " & vtxtCode & "," & _
                      " Location = " & VTXTLocation & "," & _
                      " ContactNo= " & vtxtContactNo & "," & _
                      " ContactPerson = " & vtxtContactPerson & "," & _
                      " Company = " & vtxtCompany & _
                      " where id = " & labID.Caption


        gconDMIS.Execute (SQL_STATEMENT)
        '********NEW LOG AUDIT***********
        NEW_LogAudit "E", "FINANCING COMPANY", SQL_STATEMENT, N2Str2Null(labID), "", "CODE :" & N2Str2Null(txtCode), "", ""
        '********NEW LOG AUDIT***********



        LogAudit "E", "FINANCIAL COMPANY", txtCompany
    End If
    rsRefresh
    FillSearchGrid ""
    If AddorEdit = "EDIT" Then
        rsFinCom.Find "id =" & labID
    End If
    cmdCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSaveDetail_Click()
    Dim UrbanRate                                                     As Double
    Dim RuralRate                                                     As Double
    Dim DownPayment                                                   As Double
    Dim TERM                                                          As Integer
    Dim i                                                             As Long

    On Error GoTo ErrorCode:

    UrbanRate = NumericVal(txtUrbanRate)
    RuralRate = NumericVal(txtRuralRate)
    DownPayment = NumericVal(txtDownpayment)
    TERM = NumericVal(txtTerms)

    If labFincomRateID <> 0 Then
        gconDMIS.Execute ("UPDate SMIS_FINCOM_Rate set RPerct=" & RuralRate & " , Uperct=" & UrbanRate & ", DownPayment=" & DownPayment & " Where ID=" & labFincomRateID)
        MessagePop RecSave, " Record Updated", txtCompany & " Interest Rate Has been Updated"
    Else
        If grdTags.Rows <> 1 Then
            For i = 1 To grdTags.Rows - 1
                If grdTags.TextMatrix(i, 0) = TERM Then
                    MessagePop RecSaveError, "Duplicate  Entry ", " Term Already Exists"
                    On Error Resume Next
                    txtTerms.SetFocus
                    Exit Sub
                End If
            Next
        End If
        gconDMIS.Execute ("INSERT INTO SMIS_FINCOM_Rate (FINCOMID,TERM,RPerct,UPerct,DownPayment ) values " _
                        & " (" & labID & " , " & TERM & ", " & RuralRate & ", " & UrbanRate & "," & DownPayment & ")")
        MessagePop RecSave, " Record Added", txtCompany & " Interest Rate Has been Added"
    End If
    StoreMemVars
    ShowHidePictureBox2 picFinRateDetail, False
    On Error Resume Next
    grdTags.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelect_Click()
    SelectCustomer = "Financing"
    frmCustomerSearch1.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And picFinRateDetail.Visible = True Then
        cmdExitDetail_Click
    ElseIf KeyCode = vbKeyEscape And picFinRateDetail.Visible = False Then
        cmdCancel_Click
    Else
        If picAdds.Visible = True And KeyCode = vbKeyEscape Then
            Unload Me
        Else
            MoveKeyPress KeyCode
        End If
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (FINANCING COMPANY)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labID.Caption), "FINANCING COMPANY")
            'End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    Call AddColumnHeader("Code, Description", lstCompany)
    Call ResizeColumnHeader(lstCompany, "25,70")

    Frame1.Enabled = False
    cmdAddDetails.Enabled = False
    txtSearch.Text = ""
    picAdds.Visible = True
    picSaves.Visible = False
    initMemvars
    StoreMemVars
    FillSearchGrid ""
    Screen.MousePointer = 0
End Sub

Private Sub grdTags_DblClick()
    If AddorEdit = "EDIT" Then
        If grdTags.Row = 0 Then: Exit Sub
        cmdDeleteDetail.Enabled = True
        CAP2.Caption = ":::Edit::"
        labFincomRateID = grdTags.TextMatrix(grdTags.Row, 4)
        txtTerms.Text = grdTags.TextMatrix(grdTags.Row, 0)
        txtTerms.Enabled = False
        txtRuralRate.Text = FormatNumber(NumericVal(grdTags.TextMatrix(grdTags.Row, 1)))
        txtUrbanRate.Text = FormatNumber(NumericVal(grdTags.TextMatrix(grdTags.Row, 2)))
        txtDownpayment.Text = FormatNumber(NumericVal(grdTags.TextMatrix(grdTags.Row, 3)))
        ShowHidePictureBox2 picFinRateDetail, True
        On Error Resume Next
        txtRuralRate.SetFocus
    End If
End Sub

Private Sub grdTags_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdTags_DblClick
End Sub

Private Sub lstCompany_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCompany
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

Private Sub lstCompany_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCompany_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsFinCom.MoveFirst
    rsFinCom.Find ("ID=" & Item.ListSubItems(2).Text)
    Frame1.Enabled = False
    cmdAddDetails.Enabled = False
    StoreMemVars
End Sub

Private Sub lstCompany_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstCompany_DblClick
End Sub

Private Sub optCode_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optCompany_Click()
    FillSearchGrid (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub picFinRateDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ShowHidePictureBox2 picFinRateDetail, False
        On Error Resume Next
        grdTags.SetFocus
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    If AddorEdit = "ADD" And Len(Trim(txtCode)) > 0 Then
        If gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_FINCOM WHERE CODE=" & N2Str2Null(txtCode)).Fields(0) >= 1 Then
            MessagePop InfoVoid, "Duplicate Entry", "Bank Code already Asssigned"
            Cancel = True
        End If
    End If
End Sub

Private Sub txtDownpayment_GotFocus()
    If NumericVal(txtDownpayment.Text) <= 0 Then txtDownpayment = ""
End Sub

Private Sub txtDownpayment_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtDownPayment_LostFocus()
    If NumericVal(txtDownpayment.Text) <= 0 Then txtDownpayment = "0.00"
    txtDownpayment = FormatNumber(txtDownpayment)
End Sub

Private Sub txtRuralRate_GotFocus()
    If NumericVal(txtRuralRate.Text) <= 0 Then txtRuralRate = ""
End Sub

Private Sub txtRuralRate_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtRuralRate_LostFocus()
    If NumericVal(txtRuralRate.Text) <= 0 Then txtRuralRate = "0.00"
    txtRuralRate = FormatNumber(txtRuralRate)
End Sub

Private Sub txtSearch_Change()
    FillSearchGrid (txtSearch.Text)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyDown Then
    '    If lstCompany.ListItems.Count > 0 And lstCompany.Enabled = True Then: lstCompany.SetFocus
    'End If
End Sub

Private Sub txtTerms_GotFocus()
    If NumericVal(txtTerms.Text) <= 0 Then txtTerms = ""
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtTerms_LostFocus()
    If NumericVal(txtTerms.Text) <= 0 Then txtTerms = "0.00"
    txtTerms = FormatNumber(txtTerms)
End Sub

Private Sub txtUrbanRate_GotFocus()
    If NumericVal(txtUrbanRate.Text) <= 0 Then txtUrbanRate = ""
End Sub

Private Sub txtUrbanRate_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtUrbanRate_LostFocus()
    If NumericVal(txtUrbanRate.Text) <= 0 Then txtUrbanRate = "0.00"
    txtUrbanRate = FormatNumber(txtUrbanRate)
End Sub

