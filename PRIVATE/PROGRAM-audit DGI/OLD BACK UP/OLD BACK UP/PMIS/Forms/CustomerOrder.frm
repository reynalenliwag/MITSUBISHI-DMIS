VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISTrans_CustomerOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Order"
   ClientHeight    =   7125
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11505
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CustomerOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11505
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   123
      Top             =   6780
      Width           =   11505
      Begin VB.Label LAB_ADB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   8400
         TabIndex        =   134
         Top             =   0
         Width           =   3075
      End
      Begin VB.Label labDetails 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   5220
         TabIndex        =   131
         Top             =   0
         Width           =   3165
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Inv #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3510
         TabIndex        =   130
         Top             =   0
         Width           =   765
      End
      Begin VB.Label labinvNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4290
         TabIndex        =   124
         Top             =   0
         Width           =   915
      End
      Begin VB.Label labSJ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   2610
         TabIndex        =   128
         Top             =   0
         Width           =   885
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SJ #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   127
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OR #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labORNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   870
         TabIndex        =   125
         Top             =   0
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   2820
      Top             =   4890
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Parts Issuance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2700
      ScaleHeight     =   255
      ScaleWidth      =   8685
      TabIndex        =   75
      Top             =   5520
      Width           =   8715
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   6360
         TabIndex        =   80
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4380
         TabIndex        =   79
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2790
         TabIndex        =   78
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1440
         TabIndex        =   77
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   76
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6645
      Left            =   60
      TabIndex        =   67
      Top             =   0
      Width           =   2595
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FFFFFF&
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
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1260
         Width           =   2475
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   69
         Top             =   630
         Width           =   2385
      End
      Begin VB.OptionButton optTranno 
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   68
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4965
         Left            =   30
         TabIndex        =   71
         Top             =   1620
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8758
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomerOrder.frx":08CA
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tranno"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   129
         Top             =   900
         Width           =   2205
      End
      Begin VB.Label Label18 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   72
         Top             =   150
         Width           =   1455
      End
   End
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2820
      TabIndex        =   27
      Top             =   -2790
      Width           =   8565
      ExtentX         =   15108
      ExtentY         =   4630
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   2700
      ScaleHeight     =   2190
      ScaleWidth      =   8715
      TabIndex        =   42
      Top             =   3285
      Width           =   8745
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8100
         Top             =   120
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2085
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   3678
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   -2147483633
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   2640
      ScaleHeight     =   915
      ScaleWidth      =   8835
      TabIndex        =   90
      Top             =   5940
      Width           =   8835
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
         Left            =   7980
         MouseIcon       =   "CustomerOrder.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
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
         Left            =   7200
         MouseIcon       =   "CustomerOrder.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelCO 
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6420
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CustomerOrder.frx":139C
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Entry"
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
         Left            =   5640
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CustomerOrder.frx":1828
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   4860
         MouseIcon       =   "CustomerOrder.frx":1C9F
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":1DF1
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
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
         Left            =   4080
         MouseIcon       =   "CustomerOrder.frx":214D
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":229F
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
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
         Left            =   3300
         MouseIcon       =   "CustomerOrder.frx":25B2
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":2704
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   795
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
         Left            =   2520
         MouseIcon       =   "CustomerOrder.frx":2A54
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":2BA6
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
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
         MouseIcon       =   "CustomerOrder.frx":2F04
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":3056
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
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
         Left            =   960
         MouseIcon       =   "CustomerOrder.frx":3350
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":34A2
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
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
         Left            =   180
         MouseIcon       =   "CustomerOrder.frx":37FA
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":394C
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9615
      ScaleHeight     =   885
      ScaleWidth      =   2220
      TabIndex        =   87
      Top             =   5895
      Width           =   2220
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         Left            =   990
         MouseIcon       =   "CustomerOrder.frx":3CAB
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":3DFD
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   795
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
         Left            =   210
         MouseIcon       =   "CustomerOrder.frx":413B
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":428D
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.PictureBox Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   2730
      ScaleHeight     =   3135
      ScaleWidth      =   8685
      TabIndex        =   28
      Top             =   90
      Width           =   8715
      Begin VB.TextBox txtPRtranno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   133
         Top             =   540
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton Command3 
         Caption         =   ".."
         Height          =   345
         Left            =   4410
         TabIndex        =   132
         Top             =   60
         Width           =   375
      End
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   2460
         TabIndex        =   122
         Top             =   570
         Width           =   255
      End
      Begin VB.Frame fraPayType 
         Caption         =   "Payment Type"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4560
         TabIndex        =   111
         Top             =   2430
         Width           =   4005
         Begin VB.OptionButton optCHARGE 
            Caption         =   "CHARGE"
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
            Left            =   2550
            TabIndex        =   113
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton optCASH 
            Caption         =   "CASH"
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
            Left            =   1530
            TabIndex        =   112
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.ComboBox cboRefPRSNo 
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2430
         TabIndex        =   8
         Text            =   "cboRefPRSNo"
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2370
         Width           =   1995
      End
      Begin VB.CommandButton c 
         Caption         =   "F1 - Assign PIS Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   82
         Top             =   60
         Width           =   2175
      End
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   6930
         TabIndex        =   81
         Top             =   60
         Width           =   255
      End
      Begin VB.TextBox txtReferencePIS 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   360
         Left            =   5160
         TabIndex        =   1
         Text            =   "PIWGC06H360"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
         Width           =   1785
      End
      Begin VB.ComboBox cboChargeTo 
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
         Height          =   330
         Left            =   5550
         TabIndex        =   11
         Text            =   "cboChargeTo"
         ToolTipText     =   "Select option from list."
         Top             =   -405
         Width           =   1785
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   615
         Left            =   4560
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Type your message or remarks."
         Top             =   1740
         Width           =   4035
      End
      Begin VB.TextBox txtCustName 
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
         Height          =   945
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         ToolTipText     =   "Type complete name of customer."
         Top             =   1380
         Width           =   4365
      End
      Begin VB.TextBox txtTranDate 
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Type the date of transaction in mm/dd/yyyy format (e.g 7/5/2004)"
         Top             =   570
         Width           =   1245
      End
      Begin VB.TextBox txtDS1 
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Type percentage to be added in the total amount. Do not include percent sign (e.g. 10, 15)"
         Top             =   945
         Width           =   525
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   1740
         Picture         =   "CustomerOrder.frx":45DD
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   -540
         Width           =   435
         Begin VB.TextBox txtTranType 
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
            Left            =   0
            MaxLength       =   3
            TabIndex        =   61
            Top             =   60
            Width           =   525
         End
      End
      Begin VB.TextBox txtDS_Desc1 
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
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   13
         ToolTipText     =   "Input the type of the added amount."
         Top             =   945
         Width           =   1365
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3420
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Input customer code (e.g. S01163)"
         Top             =   960
         Width           =   1005
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3420
         MaxLength       =   7
         TabIndex        =   4
         ToolTipText     =   "Type the transaction terms."
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox txtChargeTo 
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   -375
         Width           =   495
      End
      Begin VB.TextBox txtTranNo 
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
         Left            =   3390
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   60
         Width           =   1005
      End
      Begin VB.ComboBox cboSMName 
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
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2760
         Width           =   3345
      End
      Begin VB.ComboBox cboSalesMan 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Text            =   "cboSalesMan"
         Top             =   1620
         Width           =   765
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7110
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   59
         Top             =   510
         Width           =   1515
         Begin VB.TextBox txtNetInvAmt 
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
            Height          =   345
            Left            =   90
            MaxLength       =   15
            TabIndex        =   64
            Top             =   810
            Width           =   1395
         End
         Begin VB.TextBox txtDS_Amt1 
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
            Height          =   345
            Left            =   90
            MaxLength       =   15
            TabIndex        =   63
            Top             =   440
            Width           =   1395
         End
         Begin VB.TextBox txtTTLInvAmt 
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
            Height          =   345
            Left            =   90
            MaxLength       =   15
            TabIndex        =   62
            Top             =   60
            Width           =   1395
         End
      End
      Begin VB.TextBox txtRONO 
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
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelectCustomer 
         Caption         =   "F2 - Select Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   73
         Top             =   960
         Width           =   2685
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference PRS Number :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   90
         TabIndex        =   110
         Top             =   2400
         Width           =   2355
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PIS No."
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
         Left            =   4470
         TabIndex        =   74
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   4290
         TabIndex        =   66
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Height          =   285
         Left            =   5340
         TabIndex        =   65
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "NET Amount"
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
         Left            =   5940
         TabIndex        =   31
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Man"
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
         TabIndex        =   41
         Top             =   2790
         Width           =   975
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
         Left            =   3840
         TabIndex        =   40
         Top             =   990
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL Amount"
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
         Left            =   5445
         TabIndex        =   39
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label7 
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
         Left            =   2250
         TabIndex        =   38
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label6 
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
         Left            =   2760
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. Date"
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
         TabIndex        =   36
         Top             =   600
         Width           =   1005
      End
      Begin VB.Label labChargeTo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
         Enabled         =   0   'False
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
         Left            =   4560
         TabIndex        =   35
         Top             =   -390
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. No."
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
         Left            =   2280
         TabIndex        =   34
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
         TabIndex        =   33
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   4650
         TabIndex        =   32
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label labRONO 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RO Number"
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
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label labPosted 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLED"
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
         Height          =   315
         Left            =   7200
         TabIndex        =   29
         Top             =   90
         Width           =   1425
      End
   End
   Begin VB.PictureBox fraAddTran 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   3810
      ScaleHeight     =   3555
      ScaleWidth      =   6825
      TabIndex        =   43
      Top             =   930
      Width           =   6855
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   255
         Left            =   3150
         TabIndex        =   119
         Top             =   1890
         Width           =   285
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parts Details"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3345
         Left            =   3840
         TabIndex        =   102
         Top             =   120
         Width           =   2865
         Begin VB.Frame Frame5 
            Caption         =   "Model Codes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   150
            TabIndex        =   117
            Top             =   2400
            Width           =   2595
            Begin VB.TextBox txtModelCode 
               BackColor       =   &H00FFFFFF&
               CausesValidation=   0   'False
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
               Height          =   375
               Left            =   120
               MaxLength       =   6
               TabIndex        =   118
               ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
               Top             =   270
               Width           =   2325
            End
         End
         Begin VB.CheckBox chkAvailableOnStock 
            Alignment       =   1  'Right Justify
            Caption         =   "Available on Stock"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   120
            TabIndex        =   116
            Top             =   270
            Width           =   2595
         End
         Begin VB.Frame Frame3 
            Height          =   975
            Left            =   150
            TabIndex        =   103
            Top             =   630
            Width           =   2595
            Begin VB.OptionButton optConsigned 
               Caption         =   "Consigned"
               CausesValidation=   0   'False
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
               TabIndex        =   106
               Top             =   660
               Width           =   1845
            End
            Begin VB.OptionButton optImported 
               Caption         =   "Imported"
               CausesValidation=   0   'False
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
               TabIndex        =   105
               Top             =   390
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optLocalPurchase 
               Caption         =   "Local Purchases"
               CausesValidation=   0   'False
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
               TabIndex        =   104
               Top             =   150
               Width           =   1845
            End
         End
         Begin VB.Frame Frame4 
            Height          =   765
            Left            =   150
            TabIndex        =   107
            Top             =   1590
            Width           =   2595
            Begin VB.OptionButton optGenuine 
               Caption         =   "Genuine"
               CausesValidation=   0   'False
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
               TabIndex        =   109
               Top             =   180
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optNonGenuine 
               Caption         =   "Non-Genuine"
               CausesValidation=   0   'False
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
               TabIndex        =   108
               Top             =   420
               Width           =   1845
            End
         End
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         CausesValidation=   0   'False
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
         Left            =   2880
         MouseIcon       =   "CustomerOrder.frx":7319
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":746B
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Delete Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdTranCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         Left            =   2160
         MouseIcon       =   "CustomerOrder.frx":7796
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":78E8
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Cancel Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtTranDescription 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Left            =   90
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1110
         Width           =   3675
      End
      Begin VB.TextBox txtTranTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2190
         Width           =   1665
      End
      Begin VB.TextBox txtTranUPrice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   20
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1860
         Width           =   1665
      End
      Begin VB.TextBox txtTranQty 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1500
         Width           =   705
      End
      Begin VB.TextBox txtTranItemNo 
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   765
      End
      Begin VB.ComboBox cboTranPartNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         ItemData        =   "CustomerOrder.frx":7C26
         Left            =   1440
         List            =   "CustomerOrder.frx":7C28
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtPartID 
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
         Left            =   1470
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   480
         Width           =   585
      End
      Begin VB.CommandButton cmdTranSave 
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
         Left            =   1440
         MouseIcon       =   "CustomerOrder.frx":7C2A
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":7D7C
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Save Entry"
         Top             =   2640
         Width           =   735
      End
      Begin VB.Frame fraCostToCost 
         Height          =   405
         Left            =   2220
         TabIndex        =   120
         Top             =   1410
         Width           =   1575
         Begin VB.CheckBox Check1 
            Caption         =   "Cost to Cost"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   90
            TabIndex        =   121
            Top             =   150
            Width           =   1395
         End
      End
      Begin VB.TextBox txtTranUCost 
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
         Left            =   2190
         MaxLength       =   10
         TabIndex        =   114
         Text            =   "1000.00"
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1500
         Width           =   1575
      End
      Begin VB.Label labTranUCost 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
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
         Left            =   2250
         TabIndex        =   115
         Top             =   1530
         Width           =   615
      End
      Begin VB.Label labPartNo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Description"
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
         Left            =   1470
         TabIndex        =   58
         Top             =   1860
         Width           =   1275
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
         Left            =   1500
         TabIndex        =   56
         Top             =   1890
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Extend Price"
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
         Left            =   120
         TabIndex        =   50
         Top             =   2250
         Width           =   1305
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   840
         TabIndex        =   49
         Top             =   1890
         Width           =   615
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   510
         TabIndex        =   48
         Top             =   1530
         Width           =   915
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Left            =   570
         TabIndex        =   47
         Top             =   510
         Width           =   855
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
         Left            =   570
         TabIndex        =   46
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   870
         Width           =   1275
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
         Left            =   1560
         TabIndex        =   57
         Top             =   1860
         Width           =   975
      End
   End
   Begin VB.PictureBox fraSignatories 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   4185
      ScaleHeight     =   2325
      ScaleWidth      =   4380
      TabIndex        =   51
      Top             =   1815
      Width           =   4410
      Begin VB.CommandButton cmdPrintRIV 
         Caption         =   "&Print PIS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   3000
         MouseIcon       =   "CustomerOrder.frx":80CC
         MousePointer    =   99  'Custom
         Picture         =   "CustomerOrder.frx":821E
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   1560
         Width           =   855
      End
      Begin VB.CheckBox chkPreview 
         BackColor       =   &H00DEDFDE&
         Height          =   255
         Left            =   4020
         TabIndex        =   26
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtApprovedBy 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   780
         Width           =   2835
      End
      Begin VB.TextBox txtRequestedBy 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1140
         Width           =   2835
      End
      Begin VB.TextBox txtIssuedBy 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   420
         Width           =   2835
      End
      Begin VB.TextBox txtPreparedBy 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   60
         Width           =   2835
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
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
         Left            =   90
         TabIndex        =   55
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Received By"
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
         Left            =   90
         TabIndex        =   54
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Left            =   90
         TabIndex        =   53
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
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
         Left            =   90
         TabIndex        =   52
         Top             =   90
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmPMISTrans_CustomerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd                                           As ADODB.Recordset
Dim RSTDAYTRAN                                         As ADODB.Recordset
Dim RSPARTMAS                                          As ADODB.Recordset
Dim RSSALESMAN                                         As ADODB.Recordset
Dim RSCUNTER                                           As ADODB.Recordset
Dim RSPROFILE                                          As ADODB.Recordset
Dim rsSignatories                                      As ADODB.Recordset
Dim RSREPOR                                            As ADODB.Recordset
Dim RSCUSTOMER                                         As ADODB.Recordset

Dim KCNT                                               As Integer
Dim ADDOREDIT                                          As String
Dim ORD_TOTUPRICE                                      As Double
Dim ORD_TOTINVAMT                                      As Double
Dim ORD_TOTVAT                                         As Double
Dim ORD_TOTQTY                                         As Double
Dim PREVORDTYPE                                        As String
Dim PREVORDNO                                          As String
Dim REPOR_STATUS                                       As String
Dim LOCALACESS                                         As String





Sub ADBPRINTING()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS' ", gconDMIS
    Open App.Path & "\ADB.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = 'ADB' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0
        If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
            cntCOPY = 4
        Else
            cntCOPY = 2
        End If
        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ADVANCED BILL VOUCHER</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsOrd_Hd!TranType) & "-" & Null2String(rsOrd_Hd!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Charge To: " & Null2String(rsOrd_Hd!chargeto) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Ref RO# : " & Null2String(rsOrd_Hd!RoNo) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL RIV</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\ADB.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\ADB.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            'If chkPreview.Value = 1 Then
            '   browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            'Else
            '   browRIV.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
            'End If
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub BringToFront()
    Picture1.Enabled = False
    fraDetails.Enabled = False
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True
End Sub





Private Sub c_Click()
    cmdPISNum_Click
End Sub



Private Sub cboRefPRSNo_GotFocus()
    Dim rsPRS                                          As ADODB.Recordset
    Dim rsPRS_HDDup                                    As ADODB.Recordset
    Set rsPRS = New ADODB.Recordset

    If COUNTERTYPE = "RIV" Or COUNTERTYPE = "ADB" Then
        rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = 'P' and SALES_ORIGIN = 'S' order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    Else
        rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = 'P' and (SALES_ORIGIN = 'W' or SALES_ORIGIN = 'O' or SALES_ORIGIN = 'M')  order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If

    If Not rsPRS.EOF And Not rsPRS.BOF Then
        rsPRS.MoveFirst: cboRefPRSNo.Clear
        Do While Not rsPRS.EOF
            Set rsPRS_HDDup = New ADODB.Recordset
            rsPRS_HDDup.Open "select refpisno from PMIS_Ord_Hd where TRANTYPE <> 'PRS' AND [TYPE] = 'P' AND refprsno = '" & Null2String(rsPRS!refpisno) & "'", gconDMIS
            If Not rsPRS_HDDup.EOF And Not rsPRS_HDDup.BOF Then
            Else
                cboRefPRSNo.AddItem Null2String(rsPRS!refpisno)
            End If
            rsPRS.MoveNext
        Loop
    End If
End Sub

Private Sub cboRefPRSNo_LostFocus()
    If LTrim(RTrim(cboRefPRSNo)) = "" Then
        MessagePop InfoVoid, "Blank Fields Detected!", "Please Input Valid Requisition Number." & vbCrLf, 3000
        Exit Sub
    End If

    If ADDOREDIT = "ADD" Then
        Dim rsRR_HDDup                                 As ADODB.Recordset
        Set rsRR_HDDup = New ADODB.Recordset
        rsRR_HDDup.Open "select refpisno,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND refprsno = '" & cboRefPRSNo.Text & "'", gconDMIS
        If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
            MsgBox "PRS Number Already Received", vbInformation, "Invalid PRS Number"
            Exit Sub
        Else
            Set rsRR_HDDup = New ADODB.Recordset

            rsRR_HDDup.Open "select tranno,DS1 ,CustCode , RONO, CUSTNAME from PMIS_vw_PRS where [TYPE] = 'P' AND refpisno = '" & cboRefPRSNo.Text & "'", gconDMIS
            'CARE OF CEBU

            If Not rsRR_HDDup.EOF Or Not rsRR_HDDup.BOF Then
                txtCustName = Null2String(rsRR_HDDup!custname)
                txtCustCode = Null2String(rsRR_HDDup!custcode)
                txtRONO = Null2String(rsRR_HDDup!RoNo)
            End If

            If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then

                KCNT = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
                Dim STOCKDESCription                   As String
                Set RSTDAYTRAN = New ADODB.Recordset
                cleargrid grdDetails
                RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsRR_HDDup!TRANNO) & " and trantype = 'PRS' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                    cboChargeTo.Enabled = False: Screen.MousePointer = 11: RSTDAYTRAN.MoveFirst
                    Do While Not RSTDAYTRAN.EOF
                        KCNT = KCNT + 1
                        STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP))
                        grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
                                           Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                                           STOCKDESCription & Chr(9) & _
                                           N2Str2IntZero(RSTDAYTRAN!tranqty) & Chr(9) & _
                                           Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                                           Format(N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
                        ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                        ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                        ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                        RSTDAYTRAN.MoveNext
                    Loop
                    txtTTLInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
                    If N2Str2Zero(rsRR_HDDup!ds1) <> 0 Then
                        txtDS1.Text = N2Str2Zero(rsRR_HDDup!ds1)
                        If txtDS_Desc1.Text = "" Then
                            txtDS_Desc1.Text = "DISCOUNT"
                        End If
                        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(ORD_TOTUPRICE) * (NumericVal(txtDS1.Text) / 100))
                        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(ORD_TOTUPRICE) - NumericVal(txtDS_Amt1.Text))
                    Else
                        txtDS1.Text = N2Str2Zero(rsRR_HDDup!ds1)
                        txtDS_Desc1.Text = ""
                        txtDS_Amt1.Text = "0.00"
                        txtNetInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
                    End If
                    ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
                    If KCNT <> 0 Then grdDetails.RemoveItem 1
                    Screen.MousePointer = 0
                End If
            Else

                MessagePop InfoFriend, "Invalid Parts Requisition Number!", "Please Input Valid Requisition Number." & vbCrLf & "You will be not able to Proceed.", 3000


                If ADDOREDIT = "ADD" Then
                    cleargrid grdDetails
                End If
            End If
        End If
    End If
End Sub

Private Sub cboSMName_Click()
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan where signname = " & N2Str2Null(cboSMName.Text), gconDMIS
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        cboSalesMan.Text = Null2String(RSSALESMAN!empno)
    End If
End Sub

Private Sub cboSMName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSave.Value = True
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        Check1.Enabled = True
    Else
        Check1.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    cboTranPartNo_Change
End Sub
Private Sub cboTranPartNo_Click()
    cboTranPartNo_Change
End Sub

Function SetSTOCKDESC2(pid As Variant)
    If COUNTERTYPE = "ADB" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select PARTNUMBER,descriptio,dnpp,srp from PMIS_Dnpp where PARTNUMBER = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!DESCRIPTIO)
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!DNPP))
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
        Else
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        End If
    Else

        If pid <> "" Then
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select STOCKNO,id,STOCKDESC,srp,mac,dnp from PMIS_PARTMAS where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
                If txtTranType.Text = "DR" Then
                    If cboChargeTo.Text = "PARTS CLAIM" Then
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                    Else
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                    End If
                Else

                    If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                    ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                    Else
                        Dim rsPRS_Header               As ADODB.Recordset
                        Dim rsPRS_Details              As ADODB.Recordset
                        Set rsPRS_Header = New ADODB.Recordset
                        Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
                        If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
                            Set rsPRS_Details = New ADODB.Recordset
                            Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
                            If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
                                txtTranQty.Text = N2Str2Zero(rsPRS_Details!tranqty)
                            End If
                        End If

                        'txtTranUCost.Text = ComputeMacasofDate(Null2String(RSPARTMAS!STOCKNO), txtTranDate)
                        txtTranUCost.Text = ToDoubleNumber(RSPARTMAS!Mac)
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))

                    End If
                End If
            Else
                txtTranUPrice.Text = "0.00"
                txtTranUCost.Text = 0
            End If
        End If
    End If
End Function

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub



Private Sub Check1_Click()
    If Module_Access(LOGID, "APPLY PARTS COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    If Check1.Value = 1 Then
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        txtTranUPrice.Text = txtTranUCost.Text
    Else
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Function CheckIfRoExists(XXX As String) As String
    Dim rsRo_det                                       As ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(XXX))
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        CheckIfRoExists = UCase(Null2String(rsRo_det!REP_OR))
    End If
    Set rsRo_det = Nothing
End Function


Function CheckIfROBilled(XXX As String) As String
    Dim rsRo_det                                       As ADODB.Recordset
    Set rsRo_det = gconDMIS.Execute("SELECT INVOICE FROM CSMS_REPOR WHERE INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(XXX))
    If Not rsRo_det.EOF And Not rsRo_det.BOF Then
        CheckIfROBilled = UCase(Null2String(rsRo_det!Invoice))
    End If
    Set rsRo_det = Nothing
End Function

Sub CHGPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If

    'UPDATE : JBF 01/29/09
    '    If NumericVal(txtDS1.Text) = 0 Then
    '        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    '        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        Else
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '            Screen.MousePointer = 0
    '    Else
    '            Screen.MousePointer = 11
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDISC_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        Else
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '            Screen.MousePointer = 0
    '    End If

End Sub

Sub CHGPRINTING_OTC()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\PCHG.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = 'CHG' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0
        If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 1
        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS ISSUANCE SLIP (COUNTER-CHG)</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "COUNTER PIS-" & Null2String(rsOrd_Hd!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL PIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                'Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\PCHG.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\PCHG.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LOCALACESS) = False Then Exit Sub
    ADDOREDIT = "ADD"
    InitMemVars
    PisValidation
    Command3.Enabled = True
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    ADDOREDIT = "ADD"
    cmdTranDelete.Enabled = False
    InitParts
    On Error Resume Next
    cboTranPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    txtTranDate.Enabled = False
    StoreMemvars

    txtPRtranno.Visible = False
    Command3.Enabled = False
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LOCALACESS) = False Then Exit Sub

    On Error GoTo Errorcode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCURONHAND, PCurTISSQTY, PCURISSUANCES     As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = " & N2Str2Null(rsOrd_Hd!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!tranqty)
                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!tranqty)
                    PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) - N2Str2Zero(rsTdaytranDup!tranqty)
                    If Null2String(rsOrd_Hd!STATUS) = "P" Then
                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                          " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                          " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "PART NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        Else
                            SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                          " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                          " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW LOG AUDIT-------------------------------------------------
                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "PART NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), COUNTERTYPE, "")
                            'NEW LOG AUDIT-------------------------------------------------
                        End If
                        SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                      " onhand = " & PCURONHAND & "," & _
                                      " tissqty = " & PCurTISSQTY & "," & _
                                      " issuances = " & PCURISSUANCES & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        'NEW LOG AUDIT-------------------------------------------------
                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "PART NO: " & Null2String(N2Str2Null(rsTdaytranDup!STOCK_ORD)), "", "")
                        'NEW LOG AUDIT-------------------------------------------------
                    End If
                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                  " status = 'C'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & rsTdaytranDup!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""
                End If

                rsTdaytranDup.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " status = 'C'," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""
        rsRefresh
        On Error Resume Next
        rsOrd_Hd.Find "id =" & labID.Caption
        StoreMemvars
    End If
    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing

    Exit Sub

Errorcode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCALACESS) = False Then Exit Sub
    ADDOREDIT = "EDIT"
    PREVORDTYPE = txtTranType.Text
    PREVORDNO = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    cmdEditTranDate.Enabled = True
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    On Error Resume Next
    txtCustName.SetFocus
    Command3.Enabled = True
End Sub

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", LOCALACESS) = False Then Exit Sub
    txtTranDate.Enabled = True
    txtTranDate.Locked = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsOrd_Hd.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    rsOrd_Hd.MoveLast
    StoreMemvars
End Sub

Private Sub cmdNext_Click()
    rsOrd_Hd.MoveNext
    If rsOrd_Hd.EOF Then
        rsOrd_Hd.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPISNum_Click()
    With frmPMISPIFormation
        If ADDOREDIT = "EDIT" Then
            .txtedit = "EDIT"
        Else
            .txtedit = ""
        End If
        .lbl2 = Mid(txtReferencePIS, 3, 1)
        .lbl3 = Mid(txtReferencePIS, 4, 1)
        .lbl4 = Mid(txtReferencePIS, 5, 1)
        .lbl9.Text = Mid(txtReferencePIS, 9, 3)
        .lbl11 = Mid(txtReferencePIS, 12, 1)
        If .lbl2.Caption = "S" Then
            .optS.Value = True
        ElseIf .lbl2.Caption = "W" Then
            .optW.Value = True
        ElseIf .lbl2.Caption = "M" Then
            .optM.Value = True
        ElseIf .lbl2.Caption = "J" Then
            .optJ.Value = True
        ElseIf .lbl2.Caption = "O" Then
            .optO.Value = True
        End If
        If .lbl3.Caption = "G" Then
            .optG.Value = True
        ElseIf .lbl3.Caption = "B" Then
            .optB.Value = True
        End If
        If .lbl4.Caption = "C" Then
            .optC.Value = True
        ElseIf .lbl4.Caption = "I" Then
            .optI.Value = True
        ElseIf .lbl4.Caption = "W" Then
            .optW2.Value = True
        End If
        If .lbl11.Caption = "1" Then
            .opt1.Value = True
        ElseIf .lbl11.Caption = "2" Then
            .opt2.Value = True
        ElseIf .lbl11.Caption = "0" Then
            .opt0.Value = True
        End If
    End With
    frmPMISPIFormation.Show 1

    If txtTranType = "ADB" Then
        If Mid(txtReferencePIS, 3, 1) = "S" Then
            txtRONO.Visible = True
            txtRONO.ZOrder 0
            cmdSelectCustomer.Visible = False
        Else
            txtRONO.Visible = False
            cmdSelectCustomer.Visible = True
            cmdSelectCustomer.ZOrder 0
        End If
    End If

    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub txtPRtranno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTranNo.Text = txtPRtranno
        txtPRtranno.Visible = False
    End If
End Sub

Private Sub cmdPost_Click()


    Dim rsPrtMas                                       As New ADODB.Recordset
    Dim rsTdytran                                      As New ADODB.Recordset
    Dim blnStockremove                                 As Boolean
    Dim strPartno                                      As String
    Dim onhandasof                                     As Long
    Dim MACASOF                                        As Double
    blnStockremove = False
    If Function_Access(LOGID, "Acess_Post", LOCALACESS) = False Then Exit Sub

    If txtTranType <> "ADB" Then
        'CHECK ON HAND AS OF DATE
        Dim rsTran                                     As ADODB.Recordset
        Set rsTran = gconDMIS.Execute("SELECT STOCK_ORD, TRANQTY FROM PMIS_TDAYTRAN WHERE TRANTYPE=" & N2Str2Null(txtTranType) & " AND TYPE='P' AND TRANNO=" & N2Str2Null(txtTranNo))
        While Not rsTran.EOF
            onhandasof = COMPUTE_ONHANDASOFDATE(txtTranDate, Null2String(rsTran!STOCK_ORD), "P")
            If onhandasof <= 0 Then
                MsgBox "Zero Onhand Balance for the Part Number " & vbCrLf & Null2String(rsTran!STOCK_ORD) & " As of  " & txtTranDate & vbCrLf & "Posting of This Transaction is Allowed.", vbInformation
                Exit Sub
            End If

            If onhandasof - N2Str2IntZero(rsTran!tranqty) < 0 Then
                MsgBox "Negative Issuances for the Part Number " & vbCrLf & Null2String(rsTran!STOCK_ORD) & " As of  " & txtTranDate & vbCrLf & "Posting of This Transaction is Allowed." & vbCrLf & "", vbInformation
                Exit Sub
            End If
            rsTran.MoveNext
        Wend
    End If




    If txtTranType.Text = "RIV" Or txtTranType.Text = "ADB" Then

        If CheckIfROBilled(txtRONO.Text) <> "" Then
            MsgBox "Warning: This RO is Already been billed for this issuance" & vbCrLf & "Posting of Transaction Cannot be done for this RO", vbCritical, "Repair Order Already Billed"
            Exit Sub
        End If
    End If

    'On Error GoTo ErrorCode:
    '====================================================================================================
    Dim fild                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    If fild = "" Or fild = "No Entry" Then
        MsgBox "Posting of Transaction without issuance of Part(s) is not allowed.", vbCritical, "Pls. Add Part(s)."
        Exit Sub
    End If

    '=[ EAP:033109: check parts if current onhand is not zero in posting ]=
    If txtTranType = "RIV" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'P' and trantype in('RIV') "), gconDMIS
    ElseIf txtTranType = "CSH" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'P' and trantype in('CSH') "), gconDMIS
    ElseIf txtTranType = "CHG" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'P' and trantype in('CHG') "), gconDMIS
    Else
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'P' and trantype in('DR') "), gconDMIS
    End If

    If Not (rsTdytran.BOF And rsTdytran.EOF) Then
        Do While Not rsTdytran.EOF

            rsPrtMas.Open "Select STOCKNO,onhand from PMIS_PARTMAS where STOCKNO = '" & rsTdytran!STOCK_ORD & "' ", gconDMIS
            '=[ EAP:040209: this will remove the partnumber without stock in the transaction. ]=
            If Not (rsPrtMas.BOF And rsPrtMas.EOF) Then
                If rsPrtMas!ONHAND <= 0 Then
                    MsgBox "Partnumber# " & rsTdytran!STOCK_ORD & " will be remove from the transaction Out of Stock"
                    SQL_STATEMENT = "delete from PMIS_TdayTran where Id = '" & rsTdytran!ID & "' "
                    gconDMIS.Execute SQL_STATEMENT
                    blnStockremove = True
                ElseIf rsPrtMas!ONHAND < rsTdytran!tranqty Then
                    MsgBox "Some Part Number Onhand is less thatn your Requested Quantity", vbInformation
                    Exit Sub
                End If
                rsPrtMas.MoveNext
            End If
            rsPrtMas.Close

            rsTdytran.MoveNext
        Loop
    End If

    '=[ EAP:040209: if there's a partnumber that has been removed. transaction will not be posted ]=
    If blnStockremove Then
        cmdTranCancel.Value = True
        rsRefresh
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = True Then
        Dim PCURONHAND, PCurTISSQTY, PCURISSUANCES     As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = " & N2Str2Null(rsOrd_Hd!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly

        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                '====================================================
                'updating code: JAA - 06172008  -- Take Out the Validation for searching for ACTIVE Parts only.
                'rsPartmasDup.Open "select STOCKNO,onhand,TISSQTY,issuances,REQSERVED,S_REQSERVED,NON_HARI from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                rsPartmasDup.Open "select STOCKNO,onhand,TISSQTY,issuances,REQSERVED,S_REQSERVED,NON_HARI from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD) & "", gconDMIS
                '====================================================
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    '====================================================
                    'UPDATING CODE: JAA - 09082008  -- DO NOT DEDUCT STOCK FROM MASTER FILE.
                    If COUNTERTYPE <> "ADB" Then
                        PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) - N2Str2Zero(rsTdaytranDup!tranqty)
                        PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) + N2Str2Zero(rsTdaytranDup!tranqty)
                        PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) + N2Str2Zero(rsTdaytranDup!tranqty)

                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                          " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) + N2Str2Zero(rsTdaytranDup!tranqty) & _
                                          " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                            '===================================================================
                        Else
                            SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                          " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) + N2Str2Zero(rsTdaytranDup!tranqty) & _
                                          " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                            NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                        End If

                        SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                      " onhand = " & PCURONHAND & "," & _
                                      " tissqty = " & PCurTISSQTY & "," & _
                                      " issuances = " & PCURISSUANCES & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        NEW_LogAudit "E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                    End If
                    'MACASOF = ComputeMacasofDate(Null2String(rsTdaytranDup!STOCK_ORD), txtTranDate)
                    SQL_STATEMENT = "update PMIS_TdayTran set " & _
                                  " status = 'P'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & rsTdaytranDup!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "PP", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""

                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " status = 'P'," & _
                      " totalqty = " & ORD_TOTQTY & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labID.Caption

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""
        rsRefresh
        rsOrd_Hd.Find "id =" & labID.Caption
        StoreMemvars

        Set rsTdaytranDup = Nothing
        Set rsPartmasDup = Nothing

        If txtTranType.Text = "RIV" Or txtTranType.Text = "ADB" Then
            ImportParts txtRONO
        End If
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    rsOrd_Hd.MovePrevious
    If rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACESS) = False Then Exit Sub

    If rsOrd_Hd!TranType = "ADB" Or rsOrd_Hd!TranType = "RIV" Then
        If MsgQuestionBox("Parts Issuance Slip will be printed. You want to print it in a Blank form?", "Confirm Printing...") = True Then
            If COMPANY_CODE = "HCI" Then
                cmdPrintRIV_Click
            Else
                fraSignatories.Visible = True
                fraSignatories.ZOrder 0
                txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
                txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
                txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
                On Error Resume Next
                txtRequestedBy.SetFocus
            End If
        Else

            SERVICEPISPRINTING
        End If
    End If

    If rsOrd_Hd!TranType = "CSH" Then
        If MsgQuestionBox("Parts Issuance Slip (CSH) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            '==============================================
            'updating code:     JAA  - 02052008
            If COMPANY_CODE = "HMH" Then
                If MsgQuestionBox("Print Parts Issuance in a Blank form?", "Confirm Printing...") = True Then




                    fraSignatories.Visible = True
                    fraSignatories.ZOrder 0
                    txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
                    txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
                    txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
                    On Error Resume Next
                    txtRequestedBy.SetFocus
                Else
                    '                     '===[Update EAP: 072208 Print Audit ]===
                    '                     SaveReprintInformation strClassType + COUNTERTYPE, MODULENAME, txtTranNo.Text, "Null", LOGDATE, LOGNAME, False:
                    '                     If CANCEL_ANS = "NO" Then Exit Sub
                    '                     '****************************************

                    CSHPRINTING
                End If
            Else

                CSHPRINTING
            End If
            '==============================================
        End If
    End If
    If rsOrd_Hd!TranType = "CHG" Then
        If MsgQuestionBox("Parts Issuance Slip (CHG) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            If COMPANY_CODE = "HMH" Then
                If MsgQuestionBox("Print Parts Issuance in a Blank form?", "Confirm Printing...") = True Then



                    fraSignatories.Visible = True
                    fraSignatories.ZOrder 0
                    txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
                    txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
                    txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
                    On Error Resume Next
                    txtRequestedBy.SetFocus
                Else

                    CHGPRINTING
                End If
            Else

                CHGPRINTING
            End If
        End If
    End If
    If rsOrd_Hd!TranType = "DR" Then
        If MsgQuestionBox("DR Out Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then


            NEWDRPRINTING
        End If
    End If

    NEW_LogAudit "V", LOCALACESS, "", labID, "Parts", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    If rsOrd_Hd!TranType = "RIV" Then
        SERVICEPISPRINTING_BLANKFORM
    End If
    If rsOrd_Hd!TranType = "ADB" Then
        ADBPRINTING
    End If
    If rsOrd_Hd!TranType = "CSH" Then
        CSHPRINTING_OTC
    End If
    If rsOrd_Hd!TranType = "CHG" Then
        CHGPRINTING_OTC
    End If

    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)

    SendToBack
End Sub

Private Sub cmdSave_Click()
    Dim RSRO                                           As ADODB.Recordset
    If Len(Trim(RTrim(txtTranNo))) <> 6 Then
        MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(cboRefPRSNo.Text)) = "" Then
        MsgBox "Reference PRS Number is Required...", vbCritical, "Pls. select PRS No."
        On Error Resume Next
        cboRefPRSNo.SetFocus
        Exit Sub
    End If

    If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
        MsgBox "Invalid Reference PIS Number!", vbCritical, "PIS Required!"
        On Error Resume Next
        txtReferencePIS.SetFocus
        Exit Sub
    End If

    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical, "Transaction Date Error"
            On Error Resume Next
            txtTranDate.SetFocus
        End If
    Else
        MsgBox "Please Input Valid Date!!", vbInformation
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If

    Select Case txtTranType
        Case "CHG"
            If LTrim(RTrim(txtTerms.Text)) = "" Then
                MsgBox "Terms must have a value", vbInformation
                On Error Resume Next
                txtTerms.SetFocus
                Exit Sub
            End If
        Case "RIV"
            If gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_REPOR WHERE REP_OR=" & N2Str2Null(LTrim(RTrim(Replace(txtRONO, "'", ""))))).Fields(0).Value = 0 Then
                MsgBox "RO Number Doesn't Exists. Please Correct Repair Order Number", vbInformation
                On Error Resume Next
                txtRONO.SetFocus
                Exit Sub
            End If
            If CheckIfROBilled(txtRONO) <> "" Then
                On Error Resume Next
                MsgBox "Repair Order " & txtRONO & " is already been invoiced." & vbCrLf & "Cannot Issue any Item for particular Repair Order.", vbInformation
                txtRONO.SetFocus
                Exit Sub
            End If

            Dim RSADB                                  As ADODB.Recordset
            Set RSADB = gconDMIS.Execute("SELECT COUNT(RONO) FROM PMIS_ORD_HD WHERE TRANTYPE='ADB' AND RONO=" & N2Str2Null(txtRONO))
            If RSADB.Fields(0).Value > 0 Then
                If MsgBox("There is Advance Bill for this RO!!" & vbCrLf & " Are you Sure You will do service issuance(s)?", vbInformation + vbYesNo, "Advance Bill Deteched!!") = vbNo Then
                    On Error Resume Next
                    txtRONO.SetFocus
                    Exit Sub
                End If
            End If

        Case "ADB"
            'SERVICE ADVANCE BILL : CHECK RO EXISTS
            If Mid(txtReferencePIS, 3, 1) = "S" Then
                If gconDMIS.Execute("SELECT COUNT(*) FROM CSMS_REPOR WHERE REP_OR=" & N2Str2Null(LTrim(RTrim(Replace(txtRONO, "'", ""))))).Fields(0).Value = 0 Then
                    MsgBox "RO Number Doesn't Exists. Please Correct Repair Order Number", vbInformation
                    On Error Resume Next
                    txtRONO.SetFocus
                    Exit Sub
                End If

                If CheckIfROBilled(txtRONO) <> "" Then
                    MsgBox "Repair Order " & txtRONO & " is already been invoiced." & vbCrLf & "Cannot Issue any Item for particular Repair Order.", vbCritical
                    On Error Resume Next
                    txtRONO.SetFocus
                    Exit Sub
                End If
            End If
    End Select

    If LTrim(RTrim(txtCustCode)) = "" Then
        MsgBox "Customer Information Is Required...", vbInformation, "Pls Select Customer Information..."
        Exit Sub
    End If


    On Error GoTo Errorcode

    Dim NEXTCUNTER                                     As String
    Dim RSFINDDUP                                      As ADODB.Recordset
    Dim XSALES_ORIGIN                                  As String
    Dim XSI_TYPE                                       As String
    Dim XPAY_CLASS                                     As String
    Dim XCHAR_YEAR                                     As String
    Dim XCHAR_MONTH                                    As String
    Dim XIS_SERIES                                     As String
    Dim XTRACK_CODE                                    As String
    Dim VCBOSALESMAN                                   As String
    Dim VCBOSMNAME                                     As String
    Dim VTXTTRANTYPE                                   As String
    Dim VTXTTRANNO                                     As String
    Dim VTXTTRANDATE                                   As String
    Dim VTXTCUSTCODE                                   As String
    Dim VTXTCUSTNAME                                   As String
    Dim VTXTCHARGETO                                   As String
    Dim VTXTREP_OR                                     As String
    Dim VTXTREFPRSNO                                   As String
    Dim VTXTRONO                                       As String
    Dim VtxtTerms                                      As String
    Dim VStatus                                        As String
    Dim VTXTTTLINVAMT                                  As Double
    Dim VTXTDS1                                        As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1                                    As Double
    Dim VTXTNETINVAMT                                  As Double
    Dim VTXTRemarks                                    As String
    Dim Vusercode                                      As String
    Dim VLastUpdate                                    As String
    Dim VIN_PROCESS                                    As String
    Dim VTXTREFERENCEPIS                               As String



    Set RSFINDDUP = gconDMIS.Execute("SELECT ISNULL(SUM(TQ),0) AS TQ  FROM ( " & vbCrLf & _
                                   " SELECT COUNT(*) TQ FROM PMIS_ORD_HD WHERE [TYPE] = 'P' AND TRANTYPE = '" & txtTranType & "' AND TRANNO = '" & txtTranNo & "' " & vbCrLf & _
                                   " UNION " & vbCrLf & _
                                   " SELECT COUNT(*) FROM PMIS_ORD_HIST WHERE [TYPE] = 'P' AND TRANTYPE = '" & txtTranType & "' AND TRANNO = '" & txtTranNo & "')T ")
    If ADDOREDIT = "ADD" Then
        If RSFINDDUP.Fields(0).Value > 0 Then
            MsgBox "Transaction No. already exist!", vbCritical, "Duplicate Transaction Number"
            On Error Resume Next
            txtTranNo.SetFocus
            Exit Sub
        End If

    Else
        If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(rsOrd_Hd!TRANNO))) Then
            If RSFINDDUP.Fields(0).Value > 1 Then
                MsgBox "Transaction No. already exist!", vbCritical, "Duplicate Transaction Number"
                On Error Resume Next
                Exit Sub
            End If
        End If
    End If



    VCBOSALESMAN = N2Str2Null(cboSalesMan.Text)
    VCBOSMNAME = N2Str2Null(cboSMName.Text)

    If Left(txtTranNo.Text, 1) <> "P" Then
        NEXTCUNTER = NumericVal(txtTranNo.Text) + 1
    End If

    VTXTTRANTYPE = N2Str2Null(txtTranType.Text)
    VTXTTRANNO = N2Str2Null(txtTranNo.Text)
    VTXTTRANDATE = N2Date2Null(txtTranDate.Text)
    VTXTCUSTCODE = N2Str2Null(txtCustCode.Text)
    VTXTCUSTNAME = N2Str2Null(txtCustName.Text)
    VTXTREFERENCEPIS = N2Str2Null(txtReferencePIS.Text)
    VTXTREFPRSNO = N2Str2Null(cboRefPRSNo.Text)
    VIN_PROCESS = "'Y'"
    VTXTCHARGETO = "'VAR'"

    Dim RRTRANDATE                                     As String
    Dim RRTRANNO                                       As String
    Dim RRTRANTYPE                                     As String
    Dim RRITEMNO                                       As String
    Dim RRSTOCK_ORD                                    As String
    Dim RRSTOCK_SUP                                    As String
    Dim RRTRANQTY                                      As Integer
    Dim RRTRANUCOST                                    As Double
    Dim RRTRANINVAMT                                   As Double
    Dim RRIN_OUT                                       As String
    Dim RRSTATUS                                       As String

    VTXTRONO = N2Str2Null(txtRONO.Text)
    If Len(txtRONO.Text) = 7 Then
        VTXTREP_OR = "'" & Left(txtRONO.Text, 1) & "-" & Right(txtRONO.Text, 6) & "'"
    Else
        VTXTREP_OR = "NULL"
    End If

    VtxtTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLINVAMT = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNETINVAMT = NumericVal(txtNetInvAmt.Text)
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"
    XSALES_ORIGIN = N2Str2Null(Mid(txtReferencePIS, 3, 1))
    XSI_TYPE = N2Str2Null(Mid(txtReferencePIS, 4, 1))
    XPAY_CLASS = N2Str2Null(Mid(txtReferencePIS, 5, 1))
    XCHAR_YEAR = N2Str2Null(Mid(txtReferencePIS, 6, 2))
    XCHAR_MONTH = N2Str2Null(Mid(txtReferencePIS, 8, 1))
    XIS_SERIES = N2Str2Null(Mid(txtReferencePIS, 9, 3))
    XTRACK_CODE = N2Str2Null(Mid(txtReferencePIS, 12, 1))
    VStatus = "'N'"
    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = Replace(txtRemarks.Text, Chr(13), "")
        VTXTRemarks = Replace(txtRemarks.Text, Chr(9), "")
        VTXTRemarks = Replace(Trim(txtRemarks.Text), Chr(27), "")
        VTXTRemarks = N2Str2Null(VTXTRemarks)
    End If

    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "INSERT INTO PMIS_ORD_HD" & _
                      " (TYPE,TRANTYPE,TRANNO,TRANDATE,CUSTCODE,CUSTNAME,CHARGETO,REFPRSNO,RONO,REP_OR,SALESMAN,SMNAME,TERMS,TTLINVAMT,DS1,DS_DESC1,DS_AMT1,NETINVAMT,REMARKS,STATUS,USERCODE,LASTUPDATE,IN_PROCESS,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " VALUES ('P'," & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & "," & VTXTREFPRSNO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("A", LOCALACESS, SQL_STATEMENT, FindTransactionID(txtTranNo, "TRANNO", "PMIS_ORD_HD", "DETAILS", N2Str2Null("P"), "TYPE"), "PARTS", txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " TRANTYPE = " & VTXTTRANTYPE & "," & _
                      " TRANNO = " & VTXTTRANNO & "," & _
                      " TRANDATE = " & VTXTTRANDATE & "," & _
                      " CUSTCODE = " & VTXTCUSTCODE & "," & _
                      " CUSTNAME = " & VTXTCUSTNAME & "," & _
                      " CHARGETO = " & VTXTCHARGETO & "," & _
                      " REFPRSNO = " & VTXTREFPRSNO & "," & _
                      " RONO = " & VTXTRONO & "," & _
                      " REP_OR = " & VTXTREP_OR & "," & _
                      " SALESMAN = " & VCBOSALESMAN & "," & _
                      " SMNAME = " & VCBOSMNAME & "," & _
                      " TERMS = " & VtxtTerms & "," & _
                      " TTLINVAMT = " & VTXTTTLINVAMT & "," & _
                      " DS1 = " & VTXTDS1 & "," & _
                      " DS_DESC1 = " & VTXTDS_Desc1 & "," & _
                      " DS_AMT1 = " & VTXTDS_Amt1 & "," & _
                      " NETINVAMT = " & VTXTNETINVAMT & "," & _
                      " REMARKS = " & VTXTRemarks & ", " & _
                      " STATUS = " & VStatus & ", " & _
                      " USERCODE = " & Vusercode & ", " & _
                      " IN_PROCESS = " & VIN_PROCESS & ", " & _
                      " REFPISNO = " & VTXTREFERENCEPIS & ", " & _
                      " LASTUPDATE = " & VLastUpdate & _
                      " WHERE ID = " & labID.Caption

        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "PARTS", "TRAN NO: " & txtTranNo, COUNTERTYPE, "")

        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " SALES_ORIGIN = " & XSALES_ORIGIN & "," & _
                      " SI_TYPE = " & XSI_TYPE & "," & _
                      " PAY_CLASS = " & XPAY_CLASS & "," & _
                      " CHAR_YEAR = " & XCHAR_YEAR & "," & _
                      " CHAR_MONTH = " & XCHAR_MONTH & "," & _
                      " IS_SERIES = " & XIS_SERIES & "," & _
                      " TRACK_CODE = " & XTRACK_CODE & "" & _
                      " WHERE ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")

        SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                      " TRANTYPE = " & VTXTTRANTYPE & "," & _
                      " TRANDATE = " & VTXTTRANDATE & "," & _
                      " TRANNO = " & VTXTTRANNO & _
                      " WHERE [TYPE] = 'P' AND TRANTYPE = '" & PREVORDTYPE & "' AND TRANNO = '" & Null2String(rsOrd_Hd!TRANNO) & "'"
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("EE", LOCALACESS, SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, "")
        ShowSuccessFullyUpdated
    End If
    If ADDOREDIT = "ADD" Then
        If Left(txtTranNo.Text, 1) = "P" Then
            'do nOthing
        Else
            SQL_STATEMENT = "update PMIS_Counter set nextnumber = '" & NEXTCUNTER & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' where [TYPE] = 'P' AND modul = " & VTXTTRANTYPE
            gconDMIS.Execute SQL_STATEMENT
        End If
        Call NEW_LogAudit("E", "PARTS COUNTER", SQL_STATEMENT, FindTransactionID(VTXTTRANTYPE, "MODUL", "PMIS_Counter", "DETAILS", N2Str2Null("P"), "TYPE"), "", "MODUL: " & Null2String(VTXTTRANTYPE), "", "")
        Call FillGrid
    Else
        rsRefresh
        rsOrd_Hd.Find "TRANNO = " & VTXTTRANNO
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                      " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                      " NETINVAMT = " & ORD_TOTINVAMT & _
                      " WHERE [TYPE] = 'P' AND TRANNO = " & VTXTTRANNO & " AND TRANTYPE = " & VTXTTRANTYPE
        gconDMIS.Execute SQL_STATEMENT
        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "", "TRAN NO: " & txtTranNo, "", "")
    End If

    fraDetails.Enabled = True
    rsRefresh
    rsOrd_Hd.Find "tranno = " & VTXTTRANNO
    cmdCancel.Value = True

    On Error GoTo Errorcode
    If ADDOREDIT = "ADD" Then
        Dim rsTdaytranDup                              As ADODB.Recordset
        Dim rstdaytranDUp2                             As ADODB.Recordset
        Dim RSPRS_HD                                   As ADODB.Recordset
        Dim rsPartMasClone                             As ADODB.Recordset
        Dim ISS_CNT                                    As Integer
        Dim VMACSTOCKNO                                As Double

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO), gconDMIS

        If rsTdaytranDup.EOF And rsTdaytranDup.BOF Then
            rsTdaytranDup.Close
            Set RSPRS_HD = gconDMIS.Execute("Select * from PMIS_vw_PRS where refpisno = '" & cboRefPRSNo.Text & "'")
            If Not RSPRS_HD.EOF And Not RSPRS_HD.BOF Then
                Set rstdaytranDUp2 = New ADODB.Recordset
                rstdaytranDUp2.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,tranuprice from PMIS_TdayTran where trantype = 'PRS' and tranno = " & N2Str2Null(RSPRS_HD!TRANNO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rstdaytranDUp2.EOF And Not rstdaytranDUp2.BOF Then
                    rstdaytranDUp2.MoveFirst: ISS_CNT = 0
                    Do While Not rstdaytranDUp2.EOF
                        Set rsPartMasClone = gconDMIS.Execute("Select STOCKNO,ONHAND,NON_HARI,MAC from PMIS_StockMas where TYPE = 'P' and STOCKNO = " & N2Str2Null(rstdaytranDUp2!STOCK_ORD))
                        If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
                            If COUNTERTYPE = "ADB" Then
                                ISS_CNT = ISS_CNT + 1
                                '===================================
                                'UPDATING CODE:     JAA - 09082008          - INCLUDE MAC UPON SAVING OF TRANSACTION
                                'VMACSTOCKNO = N2Str2Zero(rsPartMasClone!Mac)
                                '===================================
                                VMACSTOCKNO = 0
                                RRTRANDATE = N2Str2Null(rsOrd_Hd!trandate)
                                RRTRANTYPE = "'" & COUNTERTYPE & "'"
                                RRTRANNO = N2Str2Null(rsOrd_Hd!TRANNO)
                                RRITEMNO = N2Str2Null(Format(Null2String(rstdaytranDUp2!itemno), "0000"))
                                RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!STOCK_ORD)
                                RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!STOCK_SUP)
                                RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!tranqty)
                                RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANUPRICE)
                                RRTRANUCOST = 0
                                RRIN_OUT = "'O'"

                                RRSTATUS = "'N'"
                                SQL_STATEMENT = "INSERT INTO PMIS_TDAYTRAN " & _
                                                "(TYPE,MAC,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,TRANUPRICE,LASTUPDATE,USERCODE,STATUS,IN_OUT,NON_HARI)" & _
                                              " VALUES ('P'," & VMACSTOCKNO & "," & RRTRANDATE & ", '" & COUNTERTYPE & "', " & RRTRANNO & "," & _
                                              " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                              " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                              " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & "," & N2Str2Null(rsPartMasClone!NON_HARI) & ")"
                                gconDMIS.Execute SQL_STATEMENT
                            Else
                                If N2Str2Zero(rsPartMasClone!ONHAND) > 0 Then
                                    ISS_CNT = ISS_CNT + 1
                                    '===================================
                                    'UPDATING CODE:     JAA - 09082008          - INCLUDE MAC UPON SAVING OF TRANSACTION
                                    VMACSTOCKNO = N2Str2Zero(rsPartMasClone!Mac)
                                    '===================================
                                    RRTRANDATE = N2Str2Null(rsOrd_Hd!trandate)
                                    RRTRANTYPE = "'" & COUNTERTYPE & "'"
                                    RRTRANNO = N2Str2Null(rsOrd_Hd!TRANNO)
                                    RRITEMNO = N2Str2Null(Format(Null2String(rstdaytranDUp2!itemno), "0000"))
                                    RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!STOCK_ORD)
                                    RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!STOCK_SUP)
                                    If N2Str2Zero(rsPartMasClone!ONHAND) < N2Str2IntZero(rstdaytranDUp2!tranqty) Then
                                        MsgBox "Warning: Requested Quantity on " + N2Str2Null(rstdaytranDUp2!STOCK_ORD) + " is greater than available stock!" & vbCrLf & "System will default the available stock only.", vbInformation, "Requested Exceeds available stock on-hand"
                                        RRTRANQTY = N2Str2Zero(rsPartMasClone!ONHAND)
                                    Else
                                        RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!tranqty)
                                    End If
                                    RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANUPRICE)
                                    RRTRANUCOST = N2Str2Zero(VMACSTOCKNO)
                                    RRIN_OUT = "'O'"
                                    RRSTATUS = "'N'"
                                    SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                                    "(TYPE,mac,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out,NON_HARI)" & _
                                                  " values ('P'," & VMACSTOCKNO & "," & RRTRANDATE & ", '" & COUNTERTYPE & "', " & RRTRANNO & "," & _
                                                  " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                                  " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                                  " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & "," & N2Str2Null(rsPartMasClone!NON_HARI) & ")"
                                    gconDMIS.Execute SQL_STATEMENT
                                Else
                                    MsgBox "Requested Part No. " & Null2String(rstdaytranDUp2!STOCK_ORD) & " doesn't have Stock in your Master File", vbInformation, "Cannot Add Parts!"
                                    FillDetails
                                End If
                            End If
                            NEW_LogAudit "A", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, ""
                        Else
                            MsgBox "Requested Part No. " & Null2String(rstdaytranDUp2!STOCK_ORD) & " is not yet active in your Master File", vbInformation, "Cannot Add Parts!"
                            FillDetails
                        End If
                        rstdaytranDUp2.MoveNext
                    Loop
                End If
            End If
            cleargrid grdDetails
            FillDetails

            SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                          " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                          " netinvamt = " & ORD_TOTINVAMT & _
                          " where [TYPE] = 'P' AND tranno = " & VTXTTRANNO & " and trantype = " & VTXTTRANTYPE
            gconDMIS.Execute SQL_STATEMENT
            '=============================
            'updating code: JAA - 05242008
            If COUNTERTYPE = "DR" Then
                cmdAddTran_Click
            End If
            '=============================
        Else
            cleargrid grdDetails
            FillDetails
            cmdAddTran_Click
        End If
    End If

    FillGrid
    'NO NEED WE HAVE SEPERATE MODULE NOW AXP
    '    If ADDOREDIT = "ADD" Then
    '        InsertAdvanceBill
    '    End If
    Exit Sub

Errorcode:
    MsgBox err.Description
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemvars
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo Errorcode:

    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_TdayTran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", LOCALACESS, SQL_STATEMENT, labID, "Parts", "PART NO: " & cboTranPartNo, COUNTERTYPE, labDetID
        '        gconDMIS.Execute (" delete from CSMS_RO_Det where rep_or = 'txtRONO' and detcde = 'cboTranPartNo' ")
        ShowDeletedMsg
    End If
    Dim cnt                                            As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "select id,itemno from PMIS_TdayTran where [TYPE] = 'P' AND trantype = " & N2Str2Null(COUNTERTYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " order by itemno asc", gconDMIS
    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        cnt = 0
        Do While Not rsTdaytranDup.EOF
            cnt = cnt + 1
            SQL_STATEMENT = "update PMIS_TdayTran set itemno = " & Format(cnt, "0000") & " where id = " & rsTdaytranDup!ID
            gconDMIS.Execute SQL_STATEMENT
            rsTdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                  " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                  " netinvamt = " & ORD_TOTINVAMT & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "PARTS", "TRAN NO: " & txtTranNo, COUNTERTYPE, "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------

    rsRefresh
    On Error Resume Next
    rsOrd_Hd.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub


Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If ADDOREDIT = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset

        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_TdayTran where [TYPE] = 'P' AND STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(rsOrd_Hd!TRANNO) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Part Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Exit Sub
        End If
        Set rsTDaytranClone = Nothing
    End If

    Dim ORDTRANDATE, ORDTRANNO, ORDTRANTYPE            As String
    Dim ORDITEMNO, ORDSTOCK_ORD, ORDSTOCK_SUP          As String
    Dim ORDTRANQTY                                     As Integer
    Dim ORDTRANUCOST                                   As Double
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double
    Dim ORDMAC                                         As Double
    Dim CRITICAL_QUESTION                              As String
    Dim onhandasof                                     As Long

    If txtTranType.Text <> "ADB" Then
        Dim CurONHAND                                  As Long
        Dim CurSAFESTOCK                               As Long
        Dim CurTISSQTY                                 As Long
        Dim curRESSERVICE                              As Long
        Dim curIssuances                               As Long
        Dim PrevCurOrdQty                              As Long

        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances,MAC,NON_HARI from PMIS_PARTMAS where STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y'", gconDMIS
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            onhandasof = COMPUTE_ONHANDASOFDATE(txtTranDate, Null2String(RSPARTMAS!STOCKNO), "P")
            CurONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
            CurSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
            CurTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            curRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
            curIssuances = N2Str2IntZero(RSPARTMAS!ISSUANCES)

            If onhandasof <= 0 Then
                MsgBox "On hand as of " & txtTranDate & " is " & onhandasof & vbCrLf & "Cannot Issue Item Stock Ledger Won't Be Balanced.", vbInformation, "Invalid Date Or Quantity"
                Exit Sub
            End If
            If onhandasof - txtTranQty < 0 Then
                MsgBox "Qty Ordered Exceeds Current Stock for the Date " & txtTranDate, vbInformation
                Exit Sub
            End If

            ORDMAC = NumericVal(RSPARTMAS!Mac)
            'ORDMAC = ComputeMacasofDate(cboTranPartNo, txtTranDate)

            If ORDMAC <= 0 Then
                Screen.MousePointer = 0
                MsgBox "Warning: This Part Number has Zero Cost! Pls Check in Parts Master File or Process Update Master File to Proceed.", vbCritical, "Stock Has Zero Cost"
                Screen.MousePointer = 0
                Exit Sub
            Else
                txtTranUCost.Text = ORDMAC
            End If

            If ADDOREDIT <> "ADD" Then
                PrevCurOrdQty = NumericVal(labPrevOrdQty.Caption)
                CurTISSQTY = CurTISSQTY - PrevCurOrdQty
                curIssuances = curIssuances - PrevCurOrdQty
            End If

            If CurONHAND <= 0 Then
                Screen.MousePointer = 0
                MsgSpeechBox "Out of Stock!"
                Exit Sub
            End If

            If txtTranType.Text = "CSH" Or txtTranType.Text = "CHG" Then
                If CurONHAND <= curRESSERVICE Then
                    Screen.MousePointer = 0
                    If MsgQuestionBox("Stock is Reserved for Service... Continue Anyway?", "Stock Status Alert!") = False Then
                        Exit Sub
                    End If
                    CRITICAL_QUESTION = "Stock is Reserved for Service... Continue Anyway?"
                    Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, COUNTERTYPE, "")
                    MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
                End If
            End If

            If NumericVal(txtTranQty.Text) > CurONHAND Then
                Screen.MousePointer = 0
                MsgSpeechBox "Qty Ordered Exceeds Current Stock!"
                On Error Resume Next
                txtTranQty.SetFocus
                Exit Sub
            Else
                CurONHAND = CurONHAND - NumericVal(txtTranQty.Text)
            End If

            If CurONHAND < CurSAFESTOCK Then
                Screen.MousePointer = 0
                If MsgQuestionBox("Current On-hand is now below the Safety Stock Level... Proceed Anyway?", "Safety Stock Alert!") = False Then
                    Exit Sub
                End If
                CRITICAL_QUESTION = "Current On-hand is now below the Safety Stock Level... Proceed Anyway?"

                Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " " & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, COUNTERTYPE, "")
                MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
                Screen.MousePointer = 11
            End If
        Else
            Screen.MousePointer = 0
            MsgSpeechBox "Part Number Not Found!"
            Exit Sub
        End If
    End If

    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
    ORDTRANTYPE = N2Str2Null(txtTranType.Text)
    ORDTRANNO = N2Str2Null(txtTranNo.Text)
    ORDITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))

    ORDSTOCK_ORD = N2Str2Null(cboTranPartNo.Text)
    If txtTranType.Text = "ADB" Then ORDSTOCK_SUP = N2Str2Null(Left(txtTranDescription.Text, 100)) Else ORDSTOCK_SUP = N2Str2Null(cboTranPartNo.Text)
    ORDTRANQTY = NumericVal(txtTranQty.Text)
    'FML - 0423208 - MODIFIED TRANUCOST TO DEFAULT MAC
    'ORDTRANUCOST = NumericVal(txtTranUCost.Text)
    ORDTRANUCOST = ORDMAC
    ORDTRANINVAMT = NumericVal(txtTranUPrice.Text)
    If Round(ORDTRANINVAMT, 2) < Round(ORDTRANUCOST, 2) Then
        If COMPANY_CODE = "HAS" Then
            If Mid(txtReferencePIS.Text, 3, 1) = "W" Then
                If MsgBox("Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & " Do you want to Proceed", vbQuestion + vbYesNo, "PMIS") = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
                CRITICAL_QUESTION = "Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & " Do you want to Proceed"
                Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " " & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, COUNTERTYPE, "")
                MsgBox "User Action has been Log to Audit Trail", vbInformation, "Audit Trail Information"
            End If
        Else
            Screen.MousePointer = 0
            MsgBox "Warning: Issuance Unit Price for this Part Number is less than its Cost!" & vbCrLf & "System will not allow this transaction to Proceed.", vbCritical, "Unit Price is Below Cost"
            Exit Sub
        End If
    End If

    If txtTranType.Text = "ADB" Then ORDIN_OUT = "'A'" Else ORDIN_OUT = "'O'"
    ORDSTATUS = "'N'"

    'UPDATE BY : MJP 05-20-2008
    Dim HARI_NONHARI                                   As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select NON_HARI from PMIS_PARTMAS where STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y'")
    If Not RSTMP.EOF And Not RSTMP.BOF Then
        HARI_NONHARI = N2Str2Null(RSTMP!NON_HARI)
    Else
        HARI_NONHARI = N2Str2Null("")
    End If
    'UPDATE BY : MJP 05-20-2008


    'UPDATING CODE: JAA - 05152008 - I ADDED NON_HARI VALUE in INSERTING INTO TDAYTRAN TO UPDATE THE NON_HARI WHENEVER USER WILL ADD NEW PARTS
    If ADDOREDIT = "ADD" Then

        SQL_STATEMENT = "INSERT INTO PMIS_TDAYTRAN " & _
                        "(TYPE,TRANDATE,TRANTYPE,TRANNO,ITEMNO,STOCK_ORD,STOCK_SUP,TRANQTY,TRANUCOST,MAC,TRANUPRICE,LASTUPDATE,USERCODE,STATUS,IN_OUT,NON_HARI)" & _
                      " VALUES ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDMAC & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & "," & HARI_NONHARI & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(txtTranNo), "TRANNO", "PMIS_ORD_HD", "DETAILS", N2Str2Null(Null2String(ORDTRANTYPE)), "TRANTYPE"), "Parts", "PART NO: " & cboTranPartNo, COUNTERTYPE, labDetID
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                      " TRANDATE = " & ORDTRANDATE & "," & _
                      " TRANTYPE = " & ORDTRANTYPE & "," & _
                      " TRANNO = " & ORDTRANNO & "," & _
                      " ITEMNO = " & ORDITEMNO & "," & _
                      " STOCK_ORD = " & ORDSTOCK_ORD & "," & _
                      " STOCK_SUP = " & ORDSTOCK_SUP & "," & _
                      " MAC= " & ORDMAC & "," & _
                      " TRANQTY = " & ORDTRANQTY & "," & _
                      " TRANUCOST = " & ORDTRANUCOST & "," & _
                      " TRANUPRICE = " & ORDTRANINVAMT & "," & _
                      " LASTUPDATE = '" & LOGDATE & "'," & _
                      " STATUS = " & ORDSTATUS & "," & _
                      " IN_OUT = " & ORDIN_OUT & "," & _
                      " USERCODE = " & N2Str2Null(LOGCODE) & "" & _
                      " WHERE ID = " & labDetID.Caption

        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "PARTS", "TRAN NO: " & txtTranNo, COUNTERTYPE, labDetID
        ShowSuccessFullyUpdated
    End If
    cleargrid grdDetails
    FillDetails
    SQL_STATEMENT = "UPDATE PMIS_ORD_HD SET" & _
                  " TOTALQTY = " & ORD_TOTQTY & "," & _
                  " TTLINVAMT = " & ORD_TOTUPRICE & "," & _
                  " NETINVAMT = " & ORD_TOTINVAMT & _
                  " WHERE ID = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    'NEW LOG AUDIT------------------------------------------------------------
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "P", "TRAN NO: " & txtTranNo, "", "")
    'NEW LOG AUDIT------------------------------------------------------------

    Dim rsPRS_Header                                   As ADODB.Recordset
    Dim rsPRS_Details                                  As ADODB.Recordset
    Set rsPRS_Header = New ADODB.Recordset
    Set rsPRS_Header = gconDMIS.Execute("Select * from PMIS_vw_PRS where REFPISNO = '" & cboRefPRSNo.Text & "'")
    If Not rsPRS_Header.EOF And Not rsPRS_Header.BOF Then
        Set rsPRS_Details = New ADODB.Recordset
        Set rsPRS_Details = gconDMIS.Execute("Select * from PMIS_vw_PRS_Tran Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text))
        If Not rsPRS_Details.EOF And Not rsPRS_Details.BOF Then
            SQL_STATEMENT = "Update PMIS_vw_PRS_Tran set TRemarks = 'SERVED'  Where Tranno = " & N2Str2Null(rsPRS_Header!TRANNO) & " AND STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text)
            gconDMIS.Execute SQL_STATEMENT
            '====================================================================
            'If AddorEdit = "ADD" Then
            '    NEW_LogAudit "AA", LOCALACESS, SQL_STATEMENT, labID, "Parts", txtTranNo, COUNTERTYPE, labDetID
            'Else
            NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo, COUNTERTYPE, labDetID
            'End If
            '====================================================================
        Else
        End If
    Else
    End If


    rsRefresh
    On Error Resume Next
    rsOrd_Hd.Find "id = " & labID.Caption
    StoreMemvars
    Screen.MousePointer = 0
    If ADDOREDIT = "ADD" Then
        cmdAddTran_Click
        fraDetails.Enabled = False
        Picture1.Enabled = False
    Else
        cmdTranCancel.Value = True
    End If
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub



Private Sub cmdSelectCustomer_Click()
    frmCustomerSearch.Show 1
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If Module_Access(LOGID, "EDIT PARTS ISSUANCE AMOUNT", "SYSTEM") = False Then Exit Sub
    txtTranUPrice.Enabled = True
End Sub

Sub CSHPRINTING()


    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If

    'UPDATE : JBF 01/29/09
    '    If NumericVal(txtDS1.Text) = 0 Then
    '         rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    '         rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        Else
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '            Screen.MousePointer = 0
    '    Else
    '
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDISC_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        Else
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '            Screen.MousePointer = 0
    '    End If

End Sub

Sub CSHPRINTING_OTC()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\PCSH.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = 'CSH' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0
        If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 1
        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS ISSUANCE SLIP (COUNTER-CSH)</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "COUNTER PIS-" & Null2String(rsOrd_Hd!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL PIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                'Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\PCSH.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\PCSH.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub DRPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "DR.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub FillCboSalesMan()
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan order by signname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        RSSALESMAN.MoveFirst: cboSalesMan.Clear: cboSMName.Clear
        Do While Not RSSALESMAN.EOF
            cboSalesMan.AddItem Null2String(RSSALESMAN!empno)
            cboSMName.AddItem Null2String(RSSALESMAN!signname)
            RSSALESMAN.MoveNext
        Loop
    Else
        cboSalesMan.Clear: cboSMName.Clear
    End If
End Sub

Sub FillDetails()
    KCNT = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            KCNT = KCNT + 1

            STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))

            grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(RSTDAYTRAN!tranqty) & Chr(9) & _
                               Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            RSTDAYTRAN.MoveNext
        Loop
        If NumericVal(txtDS1.Text) <> 0 Then
            If txtDS_Desc1.Text = "" Then
                txtDS_Desc1.Text = "DISCOUNT"
            End If
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
        Else
            '================================
            'updating code:    JAA - 02022008
            '            txtDS_Desc1.Text = ""
            '            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            '            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = 0
            txtTTLInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
            txtNetInvAmt.Text = ToDoubleNumber(ORD_TOTINVAMT)
            '================================
        End If
        ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
        If KCNT <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cboChargeTo.Enabled = True
        cleargrid grdDetails
    End If
End Sub

Sub FillGrid()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select top 20 Tranno,tranno x from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' order by Tranno desc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select top 20 rono,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and rono is not null order by tranno desc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid3()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select Custname,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' order by CUSTNAME asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Function FillSalesMan(XXX As String) As String
    Set RSSALESMAN = New ADODB.Recordset
    RSSALESMAN.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSALESMAN.EOF And Not RSSALESMAN.BOF Then
        FillSalesMan = Null2String(RSSALESMAN!signname)
        cboSalesMan.Text = Null2String(RSSALESMAN!empno)
    Else
        cboSalesMan.Text = ""
    End If
End Function

Sub FillSearchCusTomer(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select top 20 custname, tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and CUSTNAME  like '" & XXX & "%' order by CUSTNAME")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select top 20 tranno, tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and tranno like '" & XXX & "%'")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select top 20 Rono, tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & COUNTERTYPE & "' and rono like '" & XXX & "%' order by tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemvars
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "GENERATE NON INVOICE NUMBER", "DATA ENTRY") = False Then Exit Sub

    txtPRtranno.Visible = True
    txtPRtranno.SetFocus
    txtPRtranno.Locked = True
    Dim sqltxt                                         As String
    Dim RSTMP                                          As New ADODB.Recordset
    Dim ISSCOUNTER                                     As Integer

    On Error GoTo Errorcode
    If txtTranType = "CSH" Then
        sqltxt = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'CSH')  AND LEFT(TRANNO,1) = 'P'"
        sqltxt = sqltxt & "AND [TYPE] = 'P'"

    ElseIf txtTranType = "RIV" Then
        sqltxt = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'RIV')  AND LEFT(TRANNO,1) = 'P'"
        sqltxt = sqltxt & "AND [TYPE] = 'P'"

    ElseIf txtTranType = "CHG" Then
        sqltxt = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'CHG')  AND LEFT(TRANNO,1) = 'P'"
        sqltxt = sqltxt & "AND [TYPE] = 'P'"

    ElseIf txtTranType = "DR" Then

        sqltxt = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'DR')  AND LEFT(TRANNO,1) = 'P'"
        sqltxt = sqltxt & "AND [TYPE] = 'P'"

    ElseIf txtTranType = "ADB" Then
        sqltxt = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'ADB')  AND LEFT(TRANNO,1) = 'P'"
        sqltxt = sqltxt & "AND [TYPE] = 'P'"
    End If

    Set RSTMP = gconDMIS.Execute(sqltxt)
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        ISSCOUNTER = NumericVal(RSTMP!BILANG)
    End If

    ISSCOUNTER = ISSCOUNTER + 1
    txtPRtranno.Text = "P" & Format(ISSCOUNTER, "00000")

    Set RSTMP = Nothing
Errorcode:
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim fild                                           As String
    Dim PCURONHAND                                     As Integer
    Dim PCurTISSQTY                                    As Integer
    Dim PCURISSUANCES                                  As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Dim rsPartmasDup                                   As ADODB.Recordset
    Dim INVNUMBER                                      As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If picDetails.Visible = False Then Exit Sub

            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Parts Issuance)"
            '====================================================================
            If COUNTERTYPE = "CSH" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS ISSUANCE COUNTER CASH")
            ElseIf COUNTERTYPE = "CHG" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS ISSUANCE COUNTER CHARGE")
            ElseIf COUNTERTYPE = "DR" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS DR OUT ISSUANCE")
            ElseIf COUNTERTYPE = "RIV" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS ISSUANCE SERVICE ISSUANCE")
            Else
                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS ADVANCE BILL DATA ENTRY")
            End If
            '====================================================================

        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemvars
            End If
            txtPRtranno.Visible = False
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsOrd_Hd!STATUS) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!STATUS) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!STATUS) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                    picDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!STATUS) <> "P" And Null2String(rsOrd_Hd!STATUS) <> "C" And Null2String(rsOrd_Hd!STATUS) <> "B" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If
        Case vbKeyF5
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!STATUS) <> "P" And Null2String(rsOrd_Hd!STATUS) <> "C" And Null2String(rsOrd_Hd!STATUS) <> "B" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True Then
                cmdPost_Click
            End If
        Case vbKeyF11
            If Picture1.Visible = True And (labSJ = "" And labORNo = "") And Null2String(rsOrd_Hd!STATUS) = "C" Then
                If MsgBox("Are you sure you want to uncancel this transaction", vbInformation + vbYesNo) = vbYes Then
                    gconDMIS.Execute ("UPDATE PMIS_ORD_HD SET STATUS=NULL WHERE ID=" & labID)
                    rsRefresh
                    rsOrd_Hd.Find ("id=" & labID)
                    StoreMemvars
                End If
            End If
        Case vbKeyF12

            If Picture1.Visible = True And (labSJ = "" And labORNo = "") Then
                If Null2String(rsOrd_Hd!STATUS) = "B" Then
                    MsgBox "RO # " & txtRONO & " is already been invoiced. " & vbCrLf & "To Unpost this Transaction Service Invoice Should Be Cancelled First.", vbCritical
                    rsRefresh
                    rsOrd_Hd.Find ("ID=" & labID)
                    StoreMemvars
                    Exit Sub
                ElseIf Null2String(rsOrd_Hd!STATUS) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPost", LOCALACESS) = False Then Exit Sub

                    If txtTranType = "RIV" Or txtTranType = "ADB" Then
                        INVNUMBER = CheckIfROBilled(txtRONO)
                        If LTrim(RTrim(INVNUMBER)) <> "" Then
                            MsgBox "RO # " & txtRONO & " is already been invoiced. Service Invoice # " & INVNUMBER & vbCrLf & "Cannot Unpost Current Transaction." & vbCrLf & "To Unpost this Transaction Service Invoice Should Be Cancelled.", vbCritical
                            rsRefresh
                            rsOrd_Hd.Find ("ID=" & labID)
                            StoreMemvars
                            Exit Sub
                        End If
                        If LAB_ADB = "" Then
                            If MsgBox("Unposting of this transaction will remove issuance of parts in CarService" + vbCrLf & "Are you sure you want to Unpost this Transaction?", vbCritical + vbYesNo) = vbNo Then
                                Exit Sub
                            End If
                        End If
                    End If
                    Set rsTdaytranDup = New ADODB.Recordset
                    rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = " & N2Str2Null(rsOrd_Hd!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
                    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                        rsTdaytranDup.MoveFirst
                        Do While Not rsTdaytranDup.EOF
                            Set rsPartmasDup = New ADODB.Recordset
                            rsPartmasDup.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                            If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                                If COUNTERTYPE <> "ADB" Then
                                    PCURONHAND = N2Str2IntZero(rsPartmasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!tranqty)
                                    PCurTISSQTY = N2Str2IntZero(rsPartmasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!tranqty)
                                    PCURISSUANCES = N2Str2IntZero(rsPartmasDup!ISSUANCES) - N2Str2Zero(rsTdaytranDup!tranqty)
                                    If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                                        SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                                      " REQSERVED = " & N2Str2IntZero(rsPartmasDup!REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                                      " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                        'NEW LOG AUDIT-----------------------------------------
                                        Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "TRAN NO : " & Null2String(rsOrd_Hd!TranType) & " - UNPOST", COUNTERTYPE, "")
                                        'NEW LOG AUDIT-----------------------------------------
                                    Else
                                        SQL_STATEMENT = "update PMIS_STOCKMAS set" & _
                                                      " S_REQSERVED = " & N2Str2IntZero(rsPartmasDup!S_REQServed) - N2Str2Zero(rsTdaytranDup!tranqty) & _
                                                      " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                        'NEW LOG AUDIT-----------------------------------------
                                        Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "TRAN NO : " & Null2String(rsOrd_Hd!TranType & " - UNPOST"), COUNTERTYPE, "")
                                        'NEW LOG AUDIT-----------------------------------------
                                    End If


                                    SQL_STATEMENT = "UPDATE PMIS_STOCKMAS SET" & _
                                                  " ONHAND = " & PCURONHAND & "," & _
                                                  " TISSQTY = " & PCurTISSQTY & "," & _
                                                  " ISSUANCES = " & PCURISSUANCES & "," & _
                                                  " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                                  " LASTUPDATE = '" & LOGDATE & "'" & _
                                                  " WHERE STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    gconDMIS.Execute SQL_STATEMENT

                                    'NEW LOG AUDIT-----------------------------------------
                                    Call NEW_LogAudit("E", "PARTS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "TRAN NO : " & Null2String(rsOrd_Hd!TranType) & " - UNPOST", COUNTERTYPE, "")
                                    'NEW LOG AUDIT-----------------------------------------
                                End If

                                SQL_STATEMENT = "UPDATE PMIS_TDAYTRAN SET" & _
                                              " STATUS = 'N'," & _
                                              " USERCODE = " & N2Str2Null(LOGCODE) & "," & _
                                              " LASTUPDATE = '" & LOGDATE & "'" & _
                                              " WHERE ID = " & rsTdaytranDup!ID
                                gconDMIS.Execute SQL_STATEMENT
                                NEW_LogAudit "UU", LOCALACESS, SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                            End If
                            rsTdaytranDup.MoveNext
                        Loop
                    End If
                    SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                                  " status = 'N'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & labID.Caption
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "U", LOCALACESS, SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""
                    rsRefresh
                    rsOrd_Hd.Find "id =" & labID.Caption
                    StoreMemvars
                    If txtTranType = "RIV" Or txtTranType = "ADB" Then
                        ImportParts txtRONO
                    End If
                End If
                Set rsTdaytranDup = Nothing
                Set rsPartmasDup = Nothing
            End If

        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    LOCALACESS = ""

    If COUNTERTYPE = "RIV" Then
        LOCALACESS = "PARTS ISSUANCE SERVICE ISSUANCE"
    ElseIf COUNTERTYPE = "ADB" Then
        LOCALACESS = "PARTS ADVANCE BILL DATA ENTRY"
    ElseIf COUNTERTYPE = "CSH" Then
        LOCALACESS = "PARTS ISSUANCE COUNTER CASH"
    ElseIf COUNTERTYPE = "CHG" Then
        LOCALACESS = "PARTS ISSUANCE COUNTER CHARGE"
    ElseIf COUNTERTYPE = "DR" Then
        LOCALACESS = "PARTS DR OUT ISSUANCE"
    End If


    CenterMe frmMain, Me, 1: PMIS_ORDER_SHOW = True
    textSearch.Text = "":                             'Picture5.ZOrder 0
    If COUNTERTYPE <> "RIV" And COUNTERTYPE <> "ADB" Then
        cmdSelectCustomer.Visible = True
        optRONo.Enabled = False
    Else
        cmdSelectCustomer.Visible = False
    End If

    If COUNTERTYPE = "CSH" Then optCASH.Value = True
    If COUNTERTYPE = "CHG" Then optCHARGE.Value = True
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False



    InitMemVars
    If LOGLEVEL = "ADM" Then
    Else
        If COUNTERTYPE = "ADB" Then
        Else
            txtTranUPrice.Enabled = False
        End If
    End If
    rsRefresh
    On Error Resume Next
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then rsOrd_Hd.MoveLast
    StoreMemvars
    Screen.MousePointer = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISTrans_CustomerOrder = Nothing
    COUNTERTYPE = ""
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                           As String
    If Null2String(rsOrd_Hd!STATUS) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"

        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"

    ElseIf Null2String(rsOrd_Hd!STATUS) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"

        MsgBox "Transactions are Already Billed-Out" & vbCrLf & _
               "and Cannot be Changed", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(rsOrd_Hd!STATUS) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBox "Transactions are Already Posted" & vbCrLf & _
               "and Cannot Be Changed!", vbInformation, "Edit Not Allowed!"
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            ADDOREDIT = "EDIT"
            cmdTranDelete.Enabled = True
            BringToFront
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Parts!"
            Exit Sub
        End If
    End If
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
End Sub

Sub InitCbo()
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select id,STOCKNO,STOCKDESC from PMIS_PARTMAS where ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        cboTranPartNo.Clear
        Do While Not RSPARTMAS.EOF
            cboTranPartNo.AddItem Null2String(RSPARTMAS!STOCKNO)
            RSPARTMAS.MoveNext
        Loop
    End If
    FillCboSalesMan
End Sub

Sub InitCboChargeToCounter()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "COMPANY"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.AddItem "PARTS CLAIM"
    cboChargeTo.Text = "VARIOUS"
End Sub

Sub InitCboChargeToWarehouse()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "COMPANY"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.AddItem "PARTS CLAIM"
    cboChargeTo.Text = "MECHANICAL"
End Sub

Sub InitGrid()
    With grdDetails
        .Rows = 7
        .ColWidth(0) = 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColAlignment(2) = 2
        .ColWidth(3) = 2200
        .ColWidth(4) = 1000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1300
        .Row = 0
        .Col = 1
        .Text = "Item"
        .Col = 2
        .Text = "Part Number"
        .Col = 3
        .Text = "Description"
        .Col = 4
        .Text = "QTY"
        .Col = 5
        .Text = "Price"
        .Col = 6
        .Text = "Extend Price"
    End With
End Sub

Sub InitMemVars()
    labSJ = "": labORNo = "": labinvNo = "": labDetails = ""
    If COUNTERTYPE = "RIV" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'RIV'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = True
        txtTerms.Enabled = False
    End If
    If COUNTERTYPE = "CSH" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'CSH'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = False
    End If
    If COUNTERTYPE = "CHG" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'CHG'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = True
    End If
    If COUNTERTYPE = "DR" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'DR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = False
        txtTerms.Enabled = True
    End If
    If COUNTERTYPE = "ADB" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'P' AND modul = 'ADB'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtTranNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = True
        txtTerms.Enabled = False
    End If
    txtTranDate.Text = LOGDATE
    txtCustCode.Text = ""
    txtCustName.Text = ""
    txtChargeTo.Text = "VAR"
    txtReferencePIS.Text = ""
    cboRefPRSNo.Clear
    txtRONO.Text = ""
    txtTerms.Text = ""
    txtTTLInvAmt.Text = "0.00"
    txtDS1.Text = "0"
    txtDS_Desc1.Text = "0.00"
    txtDS_Amt1.Text = "0.00"
    txtNetInvAmt.Text = "0.00"
    txtRemarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    InitCbo
    InitGrid
    txtTranDate.Enabled = False
    cleargrid grdDetails
    SendToBack
    InitSignatories
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(KCNT + 1, "0000")
    cboTranPartNo.Text = ""
    txtTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUCost.Text = "0.00"
    txtTranUPrice.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
    'If COUNTERTYPE = "ADB" Then
    '    labTranUCost.Visible = False: txtTranUCost.Visible = False
    'Else
    labTranUCost.Visible = False: txtTranUCost.Visible = False
    'End If
    Check1.Enabled = False
End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
End Sub
Function CheckIfitsAdvanceBilling(RoNo As String, StockType As String) As Boolean
    Dim rsAdvanceBill                                  As ADODB.Recordset
    Set rsAdvanceBill = New ADODB.Recordset
    rsAdvanceBill.Open "select COUNT(*) from PMIS_Ord_Hd inner join PMIS_TDAYTRAN on PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO " & _
                     " and PMIS_ORD_HD.TRANTYPE = PMIS_TDAYTRAN.TRANTYPE where PMIS_ORD_HD.[TYPE] = '" & STOCK_TYPE & "' AND PMIS_ORD_HD.trantype = 'ADB' and PMIS_ord_hd.rono = '" & RoNo & "' and pmis_tdaytran.[type] = 'P' ", gconDMIS
    If Not rsAdvanceBill.EOF And Not rsAdvanceBill.BOF Then
        If MsgQuestionBox("Advance Bill for Repair Order: " & txtRONO.Text & " is Available " & vbCrLf & _
                          "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then
            CheckIfitsAdvanceBilling = True
        End If

    End If
    Set rsAdvanceBill = Nothing
End Function

Sub InsertAdvanceBill()
    Dim ORDTRANDATE, ORDTRANNO, ORDTRANTYPE            As String
    Dim ORDITEMNO, ORDSTOCK_ORD, ORDSTOCK_SUP          As String
    Dim ORDTRANQTY                                     As Integer
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double

    Dim CurONHAND, CurSAFESTOCK, CurTISSQTY            As Integer
    Dim curRESSERVICE, curIssuances                    As Integer

    If txtTranType.Text = "RIV" Then
        Dim rsAdvanceBill                              As ADODB.Recordset
        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HD.rono,PMIS_ORD_HD.trandate,PMIS_ORD_HD.trantype,PMIS_ORD_HD.tranno,PMIS_TDAYTRAN.trantype,PMIS_TDAYTRAN.tranno,PMIS_TDAYTRAN.itemno,PMIS_TDAYTRAN.STOCK_ORD,PMIS_TDAYTRAN.tranqty,PMIS_TDAYTRAN.tranuprice from PMIS_Ord_Hd inner join PMIS_TDAYTRAN on PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO and PMIS_ORD_HD.TRANTYPE = PMIS_TDAYTRAN.TRANTYPE where PMIS_ORD_HD.[TYPE] = 'P' AND PMIS_ORD_HD.trantype = 'ADB' and PMIS_ord_hd.rono = '" & txtRONO.Text & "' and pmis_tdaytran.[type] = 'P' ", gconDMIS
        If Not rsAdvanceBill.EOF And Not rsAdvanceBill.BOF Then
            If MsgQuestionBox("Advance Bill for Repair Order: " & txtRONO.Text & " is Available " & vbCrLf & _
                              "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then


                rsAdvanceBill.MoveFirst
                Do While Not rsAdvanceBill.EOF
                    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
                    ORDTRANTYPE = "'RIV'"
                    ORDTRANNO = "'" & txtTranNo.Text & "'"
                    ORDITEMNO = N2Str2Null(Format(rsAdvanceBill!itemno, "0000"))
                    ORDSTOCK_ORD = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDSTOCK_SUP = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!tranqty)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set RSPARTMAS = New ADODB.Recordset
                    RSPARTMAS.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                        CurONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
                        CurSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
                        CurTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
                        curRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
                        curIssuances = N2Str2IntZero(RSPARTMAS!ISSUANCES)

                        If CurONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                        End If

                        If ORDTRANQTY > CurONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(RSPARTMAS!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CurONHAND = CurONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If





                    SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                    "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                  " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                  " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                  " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                  " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "AA", "PARTS ADVANCE BILL DATA ENTRY", SQL_STATEMENT, labID, "Parts", "", txtTranType, ""

                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    SQL_STATEMENT = "update PMIS_Ord_Hd set " & _
                                  " status2='R',  ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                  " netinvamt = " & ORD_TOTINVAMT & _
                                  " where id = " & labID.Caption

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "E", "PARTS ADVANCE BILL DATA ENTRY", SQL_STATEMENT, labID, "Parts", "", txtTranType, ""
                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If



        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HIST.rono,PMIS_ORD_HIST.trandate,PMIS_ORD_HIST.trantype,PMIS_ORD_HIST.tranno,PMIS_DAYTRAN.trantype,PMIS_DAYTRAN.tranno,PMIS_DAYTRAN.itemno,PMIS_DAYTRAN.STOCK_ORD,PMIS_DAYTRAN.tranqty,PMIS_DAYTRAN.tranuprice from PMIS_Ord_Hist inner join PMIS_DAYTRAN on PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO and PMIS_ORD_HIST.TRANTYPE = PMIS_DAYTRAN.TRANTYPE where PMIS_ORD_HIST.[TYPE] = 'P' AND PMIS_ORD_HIST.trantype = 'ADB' and PMIS_ord_hIST.rono = '" & txtRONO.Text & "'", gconDMIS
        If Not rsAdvanceBill.EOF And Not rsAdvanceBill.BOF Then
            If MsgQuestionBox("Advance Bill for Repair Order: " & txtRONO.Text & " is Available " & vbCrLf & _
                              "Would you like to Insert this Transaction?", "Available Advance Bill") = True Then

                rsAdvanceBill.MoveFirst
                Do While Not rsAdvanceBill.EOF
                    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
                    ORDTRANTYPE = "'RIV'"
                    ORDTRANNO = "'" & txtTranNo.Text & "'"
                    ORDITEMNO = N2Str2Null(Format(rsAdvanceBill!itemno, "0000"))
                    ORDSTOCK_ORD = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDSTOCK_SUP = N2Str2Null(rsAdvanceBill!STOCK_ORD)
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!tranqty)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set RSPARTMAS = New ADODB.Recordset
                    RSPARTMAS.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_PARTMAS where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                        CurONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
                        CurSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
                        CurTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
                        curRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
                        curIssuances = N2Str2IntZero(RSPARTMAS!ISSUANCES)

                        If CurONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                            If MsgQuestionBox("Warning: Error Has been encountered... Continue Anyway?", "Error Encountered") = False Then
                                Exit Sub
                            End If
                        End If
                        If ORDTRANQTY > CurONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(rsAdvanceBill!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CurONHAND = CurONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If

                    SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                    "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                  " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                  " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                  " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                  " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "A", "PARTS ADVANCE BILL DATA ENTRY", SQL_STATEMENT, labID, "Parts", "", txtTranType, ""

                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                                  " status2='R', ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                  " netinvamt = " & ORD_TOTINVAMT & _
                                  " where id = " & labID.Caption

                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "E", "PARTS ADVANCE BILL DATA ENTRY", SQL_STATEMENT, labID, "Parts", txtTranNo, txtTranType, ""
                    '
                    '                    SQL_STATEMENT = "update PMIS_PARTMAS set" & _
                                         '                                  " onhand = " & CurONHAND & "," & _
                                         '                                  " TISSQTY = " & CurTISSQTY + ORDTRANQTY & ", " & _
                                         '                                  " issuances = " & curIssuances + ORDTRANQTY & _
                                         '                                  " where STOCKNO = " & ORDSTOCK_SUP
                    '
                    '                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "EE", "PARTS ADVANCE BILL DATA ENTRY", SQL_STATEMENT, labID, "Parts", "", txtTranType, ""

                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Private Sub lstOrd_Hd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOrd_Hd
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstOrd_Hd_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstOrd_Hd_GotFocus()
    On Error Resume Next
    lstOrd_Hd_ItemClick lstOrd_Hd.SelectedItem
End Sub

Private Sub lstOrd_Hd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    If optTranno.Value = True Then
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    Else
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemvars
End Sub

Private Sub lstOrd_Hd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Sub NEWDRPRINTING()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\DR.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' and tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = 'DR' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0

        '===========================
        'updating code:     JAA - 02052008   - To trace the number of copy to be printed
        If COMPANY_CODE = "HAI" Then
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 2
        Else
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 1
        End If
        '===========================


        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS DELIVERY RECEIPT</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsOrd_Hd!TranType) & "-" & Null2String(rsOrd_Hd!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"

            If COMPANY_CODE = "HBK" Then
                Print #1, "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
            End If

            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL DR</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\DR.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\DR.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub optCASH_Click()
    COUNTERTYPE = "CSH"
End Sub

Private Sub optCHARGE_Click()
    COUNTERTYPE = "CHG"
End Sub

Private Sub Option1_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "CUSTOMER NAME"
    If textSearch = "" Then FillGrid3 Else FillSearchCusTomer (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optRONo_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "RO Number"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optTranno_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "Tran. No."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Sub PISPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "PIS.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "PISDisc.RPT", "{ord_hd.TYPE} = 'P' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub rsRefresh()

    If LOGLEVEL = "RIV USER" Then
        If COUNTERTYPE = "ADB" Then
            Me.Caption = "ADVANCE BILL DATA ENTRY"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'ADB' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "RIV" Then
            Me.Caption = "Parts Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where ISNULL(STATUS2,'')<>'R' AND  [TYPE] = 'P' AND trantype = 'RIV' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        InitCboChargeToWarehouse
    Else

        If COUNTERTYPE = "CSH" Then
            Me.Caption = "Parts Issuance Slip Data Entry (Over the Counter)"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'CSH' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "CHG" Then
            Me.Caption = "Charge Counter Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'CHG' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "RIV" Then
            Me.Caption = "Parts Issuance Slip Data Entry (Service Requisition)"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where ISNULL(STATUS2,'')<>'R' AND   [TYPE] = 'P' AND trantype = 'RIV' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "DR" Then
            Me.Caption = "DR Out Issuance Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'DR' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        If COUNTERTYPE = "ADB" Then
            Me.Caption = "Advance Bill Data Entry"
            Set rsOrd_Hd = New ADODB.Recordset
            rsOrd_Hd.Open "select * from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = 'ADB' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        InitCboChargeToCounter
    End If
End Sub

Sub SendToBack()
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False
    Picture1.Enabled = True
    fraDetails.Enabled = True
    picDetails.Enabled = True
End Sub

Sub SERVICEPISPRINTING()

    Screen.MousePointer = 11
    If NumericVal(txtDS1.Text) = 0 Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

        If rsOrd_Hd!TranType = "RIV" Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
    Else
        If rsOrd_Hd!TranType = "RIV" Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
    End If


    Screen.MousePointer = 0

    ' update code : JBF 01/29/09
    '    Screen.MousePointer = 11
    '    If NumericVal(txtDS1.Text) = 0 Then
    '            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    '            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '            If COMPANY_CODE = "HCI" Then
    '                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            Else
    '                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            End If
    '      Else
    '            If COMPANY_CODE = "HCI" Then
    '                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            Else
    '                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Parts.rpt", "{ord_hd.TYPE} = 'P' and {ord_hd.TRANTYPE} = 'ADB' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '            End If
    '     End If
    '     Screen.MousePointer = 0
End Sub

Sub SERVICEPISPRINTING_BLANKFORM()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS

    Open App.Path & "\PIS.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'P' AND tranno = " & N2Str2Null(rsOrd_Hd!TRANNO) & " and trantype = 'RIV' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0

        If COMPANY_CODE = "HAI" Then
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 2
        Else
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then cntCOPY = 4 Else cntCOPY = 1
        End If


        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>PARTS ISSUANCE SLIP</strong></font></td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "<td align=center width=60%>&nbsp;</td>"
            Print #1, "<td align=left width=20%>&nbsp;</td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"

            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Repair Order Number:&nbsp;</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsOrd_Hd!RoNo) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%>&nbsp;</td>"
            Print #1, "</tr>"

            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "SERVICE PIS-" & Null2String(rsOrd_Hd!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsOrd_Hd!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsOrd_Hd!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsOrd_Hd!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>PART NUMBER</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "</tr>"
            Next
            Print #1, "</table>"
            '==================================
            'updating code:     JAA  - 02092008
            If COMPANY_CODE = "HBK" Then
                Print #1, "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
            End If
            '==================================
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            If cntCOPY = 4 And knt < 3 Then
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL PIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=8%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtPreparedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtIssuedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtApprovedBy.Text & "</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & txtRequestedBy.Text & "</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>---------------------------------</FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Received By</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            If knt <> 2 And knt <> 4 Then
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
                'Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
                Print #1, "<table>"
                Print #1, "<tr>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "<td width=25%><FONT SIZE=3>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
        Next
        Print #1, "</body></html>"
        Close #1
        On Error Resume Next
        Open App.Path & "\PIS.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"

            browRIV.Refresh
            browRIV.Navigate App.Path & "\PIS.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub SetCustInfo(rep As String)
    rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set RSREPOR = New ADODB.Recordset
    RSREPOR.Open "select rep_or,niym,acct_no,invoice,plate_no,dte_rel from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS


    If Not RSREPOR.EOF And Not RSREPOR.BOF Then
        '==========================================
        'updating code:     JAA - 02082008
        '        If Null2String(RSREPOR!dte_rel) <> "" Then
        '            'If Null2String(rsREPOR!invoice) <> "" Then
        '            '==========================================
        '            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                     '                 " Parts Issuance for this Repair Order must have a Reference Advanced Bill!", vbCritical, "Critical Issue!"
        '            If MsgBox("Would You Like to Continue?", vbQuestion + vbYesNo, "Continue...") = vbNo Then
        '                On Error Resume Next
        '                txtRONO.SetFocus
        '                Exit Sub
        '            Else
        '                MsgBox "Pls. Input Reference Number from Remarks Field..."
        '                On Error Resume Next
        '                txtRemarks.SetFocus
        '            End If
        '
        '        End If
        '==========================================
        'updating code:     JAA - 02082008
        If Null2String(RSREPOR!Invoice) <> "" Then
            REPOR_STATUS = "Billed-Out"
        End If
        '==========================================
        txtCustName.Text = Null2String(RSREPOR!niym)
        txtCustCode.Text = Null2String(RSREPOR!ACCT_NO)

        Dim RSCUSTINFO                                 As ADODB.Recordset
        If Null2String(RSREPOR!plate_no) <> "" Then
            Set RSCUSTINFO = New ADODB.Recordset
            'IT MIGHT GIVE A WRONG INFO OF THE CUSTOMER
            Set RSCUSTINFO = gconDMIS.Execute("select * from CSMS_CUSVEH where Plate_NO=" & N2Str2Null(RSREPOR!plate_no) & "  And CUSCDE='" & Null2String(RSREPOR!ACCT_NO) & "'")
            If Not RSCUSTINFO.EOF Or Not RSCUSTINFO.BOF Then
                txtRemarks = "MODEL: " & Null2String(RSCUSTINFO("model")) & vbCrLf & "ENGINE#:" & Null2String(RSCUSTINFO("SERIAL")) & vbCrLf & "VIN#:" & Null2String(RSCUSTINFO("vin")) & vbCrLf & "PLATE#:" & Null2String(RSCUSTINFO("plate_no"))
            End If
        End If
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""

    End If
End Sub

Sub SetCustomer()
    Dim RSCUSTOMER                                     As ADODB.Recordset
    Set RSCUSTOMER = New ADODB.Recordset
    Set RSCUSTOMER = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not RSCUSTOMER.EOF And Not RSCUSTOMER.BOF Then
        txtCustName.Text = Null2String(RSCUSTOMER!AcctName) & vbCrLf & Null2String(RSCUSTOMER!CUSTOMERADD) & vbCrLf & Null2String(RSCUSTOMER!City)
    Else
        txtCustName = ""
    End If
End Sub

Sub SetPartDetails(XXX As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_StockMas where TYPE = 'P' and StockNo = '" & XXX & "' AND ACTIVE = 'Y'")
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        If N2Str2Zero(RSPARTMAS!ONHAND) > 0 Then chkAvailableOnStock.Value = 1 Else chkAvailableOnStock.Value = 0
        optLocalPurchase.Value = False: optImported.Value = False: optConsigned.Value = False
        optGenuine.Value = False: optNonGenuine.Value = False
        If Null2String(RSPARTMAS!PartsOrigin) = "M" Then
            optImported.Value = True
        End If
        If Null2String(RSPARTMAS!PartsOrigin) = "L" Then
            optLocalPurchase.Value = True
        End If
        If Null2String(RSPARTMAS!PartsOrigin) = "C" Then
            optConsigned.Value = True
        End If
        If Null2String(RSPARTMAS!Genuine) = "Y" Then
            optGenuine.Value = True
        Else
            optNonGenuine.Value = True
        End If
        txtModelCode.Text = Null2String(RSPARTMAS!MODELCODE)
    Else
        optLocalPurchase.Value = False
        optImported.Value = False
        optConsigned.Value = False
        optGenuine.Value = False
        optNonGenuine.Value = False
        txtModelCode.Text = ""
    End If
End Sub

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_PARTMAS where ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_PARTMAS where STOCKNO = " & N2Str2Null(DDD) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
        SetPartDetails DDD
    End If
End Function


Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT  STOCKDESC FROM PMIS_PARTMAS WHERE STOCKNO=" & N2Str2Null(ppp), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function



Function SetSTOCKNO(pid As Variant)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_PARTMAS where id = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!Mac))
            End If
        End If
    Else
        txtTranUPrice.Text = "0.00"
        txtTranUCost.Text = 0
    End If
End Function

Sub StoreMemvars()
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        labSJ = "": labORNo = "": labDetails = "": labinvNo = ""
        labID.Caption = rsOrd_Hd!ID
        txtTranType.Text = Null2String(rsOrd_Hd!TranType)
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(rsOrd_Hd!TRANNO)
        txtTranDate.Text = Null2String(rsOrd_Hd!trandate)
        txtCustCode.Text = Null2String(rsOrd_Hd!custcode)
        txtCustName.Text = Null2String(rsOrd_Hd!custname)
        txtReferencePIS.Text = Null2String(rsOrd_Hd!refpisno)
        cboRefPRSNo.Text = Null2String(rsOrd_Hd!refpRsno)

        If Mid(txtReferencePIS, 5, 1) = "W" Then
            txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If

        If Null2String(rsOrd_Hd!chargeto) = "MEC" Then
            cboChargeTo.Text = "MECHANICAL"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "COM" Then
            cboChargeTo.Text = "COMPANY"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "WAR" Then
            cboChargeTo.Text = "WARRANTY"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "TIN" Then
            cboChargeTo.Text = "TINSMITH"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "FLE" Then
            cboChargeTo.Text = "FLEET"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "VAR" Then
            cboChargeTo.Text = "VARIOUS"
        ElseIf Null2String(rsOrd_Hd!chargeto) = "PCL" Then
            cboChargeTo.Text = "PARTS CLAIM"
        Else
            cboChargeTo.Text = ""
        End If
        txtRONO.Text = Null2String(rsOrd_Hd!RoNo)
        cboSMName.Text = FillSalesMan(Null2String(rsOrd_Hd!salesman))
        txtTerms.Text = Null2String(rsOrd_Hd!TERMS)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(rsOrd_Hd!ds1)
        txtDS_Desc1.Text = Null2String(rsOrd_Hd!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ds_amt1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!netinvamt))
        txtRemarks.Text = Null2String(rsOrd_Hd!remarks)

        If COUNTERTYPE = "RIV" Or COUNTERTYPE = "ADB" Then
            If Null2String(rsOrd_Hd!STATUS2) = "R" Then
                LAB_ADB = "ISSUANCE AGAINST ADB"
            Else
                LAB_ADB = ""
            End If
            If Null2String(rsOrd_Hd!STATUS) = "P" Or Null2String(rsOrd_Hd!STATUS) = "B" Then
                labinvNo = CheckIfROBilled(Null2String(rsOrd_Hd!RoNo))
            Else
                labinvNo = ""
            End If
            If labinvNo <> "" Then
                labPosted.Caption = "BILLED OUT"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
                labORNo = CheckORNum(labinvNo, "SI", COUNTERTYPE)
                labSJ = CheckSJNum(Null2String(labinvNo), "SI")
                If labORNo = "" And labSJ = "" Then
                    labDetails = ""
                ElseIf labORNo = "" And labSJ <> "" Then
                    labDetails = "Imported Sales Journal"
                ElseIf labORNo <> "" And labSJ = "" Then
                    labDetails = "OR Issued"
                Else
                    labDetails = "OR Issued/Journal Posted"
                End If
            Else
                If Null2String(rsOrd_Hd!STATUS) = "C" Then
                    labPosted.Caption = "CANCELLED"
                    cmdEdit.Enabled = False
                    cmdCancelCO.Enabled = False
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = False
                ElseIf Null2String(rsOrd_Hd!STATUS) = "P" Then
                    labPosted.Caption = "POSTED"
                    cmdEdit.Enabled = False
                    cmdCancelCO.Enabled = False
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = True
                Else
                    labPosted.Caption = ""
                    cmdEdit.Enabled = True
                    If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = True
                    cmdPost.Enabled = True
                    cmdPrint.Enabled = False
                End If
                If Null2String(rsOrd_Hd!In_Process) = "N" Then
                    labPosted.Caption = "RELEASED"
                    cmdEdit.Enabled = False
                    cmdCancelCO.Enabled = False
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = False
                End If
            End If
        Else
            If COUNTERTYPE = "CSH" Or COUNTERTYPE = "CHG" Then
                labinvNo = Null2String(rsOrd_Hd!TRANNO)
                labORNo = CheckORNum(Null2String(rsOrd_Hd!TRANNO), "PI", COUNTERTYPE)
                labSJ = CheckSJNum(Null2String(rsOrd_Hd!TRANNO), "PI")
            End If

            If labORNo = "" And labSJ = "" Then
                labDetails = ""
            ElseIf labORNo = "" And labSJ <> "" Then
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
                labDetails = "Imported Sales Journal"
            ElseIf labORNo <> "" And labSJ = "" Then
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
                labDetails = "OR Issued"
            Else
                labDetails = "OR Issued/Journal Posted"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
            End If

            If Null2String(rsOrd_Hd!STATUS) = "C" Then
                labPosted.Caption = "CANCELLED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            ElseIf Null2String(rsOrd_Hd!STATUS) = "P" Then
                labPosted.Caption = "POSTED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
            Else
                labPosted.Caption = ""
                cmdEdit.Enabled = True
                If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = True
                cmdPost.Enabled = True
                cmdPrint.Enabled = False
            End If
            If Null2String(rsOrd_Hd!In_Process) = "N" Then
                labPosted.Caption = "RELEASED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            End If
        End If


        cleargrid grdDetails
        FillDetails
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Function StorePartsEntry(ByVal ID As Variant)
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice,tranucost from PMIS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        labDetID.Caption = RSTDAYTRAN!ID
        labPartNo.Caption = Null2String(RSTDAYTRAN!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(RSTDAYTRAN!tranqty)
        txtTranItemNo.Text = Format(Null2String(RSTDAYTRAN!itemno), "0000")
        cboTranPartNo.Text = Null2String(RSTDAYTRAN!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(RSTDAYTRAN!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(RSTDAYTRAN!tranqty)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!tranucost))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!tranqty) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUPrice.Enabled = False
    End If
    If COUNTERTYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        'labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Function

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    ElseIf optRONo.Value = True Then
        Dim RONOStr                                    As String
        RONOStr = textSearch.Text
        If Left(RONOStr, 2) = "R-" Then
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (RONOStr)
    Else
        FillSearchCusTomer (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOrd_Hd.ListItems.Count > 0 And lstOrd_Hd.Enabled = True Then: lstOrd_Hd.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then
            labPosted.Visible = False
        Else
            labPosted.Visible = True
        End If
    End If
End Sub

Private Sub txtCustName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtDS_Desc1_Change()
    If Len(txtDS_Desc1.Text) = 1 Then
        If txtDS_Desc1.Text = "D" Then
            txtDS_Desc1.Text = "DISCOUNT"
        End If
    End If
End Sub

Private Sub txtDS1_Change()
    If NumericVal(txtDS1.Text) <> 0 Then
        If txtDS_Desc1.Text = "" Then
            txtDS_Desc1.Text = "DISCOUNT"
        End If
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    Else
        txtDS_Desc1.Text = ""
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    End If
End Sub

Private Sub txtDS1_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtDS1_LostFocus()
    If NumericVal(txtDS1.Text) <> 0 Then
        txtDS_Desc1.Text = "DISCOUNT"
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    Else
        txtDS_Desc1.Text = ""
        txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
        txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
    End If
End Sub

Private Sub txtReferencePIS_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtRONO_LostFocus()
    Dim RONOStr                                        As String
    Dim RSADB                                          As ADODB.Recordset

    RONOStr = txtRONO.Text
    If Left(RONOStr, 2) = "R-" Then
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
    End If
    txtRONO.Text = RONOStr

    If txtTranType.Text = "RIV" Or txtTranType.Text = "ADB" Then

        If CheckIfRoExists(txtRONO.Text) <> "" Then
            MsgBox "RO Number Doesn't Exists. Please Correct Repair Order Number", vbCritical, "Invalid RO Number"
            On Error Resume Next
            txtRONO.SetFocus
            Exit Sub
        End If

        If CheckIfROBilled(txtRONO.Text) <> "" Then
            'MessagePop InfoVoid, "Warning:Billed Out RO", "" & vbCrLf & "Cannot issue stock item(s) for this RO", 5000
            MsgBox "This RO is Already been billed for this issuance", vbInformation, "Billed Out RO"
            Exit Sub
        End If

        Set RSADB = gconDMIS.Execute("SELECT RONO,TRANTYPE FROM PMIS_ORD_HD WHERE TRANTYPE IN('ADB','RIV')  AND RONO=" & N2Str2Null(txtRONO))

        If Not (RSADB.EOF Or RSADB.BOF) Then
            While Not RSADB.EOF
                If Null2String(RSADB!TranType) = "ADB" Then
                    If MsgBox("There is Advance Bill for this RO!!" & vbCrLf & " Are you Sure You will do service issuance(s)?", vbInformation + vbYesNo, "Advance Bill Deteched!!") = vbNo Then
                        On Error Resume Next
                        txtRONO.SetFocus
                        Exit Sub
                    End If
                End If
                RSADB.MoveNext
            Wend
            'TO DO


        End If
    End If




    SetCustInfo (RONOStr)





End Sub



Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
    'updating code:     jaa - 10292008          - Transaction Month should be equal to current month
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            '            txtTranDate.SetFocus
        End If
    End If
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranNo_LostFocus()
    If LTrim(RTrim(txtTranNo)) = "" Then
        MessagePop InfoVoid, "Blank Fields Detected!", "Please Input Valid Transaction Number." & vbCrLf, 3000
        Exit Sub
    End If


    txtTranNo.Text = Format(txtTranNo.Text, "000000")
    Dim RSFINDDUP                                      As ADODB.Recordset
    If ADDOREDIT = "ADD" Then
        Set RSFINDDUP = New ADODB.Recordset
        RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
            MsgSpeechBox "Transaction No. already exist!"
            On Error Resume Next
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(rsOrd_Hd!TRANNO))) Then
            Set RSFINDDUP = New ADODB.Recordset
            RSFINDDUP.Open "select trantype,tranno from PMIS_Ord_Hd where [TYPE] = 'P' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Transaction No. already exist!"
                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        'EAP:012709 Validation for negative and zero issuances.
        If txtTranQty.Text <= 0 Then
            MessagePop InfoVoid, "Invalid Input", "Quantity must not have a zero or negative value"

            On Error Resume Next
            txtTranQty.SetFocus
            cmdTranSave.Enabled = False
            Exit Sub
        Else
            cmdTranSave.Enabled = True
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
        End If

    End If


End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranQty_LostFocus()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = Format(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text), MAXIMUM_DIGIT)
    Else
        txtTranQty.Text = 1
        txtTranTotalAmt.Text = Format(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text), MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtTranQty_Validate(Cancel As Boolean)
    If NumericVal(txtTranQty) <= 0 Then
        MessagePop InfoVoid, "Invalid Input", "Please Input Valid Quantity"
        Cancel = True
    End If
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtTranUPrice_Change()
    If txtTranUPrice.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub txtTranUPrice_GotFocus()
    If NumericVal(txtTranUPrice.Text) = 0 Then txtTranUPrice.Text = ""
End Sub

Private Sub txtTranUPrice_KeyPress(KeyCode As Integer)
    If (KeyCode < 48 Or KeyCode > 57) And KeyCode <> 110 And KeyCode <> 46 Then
        KeyCode = 0
    End If
End Sub

Private Sub txtTranUPrice_LostFocus()
    txtTranUPrice.Text = Format(txtTranUPrice.Text, MAXIMUM_DIGIT)
End Sub


Public Sub PisValidation()
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    cmdPost.Enabled = False
    'EAP:033109 so user cannot pressd f8 when transaction is not yet saved.
    cmdPost.Enabled = False
End Sub
