VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmPMISTrans_CustomerOrder_AC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Customer Order"
   ClientHeight    =   7065
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11505
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_CustomerOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   11505
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11505
      TabIndex        =   123
      Top             =   6720
      Width           =   11505
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
         TabIndex        =   130
         Top             =   0
         Width           =   1125
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
         TabIndex        =   129
         Top             =   0
         Width           =   855
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
         Left            =   2010
         TabIndex        =   128
         Top             =   0
         Width           =   855
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
         Left            =   2880
         TabIndex        =   127
         Top             =   0
         Width           =   1095
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
         Left            =   4860
         TabIndex        =   126
         Top             =   0
         Width           =   1095
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
         Left            =   4020
         TabIndex        =   125
         Top             =   0
         Width           =   825
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
         Left            =   5970
         TabIndex        =   124
         Top             =   0
         Width           =   5445
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2760
      ScaleHeight     =   870
      ScaleWidth      =   8715
      TabIndex        =   90
      Top             =   5865
      Width           =   8715
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
         Left            =   7860
         MouseIcon       =   "AC_CustomerOrder.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":0A1C
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
         Left            =   7080
         MouseIcon       =   "AC_CustomerOrder.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":0ED4
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
         Left            =   6300
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "AC_CustomerOrder.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
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
         Left            =   5520
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "AC_CustomerOrder.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":1818
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
         Left            =   4740
         MouseIcon       =   "AC_CustomerOrder.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":1C8F
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
         Left            =   3960
         MouseIcon       =   "AC_CustomerOrder.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":213D
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
         Left            =   3180
         MouseIcon       =   "AC_CustomerOrder.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":25A2
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
         Left            =   2400
         MouseIcon       =   "AC_CustomerOrder.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":2A44
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
         Left            =   1620
         MouseIcon       =   "AC_CustomerOrder.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":2EF4
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
         Left            =   840
         MouseIcon       =   "AC_CustomerOrder.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":3340
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
         Left            =   60
         MouseIcon       =   "AC_CustomerOrder.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
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
         Caption         =   "F5 - Delete Accs."
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
         Caption         =   "F4 - Edit Accs."
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
         Caption         =   "F3 - Add Accs."
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
      Begin VB.OptionButton Option1 
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   131
         Top             =   930
         Width           =   2205
      End
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
         Left            =   75
         MaxLength       =   35
         TabIndex        =   70
         Top             =   1290
         Width           =   2445
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   69
         Top             =   645
         Width           =   2385
      End
      Begin VB.OptionButton optTranno 
         Caption         =   "Transaction No."
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
         Left            =   180
         TabIndex        =   68
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4935
         Left            =   60
         TabIndex        =   71
         Top             =   1650
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
         MouseIcon       =   "AC_CustomerOrder.frx":3B49
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9825
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   87
      Top             =   5805
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
         Left            =   810
         MouseIcon       =   "AC_CustomerOrder.frx":3CAB
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":3DFD
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Abort Transaction"
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
         Left            =   30
         MouseIcon       =   "AC_CustomerOrder.frx":413B
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":428D
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
      Left            =   2700
      ScaleHeight     =   3135
      ScaleWidth      =   8715
      TabIndex        =   28
      Top             =   90
      Width           =   8745
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
         Height          =   315
         Left            =   4440
         TabIndex        =   133
         Top             =   420
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   375
         Left            =   4440
         TabIndex        =   132
         Top             =   60
         Width           =   405
      End
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   2460
         TabIndex        =   121
         Top             =   630
         Visible         =   0   'False
         Width           =   225
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
      Begin VB.CommandButton Command2 
         Caption         =   "F1 - Assign AIS Number"
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
         Left            =   7110
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
         Left            =   5250
         TabIndex        =   1
         Text            =   "AIWGC06H360"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
         Width           =   1875
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
         Top             =   600
         Width           =   1275
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
         Picture         =   "AC_CustomerOrder.frx":45DD
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
         Locked          =   -1  'True
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
         Height          =   345
         Left            =   3420
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
      Begin VB.CommandButton Command1 
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
         Left            =   1140
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reference ARS Number :"
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
         Caption         =   "AIS No."
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
         Left            =   4530
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
         Left            =   5430
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
         Width           =   1095
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
         Index           =   0
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
         Left            =   6900
         TabIndex        =   29
         Top             =   90
         Width           =   1725
      End
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
   Begin VB.CommandButton cmdAddTran 
      Caption         =   "Command5"
      Height          =   3705
      Left            =   5520
      TabIndex        =   134
      Top             =   1110
      Width           =   4095
   End
   Begin VB.CommandButton cmdSignatories 
      Caption         =   "Command5"
      Height          =   2565
      Left            =   4830
      TabIndex        =   135
      Top             =   1740
      Width           =   4605
   End
   Begin VB.PictureBox fraSignatories 
      Height          =   2355
      Left            =   4935
      ScaleHeight     =   2295
      ScaleWidth      =   4350
      TabIndex        =   51
      Top             =   1845
      Width           =   4410
      Begin VB.CommandButton cmdPrintRIV 
         Caption         =   "&Print ARS"
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
         Left            =   3030
         MouseIcon       =   "AC_CustomerOrder.frx":7319
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":746B
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
         Left            =   1410
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   780
         Width           =   2895
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
         Left            =   1410
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1140
         Width           =   2895
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
         Left            =   1410
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   420
         Width           =   2895
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
         Left            =   1410
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   60
         Width           =   2895
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
   Begin VB.PictureBox picHPI 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   4830
      ScaleHeight     =   1575
      ScaleWidth      =   5115
      TabIndex        =   136
      Top             =   2730
      Visible         =   0   'False
      Width           =   5145
      Begin VB.CheckBox chkvat 
         Caption         =   "with vat"
         Height          =   405
         Left            =   120
         TabIndex        =   142
         Top             =   1140
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3810
         TabIndex        =   141
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox TXT_PRINT_REC_BY 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   140
         Top             =   300
         Width           =   3675
      End
      Begin VB.TextBox TXT_PRINT_ISSUEBY 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   139
         Top             =   660
         Width           =   3675
      End
      Begin VB.CommandButton CMDHPI_PIS 
         Caption         =   "Print Issue Slip"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2610
         TabIndex        =   138
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton CMDHPI_SI 
         Caption         =   "Print Sales Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1410
         TabIndex        =   137
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "FOR HPI WE ARE USING ONLY CSH.RPT AND PIS.RPT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -30
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Received By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   90
         TabIndex        =   144
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   90
         TabIndex        =   143
         Top             =   750
         Width           =   960
      End
   End
   Begin VB.PictureBox fraAddTran 
      Height          =   3495
      Left            =   5610
      ScaleHeight     =   3435
      ScaleWidth      =   3855
      TabIndex        =   43
      Top             =   1230
      Width           =   3915
      Begin VB.CommandButton Command3 
         Caption         =   "::"
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
         Left            =   3120
         TabIndex        =   122
         Top             =   1800
         Width           =   315
      End
      Begin VB.Frame Frame2 
         Caption         =   "Parts Details"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3315
         Left            =   3870
         TabIndex        =   102
         Top             =   60
         Width           =   2865
         Begin VB.Frame Frame5 
            Caption         =   "Model Codes"
            Height          =   765
            Left            =   150
            TabIndex        =   117
            Top             =   2400
            Width           =   2595
            Begin VB.TextBox txtModelCode 
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
               Height          =   255
               Left            =   150
               TabIndex        =   106
               Top             =   660
               Width           =   1845
            End
            Begin VB.OptionButton optImported 
               Caption         =   "Imported"
               Height          =   255
               Left            =   150
               TabIndex        =   105
               Top             =   390
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optLocalPurchase 
               Caption         =   "Local Purchases"
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
               Height          =   255
               Left            =   150
               TabIndex        =   109
               Top             =   180
               Value           =   -1  'True
               Width           =   1845
            End
            Begin VB.OptionButton optNonGenuine 
               Caption         =   "Non-Genuine"
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
         MouseIcon       =   "AC_CustomerOrder.frx":77D1
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":7923
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Delete Entry"
         Top             =   2580
         Width           =   735
      End
      Begin VB.CommandButton cmdTranCancel 
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
         Left            =   2160
         MouseIcon       =   "AC_CustomerOrder.frx":7C4E
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":7DA0
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Cancel Entry"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtTranDescription 
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
         Left            =   90
         MaxLength       =   50
         TabIndex        =   18
         Top             =   1050
         Width           =   3675
      End
      Begin VB.TextBox txtTranTotalAmt 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2160
         Width           =   1665
      End
      Begin VB.TextBox txtTranUPrice 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   20
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1800
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
         Top             =   1440
         Width           =   705
      End
      Begin VB.TextBox txtTranItemNo 
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
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   1365
      End
      Begin VB.ComboBox cboTranPartNo 
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
         Left            =   1470
         Sorted          =   -1  'True
         TabIndex        =   17
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   420
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
         Top             =   420
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
         MouseIcon       =   "AC_CustomerOrder.frx":80DE
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrder.frx":8230
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Save Accessories"
         Top             =   2580
         Width           =   735
      End
      Begin VB.Frame fraCostToCost 
         Height          =   405
         Left            =   2220
         TabIndex        =   119
         Top             =   1350
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
            TabIndex        =   120
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
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   114
         Text            =   "1000.00"
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1440
         Visible         =   0   'False
         Width           =   945
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
         Top             =   1470
         Visible         =   0   'False
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
         Top             =   1800
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
         Top             =   1800
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
         Left            =   1500
         TabIndex        =   56
         Top             =   1830
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
         Top             =   2190
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
         Top             =   1830
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
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories #"
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
         TabIndex        =   47
         Top             =   450
         Width           =   1965
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
         Top             =   90
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
         Top             =   810
         Width           =   1275
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "SELECT PRINTING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   7860
      TabIndex        =   146
      Top             =   4290
      Width           =   3555
      Begin VB.CommandButton Command 
         Caption         =   "RIS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   510
         TabIndex        =   148
         Top             =   270
         Width           =   1185
      End
      Begin VB.CommandButton Command6 
         Caption         =   "PSI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1770
         TabIndex        =   147
         Top             =   270
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Press ESC to Cancel  "
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   149
         Top             =   870
         Width           =   3195
      End
   End
   Begin VB.PictureBox picoverride 
      Height          =   1095
      Left            =   5400
      ScaleHeight     =   1035
      ScaleWidth      =   3555
      TabIndex        =   150
      Top             =   2520
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox txtoverride 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000040&
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   151
         Top             =   480
         Width           =   2805
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Index           =   0
         Left            =   -360
         TabIndex        =   152
         Top             =   0
         Width           =   4515
         _Version        =   655364
         _ExtentX        =   7964
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   "Enter Code to Override"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   255
         GradientColorDark=   4210752
         ForeColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmPMISTrans_CustomerOrder_AC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HD                                           As ADODB.Recordset
Dim RSTDAYTRAN                                         As ADODB.Recordset
Dim RSPARTMAS                                          As ADODB.Recordset
Dim RSSALESMAN                                         As ADODB.Recordset
Dim RSCUNTER                                           As ADODB.Recordset
Attribute RSCUNTER.VB_VarUserMemId = 1073938435
Dim RSPROFILE                                          As ADODB.Recordset
Dim RSREPOR                                            As ADODB.Recordset
Attribute RSREPOR.VB_VarUserMemId = 1073938439
Dim rsCustomer                                         As ADODB.Recordset
Dim KCNT                                               As Integer
Attribute KCNT.VB_VarUserMemId = 1073938441
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938442
Dim ORD_TOTUPRICE                                      As Double
Attribute ORD_TOTUPRICE.VB_VarUserMemId = 1073938443
Dim ORD_TOTINVAMT                                      As Double
Dim ORD_TOTVAT                                         As Double
Dim ORD_TOTQTY                                         As Double
Dim PREVORDTYPE                                        As String
Attribute PREVORDTYPE.VB_VarUserMemId = 1073938447
Dim PREVORDNO                                          As String
Dim REPOR_STATUS                                       As String
Dim LOCALACESS                                         As String
Attribute LOCALACESS.VB_VarUserMemId = 1073938435
Dim ichg                                               As Boolean
Dim ictr                                               As Integer

Function CheckIfROBilled(XXX As String) As String
    Dim rsRO_DET                                       As ADODB.Recordset
    Set rsRO_DET = New ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("Select INVOICE from CSMS_REPOR where INVOICE IS NOT NULL AND REP_OR = " & N2Str2Null(XXX))
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        CheckIfROBilled = Null2String(rsRO_DET!invoice)
    End If
    Set rsRO_DET = Nothing
End Function

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

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_Accessories where STOCKNO= '" & ppp & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If
        Else
            '            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
            '                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
            '                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            '            Else
            '                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
            '                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            '            End If
            '==[ Update EAP:072508 Waranty and Internal ]==
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If


        End If
    Else
        If COUNTERTYPE = "ADB" Then
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC = Null2String(RSPARTMAS!DESCRIPTIO)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        Else
            txtTranUPrice.Text = 0
            txtTranUCost.Text = 0
        End If
    End If
End Function

Function SetSTOCKDESC2(pid As Variant)
    If COUNTERTYPE = "ADB" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!DESCRIPTIO)
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
        Else
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Accessories where STOCKNO = " & N2Str2Null(cboTranPartNo.Text) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        End If
    Else
        If pid <> "" Then
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Accessories where id = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
                If txtTranType.Text = "DR" Then
                    If cboChargeTo.Text = "PARTS CLAIM" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * ConvertToBIRDecimalFormat(VAT_RATE))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    Else
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    End If
                Else
                    '                    If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                    '                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                    '                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    '                    Else
                    '                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                    '                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    '                    End If
                    '==[ Update:EAP:072508: Waranty and Internal ]==
                    If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    ElseIf Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    Else
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    End If

                End If
            Else
                txtTranUPrice.Text = "0.00"
                txtTranUCost.Text = 0
            End If
        End If
    End If
End Function

Function SetSTOCKNO(pid As Variant)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_Accessories where id = " & pid & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            End If
        End If
    Else
        txtTranUPrice.Text = "0.00"
        txtTranUCost.Text = 0
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_Accessories where STOCKNO = " & N2Str2Null(DDD) & " AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
        SetPartDetails DDD
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_Accessories where ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select srp,STOCKNO,mac,dnp from PMIS_Accessories where STOCKNO = '" & ppp & "' AND ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            If txtTranType.Text = "DR" Then
                If cboChargeTo.Text = "PARTS CLAIM" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * ConvertToBIRDecimalFormat(VAT_RATE))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                End If
            Else
                If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC) * 1.12)
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                End If
            End If
        End If
        SetPartDetails ppp
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice,tranucost from PMIS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        labDetID.Caption = RSTDAYTRAN!ID
        labPartNo.Caption = Null2String(RSTDAYTRAN!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
        txtTranItemNo.Text = Format(Null2String(RSTDAYTRAN!itemno), "0000")
        cboTranPartNo.Text = Null2String(RSTDAYTRAN!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(RSTDAYTRAN!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(RSTDAYTRAN!TRANQTY)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
        txtTranUPrice.Enabled = False
    End If
    If COUNTERTYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Function

Sub SERVICEPISPRINTING()

    If COMPANY_CODE = "HQA" Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    End If
    
    If COMPANY_CODE = "HLI" Then
        rptCustomerOrder.Formulas(1) = "Issuedby = '" & cboSMName & "'"
        Screen.MousePointer = 11
        rptCustomerOrder.WindowTitle = "ACCESSSOIRES SERVICE-ISSUANCE"
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0

    Else
        Screen.MousePointer = 11
        rptCustomerOrder.WindowTitle = "ACCESSSOIRES SERVICE-ISSUANCE"
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Acc.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
    'updated code: JBF 01/29/09
    'for printing for HCI
    '    If NumericVal(txtDS1.Text) = 0 Then
    '            Screen.MousePointer = 11
    '            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
    '            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    '
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Acc.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        Else
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Acc.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '    Else
    '
    '        If COMPANY_CODE = "HCI" Then
    '            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_Partsdisc.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'RIV' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
    '        End If
    '
    '    End If
    '     Screen.MousePointer = 0

End Sub

Sub CHGPRINTING()

    rptCustomerOrder.WindowTitle = "ACCESSSOIRES CHARGE-ISSUANCE"
    
    If COMPANY_CODE = "HQA" Then
         rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
         rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    End If
    If COMPANY_CODE = "HLI" Then
        rptCustomerOrder.Formulas(1) = "Issuedby = '" & cboSMName & "'"
    End If
    
    If NumericVal(txtDS1.Text) = 0 Then
         If COMPANY_CODE = "HCI" Then
            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            Screen.MousePointer = 11
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
        Screen.MousePointer = 0
    Else
        If COMPANY_CODE = "HLI" Then
            Screen.MousePointer = 11
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            
        
         ElseIf COMPANY_CODE = "HCI" Then
            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
             PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDISC_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            Screen.MousePointer = 11
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
           
        End If
        Screen.MousePointer = 0
    End If
End Sub

Sub CSHPRINTING()

    rptCustomerOrder.WindowTitle = "ACCESSSOIRES COUNTER-ISSUANCE"
    
    If COMPANY_CODE = "HQA" Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    End If
    
    If COMPANY_CODE = "HLI" Then
        rptCustomerOrder.Formulas(1) = "Issuedby = '" & cboSMName & "'"
    End If
    
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        'commented by : JBF 06/15/2009
        'PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranno.Text), DMIS_REPORT_Connection, 1
        'some parts items are  consolidated during printing which cause error
        If COMPANY_CODE = "HCI" Then
            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        End If
            Screen.MousePointer = 0
    Else
        If COMPANY_CODE = "HLI" Then
            Screen.MousePointer = 11
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
           
         ElseIf COMPANY_CODE = "HCI" Then
        
            rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDISC_internalPrintOut.RPT", "{ord_hd.TRANTYPE} = 'CSH' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            Screen.MousePointer = 11
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            
        End If
        Screen.MousePointer = 0
    End If
End Sub

Sub RIVPRINTING()
    rptCustomerOrder.WindowTitle = "ACCESSSOIRES SERVICE-ISSUANCE"
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIVDisc.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub CSHPRINTING_OTC()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\ACSH.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'CSH' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ACCESSORIES ISSUANCE SLIP (COUNTER-CSH)</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "COUNTER AIS-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
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
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
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
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL AIS</FONT></td>"
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
        Open App.Path & "\ACSH.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\ACSH.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub CHGPRINTING_OTC()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\ACHG.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'CHG' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ACCESSORIES ISSUANCE SLIP (COUNTER-CHG)</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "COUNTER AIS-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
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
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
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
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL AIS</FONT></td>"
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
        Open App.Path & "\ACHG.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\ACHG.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub SERVICEPISPRINTING_BLANKFORM()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\AIS.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'RIV' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ACCESSORIES ISSUANCE SLIP</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Repair Order Number:&nbsp;</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(RSORD_HD!RONO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%>&nbsp;</td>"
            Print #1, "</tr>"


            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "SERVICE AIS-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=5%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=20%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ACCESSORIES CODE</FONT></td>"
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
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
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

            If COMPANY_CODE = "HBK" Then
                Print #1, "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
            End If

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
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL AIS</FONT></td>"
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
        Open App.Path & "\AIS.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\AIS.HTML"
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
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "DR.RPT", "{ord_hd.type} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
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
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where TYPE = 'A' and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'DR' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ACCESSORIES DELIVERY RECEIPT</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(RSORD_HD!TranType) & "-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
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
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
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

            '==================================
            'updating code:     JAA  - 05242008
            If COMPANY_CODE = "HBK" Then
                Print #1, "<br><br><br><br><br><br><br><br><br><br><br><br><br><br><br><br>"
            End If
            '==================================

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

Sub ADBPRINTING()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile", gconDMIS
    Open PMIS_REPORT_PATH & "ADB.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = 'ADB' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(RSORD_HD!TranType) & "-" & Null2String(RSORD_HD!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HD!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HD!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Charge To: " & Null2String(RSORD_HD!chargeto) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HD!CUSTNAME) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Ref RO# : " & Null2String(RSORD_HD!RONO) & "</b></FONT></td>"
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
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
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
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL ADB</FONT></td>"
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
        Open PMIS_REPORT_PATH & "ADB.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate PMIS_REPORT_PATH & "ADB.HTML"
            DoEvents
            If chkPreview.Value = 1 Then
                browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Else
                browRIV.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
            End If
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    RSORD_HD.Bookmark = rsFind(RSORD_HD.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    If COUNTERTYPE = "CSH" Then
        Me.Caption = "Accessories Issuance Slip (CSH) Data Entry (Over the Counter)"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = 'A' AND trantype = 'CSH' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If COUNTERTYPE = "CHG" Then
        Me.Caption = "Accessories Issuance Slip (CHG) Data Entry (Over the Counter)"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = 'A' AND trantype = 'CHG' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If COUNTERTYPE = "RIV" Then
        Me.Caption = "Accessories Issuance Slip Data Entry (Service Requisition)"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = 'A' AND trantype = 'RIV' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If COUNTERTYPE = "DR" Then
        Me.Caption = "DR Out Issuance Data Entry"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = 'A' AND trantype = 'DR' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If COUNTERTYPE = "ADB" Then
        Me.Caption = "Advance Bill Data Entry"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_Ord_Hd where [TYPE] = 'A' AND trantype = 'ADB' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    InitCboChargeToCounter
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

Sub initMemvars()
    labSJ = "": labORNo = "": labinvNo = "": labDetails = ""

    If COUNTERTYPE = "RIV" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'A' AND modul = 'RIV'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'A' AND modul = 'CSH'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'A' AND modul = 'CHG'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'A' AND modul = 'DR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'A' AND modul = 'ADB'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    cleargrid grdDetails
    SendToBack
    InitSignatories
    txtTranDate.Enabled = False
End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
End Sub

Sub StoreMemVars()
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        labID.Caption = RSORD_HD!ID
        labSJ = "": labORNo = "": labDetails = "": labinvNo = ""
        txtTranType.Text = Null2String(RSORD_HD!TranType)
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(RSORD_HD!TRANNO)
        txtTranDate.Text = Null2String(RSORD_HD!trandate)
        txtCustCode.Text = Null2String(RSORD_HD!CUSTCODE)
        txtCustName.Text = Null2String(RSORD_HD!CUSTNAME)
        txtReferencePIS.Text = Null2String(RSORD_HD!REFPISNO)
        cboRefPRSNo.Text = Null2String(RSORD_HD!refpRsno)

        If Mid(txtReferencePIS, 5, 1) = "W" Then
            txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If

        If Null2String(RSORD_HD!chargeto) = "MEC" Then
            cboChargeTo.Text = "MECHANICAL"
        ElseIf Null2String(RSORD_HD!chargeto) = "COM" Then
            cboChargeTo.Text = "COMPANY"
        ElseIf Null2String(RSORD_HD!chargeto) = "WAR" Then
            cboChargeTo.Text = "WARRANTY"
        ElseIf Null2String(RSORD_HD!chargeto) = "TIN" Then
            cboChargeTo.Text = "TINSMITH"
        ElseIf Null2String(RSORD_HD!chargeto) = "FLE" Then
            cboChargeTo.Text = "FLEET"
        ElseIf Null2String(RSORD_HD!chargeto) = "VAR" Then
            cboChargeTo.Text = "VARIOUS"
        ElseIf Null2String(RSORD_HD!chargeto) = "PCL" Then
            cboChargeTo.Text = "PARTS CLAIM"
        Else
            cboChargeTo.Text = ""
        End If
        txtRONO.Text = Null2String(RSORD_HD!RONO)
        cboSMName.Text = FillSalesMan(Null2String(RSORD_HD!salesman))
        txtTerms.Text = Null2String(RSORD_HD!TERMS)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(RSORD_HD!ds1)
        txtDS_Desc1.Text = Null2String(RSORD_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!DS_AMT1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HD!NETINVAMT))
        txtRemarks.Text = Null2String(RSORD_HD!REMARKS)


        If COUNTERTYPE = "RIV" Or COUNTERTYPE = "ADB" Then
            If Null2String(RSORD_HD!Status) = "P" Or Null2String(RSORD_HD!Status) = "B" Then
                labinvNo = CheckIfROBilled(Null2String(RSORD_HD!RONO))
            Else
                labinvNo = ""
            End If
            If labinvNo <> "" Then
                labPosted.Caption = "BILLED OUT"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
                labORNo = CheckORNum(labinvNo, "AI", COUNTERTYPE)
                labSJ = CheckSJNum(Null2String(labinvNo), "AI")
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
                If Null2String(RSORD_HD!Status) = "C" Then
                    labPosted.Caption = "CANCELLED"
                    cmdEdit.Enabled = False
                    cmdCancelCO.Enabled = False
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = False
                ElseIf Null2String(RSORD_HD!Status) = "P" Then
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
                If Null2String(RSORD_HD!In_Process) = "N" Then
                    labPosted.Caption = "RELEASED"
                    cmdEdit.Enabled = False
                    cmdCancelCO.Enabled = False
                    cmdPost.Enabled = False
                    cmdPrint.Enabled = False
                End If
            End If
        Else
            If COUNTERTYPE = "CSH" Or COUNTERTYPE = "CHG" Then
                labinvNo = Null2String(RSORD_HD!TRANNO)
                labORNo = CheckORNum(Null2String(RSORD_HD!TRANNO), "AI", COUNTERTYPE)
                labSJ = CheckSJNum(Null2String(RSORD_HD!TRANNO), "AI")
            End If

            'labinvNo = Null2String(rsOrd_Hd!TRANNO)
            'labORNo = CheckORNum(Null2String(rsOrd_Hd!TRANNO), "AI")
            'labSJ = CheckSJNum(Null2String(rsOrd_Hd!TRANNO), "AI")

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

            If Null2String(RSORD_HD!Status) = "C" Then
                labPosted.Caption = "CANCELLED"
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            ElseIf Null2String(RSORD_HD!Status) = "P" Then
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
            If Null2String(RSORD_HD!In_Process) = "N" Then
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
        .Text = "Accessories No."
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

Sub FillDetails()
    On Error Resume Next
    KCNT = 0
    ORD_TOTUPRICE = 0
    ORD_TOTINVAMT = 0
    ORD_TOTVAT = 0
    ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        RSTDAYTRAN.MoveFirst
        Do While Not RSTDAYTRAN.EOF
            KCNT = KCNT + 1
            If txtTranType.Text = "ADB" Then
                STOCKDESCription = Null2String(RSTDAYTRAN!STOCK_SUP)
            Else
                STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP))
            End If
            grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(RSTDAYTRAN!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
            RSTDAYTRAN.MoveNext
        Loop
        If NumericVal(txtDS1.Text) <> 0 Then
            If txtDS_Desc1.Text = "" Then
                txtDS_Desc1.Text = "DISCOUNT"
            End If
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
        Else
            '==========================================
            'UPDATING CODE:        JAA - 02022008
            '            txtDS_Desc1.Text = ""
            '            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            '            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = 0
            txtTTLInvAmt.Text = ToDoubleNumber(ORD_TOTUPRICE)
            txtNetInvAmt.Text = ToDoubleNumber(ORD_TOTINVAMT)
            '==========================================
        End If
        ORD_TOTINVAMT = ORD_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
        If KCNT <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cboChargeTo.Enabled = True
        cleargrid grdDetails
    End If
End Sub

Sub InitCbo()
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select id,STOCKNO,STOCKDESC from PMIS_Accessories where ACTIVE = 'Y'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Sub SetCustInfo(rep As String)
    rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set RSREPOR = New ADODB.Recordset
    RSREPOR.Open "select rep_or,niym,acct_no,invoice,plate_no,dte_rel from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS
    If Not RSREPOR.EOF And Not RSREPOR.BOF Then
        '==========================================
        'updating code:     JAA - 02082008
        'If Null2String(rsREPOR!invoice) <> "" Then
        If Null2String(RSREPOR!dte_rel) <> "" Then
            '==========================================
            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                 " Accessories Issuance for this Repair Order must have a Reference Advanced Bill!", vbCritical, "Critical Issue!"
            If MsgBox("Would You Like to Continue?", vbQuestion + vbYesNo, "Continue...") = vbNo Then
                On Error Resume Next
                txtRONO.SetFocus
                Exit Sub
            Else
                MsgBox "Pls. Input Reference Number from Remarks Field..."
                On Error Resume Next
                txtRemarks.SetFocus
            End If
        End If
        '==========================================
        'updating code:     JAA - 02082008
        If Null2String(RSREPOR!invoice) <> "" Then
            REPOR_STATUS = "Billed-Out"
        End If
        '==========================================
        txtCustName.Text = Null2String(RSREPOR!niym)
        txtCustCode.Text = Null2String(RSREPOR!ACCT_NO)

        Dim RSCUSTINFO                                 As ADODB.Recordset
        If Null2String(RSREPOR!PLATE_NO) <> "" Then
            Set RSCUSTINFO = New ADODB.Recordset
            Set RSCUSTINFO = gconDMIS.Execute("select * from CSMS_CUSVEH where Plate_NO=" & N2Str2Null(RSREPOR!PLATE_NO))
            If Not RSCUSTINFO.EOF Or Not RSCUSTINFO.BOF Then
                txtRemarks = "MODEL: " & Null2String(RSCUSTINFO("model")) & vbCrLf & "ENGINE#:" & Null2String(RSCUSTINFO("SERIAL")) & vbCrLf & "VIN#:" & Null2String(RSCUSTINFO("vin")) & vbCrLf & "PLATE#:" & Null2String(RSCUSTINFO("plate_no"))
            End If
        End If
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""
    End If
End Sub

Sub InsertAdvanceBill()
    Dim ORDTRANDATE, ORDTRANNO, ORDTRANTYPE            As String
    Dim ORDITEMNO, ORDSTOCK_ORD, ORDSTOCK_SUP          As String
    Dim ORDTRANQTY                                     As Integer
    Dim ORDSTATUS, ORDIN_OUT                           As String
    Dim ORDTRANINVAMT                                  As Double

    Dim CURONHAND, CURSAFESTOCK, CURTISSQTY            As Integer
    Dim CURRESSERVICE, CURISSUANCES                    As Integer

    If txtTranType.Text = "RIV" Then
        Dim rsAdvanceBill                              As ADODB.Recordset
        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HD.rono,PMIS_ORD_HD.trandate,PMIS_ORD_HD.trantype,PMIS_ORD_HD.tranno,PMIS_TDAYTRAN.trantype,PMIS_TDAYTRAN.tranno,PMIS_TDAYTRAN.itemno,PMIS_TDAYTRAN.STOCK_ORD,PMIS_TDAYTRAN.tranqty,PMIS_TDAYTRAN.tranuprice from PMIS_Ord_Hd inner join PMIS_TDAYTRAN on PMIS_ORD_HD.TRANNO = PMIS_TDAYTRAN.TRANNO and PMIS_ORD_HD.TRANTYPE = PMIS_TDAYTRAN.TRANTYPE where PMIS_ORD_HD.trantype = 'ADB' and PMIS_ord_hd.rono = '" & txtRONO.Text & "' and pmis_tdaytran.[type] = 'A' ", gconDMIS
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
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!TRANQTY)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set RSPARTMAS = New ADODB.Recordset
                    RSPARTMAS.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_Accessories where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                        CURONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
                        CURSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
                        CURTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
                        CURRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
                        CURISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)

                        If CURONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                            If MsgQuestionBox("Warning: Error Has been encountered... Continue Anyway?", "Error Encountered") = False Then
                                Exit Sub
                            End If
                        End If
                        If ORDTRANQTY > CURONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(rsAdvanceBill!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CURONHAND = CURONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If

                    gconDMIS.Execute "insert into PMIS_TdayTran " & _
                                     "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                   " values ('A'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                   " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                   " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                   " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                   " netinvamt = " & ORD_TOTINVAMT & _
                                   " where id = " & labID.Caption
                    '                    gconDMIS.Execute "update PMIS_Accessories set" & _
                                         '                                   " onhand = " & CurONHAND & "," & _
                                         '                                   " TISSQTY = " & CurTISSQTY + ORDTRANQTY & ", " & _
                                         '                                   " issuances = " & curIssuances + ORDTRANQTY & _
                                         '                                   " where STOCKNO = " & ORDSTOCK_SUP
                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If

        Set rsAdvanceBill = New ADODB.Recordset
        rsAdvanceBill.Open "select PMIS_ORD_HIST.rono,PMIS_ORD_HIST.trandate,PMIS_ORD_HIST.trantype,PMIS_ORD_HIST.tranno,PMIS_DAYTRAN.trantype,PMIS_DAYTRAN.tranno,PMIS_DAYTRAN.itemno,PMIS_DAYTRAN.STOCK_ORD,PMIS_DAYTRAN.tranqty,PMIS_DAYTRAN.tranuprice from PMIS_Ord_Hist inner join PMIS_DAYTRAN on PMIS_ORD_HIST.TRANNO = PMIS_DAYTRAN.TRANNO and PMIS_ORD_HIST.TRANTYPE = PMIS_DAYTRAN.TRANTYPE where PMIS_ORD_HIST.trantype = 'ADB' and PMIS_ord_hIST.rono = '" & txtRONO.Text & "'", gconDMIS
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
                    ORDTRANQTY = N2Str2IntZero(rsAdvanceBill!TRANQTY)
                    ORDTRANINVAMT = N2Str2Zero(rsAdvanceBill!TRANUPRICE)
                    ORDIN_OUT = "'O'"
                    ORDSTATUS = "'N'"

                    Set RSPARTMAS = New ADODB.Recordset
                    RSPARTMAS.Open "Select STOCKNO,onhand,sstock,resservice,TISSQTY,issuances from PMIS_Accessories where STOCKNO = " & ORDSTOCK_ORD & " AND ACTIVE = 'Y'", gconDMIS
                    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                        CURONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
                        CURSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
                        CURTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
                        CURRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
                        CURISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)

                        If CURONHAND <= 0 Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " is Out of Stock!"
                            If MsgQuestionBox("Warning: Error Has been encountered... Continue Anyway?", "Error Encountered") = False Then
                                Exit Sub
                            End If
                        End If
                        If ORDTRANQTY > CURONHAND Then
                            Screen.MousePointer = 0
                            MsgSpeechBox "Part Number " & Null2String(rsAdvanceBill!STOCKNO) & " Qty Ordered Exceeds Current Stock!" & vbCrLf & _
                                         "This Transaction will not be Included in RIV Transaction..."
                        Else
                            CURONHAND = CURONHAND - ORDTRANQTY
                        End If
                    Else
                        Screen.MousePointer = 0
                        MsgSpeechBox "Part Number: " & Null2String(rsAdvanceBill!STOCK_ORD) & " Not Found!"
                    End If

                    gconDMIS.Execute "insert into PMIS_TdayTran " & _
                                     "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                                   " values ('A'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                                   " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                                   " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                                   " " & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
                    cleargrid grdDetails
                    DoEvents
                    FillDetails
                    gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                                   " netinvamt = " & ORD_TOTINVAMT & _
                                   " where id = " & labID.Caption
                    '                    gconDMIS.Execute "update PMIS_Accessories set" & _
                                         '                                   " onhand = " & CurONHAND & "," & _
                                         '                                   " TISSQTY = " & CurTISSQTY + ORDTRANQTY & ", " & _
                                         '                                   " issuances = " & curIssuances + ORDTRANQTY & _
                                         '                                   " where STOCKNO = " & ORDSTOCK_SUP
                    rsAdvanceBill.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Sub SetPartDetails(XXX As String)
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    Set RSPARTMAS = gconDMIS.Execute("Select * from PMIS_Accessories where PartNo = '" & XXX & "' AND ACTIVE = 'Y'")
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

Sub InitParts()
    txtTranItemNo.Text = Format(KCNT + 1, "0000")
    cboTranPartNo.Text = ""
    txtTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUCost.Text = "0.00"
    txtTranUPrice.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
    If COUNTERTYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
    Check1.Enabled = False
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1
    cmdAddTran.Visible = False
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    cmdSignatories.ZOrder 1
    cmdSignatories.Visible = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False
    Picture1.Enabled = True
    fraDetails.Enabled = True
    Frame6.ZOrder 1
    Frame6.Visible = False
End Sub

Sub BringToFront()
    cmdAddTran.ZOrder 0
    cmdAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True

    Picture1.Enabled = False
    fraDetails.Enabled = False
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtCustName.Text = Null2String(rsCustomer!AcctName) & vbCrLf & Null2String(rsCustomer!CUSTOMERADD) & vbCrLf & Null2String(rsCustomer!City)
    End If
End Sub


Private Sub cboRefPRSNo_Click()
    cboRefPRSNo_LostFocus
End Sub

Private Sub cboRefPRSNo_GotFocus()
    Dim rsPRS                                          As ADODB.Recordset
    Dim rsPRS_HDDup                                    As ADODB.Recordset
    
    Set rsPRS = New ADODB.Recordset
    If COUNTERTYPE = "RIV" Or COUNTERTYPE = "ADB" Then
       rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = 'A' and SALES_ORIGIN ='S' order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    ElseIf COUNTERTYPE = "CSH" Or COUNTERTYPE = "CHG" Or COUNTERTYPE = "DR" Then
        rsPRS.Open "Select tranno,refpisno from PMIS_vw_PRS WHERE [TYPE] = 'A' and (SALES_ORIGIN ='W' or SALES_ORIGIN ='O' OR SALES_ORIGIN ='M' or SALES_ORIGIN ='J') order by Tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    
    If Not rsPRS.EOF And Not rsPRS.BOF Then
        rsPRS.MoveFirst: cboRefPRSNo.Clear
        Do While Not rsPRS.EOF
            Set rsPRS_HDDup = New ADODB.Recordset
            rsPRS_HDDup.Open "select refpisno from PMIS_Ord_Hd where TRANTYPE <> 'ARS' AND [TYPE] = 'A' AND refprsno = '" & Null2String(rsPRS!REFPISNO) & "'", gconDMIS
            If Not rsPRS_HDDup.EOF And Not rsPRS_HDDup.BOF Then
            Else
                cboRefPRSNo.AddItem Null2String(rsPRS!REFPISNO)
            End If
            rsPRS.MoveNext
        Loop
    End If
End Sub

Private Sub cboRefPRSNo_LostFocus()
    If AddorEdit = "ADD" Then
        Dim rsRR_HDDup                                 As ADODB.Recordset
        Set rsRR_HDDup = New ADODB.Recordset
        rsRR_HDDup.Open "select refpisno,tranno from PMIS_Ord_Hd where [TYPE] = 'A' AND refprsno = '" & cboRefPRSNo.Text & "'", gconDMIS
        If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
            MsgBox "ARS Number Already Received", vbInformation, "Invalid ARS Number"
            Exit Sub
        Else
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select tranno,DS1,custname,custcode,rono from PMIS_vw_PRS where [TYPE] = 'A' AND refpisno = '" & cboRefPRSNo.Text & "'", gconDMIS

            If Not rsRR_HDDup.EOF Or Not rsRR_HDDup.BOF Then
                txtCustName = Null2String(rsRR_HDDup!CUSTNAME)
                txtCustCode = Null2String(rsRR_HDDup!CUSTCODE)
                txtRONO = Null2String(rsRR_HDDup!RONO)
            End If

            If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
                KCNT = 0: ORD_TOTUPRICE = 0: ORD_TOTINVAMT = 0: ORD_TOTVAT = 0: ORD_TOTQTY = 0
                Dim STOCKDESCription                   As String
                Set RSTDAYTRAN = New ADODB.Recordset: cleargrid grdDetails
                RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(rsRR_HDDup!TRANNO) & " and trantype = 'ARS' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
                    cboChargeTo.Enabled = False: Screen.MousePointer = 11: RSTDAYTRAN.MoveFirst
                    Do While Not RSTDAYTRAN.EOF
                        KCNT = KCNT + 1
                        STOCKDESCription = SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_ORD))
                        grdDetails.AddItem RSTDAYTRAN!ID & Chr(9) & Format(Null2String(RSTDAYTRAN!itemno), "0000") & Chr(9) & _
                                           Null2String(RSTDAYTRAN!STOCK_ORD) & Chr(9) & _
                                           STOCKDESCription & Chr(9) & _
                                           N2Str2IntZero(RSTDAYTRAN!TRANQTY) & Chr(9) & _
                                           Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                                           Format(N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
                        ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                        ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
                        ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE))
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
                MsgSpeechBox "Invalid Accessories Requisition Number!": If AddorEdit = "ADD" Then cleargrid grdDetails
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
    End If
End Sub

Private Sub cboTranPartNo_Click()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        Check1.Enabled = True
    Else
        Check1.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    Dim rschek                                         As New ADODB.Recordset

    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)

        Set rschek = New ADODB.Recordset
        rschek.Open "Select * from PMIS_Accessories where active = 'Y' and type = 'A'and stockno = " & N2Str2Null(cboTranPartNo) & "", gconDMIS, adOpenKeyset, adLockReadOnly

        If Not rschek.EOF And Not rschek.BOF Then
        Else
            MsgBox "Sorry partnumber is not in the list pls try again!", vbCritical
            cboTranPartNo = ""
            cboTranPartNo.SetFocus
        End If


    End If
End Sub

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub Check1_Click()
    If Module_Access(LOGID, "APPLY ACCESSORIES COST TO COST AMOUNT", "SYSTEM") = False Then Check1.Value = 0: Exit Sub
    If Check1.Value = 1 Then
        txtTranUPrice.Text = txtTranUCost.Text
    Else
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    cmdAddTran.Visible = True
    cmdAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    AddorEdit = "ADD"
    cmdTranDelete.Visible = False
    InitParts
    On Error Resume Next
    cboTranPartNo.SetFocus
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LOCALACESS) = False Then Exit Sub

    On Error GoTo Errorcode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCURONHAND, PCURTISSQTY, PCURISSUANCES     As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_Accessories where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                    If Null2String(RSORD_HD!Status) = "P" Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            gconDMIS.Execute "update PMIS_Accessories set" & _
                                           " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                           " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        Else
                            gconDMIS.Execute "update PMIS_Accessories set" & _
                                           " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                           " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        End If
                        gconDMIS.Execute "update PMIS_Accessories set" & _
                                       " onhand = " & PCURONHAND & "," & _
                                       " tissqty = " & PCURTISSQTY & "," & _
                                       " issuances = " & PCURISSUANCES & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                    End If
                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                  " status = 'C'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""
                End If
                RSTDAYTRANDUP.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " status = 'C'," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""
        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labID.Caption
        StoreMemVars
    End If
    Set RSTDAYTRANDUP = Nothing
    Set RSPARTMASDUP = Nothing

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", LOCALACESS) = False Then Exit Sub
    txtTranDate.Enabled = True
    txtTranDate.Locked = False

End Sub

Private Sub CMDHPI_PIS_Click()
     'For the Issuance on Service, Advance bill, DR out,cash,charge
    If RSORD_HD!TranType = "RIV" Or RSORD_HD!TranType = "CSH" Or RSORD_HD!TranType = "ADB" Or RSORD_HD!TranType = "DR" Or RSORD_HD!TranType = "CHG" Then
        'this is for the service issuance
        rptCustomerOrder.Reset
        If Mid(Trim(txtReferencePIS), 3, 1) = "S" Then
            
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(1) = "RECEIVEDBY= '" & TXT_PRINT_REC_BY & "'"
            rptCustomerOrder.Formulas(2) = "ISSUEDBY= '" & TXT_PRINT_ISSUEBY & "'"

            If Mid(Trim(txtReferencePIS), 5, 1) = "W" Then
                Dim warvar As String
                warvar = "0.00"
                'for warranty only
                rptCustomerOrder.Formulas(3) = "Warranty= '" & warvar & "'"
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_PIS_ACC_war.RPT", "({ord_hd.TRANTYPE} = 'RIV' or {ord_hd.trantype} = 'ADB' or {ord_hd.trantype} = 'DR'  ) and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            
            Else
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_PIS_ACC.RPT", "({ord_hd.TRANTYPE} = 'RIV' or {ord_hd.trantype} = 'ADB' or {ord_hd.trantype} = 'DR'  ) and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            End If
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "PREPARED BY", TXT_PRINT_ISSUEBY)
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "ISSUED BY", TXT_PRINT_REC_BY)
            Screen.MousePointer = 0

        'this is for the walk- in customer ,jobbers, etc.
        Else
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(1) = "RECEIVEDBY= '" & TXT_PRINT_REC_BY & "'"
            rptCustomerOrder.Formulas(2) = "ISSUEDBY= '" & TXT_PRINT_ISSUEBY & "'"
            If chkvat.Value = 1 Then
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_ACCsin.RPT", "({ord_hd.TRANTYPE} = 'CSH' or {ord_hd.TRANTYPE} = 'CHG' or {ord_hd.TRANTYPE} = 'DR') and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Else
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_ACCsin1.RPT", "({ord_hd.TRANTYPE} = 'CSH' or {ord_hd.TRANTYPE} = 'CHG' or {ord_hd.TRANTYPE} = 'DR') and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            End If
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "PREPARED BY", TXT_PRINT_ISSUEBY)
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "ISSUED BY", TXT_PRINT_REC_BY)
            Screen.MousePointer = 0
        End If
    End If
End Sub


Private Sub CMDHPI_SI_Click()
      If RSORD_HD!TranType = "RIV" Or RSORD_HD!TranType = "CSH" Or RSORD_HD!TranType = "ADB" Or RSORD_HD!TranType = "DR" Or RSORD_HD!TranType = "CHG" Then
       rptCustomerOrder.Reset
        If Mid(RSORD_HD!REFPISNO, 3, 1) = "S" Then
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(1) = "RECEIVEDBY= '" & TXT_PRINT_REC_BY & "'"
            rptCustomerOrder.Formulas(2) = "ISSUEDBY= '" & TXT_PRINT_ISSUEBY & "'"

            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RIV_PIS_ACC.RPT", "({ord_hd.TRANTYPE} = 'RIV' or {ord_hd.trantype} = 'ADB' or {ord_hd.trantype} = 'DR'  ) and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "PREPARED BY", TXT_PRINT_ISSUEBY)
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "ISSUED BY", TXT_PRINT_REC_BY)
            Screen.MousePointer = 0
        Else
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(1) = "RECEIVEDBY= '" & TXT_PRINT_REC_BY & "'"
            rptCustomerOrder.Formulas(2) = "ISSUEDBY= '" & TXT_PRINT_ISSUEBY & "'"
            
            If chkvat.Value = 1 Then
                 PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_ACC.RPT", "({ord_hd.TRANTYPE} = 'CSH' or {ord_hd.TRANTYPE} = 'CHG' or {ord_hd.TRANTYPE} = 'DR') and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Else
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_ACC1.RPT", "({ord_hd.TRANTYPE} = 'CSH' or {ord_hd.TRANTYPE} = 'CHG' or {ord_hd.TRANTYPE} = 'DR') and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            End If
            
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "PREPARED BY", TXT_PRINT_ISSUEBY)
            Call SaveSetting("DMIS", "PARTS ISSUANCE", "ISSUED BY", TXT_PRINT_REC_BY)
            Screen.MousePointer = 0
        End If
    End If
End Sub

Private Sub cmdPISNum_Click()
    With frmPMISAC_AIFormation
        If AddorEdit = "EDIT" Then
            .txtedit = "EDIT"
        Else
            .txtedit = ""
        End If
        .dtTranDate.Enabled = False
         .dtTranDate = txtTranDate
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
    frmPMISAC_AIFormation.Show 1
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub cmdPost_Click()

    Dim rsPrtMas                                       As New ADODB.Recordset
    Dim rsTdytran                                      As New ADODB.Recordset
    Dim blnStockremove                                 As Boolean
    Dim strPartno                                      As String


    If Function_Access(LOGID, "Acess_Post", LOCALACESS) = False Then Exit Sub

    On Error GoTo Errorcode:

    '====================================================================================================
    'updating code: JAA - 07082008     'Do not allow posting of transaction without issuance of Accessories
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD = "" Or FILD = "No Entry" Then
        MsgBox "Posting of Transaction without issuance of Accessories is not allowed.", vbCritical, "Pls. Add Accessories."
        Exit Sub
    End If
    '====================================================================================================

    '=[ EAP:033109: check parts if current onhand is not zero in posting ]=
    If txtTranType = "RIV" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'A' and trantype in('RIV') "), gconDMIS
    ElseIf txtTranType = "CSH" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'A' and trantype in('CSH') "), gconDMIS
    ElseIf txtTranType = "CHG" Then
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'A' and trantype in('CHG') "), gconDMIS
    Else
        rsTdytran.Open ("select stock_ord,tranqty, ID from pmis_tdaytran where tranno = '" & txtTranNo & "' and type = 'A' and trantype in('DR') "), gconDMIS
    End If

    If Not (rsTdytran.BOF And rsTdytran.EOF) Then
        Do While Not rsTdytran.EOF

            rsPrtMas.Open "Select STOCKNO,onhand,srp from PMIS_ACCESSORIES where STOCKNO = '" & rsTdytran!STOCK_ORD & "' ", gconDMIS
            '=[ EAP:040209: this will remove the partnumber without stock in the transaction. ]=
            If Not (rsPrtMas.BOF And rsPrtMas.EOF) Then
                If rsPrtMas!ONHAND <= 0 Then
                    MsgBox "Partnumber# " & rsTdytran!STOCK_ORD & " will be remove from the transaction Out of Stock"
                    SQL_STATEMENT = "delete from PMIS_TdayTran where Id = '" & rsTdytran!ID & "' "
                    gconDMIS.Execute SQL_STATEMENT
                    blnStockremove = True
                ElseIf rsPrtMas!ONHAND < rsTdytran!TRANQTY Then
                    MsgBox "SOME PARTNUMBER ONHAND IS LESS THAN YOUR REQUEST QUANTITY", vbInformation
                    Exit Sub
                ElseIf rsPrtMas!SRP = 0 Then
                    MsgBox "THERE'S A ZERO SRP IN THE TRANSACTION, CANNOT POST ", vbInformation
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
        Dim PCURONHAND As Long
        Dim PCURTISSQTY As Long
        Dim PCURISSUANCES As Long
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED,NON_HARI from PMIS_Accessories where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    '====================================================
                    'updating code: JAA - 09082008  -- Do not deduct stock from Master File.
                    If COUNTERTYPE <> "ADB" Then
                        PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            SQL_STATEMENT = "update PMIS_Accessories set" & _
                                          " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                        Else
                            SQL_STATEMENT = "update PMIS_Accessories set" & _
                                          " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Accessories"), "", "TRAN NO: " & txtTranNo & " POSTED", "", "")

                        SQL_STATEMENT = "update PMIS_Accessories set" & _
                                      " onhand = " & PCURONHAND & "," & _
                                      " tissqty = " & PCURTISSQTY & "," & _
                                      " issuances = " & PCURISSUANCES & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Accessories"), "", "TRAN NO: " & txtTranNo & " POSTED", "", "")
                    End If

                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                  " status = 'P'," & _
                                  " NON_HARI = " & N2Str2Null(RSPARTMASDUP!NON_HARI) & "," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "PP", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""
                End If
                RSTDAYTRANDUP.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " status = 'P'," & _
                      " totalqty = " & ORD_TOTQTY & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""
        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labID.Caption
        StoreMemVars

        Set RSTDAYTRANDUP = Nothing
        Set RSPARTMASDUP = Nothing

        If txtTranType.Text = "RIV" Then
            If CheckIfROBilled(txtRONO.Text) <> "" Then
                MsgBox "Warning: This issuance will not be Exported to Billing since Repair Order is already Billed!", vbCritical, "Repair Order Already Billed"
                MsgBox "Warning: Status for this issuance will now be tag as Billed!", vbCritical, "Repair Order Already Billed"
            Else
'                 'this importing is reference to return parts from service
'                 'update by:NVB
'
'                If ImportDetails(txtRONO, "A", "4") = True Then
'                'do nothing
'                End If
                If VALID_COMPANY_CODE(COMPANY_CODE) = True Then
                    If ImportDetails(txtRONO, "A", "4") = True Then
                        'do nothing
                    End If
                Else
                    Call ImportAccessories(txtRONO)
                End If
            
            End If
        End If

    End If

    Exit Sub
Errorcode:
    'ShowVBError
    MsgBox err.Description
    Exit Sub

End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LOCALACESS) = False Then Exit Sub
    On Error GoTo Errorcode:
    
       If COMPANY_CODE = "HPI" Then
        TXT_PRINT_ISSUEBY.Text = GetSetting("DMIS", "PARTS ISSUANCE", "PREPARED BY", "")
        TXT_PRINT_REC_BY.Text = GetSetting("DMIS", "PARTS ISSUANCE", "ISSUED BY", "")
        If Mid(RSORD_HD!REFPISNO, 3, 1) = "S" Then CMDHPI_SI.Enabled = False
        
        picHPI.Visible = True
        picHPI.ZOrder 0
        Exit Sub
    End If
    rptCustomerOrder.Reset
    
    If COMPANY_CODE = "HQA" Then
        Picture1.Enabled = False
        Frame6.ZOrder 0
        Frame6.Visible = True
        If Mid(LTrim(RTrim(txtReferencePIS)), 3, 1) = "S" Then
            Command6.Enabled = False
        Else
            Command6.Enabled = True
        End If
        Exit Sub
        
    End If
    
    If COMPANY_CODE = "HLI" Then
        Call PRINT_ME
        Exit Sub
    End If
    
    If RSORD_HD!TranType = "ADB" Or RSORD_HD!TranType = "RIV" Then
        If MsgQuestionBox("Accessories Issuance Slip will be printed. You want to print it in a Blank form?", "Confirm Printing...") = True Then
            
            If COMPANY_CODE = "HCI" Then
                SERVICEPISPRINTING
            Else
                cmdSignatories.Visible = True
                cmdSignatories.ZOrder 0
                fraSignatories.Visible = True
                fraSignatories.ZOrder 0
                txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
                txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
                txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
                On Error Resume Next
                txtRequestedBy.SetFocus
            End If
        Else
            If COMPANY_CODE = "HCI" Then
                Exit Sub
            Else
                SERVICEPISPRINTING
            End If
        End If
    End If

    If RSORD_HD!TranType = "CSH" Then
        If MsgQuestionBox("Accessories Issuance Slip (CSH) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            If COMPANY_CODE = "HMH" Then
                If MsgQuestionBox("Print Accessories Issuance in a Blank form?", "Confirm Printing...") = True Then
                    cmdSignatories.Visible = True
                    cmdSignatories.ZOrder 0
                    fraSignatories.Visible = True
                    fraSignatories.ZOrder 0
                    txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
                    txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
                    txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
                    On Error Resume Next
                    txtRequestedBy.SetFocus
                Else
                    CSHPRINTING
                End If
            Else
                CSHPRINTING
            End If
        End If
    End If

    If RSORD_HD!TranType = "CHG" Then
        If MsgQuestionBox("Accessories Issuance Slip (CHG) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            If COMPANY_CODE = "HMH" Then
                If MsgQuestionBox("Print Accessories Issuance in a Blank form?", "Confirm Printing...") = True Then
                    cmdSignatories.Visible = True
                    cmdSignatories.ZOrder 0
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

    If RSORD_HD!TranType = "DR" Then
        If MsgQuestionBox("DR Out Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            If COMPANY_CODE = "HCI" Then
                rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
                rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                Screen.MousePointer = 11
                PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "DR.rpt", "{ord_hd.TYPE} = 'A' and {ord_hd.TRANTYPE} = 'DR' and {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
                Screen.MousePointer = 0
            Else
                NEWDRPRINTING
            End If
        End If
    End If

    NEW_LogAudit "V", LOCALACESS, "", labID, "Accessories", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub PRINT_ME()

     If MsgBox("Parts Issuance Slip will be printed. DO you want to continue?", vbYesNo + vbQuestion) = vbYes Then
        If RSORD_HD!TranType = "RIV" Then
            Call SERVICEPISPRINTING
        ElseIf RSORD_HD!TranType = "CSH" Then
            Call CSHPRINTING
        ElseIf RSORD_HD!TranType = "CHG" Then
            Call CHGPRINTING
        ElseIf RSORD_HD!TranType = "ADB" Then
            Call ADBPRINTING
        ElseIf RSORD_HD!TranType = "DR" Then
            Call NEWDRPRINTING
        End If
     Else
        Exit Sub
     End If
     
End Sub

Private Sub cmdPrintRIV_Click()
    If RSORD_HD!TranType = "RIV" Then
        SERVICEPISPRINTING_BLANKFORM
    End If
    If RSORD_HD!TranType = "ADB" Then
        ADBPRINTING
    End If
    If RSORD_HD!TranType = "CSH" Then
        CSHPRINTING_OTC
    End If
    If RSORD_HD!TranType = "CHG" Then
        CHGPRINTING_OTC
    End If

    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)

    SendToBack
End Sub

Private Sub cmdTranCancel_Click()
    Picture1.Enabled = True
    fraDetails.Enabled = True
    picDetails.Enabled = True
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo Errorcode:

    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete this Accessories, Are you Sure?", "Delete Accessories Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_TdayTran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", LOCALACESS, SQL_STATEMENT, labID, "Accessories", "ACC NO: " & cboTranPartNo, COUNTERTYPE, labDetID
        ShowDeletedMsg
    End If

    Dim CNT                                            As Integer
    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Set RSTDAYTRANDUP = New ADODB.Recordset
    RSTDAYTRANDUP.Open "select id,itemno from PMIS_TdayTran where [TYPE] = 'A' AND trantype = " & N2Str2Null(COUNTERTYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
    If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
        RSTDAYTRANDUP.MoveFirst
        CNT = 0
        Do While Not RSTDAYTRANDUP.EOF
            CNT = CNT + 1
            gconDMIS.Execute "update PMIS_TdayTran set itemno = " & Format(CNT, "0000") & " where id = " & RSTDAYTRANDUP!ID
            RSTDAYTRANDUP.MoveNext
        Loop
    End If
    FillDetails
    SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                  " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                  " netinvamt = " & ORD_TOTINVAMT & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "", "TRAN NO: " & txtTranNo & " REMOVE ACC", COUNTERTYPE, "")

    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Accessories No. must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_TdayTran where [TYPE] = 'A' AND STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Accessories No. already used in this transaction"
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

    Dim CURONHAND, CURSAFESTOCK, CURTISSQTY            As Integer
    Dim CURRESSERVICE, CURISSUANCES, PREVCURORDQTY     As Integer
    Dim ORDMAC
    Dim CRITICAL_QUESTION                              As String

    If txtTranType.Text <> "ADB" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "SELECT isnull(srp,0) as srp,STOCKNO,ONHAND,SSTOCK,RESSERVICE,TISSQTY,MAC ,ISSUANCES FROM PMIS_ACCESSORIES WHERE STOCKNO = '" & cboTranPartNo.Text & "' AND ACTIVE = 'Y'", gconDMIS
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            CURONHAND = N2Str2IntZero(RSPARTMAS!ONHAND)
            CURSAFESTOCK = N2Str2IntZero(RSPARTMAS!SSTOCK)
            CURTISSQTY = N2Str2IntZero(RSPARTMAS!TISSQTY)
            CURRESSERVICE = N2Str2IntZero(RSPARTMAS!RESSERVICE)
            CURISSUANCES = N2Str2IntZero(RSPARTMAS!ISSUANCES)
            ORDMAC = NumericVal(RSPARTMAS!MAC)
            If AddorEdit <> "ADD" Then
                PREVCURORDQTY = NumericVal(labPrevOrdQty.Caption)
                'commented by: NVB
                'this can cause negative issuances
                'CurONHAND = CurONHAND + PrevCurOrdQty
                CURTISSQTY = CURTISSQTY - PREVCURORDQTY
                CURISSUANCES = CURISSUANCES - PREVCURORDQTY
            End If
            If RSPARTMAS!SRP <= 0 Then
                 MsgBox "Accessory number :" & (RSPARTMAS!STOCKNO) & " has Zero srp. ", vbInformation, "Saving Error"
                 Exit Sub
            End If
            
            If CURONHAND <= 0 Then
                Screen.MousePointer = 0
                MsgSpeechBox "Out of Stock!"
                Exit Sub
            End If
            If ORDMAC <= 0 Then
                MsgBox "Warning: This Accessories Code has Zero Cost! Pls Check in Accessories Master File or Process Update Master File to Proceed.", vbCritical, "Stock Has Zero Cost"
                Screen.MousePointer = 0
                Exit Sub
            Else
                txtTranUCost.Text = ORDMAC
            End If

            If txtTranType.Text = "CSH" Or txtTranType.Text = "CHG" Then
                If CURONHAND <= CURRESSERVICE Then
                    Screen.MousePointer = 0
                    If MsgQuestionBox("Stock is Reserved for Service... Continue Anyway?", "Stock Status Alert!") = False Then
                        Exit Sub
                    End If
                    CRITICAL_QUESTION = "Stock is Reserved for Service... Continue Anyway?"
                    Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " ACC NO: " & cboTranPartNo, "", "")
                    MsgBox "User Action has bee Log in the Audit Trail", vbInformation, "Audit Trail Information"
                End If
            End If

            If NumericVal(txtTranQty.Text) > CURONHAND Then
                Screen.MousePointer = 0
                MsgSpeechBox "Qty Ordered Exceeds Current Stock!"
                On Error Resume Next
                txtTranQty.SetFocus
                Exit Sub
            Else
                CURONHAND = CURONHAND - NumericVal(txtTranQty.Text)
            End If

            If CURONHAND < CURSAFESTOCK Then
                Screen.MousePointer = 0
                If MsgQuestionBox("Current On-hand is now below the Safety Stock Level... Proceed Anyway?", "Safety Stock Alert!") = False Then
                    Exit Sub
                End If
                CRITICAL_QUESTION = "Current On-hand is now below the Safety Stock Level... Proceed Anyway?"
                Call NEW_LogAudit("MP", LOCALACESS, CRITICAL_QUESTION, labID, "", "TRAN NO: " & txtTranNo & " ACC NO: " & cboTranPartNo, COUNTERTYPE, "")
                MsgBox "User Action has been Log in the Audit Trail", vbInformation, "Audit Trail Information"
                Screen.MousePointer = 11
            End If
        Else
            Screen.MousePointer = 0
            MsgSpeechBox "Part Number not Found!"
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
    ORDTRANUCOST = NumericVal(txtTranUCost.Text)

    ORDTRANINVAMT = NumericVal(txtTranUPrice.Text)
    If txtTranType.Text = "ADB" Then ORDIN_OUT = "'A'" Else ORDIN_OUT = "'O'"
    ORDSTATUS = "'N'"

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                        "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,mac,tranuprice,lastupdate,usercode,status,in_out)" & _
                      " values ('A'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDMAC & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", LOCALACESS, SQL_STATEMENT, labID, "Accessories", "ACC NO: " & cboTranPartNo, COUNTERTYPE, labDetID

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update PMIS_TdayTran set" & _
                      " trandate = " & ORDTRANDATE & "," & _
                      " trantype = " & ORDTRANTYPE & "," & _
                      " mac= " & ORDMAC & "," & _
                      " tranno = " & ORDTRANNO & "," & _
                      " itemno = " & ORDITEMNO & "," & _
                      " STOCK_ORD = " & ORDSTOCK_ORD & "," & _
                      " STOCK_SUP = " & ORDSTOCK_SUP & "," & _
                      " tranqty = " & ORDTRANQTY & "," & _
                      " tranucost = " & ORDTRANUCOST & "," & _
                      " tranuprice = " & ORDTRANINVAMT & "," & _
                      " lastupdate = '" & LOGDATE & "'," & _
                      " status = " & ORDSTATUS & "," & _
                      " in_out = " & ORDIN_OUT & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "" & _
                      " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "Accessories", "ACC NO: " & cboTranPartNo, COUNTERTYPE, labDetID

        ShowSuccessFullyUpdated
    End If
    cleargrid grdDetails
    FillDetails
    SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                  " totalqty = " & ORD_TOTQTY & "," & _
                  " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                  " netinvamt = " & ORD_TOTINVAMT & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, labID, "", "TRAN NO: " & txtTranNo & " ADD/EDIT PARTS", COUNTERTYPE, "")

    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then
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

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LOCALACESS) = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    fraDetails.Enabled = False
    grdDetails.Enabled = False
    'EAP:033109 so user cannot pressd f8 when transaction is not yet saved.
    cmdPost.Enabled = False
    Command4.Enabled = True
    ichg = False
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    txtTranDate.Enabled = False
    StoreMemVars
    txtPRtranno.Visible = False
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LOCALACESS) = False Then Exit Sub
    AddorEdit = "EDIT"
    PREVORDTYPE = txtTranType.Text
    PREVORDNO = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    txtTranDate.Enabled = False
    On Error Resume Next
    txtCustName.SetFocus
    Command4.Enabled = True
    ichg = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSORD_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSORD_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSORD_HD.MoveNext
    If RSORD_HD.EOF Then
        RSORD_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSORD_HD.MovePrevious
    If RSORD_HD.BOF Then
        RSORD_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()

    'updating code:     jaa - 10292008          - Transaction Month should be equal to current month
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            'txtTranDate.SetFocus
        End If
    End If

    On Error GoTo Errorcode
    Dim NEXTCUNTER                                     As String
    Dim RSFINDDUP                                      As ADODB.Recordset
    Dim XSALES_ORIGIN, XSI_TYPE, XPAY_CLASS, XCHAR_YEAR, XCHAR_MONTH, XIS_SERIES, XTRACK_CODE As String

    Dim VCBOSALESMAN, VCBOSMNAME, VTXTTRANTYPE         As String
    Dim VTXTTRANNO, VTXTTRANDATE, VTXTCUSTCODE         As String
    Dim VTXTCUSTNAME, VTXTCHARGETO, VTXTRONO, VTXTREFPRSNO, VTXTREP_OR As String
    Dim VtxtTerms                                      As String
    Dim VTXTTTLINVAMT, VTXTDS1                         As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNETINVAMT                     As Double
    Dim VTXTRemarks, Vusercode                         As String
    Dim VLastUpdate                                    As String
    Dim VIN_PROCESS                                    As String
    Dim VTXTREFERENCEPIS                               As String

    'axp 02282008
    If Len(Trim(RTrim(txtTranNo))) <> 6 Then
        MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    End If


    '********************************************************************
    'updating code:     jaa - 10132008      - Require PIS in all type of Issuances
    'If Trim(txtTranType.Text) = "DR" Then
    '    'proceed
    'Else
    If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
        MsgBox "Invalid Reference AIS Number!", vbCritical, "PIS Required!"
        Exit Sub
    End If
    'End If
    '********************************************************************

    If Trim(txtTranType.Text) = "RIV" Then
        If Trim(txtRONO.Text) = "" Then
            MsgBox "RO Number is Required...", vbInformation, "Pls Input RO Number..."
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtCustCode)) = "" Then
            MsgBox "Customer Information Is Required...", vbInformation, "Pls Select Customer Information..."
            Command1.SetFocus
            Exit Sub
        End If
    End If

    '********************************************************************
    'updating code:     jaa - 10132008      - Require PIS in all type of Issuances
    'If Trim(txtTranType.Text) = "DR" Then
    'proceed
    'Else
    If RTrim(LTrim(cboRefPRSNo.Text)) = "" Then
        MsgBox "Reference ARS Number is Required...", vbInformation, "Pls. select ARS No."
        Exit Sub
    End If
    'End If
    '********************************************************************

    Select Case txtTranType
        Case "CHG"
'updated BY:    IEBV 02282011_1105pm
'descriptio:    Validation fo customer that has a credit limit fo those issuance that are CHARGE
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If Mid((txtReferencePIS.Text), 5, 1) = "C" Then
            Dim rsterms As ADODB.Recordset
            Set rsterms = New ADODB.Recordset
            Set rsterms = gconDMIS.Execute("Select isnull(creditdays,0) as creditdays, isnull(creditlimit,0) as creditlimit  from  all_customer where cuscde = '" & Me.txtCustCode.Text & "'")
            If AddorEdit = "EDIT" Then
                If txtCustCode.Text <> Null2String(RSORD_HD!CUSTCODE) Then
                    If ichg = True Then
                        'do nothing
                    Else
                        If rsterms!CREDITDAYS = 0 Then
                            MsgBox "Terms not yet configured.", vbInformation + vbOKOnly
                            Exit Sub
                        ElseIf NumericVal(txtNetInvAmt.Text) > NumericVal(rsterms!CreditLimit) Then
                            If MsgBox("Credit is over the limit,Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                            picoverride.Visible = True
                            picoverride.ZOrder 0
                            txtoverride.Text = ""
                            ictr = 3
                            Exit Sub
                        End If
                    End If
                Else
                    ichg = True
                End If
            Else
                If ichg = True Then
                    'do nothing
                Else
                    If rsterms!CREDITDAYS = 0 Then
                        MsgBox "Terms not yet configured.", vbInformation + vbOKOnly
                        Exit Sub
                    ElseIf NumericVal(txtNetInvAmt.Text) > NumericVal(rsterms!CreditLimit) Then
                        If MsgBox("Credit is over the limit,Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                        picoverride.Visible = True
                        picoverride.ZOrder 0
                        txtoverride.Text = ""
                        ictr = 3
                        Exit Sub
                    End If
                End If
            End If
        End If
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Case "RIV"
            If CheckIfRoExists(txtRONO.Text) = "" Then
                MsgBox "RO Number Doesn't Exists. Please Correct Repair Order Number", vbCritical, "Invalid RO Number"
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
            Set RSADB = gconDMIS.Execute("SELECT COUNT(RONO) FROM PMIS_ORD_HD WHERE TYPE='P' AND  TRANTYPE='ADB' AND RONO=" & N2Str2Null(txtRONO))
            If RSADB.Fields(0).Value > 0 Then
                If MsgBox("There is Advance Bill for this RO!!" & vbCrLf & " Are you Sure You will do service issuance(s)?", vbInformation + vbYesNo, "Advance Bill Deteched!!") = vbNo Then
                    On Error Resume Next
                    txtRONO.SetFocus
                    Exit Sub
                End If
            End If

    End Select

    'validation for transaction number
    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction No. must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set RSFINDDUP = New ADODB.Recordset
            'updating code: jaa - 09102008          - Check tranno if exist from current transaction and from history
            If txtTranType = "ADB" Then
                Call RSFINDDUP.Open("SELECT TRANNO  FROM PMIS_ORD_HD WHERE [TYPE] = 'A'  AND TRANTYPE='ADB' and tranno = '" & txtTranNo.Text & "' " & " union" & " SELECT TRANNO FROM PMIS_ORD_HIST WHERE [TYPE] = 'M' AND TRANTYPE='ADB' and tranno = '" & txtTranNo.Text & "'", gconDMIS, adOpenKeyset)
            Else

                RSFINDDUP.Open "select trantype,tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
                Set RSFINDDUP = New ADODB.Recordset
                'updating code: jaa - 09102008          - Check tranno if exist from current transaction and from history
                If txtTranType = "ADB" Then
                    Call RSFINDDUP.Open("SELECT TRANNO  FROM PMIS_ORD_HD WHERE [TYPE] = 'A'  AND TRANTYPE='ADB' and tranno = '" & txtTranNo.Text & "' " & " union" & " SELECT TRANNO FROM PMIS_ORD_HIST WHERE [TYPE] = 'M' AND TRANTYPE='ADB' and tranno = '" & txtTranNo.Text & "'", gconDMIS, adOpenKeyset)
                Else
                    RSFINDDUP.Open "select trantype,tranno from PMIS_vw_ISS_HISTORY where [TYPE] = 'A' AND trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
                End If
                If Not RSFINDDUP.EOF And Not RSFINDDUP.BOF Then
                    MsgSpeechBox "Transaction No. already exist!"
                    On Error Resume Next
                    txtTranNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    If txtTranDate.Text = "" Or IsDate(txtTranDate.Text) = False Then
        MsgSpeechBox "Invalid Transaction Date!"
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If
    If txtTranType.Text = "CHG" Then
        If txtTerms.Text = "" Then
            MsgSpeechBox "Terms must have a value."
            On Error Resume Next
            txtTerms.SetFocus
            Exit Sub
        End If
    End If
    VCBOSALESMAN = N2Str2Null(cboSalesMan.Text)
    VCBOSMNAME = N2Str2Null(cboSMName.Text)
    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE               As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP             As String
    Dim RRTRANQTY                                      As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                      As Double
    Dim RRIN_OUT, RRSTATUS, VStatus                    As String

    If Left(txtTranNo.Text, 1) = "A" Then
    Else
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
    If txtRemarks.Text = "Pls Type Your Message Here!" Then VTXTRemarks = "NULL" Else VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))

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

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_Ord_Hd" & _
                      " (TYPE,trantype,tranno,trandate,custcode,custname,chargeto,REFPRSNO,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " values ('A'," & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & "," & VTXTREFPRSNO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", LOCALACESS, SQL_STATEMENT, FindTransactionID(txtTranNo, "tranno", "PMIS_Ord_Hd", "DETAILS", N2Str2Null("A"), "TYPE"), "Accessories", "TRAN NO: " & txtTranNo & " - " & VTXTREFPRSNO, COUNTERTYPE, ""
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " trantype = " & VTXTTRANTYPE & "," & _
                      " tranno = " & VTXTTRANNO & "," & _
                      " trandate = " & VTXTTRANDATE & "," & _
                      " custcode = " & VTXTCUSTCODE & "," & _
                      " custname = " & VTXTCUSTNAME & "," & _
                      " chargeto = " & VTXTCHARGETO & "," & _
                      " REFPRSNO = " & VTXTREFPRSNO & "," & _
                      " rono = " & VTXTRONO & "," & _
                      " rep_or = " & VTXTREP_OR & "," & _
                      " salesman = " & VCBOSALESMAN & "," & _
                      " smname = " & VCBOSMNAME & "," & _
                      " terms = " & VtxtTerms & "," & _
                      " ttlinvamt = " & VTXTTTLINVAMT & "," & _
                      " ds1 = " & VTXTDS1 & "," & _
                      " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                      " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                      " netinvamt = " & VTXTNETINVAMT & "," & _
                      " remarks = " & VTXTRemarks & ", " & _
                      " status = " & VStatus & ", " & _
                      " usercode = " & Vusercode & ", " & _
                      " In_Process = " & VIN_PROCESS & ", " & _
                      " REFPISNO = " & VTXTREFERENCEPIS & ", " & _
                      " lastupdate = " & VLastUpdate & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""

        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                      " SALES_ORIGIN = " & XSALES_ORIGIN & "," & _
                      " SI_TYPE = " & XSI_TYPE & "," & _
                      " PAY_CLASS = " & XPAY_CLASS & "," & _
                      " CHAR_YEAR = " & XCHAR_YEAR & "," & _
                      " CHAR_MONTH = " & XCHAR_MONTH & "," & _
                      " IS_SERIES = " & XIS_SERIES & "," & _
                      " TRACK_CODE = " & XTRACK_CODE & "" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""

        SQL_STATEMENT = "update PMIS_TdayTran set" & _
                      " trantype = " & VTXTTRANTYPE & "," & _
                      " trandate = " & VTXTTRANDATE & "," & _
                      " tranno = " & VTXTTRANNO & _
                      " where [TYPE] = 'A' AND trantype = '" & PREVORDTYPE & "' and tranno = '" & Null2String(RSORD_HD!TRANNO) & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""

        ShowSuccessFullyUpdated
    End If

    If AddorEdit = "ADD" Then
        If Left(txtTranNo.Text, 1) = "A" Then

        Else
            gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NEXTCUNTER & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' where [TYPE] = 'A' AND modul = " & VTXTTRANTYPE
        End If
    Else
        rsRefresh
        RSORD_HD.Find "Tranno = " & VTXTTRANNO
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                       " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                       " netinvamt = " & ORD_TOTINVAMT & _
                       " where [TYPE] = 'A' AND tranno = " & VTXTTRANNO & " and trantype = " & VTXTTRANTYPE
    End If

    rsRefresh
    RSORD_HD.Find "tranno = " & VTXTTRANNO
    cmdCancel.Value = True
    On Error GoTo Errorcode
    If AddorEdit = "ADD" Then
        Dim RSTDAYTRANDUP, rstdaytranDUp2              As ADODB.Recordset
        Dim RSPRS_HD                                   As ADODB.Recordset
        Dim rsPartMasClone                             As ADODB.Recordset
        Dim ISS_CNT                                    As Integer
        Dim VMACSTOCKNO                                As Double
        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select trantype,tranno from PMIS_TdayTran where [TYPE] = 'A' AND trantype = '" & COUNTERTYPE & "' and tranno = " & N2Str2Null(RSORD_HD!TRANNO), gconDMIS
        If RSTDAYTRANDUP.EOF And RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.Close
            Set RSPRS_HD = New ADODB.Recordset
            Set RSPRS_HD = gconDMIS.Execute("Select * from PMIS_vw_PRS where refpisno = '" & cboRefPRSNo.Text & "'")
            If Not RSPRS_HD.EOF And Not RSPRS_HD.BOF Then
                Set rstdaytranDUp2 = New ADODB.Recordset
                rstdaytranDUp2.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,tranuprice from PMIS_TdayTran where trantype = 'ARS' and tranno = " & N2Str2Null(RSPRS_HD!TRANNO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rstdaytranDUp2.EOF And Not rstdaytranDUp2.BOF Then
                    rstdaytranDUp2.MoveFirst
                    Do While Not rstdaytranDUp2.EOF
                        Set rsPartMasClone = New ADODB.Recordset
                        Set rsPartMasClone = gconDMIS.Execute("Select STOCKNO,ONHAND,NON_HARI,mac from PMIS_StockMas where TYPE = 'A' and STOCKNO = " & N2Str2Null(rstdaytranDUp2!STOCK_ORD))
                        If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
                            If N2Str2Zero(rsPartMasClone!ONHAND) > 0 Then
                                ISS_CNT = ISS_CNT + 1
                                '===================================
                                'updating code:     jaa - 09052008          - Include MAC upon saving of transaction
                                VMACSTOCKNO = N2Str2Zero(rsPartMasClone!MAC)
                                '===================================
                                RRTRANDATE = N2Str2Null(RSORD_HD!trandate)
                                RRTRANTYPE = "'" & COUNTERTYPE & "'"
                                RRTRANNO = N2Str2Null(RSORD_HD!TRANNO)
                                RRITEMNO = N2Str2Null(Format(ISS_CNT, "0000"))
                                RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!STOCK_ORD)
                                RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!STOCK_SUP)
                                'JAA - 01/23/2008 - FOR CHECKING AND VALIDATING AVAILABLE STOCK ONLY
                                If N2Str2Zero(rsPartMasClone!ONHAND) < N2Str2IntZero(rstdaytranDUp2!TRANQTY) Then
                                    MsgBox "Warning: Requested Quantity on " + N2Str2Null(rstdaytranDUp2!STOCK_ORD) + " is greater than available stock!" & vbCrLf & "System will default the available stock only.", vbInformation, "Requested Exceeds available stock on-hand"
                                    RRTRANQTY = N2Str2Zero(rsPartMasClone!ONHAND)
                                Else
                                    RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!TRANQTY)
                                End If
                                RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANUPRICE)
                                RRTRANUCOST = N2Str2Zero(rstdaytranDUp2!TRANUCOST)
                                RRIN_OUT = "'O'"

                                RRSTATUS = "'N'"
                                SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                                "(TYPE,mac,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out,NON_HARI)" & _
                                              " values ('A'," & VMACSTOCKNO & "," & RRTRANDATE & ", '" & COUNTERTYPE & "', " & RRTRANNO & "," & _
                                              " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                              " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                              " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & "," & N2Str2Null(rsPartMasClone!NON_HARI) & ")"
                                gconDMIS.Execute SQL_STATEMENT
                                NEW_LogAudit "A", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""
                            Else
                                MsgBox "Requested Accessories Code: " & Null2String(rstdaytranDUp2!STOCK_ORD) & " doesn't have Stock in your Master File", vbInformation, "Cannot Add Accessories!"
                                'EAP:090308: to refresh item no. and start to 0001
                                FillDetails
                            End If
                        Else
                            MsgBox "Requested Accessories Code: " & Null2String(rstdaytranDUp2!STOCK_ORD) & " is not yet active in your Master File", vbInformation, "Cannot Add Accessories!"
                            'EAP:090308: to refresh item no. and start to 0001
                            FillDetails
                        End If
                        rstdaytranDUp2.MoveNext
                    Loop
                End If
            End If
            cleargrid grdDetails
            FillDetails
            gconDMIS.Execute "update PMIS_Ord_Hd set" & _
                           " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                           " netinvamt = " & ORD_TOTINVAMT & _
                           " where [TYPE] = 'A' AND tranno = " & VTXTTRANNO & " and trantype = " & VTXTTRANTYPE

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
    If AddorEdit = "ADD" Then
        InsertAdvanceBill
    End If
    X_FillSearchGrid ""
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub



Private Sub Command_Click()
    If txtTranType = "RIV" Or txtTranType = "ADB" Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        SERVICEPISPRINTING
    ElseIf txtTranType = "CSH" Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        CSHPRINTING
    ElseIf txtTranType = "CHG" Then
        rptCustomerOrder.Formulas(1) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(2) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        CHGPRINTING
    ElseIf txtTranType = "DR" Then
         NEWDRPRINTING
    End If
End Sub

Private Sub Command1_Click()
    frmAC_CustomerSearch.Show 1
End Sub

Private Sub Command2_Click()
    cmdPISNum_Click
End Sub

Private Sub Command3_Click()
    If Module_Access(LOGID, "EDIT ACCESSORIES TRANSACTION AMOUNT", "SYSTEM") = False Then Exit Sub
    txtTranUPrice.Enabled = True
End Sub

Private Sub Command4_Click()
    If Module_Access(LOGID, "GENERATE NON INVOICE NUMBER", "DATA ENTRY") = False Then Exit Sub

    txtPRtranno.Visible = True
    txtPRtranno.SetFocus
    txtPRtranno.Locked = True
    Dim SQLTXT                                         As String
    Dim rsTMP                                          As New ADODB.Recordset
    Dim ISSCOUNTER                                     As Integer

    On Error GoTo Errorcode
    If txtTranType = "CSH" Then
        SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'CSH')  AND LEFT(TRANNO,1) = 'A'"
        SQLTXT = SQLTXT & "AND [TYPE] = 'A'"

    ElseIf txtTranType = "RIV" Then
        SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'RIV')  AND LEFT(TRANNO,1) = 'A'"
        SQLTXT = SQLTXT & "AND [TYPE] = 'A'"

    ElseIf txtTranType = "CHG" Then
        SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'CHG')  AND LEFT(TRANNO,1) = 'A'"
        SQLTXT = SQLTXT & "AND [TYPE] = 'A'"

    ElseIf txtTranType = "DR" Then

        SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'DR')  AND LEFT(TRANNO,1) = 'A'"
        SQLTXT = SQLTXT & "AND [TYPE] = 'A'"

    ElseIf txtTranType = "ADB" Then
        SQLTXT = "SELECT COUNT(*) AS BILANG FROM PMIS_vw_ISS_HISTORY WHERE (TRANTYPE = 'ADB')  AND LEFT(TRANNO,1) = 'A'"
        SQLTXT = SQLTXT & "AND [TYPE] = 'A'"
    End If

    Set rsTMP = gconDMIS.Execute(SQLTXT)
    If Not (rsTMP.BOF And rsTMP.EOF) Then
        ISSCOUNTER = NumericVal(rsTMP!BILANG)
    End If

    ISSCOUNTER = ISSCOUNTER + 1
    txtPRtranno.Text = "A" & Format(ISSCOUNTER, "00000")

    Set rsTMP = Nothing
Errorcode:
End Sub



Private Sub Command5_Click()
    picHPI.Visible = False
End Sub

Private Sub Command6_Click()
        rptCustomerOrder.Formulas(1) = ""
        rptCustomerOrder.Formulas(2) = ""

'        If txtTranType = "CSH" Then
'            Screen.MousePointer = 11
'            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHINV.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
'            Screen.MousePointer = 0
'        ElseIf txtTranType = "CHG" Then
'            Screen.MousePointer = 11
'            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGINV.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
'            Screen.MousePointer = 0
'        End If

    If txtTranType = "CSH" Then
            
             If NumericVal(txtDS1.Text) = 0 Then
                   PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHINV.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
             Else
                   PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHINV_DISC.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
             End If
                 
    ElseIf txtTranType = "CHG" Then
            
            If NumericVal(txtDS1.Text) = 0 Then
                   PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGINV.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Else
                   PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGINV_DISC.RPT", "({ord_hd.TYPE} = 'P' or {ord_hd.TYPE} = 'M' or {ord_hd.TYPE} = 'A') AND ({ord_hd.STATUS} = 'P' OR {ord_hd.STATUS} = 'B')  AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            
            End If
    End If
        
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text

    If Shift = 2 Then

    End If

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Picture1.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Parts Issuance)"
            '====================================================================
            If COUNTERTYPE = "CSH" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCESSORIES ISSUANCE COUNTER CASH")
            ElseIf COUNTERTYPE = "CHG" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCESSORIES ISSUANCE COUNTER CHARGE")
            ElseIf COUNTERTYPE = "DR" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCESSORIES DR OUT ISSUANCE")
            ElseIf COUNTERTYPE = "RIV" Then
                Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCESSORIES SERVICE ISSUANCE")
            Else
                Call frmALL_AuditInquiry.DisplayHistory(labID, "ACCESSORIES ADVANCE BILL DATA ENTRY")
            End If
            '====================================================================

        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            Picture1.Enabled = True
            fraDetails.Enabled = True
            txtPRtranno.Visible = False
        Case vbKeyF1
            If Picture1.Visible = False Then Command2.Value = True
        Case vbKeyF2
            If Command1.Visible = True And Command1.Enabled = True Then Command1.Value = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(RSORD_HD!Status) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(RSORD_HD!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                    picDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!Status) <> "P" And Null2String(RSORD_HD!Status) <> "C" And Null2String(RSORD_HD!Status) <> "B" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If
        Case vbKeyF5
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HD!Status) <> "P" And Null2String(RSORD_HD!Status) <> "C" And Null2String(RSORD_HD!Status) <> "B" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True Then cmdPost.Value = True
        Case vbKeyF12
            If Picture1.Visible = True And (labSJ = "" And labORNo = "") Then
                If Null2String(RSORD_HD!Status) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPost", LOCALACESS) = False Then Exit Sub
                    'EAP:042209
                    'MsgCritical ("Unposting of this transaction will remove issuance of Accessories in CarService")
                    MsgBox "Unposting of this transaction will remove issuance of Accessories in CarService", vbCritical, "Critical"

                    If MsgQuestionBox("Are you sure you want to UnPost this Transaction?", "UnPost Transaction") = True Then
                        'EAP:042209 Remove issuance of Parts in CarService when unposting
                        Dim col1 As Integer, Col2 As Integer, col4 As Integer, lRow As Integer
                        Dim refRivAdb                  As String
                        Dim i                          As Integer

                        col1 = 1                      'itemno
                        Col2 = 2                      'partnumbe
                        col4 = 4                      'qty
                        lRow = grdDetails.Rows - 1

                        For i = 1 To lRow
                            refRivAdb = "'RIV" & Format(Null2String(txtTranNo), "000000") & Format(Null2String(grdDetails.TextMatrix(i, col1)), "000") & "'"
                            gconDMIS.Execute (" delete from csms_ro_det where rep_or = '" & txtRONO.Text & "' and livil = 4 and detcde =  '" & grdDetails.TextMatrix(i, Col2) & "' and detvol = '" & grdDetails.TextMatrix(i, col4) & "' and ref_riv_adb = " & refRivAdb & " ")
                        Next



                        Dim PCURONHAND, PCURTISSQTY, PCURISSUANCES As Integer
                        Dim RSTDAYTRANDUP, RSPARTMASDUP As ADODB.Recordset

                        Set RSTDAYTRANDUP = New ADODB.Recordset
                        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_TdayTran where [TYPE] = 'A' AND tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
                            RSTDAYTRANDUP.MoveFirst
                            Do While Not RSTDAYTRANDUP.EOF
                                Set RSPARTMASDUP = New ADODB.Recordset
                                RSPARTMASDUP.Open "select STOCKNO,onhand,tissqty,TISSQTY,issuances,REQSERVED,S_REQSERVED from PMIS_Accessories where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD) & " AND ACTIVE = 'Y'", gconDMIS
                                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                                    PCURONHAND = N2Str2IntZero(RSPARTMASDUP!ONHAND) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!TISSQTY) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    PCURISSUANCES = N2Str2IntZero(RSPARTMASDUP!ISSUANCES) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                                        SQL_STATEMENT = "update PMIS_Accessories set" & _
                                                      " REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                    Else
                                        SQL_STATEMENT = "update PMIS_Accessories set" & _
                                                      " S_REQSERVED = " & N2Str2IntZero(RSPARTMASDUP!S_REQSERVED) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY) & _
                                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                    End If
                                    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Accessories"), "", "TRAN NO: " & txtTranNo & " UNPOSTED", "", "")

                                    SQL_STATEMENT = "update PMIS_Accessories set" & _
                                                  " onhand = " & PCURONHAND & "," & _
                                                  " tissqty = " & PCURTISSQTY & "," & _
                                                  " issuances = " & PCURISSUANCES & "," & _
                                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                  " lastupdate = '" & LOGDATE & "'" & _
                                                  " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                    gconDMIS.Execute SQL_STATEMENT
                                    Call NEW_LogAudit("E", LOCALACESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Accessories"), "", "TRAN NO: " & txtTranNo & " UNPOSTED", "", "")

                                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                                  " status = 'N'," & _
                                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                  " lastupdate = '" & LOGDATE & "'" & _
                                                  " where id = " & RSTDAYTRANDUP!ID
                                    gconDMIS.Execute SQL_STATEMENT
                                    NEW_LogAudit "U", LOCALACESS, SQL_STATEMENT, labID, "Accessories", txtTranNo, COUNTERTYPE, ""

                                End If
                                RSTDAYTRANDUP.MoveNext
                            Loop
                        End If
                        SQL_STATEMENT = "update PMIS_Ord_Hd set" & _
                                      " status = 'N'," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where id = " & labID.Caption
                        gconDMIS.Execute SQL_STATEMENT
                        NEW_LogAudit "U", LOCALACESS, SQL_STATEMENT, labID, "Accessories", "TRAN NO: " & txtTranNo, COUNTERTYPE, ""

                        rsRefresh
                        On Error Resume Next
                        RSORD_HD.Find "id =" & labID.Caption
                        StoreMemVars
                    End If
                    Set RSTDAYTRANDUP = Nothing
                    Set RSPARTMASDUP = Nothing
                    
                    If txtTranType = "RIV" Then
                        If VALID_COMPANY_CODE(COMPANY_CODE) = True Then
                            If ImportDetails(txtRONO, "A", "4") = True Then
                                'do nothing
                            End If
                        Else
                            Call ImportAccessories(txtRONO)
                        End If
                    End If
                    
'                     'this importing is reference to return parts from service
'                     'update by:NVB
'                     If txtTranType = "RIV" Then
'                            If ImportDetails(txtRONO, "A", "4") = True Then
'                                'do nothing
'                            End If
'                    End If
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1: PMIS_ORDER_SHOW = True
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    textSearch.Text = ""
    If COUNTERTYPE = "RIV" Then
        LOCALACESS = "ACCESSORIES SERVICE ISSUANCE"
    ElseIf COUNTERTYPE = "ADB" Then
        LOCALACESS = "ACCESSORIES DR OUT SERVICE ISSUANCE"
    ElseIf COUNTERTYPE = "CSH" Then
        LOCALACESS = "ACCESSORIES ISSUANCE COUNTER CASH"
    ElseIf COUNTERTYPE = "CHG" Then
        LOCALACESS = "ACCESSORIES ISSUANCE COUNTER CHARGE"
    ElseIf COUNTERTYPE = "DR" Then
        LOCALACESS = "ACCESSORIES DR OUT ISSUANCE"
    ElseIf COUNTERTYPE = "ADB" Then
        LOCALACESS = "ACCESSORIES ADVANCE BILL DATA ENTRY"
    End If

    If COUNTERTYPE = "DR" Then cmdPISNum.Enabled = True
    If COUNTERTYPE = "CSH" Then optCASH.Value = True
    If COUNTERTYPE = "CHG" Then optCHARGE.Value = True

    If COUNTERTYPE <> "RIV" And COUNTERTYPE <> "ADB" Then
        Command1.Visible = True
        Command1.Enabled = True
        optRONo.Enabled = False
    Else

        Command1.Enabled = False
        Command1.Visible = False
    End If

    Frame1.Enabled = False
    Picture2.Visible = False
    initMemvars

    If LOGLEVEL = "ADM" Then
        'txtTranUPrice.Enabled = True
    Else
        If COUNTERTYPE = "ADB" Then
            'txtTranUPrice.Enabled = True
        Else
            txtTranUPrice.Enabled = False
        End If
    End If
    rsRefresh
    On Error Resume Next
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then RSORD_HD.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
    X_FillSearchGrid ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISTrans_CustomerOrder = Nothing
    UnloadForm Me
End Sub

Function CheckIfRoExists(XXX As String) As String
    Dim rsRO_DET                                       As ADODB.Recordset
    Set rsRO_DET = gconDMIS.Execute("SELECT REP_OR FROM CSMS_REPOR WHERE  REP_OR = " & N2Str2Null(XXX))
    If Not rsRO_DET.EOF And Not rsRO_DET.BOF Then
        CheckIfRoExists = UCase(Null2String(rsRO_DET!REP_OR))
    End If
    Set rsRO_DET = Nothing
End Function


Private Sub grdDetails_DblClick()
    Dim FILD                                           As String
    If Null2String(RSORD_HD!Status) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"
        MsgBoxXP "Transactions are Already Cancelled" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(RSORD_HD!Status) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"
        MsgBoxXP "Transactions are Already Billed-Out" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(RSORD_HD!Status) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBoxXP "Transactions are Already Posted" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        FILD = grdDetails.Text
        If FILD <> "" And FILD <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Visible = True
            BringToFront
            StorePartsEntry (FILD)
        Else
            MsgSpeechBox "No Entry of Accessories!"
            Exit Sub
        End If
    End If
End Sub

Private Sub lstOrd_Hd_GotFocus()
    On Error Resume Next
    lstOrd_Hd_ItemClick lstOrd_Hd.SelectedItem
End Sub

Private Sub optCASH_Click()
    COUNTERTYPE = "CSH"
End Sub

Private Sub optCHARGE_Click()
    COUNTERTYPE = "CHG"
End Sub

Private Sub Option1_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "CUSTOMER NAME"
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
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

Private Sub txtCustCode_Change()
    txtTerms = TERMS(txtCustCode)
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

Private Sub txtoverride_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ictr = 0 Then
        picoverride.Visible = False
        ichg = False
    Else
        If txtoverride.Text <> "ALONE" Then
            ictr = ictr - 1
            MsgBox "Invalid Code, You have ( " & ictr & " ) tries left!", vbInformation
            If ictr = 0 Then
                 picoverride.Visible = False
                 ichg = False
            Else
                txtoverride.Text = ""
                On Error Resume Next
                txtoverride.SetFocus
                ichg = False
                Exit Sub
            End If
        Else
            txtoverride.Text = ""
            picoverride.Visible = False
            ichg = True
            cmdSave.Value = True
        End If
    End If
ElseIf KeyAscii = 27 Then
    txtoverride.Text = ""
    picoverride.Visible = False
    ichg = False
    Exit Sub
End If

End Sub

Private Sub txtPRtranno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPRtranno.Visible = False
        txtTranNo.Text = txtPRtranno.Text
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
    RONOStr = txtRONO.Text
    If UCase(Left(RONOStr, 2)) = "R-" Then
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
        If COMPANY_CODE = "HAI" Then
        
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
    End If
    txtRONO.Text = RONOStr
    SetCustInfo (RONOStr)
End Sub

Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
    'updating code:     jaa - 10292008          - Transaction Month should be equal to current month
    If IsDate(txtTranDate) = True Then
        If DateDiff("m", txtTranDate, LOGDATE) <> 0 Then
            MsgBox "Warning: Transaction Month cannot be greater or less than the current month.", vbCritical
            txtTranDate.SetFocus
        End If
    End If
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranNo_LostFocus()
    txtTranNo = Format(txtTranNo, "000000")
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

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtTranUPrice_Change()
    If txtTranUPrice.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
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

Private Sub lstOrd_Hd_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    RSORD_HD.MoveFirst
    RSORD_HD.Find ("ID=" & ITEM.ListSubItems(1).Text)
    StoreMemVars
End Sub

Private Sub lstOrd_Hd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOrd_Hd
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

Private Sub lstOrd_Hd_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstOrd_Hd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        X_FillSearchGrid Format(textSearch.Text, "000000")
    ElseIf optRONo.Value = True Then
        Dim RONOStr                                    As String
        RONOStr = textSearch.Text
        If Left(RONOStr, 2) = "R-" Then
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
        X_FillSearchGrid RONOStr
    Else
        X_FillSearchGrid textSearch
    End If
End Sub

Sub X_FillSearchGrid(XXX As String)
    Dim RSORD_LIST                                     As ADODB.Recordset
    lstOrd_Hd.Sorted = False
    lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False

    On Error GoTo ERROR_MSG
    Set RSORD_LIST = New ADODB.Recordset
    XXX = Replace(LTrim(RTrim(XXX)), "'", "")

    If optTranno.Value = True Then
        Set RSORD_LIST = gconDMIS.Execute("SELECT TRANNO, ID FROM PMIS_ORD_HD WHERE [TYPE] = 'A' AND TRANTYPE = '" & txtTranType & "' AND TRANNO LIKE '" & XXX & "%' ORDER BY ID")
    ElseIf optRONo.Value = True Then
        Set RSORD_LIST = gconDMIS.Execute("SELECT RONO, ID FROM PMIS_ORD_HD WHERE [TYPE] = 'A' AND TRANTYPE = '" & txtTranType & "' AND RONO LIKE '" & XXX & "%'  ORDER BY TRANNO ASC")
    Else
        Set RSORD_LIST = gconDMIS.Execute("SELECT CUSTNAME, ID  FROM PMIS_ORD_HD WHERE [TYPE] = 'A' AND TRANTYPE = '" & txtTranType & "' AND CUSTNAME  LIKE '" & XXX & "%' ORDER BY CUSTNAME")
    End If

    If Not (RSORD_LIST.EOF And RSORD_LIST.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_LIST
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
    Exit Sub

ERROR_MSG:
    If err.Number = -2147217900 Then
        MsgBox "Kindly limit the use of (') character in searching", VBERROR, "Error"
    Else
        MsgBox err.Number & " : " & err.Description & ", Kindly report this error to Netspeed Helpdesk", VBERROR, "Error"
    End If
    
    err.Clear
End Sub
Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOrd_Hd.ListItems.Count > 0 And lstOrd_Hd.Enabled = True Then: lstOrd_Hd.SetFocus
    End If
End Sub

Private Sub optRONo_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "RO Number"
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optTranno_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "Tran. No."
    X_FillSearchGrid ""
    On Error Resume Next
    textSearch.SetFocus
End Sub

