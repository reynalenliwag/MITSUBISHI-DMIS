VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmPMISMAT_Receiving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Receiving Entry"
   ClientHeight    =   7935
   ClientLeft      =   855
   ClientTop       =   855
   ClientWidth     =   12315
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_Receiving.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   12315
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   12315
      TabIndex        =   117
      Top             =   7590
      Width           =   12315
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
         Left            =   3090
         TabIndex        =   120
         Top             =   0
         Width           =   9195
      End
      Begin VB.Label labAPJ 
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
         Left            =   930
         TabIndex        =   119
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " APJ #:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   118
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2760
      ScaleHeight     =   870
      ScaleWidth      =   9735
      TabIndex        =   76
      Top             =   6720
      Width           =   9735
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   8640
         MouseIcon       =   "MAT_Receiving.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   795
         Left            =   7860
         MouseIcon       =   "MAT_Receiving.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelRR 
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
         Left            =   7080
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_Receiving.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   6300
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_Receiving.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Unpost this Transaction"
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
         MouseIcon       =   "MAT_Receiving.frx":1B5D
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":1CAF
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   4740
         MouseIcon       =   "MAT_Receiving.frx":1FD4
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   3960
         MouseIcon       =   "MAT_Receiving.frx":2482
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":25D4
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
         Height          =   795
         Left            =   3180
         MouseIcon       =   "MAT_Receiving.frx":28E7
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":2A39
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
         Height          =   795
         Left            =   2400
         MouseIcon       =   "MAT_Receiving.frx":2D89
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":2EDB
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1620
         MouseIcon       =   "MAT_Receiving.frx":3239
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":338B
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   840
         MouseIcon       =   "MAT_Receiving.frx":3685
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":37D7
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   60
         MouseIcon       =   "MAT_Receiving.frx":3B2F
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":3C81
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   285
      Left            =   2220
      ScaleHeight     =   255
      ScaleWidth      =   9975
      TabIndex        =   67
      Top             =   6330
      Width           =   10005
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Mats."
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
         Left            =   120
         TabIndex        =   72
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Mats."
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
         Left            =   1740
         TabIndex        =   71
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Mats."
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
         Left            =   3360
         TabIndex        =   70
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label24 
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
         Index           =   0
         Left            =   5070
         TabIndex        =   69
         Top             =   30
         Width           =   1905
      End
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
         Left            =   7110
         TabIndex        =   68
         Top             =   30
         Width           =   2445
      End
   End
   Begin Crystal.CrystalReport rptReceiving 
      Left            =   2430
      Top             =   4860
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame fra_Search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7545
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   2115
      Begin VB.OptionButton optRRNo 
         Caption         =   "Transaction No."
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
         Left            =   120
         TabIndex        =   65
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "Sup. Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   64
         Top             =   630
         Width           =   1875
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   62
         Text            =   "TEXT"
         Top             =   960
         Width           =   1995
      End
      Begin MSComctlLib.ListView lstRR_HD 
         Height          =   6105
         Left            =   60
         TabIndex        =   63
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10769
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
         MouseIcon       =   "MAT_Receiving.frx":3FE0
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
      Begin VB.Label Label22 
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
         TabIndex        =   66
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10620
      ScaleHeight     =   855
      ScaleWidth      =   1590
      TabIndex        =   89
      Top             =   6720
      Width           =   1590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   780
         MouseIcon       =   "MAT_Receiving.frx":4142
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":4294
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   0
         MouseIcon       =   "MAT_Receiving.frx":45D2
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":4724
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   2130
      TabIndex        =   18
      Top             =   0
      Width           =   10125
      Begin VB.CommandButton cmdEditTrandate 
         Caption         =   "..."
         Height          =   315
         Left            =   5010
         TabIndex        =   116
         Top             =   210
         Width           =   285
      End
      Begin VB.ComboBox cboTerms 
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
         Left            =   3390
         TabIndex        =   7
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1050
         Width           =   1665
      End
      Begin VB.TextBox txtDS1 
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
         Height          =   375
         Left            =   5670
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox txtINVNo 
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
         Left            =   3810
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Type the Receiving Entry's Ref INV Number (e.g. 329874)"
         Top             =   2700
         Width           =   1155
      End
      Begin VB.TextBox txtDRNo 
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
         Left            =   1410
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "Type the Receiving Entry DR Number,if there's any  (e.g. 555665)"
         Top             =   2700
         Width           =   1155
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
         Height          =   1005
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "MAT_Receiving.frx":4A74
         ToolTipText     =   "Type your massage or remarks."
         Top             =   2010
         Width           =   4755
      End
      Begin VB.ComboBox cboClasscode 
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
         Left            =   5670
         TabIndex        =   4
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox txtRecvd_Code 
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
         Left            =   1350
         TabIndex        =   6
         ToolTipText     =   "Type the supplier's code (e.g. 00001) "
         Top             =   1020
         Width           =   1155
      End
      Begin VB.ComboBox cboRecvd_Desc 
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
         Left            =   90
         TabIndex        =   8
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1470
         Width           =   4965
      End
      Begin VB.TextBox txtRRNo 
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type Receiving entry number (e.g 003294)"
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtDS_Desc1 
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
         Left            =   6810
         TabIndex        =   13
         ToolTipText     =   "Input the type of the additional amount (e.g. VAT)"
         Top             =   1230
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtPODate 
         Height          =   345
         Left            =   3390
         TabIndex        =   3
         Top             =   615
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtRRDate 
         Height          =   345
         Left            =   3390
         TabIndex        =   1
         ToolTipText     =   "Type date of the receiving entry in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   90
         ScaleHeight     =   825
         ScaleWidth      =   5085
         TabIndex        =   31
         Top             =   1800
         Width           =   5085
         Begin VB.TextBox txtDetails 
            BackColor       =   &H00E0E0E0&
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
            Height          =   795
            Left            =   0
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   30
            Width           =   4965
         End
      End
      Begin VB.PictureBox Picture4 
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
         Height          =   1245
         Left            =   8400
         ScaleHeight     =   1245
         ScaleWidth      =   1545
         TabIndex        =   16
         Top             =   750
         Width           =   1545
         Begin VB.TextBox txtTTLRRAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            Left            =   30
            MaxLength       =   15
            TabIndex        =   55
            Top             =   90
            Width           =   1455
         End
         Begin VB.TextBox txtDS_Amt1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            MaxLength       =   15
            TabIndex        =   54
            Top             =   480
            Width           =   1425
         End
         Begin VB.TextBox txtNetRRAmt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
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
            MaxLength       =   15
            TabIndex        =   53
            Top             =   870
            Width           =   1455
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8670
         Top             =   2520
      End
      Begin MSMask.MaskEdBox txtRIV_Tranno 
         Height          =   345
         Left            =   5670
         TabIndex        =   5
         Top             =   810
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPONo 
         Height          =   345
         Left            =   1350
         TabIndex        =   2
         ToolTipText     =   "Type purchase order number of the receiving entry (e.g. 02774)"
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
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
         Height          =   225
         Index           =   2
         Left            =   2190
         TabIndex        =   58
         Top             =   -390
         Width           =   135
      End
      Begin VB.Label labRIV_TranNo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MIS #"
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
         Left            =   5100
         TabIndex        =   57
         Top             =   870
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   6390
         TabIndex        =   56
         Top             =   1260
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Left            =   6780
         TabIndex        =   25
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         Left            =   7260
         TabIndex        =   24
         Top             =   1680
         Width           =   1965
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref DR#"
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
         TabIndex        =   21
         Top             =   2730
         Width           =   795
      End
      Begin VB.Label Label8 
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
         Left            =   5190
         TabIndex        =   50
         Top             =   1770
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO"
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
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO Date"
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
         Left            =   2580
         TabIndex        =   20
         Top             =   690
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RR Number"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RR Date"
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
         Index           =   0
         Left            =   2580
         TabIndex        =   29
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
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
         Index           =   0
         Left            =   5640
         TabIndex        =   28
         Top             =   150
         Width           =   1305
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
         Height          =   225
         Left            =   2730
         TabIndex        =   27
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Receive From"
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
         TabIndex        =   26
         Top             =   1050
         Width           =   1275
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
         Height          =   225
         Left            =   3660
         TabIndex        =   23
         Top             =   1470
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref INV#"
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
         Left            =   2820
         TabIndex        =   19
         Top             =   2730
         Width           =   855
      End
      Begin VB.Label labRRsted 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLED INVOICE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   7080
         TabIndex        =   51
         Top             =   150
         Width           =   2955
      End
   End
   Begin VB.Frame fraDetails 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3225
      Left            =   2160
      TabIndex        =   17
      Top             =   3060
      Width           =   10125
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2985
         Left            =   60
         TabIndex        =   15
         Top             =   180
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   9
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
   Begin wizButton.cmd cmdAddTran 
      Height          =   3795
      Left            =   3450
      TabIndex        =   52
      Top             =   1470
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   6694
      TX              =   "cmd1"
      ENAB            =   0   'False
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
      MICON           =   "MAT_Receiving.frx":4A8E
   End
   Begin VB.PictureBox picPost 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   4230
      ScaleHeight     =   4845
      ScaleWidth      =   3825
      TabIndex        =   121
      Top             =   1545
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "x"
         Height          =   315
         Left            =   3480
         TabIndex        =   122
         Top             =   30
         Width           =   315
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   36
         Left            =   1980
         TabIndex        =   162
         Top             =   4575
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   35
         Left            =   1980
         TabIndex        =   161
         Top             =   4350
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   34
         Left            =   1980
         TabIndex        =   160
         Top             =   4140
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   33
         Left            =   1980
         TabIndex        =   159
         Top             =   3915
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   32
         Left            =   1980
         TabIndex        =   158
         Top             =   3705
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   31
         Left            =   1980
         TabIndex        =   157
         Top             =   3495
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   30
         Left            =   1980
         TabIndex        =   156
         Top             =   3270
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   29
         Left            =   1980
         TabIndex        =   155
         Top             =   3060
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label39 
         Caption         =   "Label39"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   1050
         TabIndex        =   154
         Top             =   390
         Width           =   2805
      End
      Begin VB.Label Label37 
         Caption         =   "Part No.: "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   153
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   28
         Left            =   1980
         TabIndex        =   152
         Top             =   2850
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   27
         Left            =   1980
         TabIndex        =   151
         Top             =   2640
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   26
         Left            =   1980
         TabIndex        =   150
         Top             =   2415
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   25
         Left            =   1980
         TabIndex        =   149
         Top             =   2205
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   1980
         TabIndex        =   148
         Top             =   1980
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   1980
         TabIndex        =   147
         Top             =   1770
         Visible         =   0   'False
         Width           =   1725
      End
      Begin XtremeShortcutBar.ShortcutCaption SC_RefTransNo 
         Height          =   375
         Left            =   0
         TabIndex        =   146
         Top             =   0
         Width           =   4215
         _Version        =   655364
         _ExtentX        =   7435
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ref. Transaction No(s)."
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorDark=   16711680
         ForeColor       =   16777215
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   145
         Top             =   690
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   144
         Top             =   930
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   143
         Top             =   1170
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   142
         Top             =   1410
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   141
         Top             =   1650
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   140
         Top             =   1890
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   139
         Top             =   2130
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   138
         Top             =   2370
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   137
         Top             =   2610
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   90
         TabIndex        =   136
         Top             =   2850
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   135
         Top             =   3090
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   90
         TabIndex        =   134
         Top             =   3300
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   90
         TabIndex        =   133
         Top             =   3510
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   132
         Top             =   3735
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   131
         Top             =   3945
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   90
         TabIndex        =   130
         Top             =   4170
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   90
         TabIndex        =   129
         Top             =   4380
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   90
         TabIndex        =   128
         Top             =   4590
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   1980
         TabIndex        =   127
         Top             =   705
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   1980
         TabIndex        =   126
         Top             =   915
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   20
         Left            =   1980
         TabIndex        =   125
         Top             =   1140
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   21
         Left            =   1980
         TabIndex        =   124
         Top             =   1350
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Label36"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   22
         Left            =   1980
         TabIndex        =   123
         Top             =   1560
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin VB.Frame fraAddTran 
      Caption         =   "Add/Edit Materials"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   3510
      TabIndex        =   32
      Top             =   1530
      Width           =   7485
      Begin VB.ComboBox cboPONO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   990
         Width           =   1725
      End
      Begin VB.CheckBox chkReceivedFromPO 
         Caption         =   "Received from PO"
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
         Height          =   285
         Left            =   120
         TabIndex        =   42
         Top             =   1050
         Width           =   1905
      End
      Begin VB.TextBox cboTranDescription 
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         MaxLength       =   15
         TabIndex        =   115
         Top             =   1770
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Height          =   825
         Left            =   3090
         MouseIcon       =   "MAT_Receiving.frx":4AAA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":4BFC
         Style           =   1  'Graphical
         TabIndex        =   114
         ToolTipText     =   "Add Materials"
         Top             =   2760
         Width           =   765
      End
      Begin VB.Frame fraUpdateMaster 
         Caption         =   "View for Master File Update"
         Height          =   2085
         Left            =   3930
         TabIndex        =   92
         Top             =   450
         Width           =   3375
         Begin VB.TextBox txtOldMAC 
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
            Left            =   630
            TabIndex        =   105
            Text            =   "0.00"
            Top             =   540
            Width           =   1260
         End
         Begin VB.TextBox txtOldDNP 
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
            Left            =   630
            TabIndex        =   104
            Text            =   "0.00"
            Top             =   900
            Width           =   1260
         End
         Begin VB.CheckBox chkUpdateSRP 
            Caption         =   "Update SRP"
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
            Height          =   195
            Left            =   5160
            TabIndex        =   103
            Top             =   1080
            Width           =   1485
         End
         Begin VB.CheckBox chkUpdateMAC 
            Caption         =   "Update MAC"
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
            Height          =   195
            Left            =   5160
            TabIndex        =   102
            Top             =   540
            Width           =   1485
         End
         Begin VB.CheckBox chkUpdateDNP 
            Caption         =   "Update DNP"
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
            Height          =   195
            Left            =   5160
            TabIndex        =   101
            Top             =   810
            Width           =   1485
         End
         Begin VB.TextBox txtOldSRP 
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
            Left            =   630
            TabIndex        =   100
            Text            =   "0.00"
            Top             =   1260
            Width           =   1260
         End
         Begin VB.TextBox txtOldOH 
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
            Left            =   630
            TabIndex        =   99
            Text            =   "0.00"
            Top             =   1620
            Width           =   1260
         End
         Begin VB.TextBox txtNewMAC 
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
            Left            =   1980
            TabIndex        =   98
            Text            =   "0.00"
            Top             =   540
            Width           =   1260
         End
         Begin VB.TextBox txtNewDNP 
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
            Left            =   1980
            TabIndex        =   97
            Text            =   "0.00"
            Top             =   900
            Width           =   1260
         End
         Begin VB.TextBox txtNewSRP 
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
            Left            =   1980
            TabIndex        =   96
            Text            =   "0.00"
            Top             =   1260
            Width           =   1260
         End
         Begin VB.TextBox txtNewOH 
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
            Left            =   1980
            TabIndex        =   95
            Text            =   "0.00"
            Top             =   1620
            Width           =   1260
         End
         Begin VB.CommandButton cmdOKUpdate 
            Caption         =   "&OK"
            Enabled         =   0   'False
            Height          =   555
            Left            =   3675
            MouseIcon       =   "MAT_Receiving.frx":5465
            MousePointer    =   99  'Custom
            Picture         =   "MAT_Receiving.frx":55B7
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   1380
            Width           =   555
         End
         Begin VB.CheckBox chkHARI_PARTS 
            Caption         =   "HARI PARTS"
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
            Height          =   210
            Left            =   120
            TabIndex        =   93
            Top             =   2400
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "OH"
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
            Left            =   150
            TabIndex        =   112
            Top             =   1650
            Width           =   1125
         End
         Begin VB.Label Label15 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "NEW"
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
            Left            =   2130
            TabIndex        =   111
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label16 
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
            Height          =   405
            Left            =   1620
            TabIndex        =   110
            Top             =   3000
            Width           =   285
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "OLD"
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
            Left            =   750
            TabIndex        =   109
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label18 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MAC"
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
            Left            =   150
            TabIndex        =   108
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label Label19 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "DNP"
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
            Left            =   150
            TabIndex        =   107
            Top             =   930
            Width           =   1125
         End
         Begin VB.Label Label20 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "SRP"
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
            Left            =   150
            TabIndex        =   106
            Top             =   1290
            Width           =   1125
         End
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   6660
         MouseIcon       =   "MAT_Receiving.frx":5852
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":59A4
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Delete Entry"
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdTranCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   5940
         MouseIcon       =   "MAT_Receiving.frx":5CCF
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":5E21
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Cancel Entry"
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox txtTranTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   47
         Top             =   3240
         Width           =   1515
      End
      Begin VB.TextBox txtUnitCost 
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   45
         Top             =   2520
         Width           =   1515
      End
      Begin VB.TextBox txtTranINVAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   46
         Top             =   2880
         Width           =   1515
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   44
         Top             =   2160
         Width           =   885
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
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   885
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
         TabIndex        =   41
         Top             =   600
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
         Left            =   1590
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   600
         Width           =   585
      End
      Begin wizButton.cmd cmdUpdateMaster 
         Height          =   2205
         Left            =   3870
         TabIndex        =   113
         Top             =   390
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3889
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
         MICON           =   "MAT_Receiving.frx":615F
      End
      Begin VB.CommandButton cmdTranSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   5220
         MouseIcon       =   "MAT_Receiving.frx":617B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Receiving.frx":62CD
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Save Entry"
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   150
         TabIndex        =   33
         Top             =   3240
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
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
         Left            =   480
         TabIndex        =   49
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label labDetID 
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
         Height          =   405
         Left            =   7260
         TabIndex        =   39
         Top             =   4050
         Width           =   285
      End
      Begin VB.Label Label30 
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
         Height          =   255
         Left            =   210
         TabIndex        =   38
         Top             =   2880
         Width           =   1245
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
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
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
         TabIndex        =   36
         Top             =   630
         Width           =   1365
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
         TabIndex        =   35
         Top             =   270
         Width           =   885
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
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1500
         Width           =   1125
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00C0C000&
      DrawMode        =   7  'Invert
      FillStyle       =   0  'Solid
      Height          =   4245
      Left            =   3000
      ScaleHeight     =   4185
      ScaleWidth      =   6255
      TabIndex        =   163
      Top             =   1965
      Visible         =   0   'False
      Width           =   6315
      Begin XtremeReportControl.ReportControl lstRefTransNo 
         Height          =   3885
         Left            =   30
         TabIndex        =   164
         Top             =   30
         Visible         =   0   'False
         Width           =   6195
         _Version        =   655364
         _ExtentX        =   10927
         _ExtentY        =   6853
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Press Esc to Exit "
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
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   165
         Top             =   3930
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      Caption         =   "- required field"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   10140
      TabIndex        =   60
      Top             =   8130
      Width           =   1305
   End
   Begin VB.Label Label2 
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
      Height          =   225
      Index           =   1
      Left            =   9480
      TabIndex        =   59
      Top             =   8010
      Width           =   135
   End
End
Attribute VB_Name = "frmPMISMAT_Receiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRR_HD, RSPO_HD, rsTdayTran                                      As ADODB.Recordset
Dim RSPARTMAS, RSSUPPLIER                                             As ADODB.Recordset
Dim RSCUNTER                                                          As ADODB.Recordset
Dim PCNT                                                              As Integer
Dim ADDOREDIT                                                         As String
Dim RR_TOTUCOST, RR_TOTINVAMT, RR_TOTVAT                              As Double
Dim RR_QTY_REC                                                        As Long
Dim PREVRRNO                                                          As String
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasSRP                             As Double
Dim PrevPmasOnHand                                                    As Integer
Dim NewPmasMAC, NewPmasDNP, NewPmasSRP                                As Double
Dim NewPmasOnHand, PrevTranQty                                        As Integer
Dim ISNONVAT                                                          As Boolean
Dim MODULE_STOCK_TYPE                                                 As String

Function GetRecClassCode(xxx)
    Select Case xxx
        Case "IBT": GetRecClassCode = "INTER BRANCH TRANSFER"
        Case "PCG": GetRecClassCode = "PURCHASED CHARGE"
        Case "PCS": GetRecClassCode = "PURCHASED CASH"
        Case "RCG": GetRecClassCode = "RETURN FROM CHARGE"
        Case "RCS": GetRecClassCode = "RETURN FROM CASH"
        Case "REP": GetRecClassCode = "REPLACEMENT"
        Case "RRV": GetRecClassCode = "RETURNED FROM SERVICE"
    End Select

End Function

Function GetRecClassification(xxx)
    Select Case xxx
        Case "INTER BRANCH TRANSFER": GetRecClassification = "IBT"
        Case "PURCHASED CHARGE": GetRecClassification = "PCG"
        Case "PURCHASED CASH": GetRecClassification = "PCS"
        Case "RETURN FROM CHARGE": GetRecClassification = "RCG"
        Case "RETURN FROM CASH": GetRecClassification = "RCS"
        Case "REPLACEMENT": GetRecClassification = "REP"
        Case "RETURNED FROM SERVICE": GetRecClassification = "RRV"

    End Select

End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from CSMS_MATMAS where ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from CSMS_MATMAS where STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select STOCKNO,mac from CSMS_MATMAS where TYPE = " & MODULE_STOCK_TYPE & " AND STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetPartPrice = Null2String(RSPARTMAS!Mac)
        End If
    End If
End Function

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC from CSMS_MATMAS where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function

Function SetSTOCKDESC2(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select id,STOCKDESC from CSMS_MATMAS where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
        End If
    End If
End Function

Function SetSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from CSMS_MATMAS where id = " & DDD, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
    End If
End Function

Function SetSupCode(nnn As String)
    Set RSSUPPLIER = New ADODB.Recordset
    RSSUPPLIER.Open "Select supname,supcode,sup_addrs,vat_percnt,NONVAT,Terms from PMIS_vw_Supplier where supname = '" & nnn & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSUPPLIER.EOF And Not RSSUPPLIER.BOF Then
        SetSupCode = Null2String(RSSUPPLIER!SupCode)
        txtDetails.Text = Null2String(RSSUPPLIER!sup_addrs)
        cboTerms.Text = Null2String(RSSUPPLIER!Terms)
        If Null2String(RSSUPPLIER!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(RSSUPPLIER!vat_percnt)
        End If
    Else
        txtDetails.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
        cboTerms.Text = ""
    End If
End Function

Function SetSupdesc(ppp As String)
    Set RSSUPPLIER = New ADODB.Recordset
    RSSUPPLIER.Open "Select supcode,supname,sup_addrs,vat_percnt,NONVAT,Terms from PMIS_vw_Supplier where supcode = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSUPPLIER.EOF And Not RSSUPPLIER.BOF Then
        SetSupdesc = Null2String(RSSUPPLIER!supname)
        txtDetails.Text = Null2String(RSSUPPLIER!sup_addrs)
        cboTerms.Text = Null2String(RSSUPPLIER!Terms)
        If Null2String(RSSUPPLIER!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(RSSUPPLIER!vat_percnt)
        End If
    Else
        txtDetails.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
        cboTerms.Text = ""
    End If
End Function

Function SetSupTerms(ppp As String) As Double
    Set RSSUPPLIER = New ADODB.Recordset
    RSSUPPLIER.Open "Select supcode,supname,sup_addrs,vat_percnt,NONVAT,TERMS from PMIS_vw_Supplier where supcode = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSUPPLIER.EOF And Not RSSUPPLIER.BOF Then
        SetSupTerms = N2Str2Zero(RSSUPPLIER!Terms)
        txtDetails.Text = Null2String(RSSUPPLIER!sup_addrs)
        If Null2String(RSSUPPLIER!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(RSSUPPLIER!vat_percnt)
        End If
    Else
        txtDetails.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
        SetSupTerms = ""
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        labDetID.Caption = rsTdayTran!ID
        txtTranItemNo.Text = Format(Null2String(rsTdayTran!itemno), "0000")
        cboTranPartNo.Text = Null2String(rsTdayTran!STOCK_ORD)
        cboTranDescription.Text = SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP))
        txtTranQty.Text = N2Str2IntZero(rsTdayTran!tranqty)
        txtTranINVAmt.Text = N2Str2Zero(rsTdayTran!TRANINVAMT)
        txtUnitCost.Text = N2Str2Zero(rsTdayTran!TRANUCOST)
        txtTranTotalAmt.Text = N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANINVAMT)
    End If
End Function

Sub BringToFront()
    cmdAddTran.ZOrder 0
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    Picture1.Enabled = False
    fra_Search.Enabled = False
End Sub

Sub FillCboRecvd()
    Set RSSUPPLIER = New ADODB.Recordset
    RSSUPPLIER.Open "select supname from PMIS_vw_Supplier ORDER BY SUPNAME ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSSUPPLIER.EOF And Not RSSUPPLIER.BOF Then
        RSSUPPLIER.MoveFirst
        cboRecvd_Desc.Clear
        Do While Not RSSUPPLIER.EOF
            cboRecvd_Desc.AddItem Null2String(RSSUPPLIER!supname)
            RSSUPPLIER.MoveNext
        Loop
    End If
End Sub

Sub FillDetails()
    On Error GoTo ERRORCODE
    Dim ALL_VERIFIED                                                  As Boolean
    PCNT = 0: RR_TOTUCOST = 0: RR_TOTINVAMT = 0: RR_TOTVAT = 0: RR_QTY_REC = 0
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt,tremarks from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        Screen.MousePointer = 11: rsTdayTran.MoveFirst: If N2Str2Null(rsRR_HD!STATUS) = "N" Then cmdPost.Enabled = False: ALL_VERIFIED = True
        Do While Not rsTdayTran.EOF
            PCNT = PCNT + 1
            ' grdDetails.AddItem rsTdayTran!ID & Chr(9) & Format(Null2String(rsTdayTran!itemno), "0000") & Chr(9) & _
              '                    Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
              '                    SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP)) & Chr(9) & _
              '                    N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
              '                    ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANINVAMT)) & Chr(9) & _
              '                    ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUCOST)) & Chr(9) & _
              '                    ToDoubleNumber(N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUCOST)) & Chr(9) & _
              '                    Null2String(rsTdayTran!TREMARKS)
            ' RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsTdayTran!tranqty)
            ' RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUCOST))
            ' RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANINVAMT))
            'Update By:EAP:07152008

            grdDetails.AddItem rsTdayTran!ID & Chr(9) & Format(Null2String(rsTdayTran!itemno), "0000") & Chr(9) & _
                               Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                               SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                               ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANINVAMT)) & Chr(9) & _
                               ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUCOST)) & Chr(9) & _
                               ToDoubleNumber(Round(NumericVal(N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUCOST)), 2)) & Chr(9) & _
                               Null2String(rsTdayTran!TREMARKS)
            RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsTdayTran!tranqty)
            RR_TOTUCOST = RR_TOTUCOST + Round(NumericVal(rsTdayTran!tranqty), 2) * Round(NumericVal(rsTdayTran!TRANUCOST), 2)
            RR_TOTINVAMT = RR_TOTINVAMT + Round(NumericVal(rsTdayTran!tranqty), 2) * Round(NumericVal(rsTdayTran!TRANINVAMT), 2)





            If N2Str2Null(rsRR_HD!STATUS) = "N" Then
                If Null2String(rsTdayTran!TREMARKS) <> "Verified" Then ALL_VERIFIED = False
            End If
            rsTdayTran.MoveNext
        Loop
        If N2Str2Null(rsRR_HD!STATUS) = "N" Then
            If ALL_VERIFIED = True Then cmdPost.Enabled = True Else cmdPost.Enabled = False
        End If
        If PCNT <> 0 Then grdDetails.RemoveItem 1
        If Null2String(rsRR_HD!classcode) = "PCS" Or Null2String(rsRR_HD!classcode) = "PCG" Then
            If ISNONVAT = True Then
                RR_TOTVAT = 0
            Else
                RR_TOTVAT = RR_TOTINVAMT - (RR_TOTINVAMT / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
        Else
            RR_TOTVAT = 0
        End If
        RR_TOTUCOST = RR_TOTINVAMT - RR_TOTVAT
        If NumericVal(RR_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtDS_Amt1.Text = RR_TOTVAT
            txtNetRRAmt.Text = NumericVal(txtTTLRRAmt.Text) + NumericVal(txtDS_Amt1.Text)
        Else
            txtDS1.Text = 0
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = 0
            txtNetRRAmt.Text = NumericVal(txtTTLRRAmt.Text)
        End If
        txtDS_Amt1.Text = Format(txtDS_Amt1.Text, MAXIMUM_DIGIT)
        txtNetRRAmt.Text = Format(txtNetRRAmt.Text, MAXIMUM_DIGIT)
        Screen.MousePointer = 0
    Else
        cleargrid grdDetails
    End If
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Sub FillGrid()
    Dim rsRR_HD                                                       As ADODB.Recordset
    lstRR_HD.Enabled = False
    lstRR_HD.Sorted = False: lstRR_HD.ListItems.Clear
    Set rsRR_HD = New ADODB.Recordset
    Set rsRR_HD = gconDMIS.Execute("select rrno,ID from PMIS_RR_Hd WHERE [TYPE] = 'M' order by rrno asc")
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstRR_HD.Enabled = True: Listview_Loadval Me.lstRR_HD.ListItems, rsRR_HD: lstRR_HD.Refresh
    Else
        lstRR_HD.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsRR_HD                                                       As ADODB.Recordset
    lstRR_HD.Sorted = False: lstRR_HD.ListItems.Clear
    lstRR_HD.Enabled = False
    Set rsRR_HD = New ADODB.Recordset
    Set rsRR_HD = gconDMIS.Execute("select recvd_from, ID from PMIS_RR_Hd where [TYPE] = 'M'  order by rrno asc")
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstRR_HD.Enabled = True: Listview_Loadval Me.lstRR_HD.ListItems, rsRR_HD: lstRR_HD.Refresh
    Else
        lstRR_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsRR_HD                                                       As ADODB.Recordset
    lstRR_HD.Enabled = False
    lstRR_HD.Sorted = False: lstRR_HD.ListItems.Clear
    Set rsRR_HD = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsRR_HD = gconDMIS.Execute("select rrno, ID from PMIS_RR_Hd where [TYPE] = 'M' AND rrno like'" & xxx & "%'")
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstRR_HD.Enabled = True: Listview_Loadval Me.lstRR_HD.ListItems, rsRR_HD: lstRR_HD.Refresh
    Else
        lstRR_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsRR_HD                                                       As ADODB.Recordset
    lstRR_HD.Enabled = False
    lstRR_HD.Sorted = False: lstRR_HD.ListItems.Clear
    Set rsRR_HD = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsRR_HD = gconDMIS.Execute("select recvd_from, ID from PMIS_RR_Hd where [TYPE] = 'M' AND recvd_from like '" & xxx & "%' order by rrno asc")
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstRR_HD.Enabled = True: Listview_Loadval Me.lstRR_HD.ListItems, rsRR_HD: lstRR_HD.Refresh
    Else
        lstRR_HD.Enabled = False
    End If
End Sub

Sub FindDupRRno(DDD As String)
    rsRR_HD.Bookmark = rsFind(rsRR_HD.Clone, "rrno", Format(DDD, "000000")).Bookmark
    StoreMemvars
End Sub

Sub InitCbo()
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "select STOCKNO,STOCKDESC from CSMS_MATMAS order BY STOCKNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        RSPARTMAS.MoveFirst
        cboTranPartNo.Clear
        Do While Not RSPARTMAS.EOF
            cboTranPartNo.AddItem Null2String(RSPARTMAS!STOCKNO)
            RSPARTMAS.MoveNext
        Loop
    End If
End Sub

Sub InitCboClasscode()
    cboClasscode.Clear
    cboClasscode.AddItem "INTER BRANCH TRANSFER"
    cboClasscode.AddItem "PURCHASED CHARGE"
    cboClasscode.AddItem "PURCHASED CASH"
    cboClasscode.AddItem "RETURN FROM CHARGE"
    cboClasscode.AddItem "RETURN FROM CASH"
    cboClasscode.AddItem "REPLACEMENT"
    cboClasscode.AddItem "RETURNED FROM SERVICE"
    cboClasscode.Text = "PURCHASED CHARGE"
End Sub

Sub InitCboPayTerm()
    Dim rsPayTerm                                                     As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    Set rsPayTerm = gconDMIS.Execute("Select * from ALL_PayTerm order by ID ASC")
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        rsPayTerm.MoveFirst: cboTerms.Clear
        Do While Not rsPayTerm.EOF
            cboTerms.AddItem Null2String(rsPayTerm!No_Days)
            rsPayTerm.MoveNext
        Loop
    End If
End Sub

Sub InitGrid()
    With grdDetails
        .ColWidth(0) = 1
        .ColWidth(1) = 800
        .ColWidth(2) = 1500
        .ColAlignment(2) = 2
        .ColWidth(3) = 2300
        .ColWidth(4) = 500
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        .ColWidth(7) = 1400
        .ColWidth(8) = 800

        .Row = 0
        .Col = 1: .Text = "Item"
        .Col = 2: .Text = "Material Code"
        .Col = 3: .Text = "Description"
        .Col = 4: .Text = "QTY"
        .Col = 5: .Text = "Inv. Amt."
        .Col = 6: .Text = "Cost"
        .Col = 7: .Text = "Total Amt."
        .Col = 8: .Text = "Verified"
    End With
End Sub

Sub initMemvars()
    txtRRNo.Text = ""
    txtPONo.Text = ""
    labAPJ = "": labDetails = ""
    Set RSCUNTER = New ADODB.Recordset
    RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'M' AND modul = 'RR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
        txtRRNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
    End If
    txtRRDate.Text = LOGDATE
    cboClasscode.Text = ""
    txtRIV_Tranno.Text = ""
    txtRecvd_Code.Text = ""
    FillCboRecvd
    txtDetails.Text = ""
    cboTerms.Text = ""
    txtPODate.Text = ""
    txtDRNo.Text = ""
    txtINVNo.Text = ""
    txtTTLRRAmt.Text = ""
    txtDS1.Text = ""
    txtDS_Desc1.Text = ""
    txtDS_Amt1.Text = ""
    txtNetRRAmt.Text = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    labRRsted.Caption = ""
    cleargrid grdDetails
    InitGrid
    InitCbo
    InitCboClasscode
    InitParts
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(PCNT + 1, "0000")
    cboTranPartNo.Text = ""
    cboTranDescription.Text = ""
    txtTranQty.Text = 1
    txtUnitCost.Text = "0.00"
    txtTranINVAmt.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
End Sub

Sub rsRefresh()
    Set rsRR_HD = New ADODB.Recordset
    rsRR_HD.Open "select * from PMIS_RR_Hd WHERE [TYPE] = 'M' order by rrno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub Send2BackConfirm()
    Frame1.Enabled = False
    Picture1.Enabled = True
    fraDetails.Enabled = True
    txtOldMAC.Text = 0
    txtOldDNP.Text = 0
    txtOldSRP.Text = 0
    txtOldOH.Text = 0
    txtNewMAC.Text = 0
    txtNewDNP.Text = 0
    txtNewSRP.Text = 0
    txtNewOH.Text = 0
    chkUpdateMAC.Value = 1
    chkUpdateDNP.Value = 1
    chkUpdateSRP.Value = 1
End Sub

Sub Send2FrontConfirm()
    Frame1.Enabled = False
    Picture1.Enabled = False
    fraDetails.Enabled = False
    txtOldMAC.Text = 0
    txtOldDNP.Text = 0
    txtOldSRP.Text = 0
    txtOldOH.Text = 0
    txtNewMAC.Text = 0
    txtNewDNP.Text = 0
    txtNewSRP.Text = 0
    txtNewOH.Text = 0
    chkUpdateMAC.Value = 1
    chkUpdateDNP.Value = 1
    chkUpdateSRP.Value = 1
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1
    fraAddTran.ZOrder 1
    fraAddTran.Enabled = False
    Send2BackConfirm
    Picture1.Enabled = True
    fra_Search.Enabled = False
End Sub

Sub ShowStockDetails()
    txtOldMAC.Text = "0.00"
    txtOldDNP.Text = "0.00"
    txtOldSRP.Text = "0.00"
    txtOldOH.Text = " 0"
    txtNewMAC.Text = "0.00"
    txtNewDNP.Text = "0.00"
    txtNewSRP.Text = "0"
    txtNewOH.Text = "0"
    chkHARI_PARTS.Value = 0
    cmdTranSave.Enabled = False
    Screen.MousePointer = 0

    Dim rsPartMasClone                                                As ADODB.Recordset
    Set rsPartMasClone = New ADODB.Recordset
    rsPartMasClone.Open "select STOCKNO,tpoqty,onorder,mac,dnp,srp,onhand,NON_HARI from CSMS_MATMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS

    If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
        PrevPmasMAC = Format(NumericVal(rsPartMasClone!Mac), MAXIMUM_DIGIT)
        PrevPmasDNP = Format(NumericVal(rsPartMasClone!dnp), MAXIMUM_DIGIT)
        PrevPmasSRP = Format(NumericVal(rsPartMasClone!SRP), MAXIMUM_DIGIT)
        PrevPmasOnHand = N2Str2Zero(rsPartMasClone!ONHAND)

        If Null2String(rsPartMasClone!NON_HARI) = "Y" Then
            chkHARI_PARTS.Value = 0
        Else
            chkHARI_PARTS.Value = 1
        End If
        txtOldMAC.Text = Format(PrevPmasMAC, MAXIMUM_DIGIT)
        txtOldDNP.Text = Format(PrevPmasDNP, MAXIMUM_DIGIT)
        txtOldSRP.Text = Format(PrevPmasSRP, MAXIMUM_DIGIT)
        txtOldOH.Text = Format(PrevPmasOnHand, DIGIT_FORMAT)
        Screen.MousePointer = 0
    End If




End Sub

Sub StoreMemvars()
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        labID.Caption = rsRR_HD!ID
        labAPJ = "": labDetails = ""
        txtRRNo.Text = Null2String(rsRR_HD!RRNO)
        txtRRDate.Text = Null2String(rsRR_HD!RRDATE)
        cboClasscode.Text = GetRecClassCode(Null2String(rsRR_HD!classcode))
        txtRIV_Tranno.Text = Null2String(rsRR_HD!RIV_Tranno)
        txtRecvd_Code.Text = Null2String(rsRR_HD!recvd_code)
        cboRecvd_Desc.Text = Null2String(rsRR_HD!recvd_from)
        txtDetails.Text = Null2String(rsRR_HD!Address)
        cboTerms.Text = Null2String(rsRR_HD!Terms)
        txtPONo.Text = Null2String(rsRR_HD!PONO)
        txtPODate.Text = Null2String(rsRR_HD!PODATE)
        txtDRNo.Text = Null2String(rsRR_HD!drno)
        txtINVNo.Text = Null2String(rsRR_HD!invno)
        txtDS1.Text = N2Str2IntZero(rsRR_HD!ds1)
        txtDS_Desc1.Text = Null2String(rsRR_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsRR_HD!ds_amt1))
        txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsRR_HD!ttlrramt))
        txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsRR_HD!netrramt))
        txtRemarks.Text = Null2String(rsRR_HD!remarks)
        labAPJ = CheckAPJNum(Null2String(rsRR_HD!RRNO), "MI")
        If Null2String(rsRR_HD!STATUS) = "P" Then
            labRRsted.Visible = True
            labRRsted.Caption = "POSTED [" & Null2String(rsRR_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
            cmdCancelRR.Enabled = False
            'If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = True
            If LOGLEVEL = "ADM" Then cmdUnPost.Enabled = True
            If labAPJ <> "" Then
                labDetails = "TRANSACTION IMPORTED TO ACCOUNTING"
                cmdPost.Enabled = False
                cmdUnPost.Enabled = False
                cmdPrint.Enabled = False
                cmdCancelRR.Enabled = False
            End If
        ElseIf Null2String(rsRR_HD!STATUS) = "C" Then
            labRRsted.Visible = True
            labRRsted.Caption = "CANCELLED [" & Null2String(rsRR_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdPrint.Enabled = False
            cmdCancelRR.Enabled = False
            cmdUnPost.Enabled = False
        Else
            labRRsted.Visible = False
            labRRsted.Caption = ""
            cmdEdit.Enabled = True
            cmdPost.Enabled = True
            cmdPrint.Enabled = False
            If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = True
            cmdUnPost.Enabled = False
        End If
        cleargrid grdDetails
        FillDetails
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Private Sub cboClasscode_Change()
    If cboClasscode.Text = "RETURNED FROM SERVICE" Then
        labRIV_TranNo.Visible = True
        txtRIV_Tranno.Visible = True
    Else
        labRIV_TranNo.Visible = False
        txtRIV_Tranno.Visible = False
    End If
End Sub

Private Sub cboClasscode_Click()
    If cboClasscode.Text = "RETURNED FROM SERVICE" Then
        labRIV_TranNo.Visible = True
        txtRIV_Tranno.Visible = True
    Else
        labRIV_TranNo.Visible = False
        txtRIV_Tranno.Visible = False
    End If
End Sub

Private Sub cboclasscode_LostFocus()
    If cboClasscode.Text <> "" Then
        cboClasscode.Text = cboClasscode.Text
        If cboClasscode.Text = "RETURNED FROM SERVICE" Then
            labRIV_TranNo.Visible = True
            txtRIV_Tranno.Visible = True
        Else
            labRIV_TranNo.Visible = False
            txtRIV_Tranno.Visible = False
        End If
    Else
        MsgBoxXP "Invalid code. Please Select Classification From The List... ", "Error Encountered", XP_OKOnly, msg_Information
    End If
End Sub

Private Sub cboPONO_Click()
    Dim rsPO_Details                                                  As ADODB.Recordset
    Set rsPO_Details = New ADODB.Recordset
    Set rsPO_Details = gconDMIS.Execute("Select * from PMIS_vw_ConfirmedPO where STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text) & " and PO_NO = " & N2Str2Null(cboPONO.Text))
    If Not rsPO_Details.EOF And Not rsPO_Details.BOF Then
        txtTranQty.Text = N2Str2Zero(rsPO_Details!Qty_Allocated)
        txtUnitCost.Text = N2Str2Zero(rsPO_Details!TRANUCOST)
    End If
End Sub

Private Sub cboRecvd_Desc_Click()
    txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DragDrop(Source As Control, X As Single, Y As Single)
    txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DropDown()
    txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_LostFocus()
    txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboTranDescription_Click()
    If cboTranDescription.Text <> "" Then
        txtPartID.Text = SetPartIDDesc(cboTranDescription.Text)
        cboTranPartNo.Text = SetSTOCKNO(txtPartID.Text)
        cboTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranDescription_LostFocus()
    cboTranDescription.Text = cboTranDescription.Text
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        cboTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        ShowStockDetails
    End If
End Sub

Private Sub cboTranPartNo_Click()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        cboTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
        ShowStockDetails
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    cboTranPartNo.Text = cboTranPartNo.Text
End Sub

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub chkReceivedFromPO_Click()
    If chkReceivedFromPO.Value = 1 Then
        cboPONO.Enabled = True
        cboPONO.BackColor = vbWhite
        Dim rsPO_Details                                              As ADODB.Recordset
        Set rsPO_Details = New ADODB.Recordset
        Set rsPO_Details = gconDMIS.Execute("Select * from PMIS_vw_ConfirmedPO where STOCK_ORD = " & N2Str2Null(cboTranPartNo.Text) & " order by PO_NO asc")
        If Not rsPO_Details.EOF And Not rsPO_Details.BOF Then
            rsPO_Details.MoveFirst: cboPONO.Clear
            Do While Not rsPO_Details.EOF
                cboPONO.AddItem Null2String(rsPO_Details!PO_NO)
                rsPO_Details.MoveNext
            Loop
        End If
    Else
        cboPONO.Enabled = False: cboPONO.Clear: cboPONO.BackColor = &HE0E0E0
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "MATERIALS RECEIVING") = False Then Exit Sub
    ADDOREDIT = "ADD"
    Frame1.Enabled = True
    fra_Search.Enabled = False
    Picture1.Visible = False
    cmdSave.Visible = True
    cmdCancel.Visible = True
    Picture2.Visible = True
    initMemvars
    txtRRDate.Enabled = False
    On Error Resume Next
End Sub

Private Sub cmdAddTran_Click()
    fra_Search.Enabled = False
    If Picture1.Visible = True Then
        SendToBack
        cmdAddTran.ZOrder 0
        fraAddTran.ZOrder 0
        cmdTranDelete.Enabled = False
        fraAddTran.Enabled = True
        ADDOREDIT = "ADD"
        InitParts
        On Error Resume Next
        cboTranPartNo.SetFocus
        Send2FrontConfirm
    End If
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    fra_Search.Enabled = True
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdCancelRR_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "MATERIALS RECEIVING") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
        Dim PCurOnOrder, PCurTRECQTY, PCurReceipts                    As Integer
        Dim PCurLast_recq                                             As Integer
        Dim rsTdaytranDup, rsPartmasDup                               As ADODB.Recordset
        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno,tranqty,STOCK_ORD,STATUS from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO), gconDMIS
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onorder,served,trecqty,receipts,last_recq,ONHAND from CSMS_MATMAS where TYPE = " & MODULE_STOCK_TYPE & " AND STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCurOnOrder = N2Str2IntZero(rsPartmasDup!onorder) + N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurTRECQTY = N2Str2IntZero(rsPartmasDup!trecqty) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurReceipts = N2Str2IntZero(rsPartmasDup!receipts) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurLast_recq = N2Str2IntZero(rsPartmasDup!last_recq) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    If Null2String(rsTdaytranDup!STATUS) = "P" Then
                        SQL_STATEMENT = "update CSMS_MATMAS set" & _
                                       " onorder = " & PCurOnOrder & "," & _
                                       " SERVED = " & N2Str2IntZero(rsPartmasDup!Served) - N2Str2IntZero(rsTdaytranDup!tranqty) & "," & _
                                       " ONHAND = " & N2Str2IntZero(rsPartmasDup!ONHAND) - N2Str2IntZero(rsTdaytranDup!tranqty) & "," & _
                                       " trecqty = " & PCurTRECQTY & "," & _
                                       " receipts = " & PCurReceipts & "," & _
                                       " last_recq = " & PCurLast_recq & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsTdaytranDup!STOCK_ORD), "STOCKNO", "CMS_MATMAS"), "", "RR NO: " & txtRRNo & " CANCEL", "", "")
                    End If
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
                       
        SQL_STATEMENT = "update PMIS_TdayTran set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where [TYPE] = 'M' AND tranno = " & N2Str2Null(rsRR_HD!RRNO) & " and trantype = 'RR'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "CC", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
        
        rsRefresh
        On Error Resume Next
        rsRR_HD.Find "id =" & labID.Caption
        StoreMemvars
    End If

    Exit Sub

ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS RECEIVING") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    grdDetails.Enabled = False
    PREVRRNO = Format(txtRRNo.Text, "000000")
    Frame1.Enabled = True
    fra_Search.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    txtRRDate.Enabled = False
End Sub

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", "MATERIALS RECEIVING") = False Then Exit Sub
    txtRRDate.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsRR_HD.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    rsRR_HD.MoveLast
    StoreMemvars
End Sub

Private Sub cmdNext_Click()
    rsRR_HD.MoveNext
    If rsRR_HD.EOF Then
        rsRR_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdOkUpdate_Click()

'****************************************************************************************************
'updating code: jaa - 10052008      - Disallow Updating of MAC,DNP,SRP (It should be done in Posting)

    If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 1 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                       " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                       " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 0 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                       " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 0 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 1 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                       " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 1 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 0 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
    If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 1 Then
        gconDMIS.Execute "update CSMS_MATMAS set" & _
                       " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                       " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                       " STOCKDESC = " & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & _
                       " where STOCKNO = " & N2Str2Null(cboTranPartNo.Text)
    End If
'****************************************************************************************************

    cleargrid grdDetails
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
        gconDMIS.Execute "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
    Else
        RR_TOTVAT = 0
        gconDMIS.Execute "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsRR_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True
    If ADDOREDIT = "ADD" Then cmdAddTran_Click
    Screen.MousePointer = 0
    Send2BackConfirm
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "MATERIALS RECEIVING") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    'updating code: JAA - 06272008     'Do not allow posting of transaction without issuance of Parts
    Dim fild                                                          As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    If fild = "" Or fild = "No Entry" Then
        MsgBox "Posting of Transaction cannot proceed. Pls. Add Material(s).", vbCritical, "Confirm Posting"
        Exit Sub
    End If
    '====================================================================================================

    Dim pmasOnOrder                                                   As Integer
    Dim pmasServed                                                    As Integer
    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        Set rsTdayTran = New ADODB.Recordset
        rsTdayTran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt,tranucost,mac from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
        If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
            rsTdayTran.MoveFirst
            Do While Not rsTdayTran.EOF
                If N2Str2Zero(rsTdayTran!TRANINVAMT) <= 0 Then
                    MsgSpeechBox "Transaction with Invoice Amount equal to Zero Encountered!"
                    Exit Sub
                End If
                rsTdayTran.MoveNext
            Loop
            rsTdayTran.MoveFirst
            Do While Not rsTdayTran.EOF
                Set RSPARTMAS = New ADODB.Recordset
                RSPARTMAS.Open "Select STOCKNO,onhand,trecqty,onorder,served,receipts,ACTIVE from CSMS_MATMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD), gconDMIS
                If Not RSPARTMAS.EOF And Not RSPARTMAS.EOF Then
                    pmasOnOrder = N2Str2Zero(RSPARTMAS!onorder)
                    pmasServed = N2Str2Zero(RSPARTMAS!Served)
                    If pmasOnOrder <= 0 Then pmasOnOrder = NumericVal(rsTdayTran!tranqty)
                    
                    '********************************************************************
                    'updating code: jaa - 10052008      - Update MAC,DNP,SRP upon Posting
                     If ISNONVAT = True Then
                    SQL_STATEMENT = "update CSMS_MATMAS set onhand = " & N2Str2Zero(RSPARTMAS!ONHAND) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(RSPARTMAS!trecqty) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " onorder = " & pmasOnOrder - NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " MAC = " & NumericVal(rsTdayTran!Mac) & ", " & _
                                   " SERVED = " & pmasServed + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(RSPARTMAS!receipts) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " last_recq = " & N2Str2Zero(rsTdayTran!tranqty) & ", " & _
                                   " last_recd = '" & LOGDATE & "', " & _
                                   " supcode = " & N2Str2Null(txtRecvd_Code.Text) & "," & _
                                   " dnp = '" & (N2Str2Zero(rsTdayTran!TRANUCOST)) & "'" & _
                                   " where STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
                    Else
                    SQL_STATEMENT = "update CSMS_MATMAS set onhand = " & N2Str2Zero(RSPARTMAS!ONHAND) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(RSPARTMAS!trecqty) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " onorder = " & pmasOnOrder - NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " MAC = " & NumericVal(rsTdayTran!Mac) & ", " & _
                                   " SERVED = " & pmasServed + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(RSPARTMAS!receipts) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                   " last_recq = " & N2Str2Zero(rsTdayTran!tranqty) & ", " & _
                                   " last_recd = '" & LOGDATE & "', " & _
                                   " supcode = " & N2Str2Null(txtRecvd_Code.Text) & "," & _
                                   " dnp = '" & (N2Str2Zero(rsTdayTran!TRANUCOST) * 1.12) & "'" & _
                                   " where STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO)

                     End If
                    
                     gconDMIS.Execute SQL_STATEMENT
                    
'                    SQL_STATEMENT = "update CSMS_MATMAS set onhand = " & N2Str2Zero(RSPARTMAS!ONHAND) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                   " trecqty = " & N2Str2Zero(RSPARTMAS!trecqty) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                   " onorder = " & pmasOnOrder - NumericVal(rsTdayTran!tranqty) & ", " & _
'                                   " MAC = " & NumericVal(rsTdayTran!Mac) & ", " & _
'                                   " SERVED = " & pmasServed + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                   " receipts = " & N2Str2Zero(RSPARTMAS!receipts) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                   " last_recq = " & N2Str2Zero(rsTdayTran!tranqty) & ", " & _
'                                   " last_recd = '" & LOGDATE & "', " & _
'                                   " supcode = " & N2Str2Null(txtRecvd_Code.Text) & _
'                                    ", dnp = '" & (Trim(rsTdayTran!TRANUCOST) * 1.12) & "' " & _
'                                   " where STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
'                    SQL_STATEMENT = "update CSMS_MATMAS set onhand = " & N2Str2Zero(rsPartMas!ONHAND) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                    " trecqty = " & N2Str2Zero(rsPartMas!trecqty) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                    " onorder = " & pmasOnOrder - NumericVal(rsTdayTran!tranqty) & ", " & _
'                                    " MAC = " & N2Str2IntZero(NewPmasMAC) & ", " & _
'                                    " DNP = " & N2Str2IntZero(NewPmasDNP) & ", " & _
'                                    " SRP = " & N2Str2IntZero(NewPmasSRP) & ", " & _
'                                    " SERVED = " & pmasServed + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                    " receipts = " & N2Str2Zero(rsPartMas!receipts) + NumericVal(rsTdayTran!tranqty) & ", " & _
'                                    " date_entered = '" & LOGDATE & "', " & _
'                                    " last_recq = " & N2Str2Zero(rsTdayTran!tranqty) & ", " & _
'                                    " last_recd = '" & LOGDATE & "', " & _
'                                    " supcode = " & N2Str2Null(txtRecvd_Code.Text) & _
                                    " where STOCKNO = " & N2Str2Null(rsPartMas!STOCKNO)
                    
                     '********************************************************************
                     
                    Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSPARTMAS!STOCKNO), "STOCKNO", "CSMS_MATMAS"), "", "RR NO: " & txtRRNo & " POSTED", "", "")
                    
                    If Null2String(RSPARTMAS!Active) = "N" Or Null2String(RSPARTMAS!Active) = "" Then
                        SQL_STATEMENT = "update CSMS_MATMAS set " & _
                                       " ACTIVE = 'Y'," & _
                                       " DATE_ENTERED = " & N2Date2Null(LOGDATE) & _
                                       " where STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
                        gconDMIS.Execute SQL_STATEMENT
                        Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSPARTMAS!STOCKNO), "STOCKNO", "CSMS_MATMAS"), "", "RR NO: " & txtRRNo & " POSTED ACTIVE", "", "")
                    End If
                    
                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                   " status = 'P'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdayTran!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "PP", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
                
                End If
                rsTdayTran.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
        
        rsRefresh
        On Error Resume Next
        rsRR_HD.Find "id =" & labID.Caption
        StoreMemvars
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdPrevious_Click()
    rsRR_HD.MovePrevious
    If rsRR_HD.BOF Then
        rsRR_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS RECEIVING") = False Then Exit Sub

    On Error GoTo ERRORCODE:
    If MsgQuestionBox("Receiving Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        rptReceiving.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReceiving.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptReceiving, PMIS_REPORT_PATH & "rrmat.rpt", "{rr_hd.type} = 'M' AND {rr_hd.rrno} = '" & txtRRNo.Text & "'", DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
        
        Call NEW_LogAudit("V", "MATERIALS RECEIVING", "", labID, "", "RR NO: " & txtRRNo, "", "")
    End If
    Exit Sub

ERRORCODE:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERRORCODE
    Dim rsRR_HDDup                                                    As ADODB.Recordset
    Dim xcboClasscode                                                 As String
    xcboClasscode = GetRecClassification(cboClasscode)
    'axp02232008
    If Len(Trim(RTrim(txtRRNo))) <> 6 Then
        MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
        On Error Resume Next
        txtRRNo.SetFocus
        Exit Sub
    End If


    If txtINVNo.Text = "" And txtDRNo.Text = "" Then
        MsgSpeechBox "Reference Invoice Number must be inputed!"
        On Error Resume Next
        txtINVNo.SetFocus
        Exit Sub
    End If
    If xcboClasscode = "PCG" Then
        If cboTerms.Text = "" Then
            MsgSpeechBox "Warning: Terms must be Inputed"
            On Error Resume Next
            cboTerms.SetFocus
            Exit Sub
        End If
    End If
    If txtRRDate.Text = "" Or IsDate(txtRRDate.Text) = False Then
        MsgSpeechBox "Invalid MRR Date!"
        On Error Resume Next
        txtRRDate.SetFocus
        Exit Sub
    End If

    'VALIDATION FOR TRANSACTION CODE
    If Trim(txtRRNo.Text) = "" Then
        MsgSpeechBox "MRR Number must not be empty"
        On Error Resume Next
        txtRRNo.SetFocus
        Exit Sub
    Else
        If ADDOREDIT = "ADD" Then
            Dim rsfindDup                                             As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select rrno from PMIS_RR_Hd where [TYPE] = 'M' AND rrno = '" & txtRRNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "MRR Number already exist!"
                On Error Resume Next
                txtRRNo.SetFocus
                Exit Sub
            End If
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select pono from PMIS_vw_RR_Trans where [TYPE] = 'M' AND pono = '" & txtPONo.Text & "' AND STATUS ='P'", gconDMIS
            If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
                MsgSpeechBox "Purchase Order Number Already Received"
                On Error Resume Next
                txtPONo.SetFocus
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtRRNo)) <> Null2String(rsRR_HD!RRNO) Then
                Set rsfindDup = New ADODB.Recordset
                rsfindDup.Open "select rrno from PMIS_RR_Hd where rrno = '" & txtRRNo.Text & "' and type='M'", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                    MsgSpeechBox "MRR Number already exist!"
                    On Error Resume Next
                    txtRRNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If



    Dim NewRRCunTer                                                   As String
    NewRRCunTer = NumericVal(txtRRNo.Text) + 1

    Dim VTXTRRNo, VTXTRRDate, Vcboclasscode                           As String
    Dim VTXTRecvd_Code, VTXTRecvd_From, VtxtAddress                   As String
    Dim Vcboterms, VTXTPONo, VTXTPODate                               As String
    Dim VTXTDRNo, VTXTINVNo                                           As String
    Dim VTXTTTLRRAmt, VTXTDS1                                         As Double
    Dim VTXTDS_Desc1                                                  As String
    Dim VTXTDS_Amt1, VTXTNetRRAmt                                     As Double
    Dim VTXTRemarks                                                   As String
    Dim VTXTRIV_Tranno                                                As String
    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE                              As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP                            As String
    Dim RRTRANQTY                                                     As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                                     As Double
    Dim RRIN_OUT, RRSTATUS                                            As String

    VTXTRRNo = N2Str2Null(txtRRNo.Text)
    VTXTRRDate = N2Date2Null(txtRRDate.Text)
    Vcboclasscode = N2Str2Null(xcboClasscode)
    VTXTRIV_Tranno = N2Str2Null(txtRIV_Tranno.Text)
    VTXTRecvd_Code = N2Str2Null(txtRecvd_Code.Text)
    VTXTRecvd_From = N2Str2Null(cboRecvd_Desc.Text)
    VtxtAddress = N2Str2Null(txtDetails.Text)
    Vcboterms = N2Str2Null(cboTerms.Text)
    VTXTPONo = N2Str2Null(txtPONo.Text)
    VTXTPODate = N2Date2Null(txtPODate.Text)
    VTXTDRNo = N2Str2Null(txtDRNo.Text)
    VTXTINVNo = N2Str2Null(txtINVNo.Text)
    VTXTTTLRRAmt = NumericVal(txtTTLRRAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetRRAmt = NumericVal(txtNetRRAmt.Text)
    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    End If

    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_RR_Hd" & _
                       " (TYPE,rrno,rrdate,classcode,RIV_Tranno,recvd_code,recvd_from,address,terms,pono,podate,drno,invno,ttlrramt,ds1,ds_desc1,ds_amt1,netrramt,usercode,lastupdate,remarks)" & _
                       " values ('M'," & VTXTRRNo & ", " & VTXTRRDate & ", " & Vcboclasscode & ", " & VTXTRIV_Tranno & _
                         ", " & VTXTRecvd_Code & ", " & VTXTRecvd_From & ", " & VtxtAddress & ", " & Vcboterms & _
                         ", " & VTXTPONo & ", " & VTXTPODate & ", " & VTXTDRNo & ", " & VTXTINVNo & _
                         ", " & VTXTTTLRRAmt & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetRRAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "MATERIALS RECEIVING", SQL_STATEMENT, FindTransactionID(txtRRNo, "rrno", "PMIS_RR_HD", "DETAILS", N2Str2Null("M"), "TYPE"), "Materials", txtRRNo & " - " & cboClasscode, "RR", ""
    Else
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " rrno = " & VTXTRRNo & "," & _
                       " rrdate = " & VTXTRRDate & "," & _
                       " classcode = " & Vcboclasscode & "," & _
                       " RIV_Tranno = " & VTXTRIV_Tranno & "," & _
                       " recvd_code = " & VTXTRecvd_Code & "," & _
                       " recvd_from = " & VTXTRecvd_From & "," & _
                       " address = " & VtxtAddress & "," & _
                       " terms = " & Vcboterms & "," & _
                       " pono = " & VTXTPONo & "," & _
                       " podate = " & VTXTPODate & "," & _
                       " drno = " & VTXTDRNo & "," & _
                       " invno = " & VTXTINVNo & "," & _
                       " ttlrramt = " & VTXTTTLRRAmt & "," & _
                       " ds1 = " & VTXTDS1 & "," & _
                       " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                       " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                       " netrramt = " & VTXTNetRRAmt & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " remarks = " & VTXTRemarks & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", txtRRNo & " - " & cboClasscode, "RR", ""
        
        SQL_STATEMENT = "update PMIS_TdayTran set" & _
                       " trandate = " & VTXTRRDate & "," & _
                       " tranno = " & VTXTRRNo & _
                       " where [TYPE] = 'M' AND trantype = 'RR' and tranno = '" & PREVRRNO & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", txtRRNo & " - " & cboClasscode, "RR", ""
        
    End If
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NewRRCunTer & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where [TYPE] = 'M' AND modul = 'RR'"
    End If
    rsRefresh
    On Error Resume Next
    rsRR_HD.Find "rrno = " & VTXTRRNo
    cmdCancel.Value = True
    On Error GoTo ERRORCODE
    If ADDOREDIT = "ADD" Then
        
        Dim rsTdaytranDup, rstdaytranDUp2                             As ADODB.Recordset
        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno from PMIS_ALLdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO), gconDMIS
        If rsTdaytranDup.EOF And rsTdaytranDup.BOF Then
            rsTdaytranDup.Close
            Set rstdaytranDUp2 = New ADODB.Recordset
            rstdaytranDUp2.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_AlldayTran where TYPE = 'M' and trantype = 'PO' and tranno = " & N2Str2Null(rsRR_HD!PONO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rstdaytranDUp2.EOF And Not rstdaytranDUp2.BOF Then
                rstdaytranDUp2.MoveFirst
                Do While Not rstdaytranDUp2.EOF
                    RRTRANDATE = N2Date2Null(txtRRDate.Text)
                    RRTRANTYPE = "'RR'"
                    RRTRANNO = N2Str2Null(rsRR_HD!RRNO)
                    RRITEMNO = Format(N2Str2Null(rstdaytranDUp2!itemno), "0000")
                    RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!STOCK_ORD)
                    RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!STOCK_SUP)
                    RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!tranqty)
                    RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANINVAMT)
                    RRTRANUCOST = N2Str2Zero(rstdaytranDUp2!TRANUCOST)
                    RRIN_OUT = "'I'"
                    RRSTATUS = "'N'"
                    
                    
                    '=================================================================================================================
                    'updating code:     jaa - 09062008            - To compute for NEW MAC, DNP and SRP whenever user Received from PO
                    If RECEIVED_FROM_PO = "YES" Then
                    
                        Dim rsPartMasClone                                                As ADODB.Recordset
                        Set rsPartMasClone = New ADODB.Recordset
                        rsPartMasClone.Open "select STOCKNO,tpoqty,onorder,mac,dnp,srp,onhand from CSMS_MATMAS where STOCKNO = " & N2Str2Null(RRSTOCK_ORD), gconDMIS
                        If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
                            
'                            If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
'                                If CheckIfNonVatSup(Trim(txtRecvd_Code.Text)) = False Then
'                                    RRTRANUCOST = RRTRANUCOST / ConvertToBIRDecimalFormat(VAT_RATE)
'                                End If
'                            End If
                            
                            PrevPmasMAC = FormatNumber(NumericVal(rsPartMasClone!Mac))
                            PrevPmasDNP = FormatNumber(NumericVal(rsPartMasClone!dnp))
                    
                            PrevPmasSRP = FormatNumber(NumericVal(rsPartMasClone!SRP))
                            PrevPmasOnHand = N2Str2Zero(rsPartMasClone!ONHAND)
                            NewPmasOnHand = RRTRANQTY
                    
                            'NewPmasDNP = RRTRANUCOST * ConvertToBIRDecimalFormat(VAT_RATE)
                            NewPmasDNP = RRTRANINVAMT
                            
                            If PrevPmasOnHand <= 0 Then
                                NewPmasMAC = Round((RRTRANUCOST * RRTRANQTY) / NewPmasOnHand, 2)
                            Else
                                NewPmasMAC = Round(((PrevPmasMAC * PrevPmasOnHand) + (RRTRANUCOST * RRTRANQTY)) / (NewPmasOnHand + PrevPmasOnHand), 2)
                            End If
                            NewPmasSRP = Format(PrevPmasSRP, MAXIMUM_DIGIT)
                                                        
                            
                            gconDMIS.Execute "Update CSMS_MATMAS set MAC = " & NewPmasMAC & ",DNP =" & NewPmasDNP & ",SRP = " & NewPmasSRP & " WHERE STOCKNO = " & N2Str2Null(RRSTOCK_ORD)
                            
                            
                            SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                            "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,MAC,traninvamt,lastupdate,usercode,status,in_out)" & _
                                            " values ('M'," & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                                            " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                            " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                            " " & RRTRANUCOST & "," & NewPmasMAC & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
                            gconDMIS.Execute SQL_STATEMENT
                         End If
                    '=================================================================================================================
                    Else
                        SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                                         "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                                       " values ('M'," & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                                       " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                       " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                       " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    NEW_LogAudit "A", "MATERIALS RECEIVING", SQL_STATEMENT, FindTransactionID(txtRRNo, "tranno", "PMIS_TdayTran", "DETAILS", N2Str2Null("M"), "TYPE"), "Materials", txtRRNo, "RR", ""
        
                    rstdaytranDUp2.MoveNext
                Loop
            End If
            cleargrid grdDetails
            FillDetails
            cmdAddTran_Click
        Else
            cleargrid grdDetails
            FillDetails
            cmdAddTran_Click
        End If
    End If

    cleargrid grdDetails
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
        gconDMIS.Execute "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
    Else
        RR_TOTVAT = 0
        gconDMIS.Execute "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
    End If
    
    If ADDOREDIT = "ADD" Then
          Picture1.Enabled = False
          fraDetails.Enabled = False
    
    Else
          Picture1.Enabled = True
          fraDetails.Enabled = True
          
    End If
    
    rsRefresh
    FillGrid
    On Error Resume Next
    rsRR_HD.Find "rrno = " & VTXTRRNo
    StoreMemvars
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTranCancel_Click()
    Picture1.Enabled = True
    fraDetails.Enabled = True
    SendToBack
    StoreMemvars
    'Update By : BTT
    fra_Search.Enabled = True
End Sub

Private Sub cmdTranDelete_Click()

    On Error GoTo ERRORCODE:

    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Materials Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_TdayTran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "MAT NO: " & cboTranPartNo, "RR", labDetID
        ShowDeletedMsg
    End If
    
    Dim cnt                                                           As Integer
    Dim rsTdaytranDup                                                 As ADODB.Recordset
    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "select id,itemno from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        cnt = 0
        Do While Not rsTdaytranDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update PMIS_TdayTran set itemno = " & Format(cnt, "0000") & " where id = " & rsTdaytranDup!ID
            rsTdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    Else
        RR_TOTVAT = 0
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    End If
    Call NEW_LogAudit("E", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRRNo & " DELETE DETAILS", "", "")
    
    rsRefresh
    On Error Resume Next
    rsRR_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    On Error GoTo ERRORCODE

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If
    If cboTranDescription.Text = "" Then
        MsgSpeechBox "Description must have a value"
        On Error Resume Next
        cboTranDescription.Enabled = True
        cboTranDescription.SetFocus
        Exit Sub
    End If
    If ADDOREDIT = "ADD" Then
        Dim rsTDaytranClone                                           As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_TdayTran where [TYPE] = 'M' AND STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = 'RR' and tranno =" & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Material Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Exit Sub
        End If
    End If

    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE                              As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP                            As String
    Dim RRTRANQTY                                                     As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                                     As Double
    Dim RRSTATUS, RRIN_OUT                                            As String
    Dim RRNEWMAC                                                      As Double
    Dim VTXTTREMARKS                                                  As String
    RRTRANDATE = N2Date2Null(txtRRDate.Text)
    RRTRANTYPE = "'RR'"
    RRTRANNO = N2Str2Null(txtRRNo.Text)
    RRITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    RRSTOCK_ORD = N2Str2Null(cboTranPartNo.Text)
    RRSTOCK_SUP = N2Str2Null(cboTranPartNo.Text)
    RRTRANQTY = NumericVal(txtTranQty.Text)
    RRTRANINVAMT = NumericVal(txtTranINVAmt.Text)
    RRTRANUCOST = NumericVal(txtUnitCost.Text)
    RRNEWMAC = NumericVal(txtNewMAC.Text)
            
    RRIN_OUT = "'I'"
    RRSTATUS = "'N'"
    VTXTTREMARKS = "'Verified'"
    Screen.MousePointer = 11
    If RRTRANINVAMT <= 0 Then
        MsgSpeechBox "Warning: Invoice Amount must not be zero"
        Screen.MousePointer = 0
        Exit Sub
    End If

    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_TdayTran " & _
                         "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,mac,traninvamt,lastupdate,usercode,status,in_out,TRemarks)" & _
                       " values ('M'," & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                       " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                       " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                       " " & RRTRANUCOST & "," & RRNEWMAC & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & "," & VTXTTREMARKS & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "MAT NO: " & cboTranPartNo, "RR", ""
        
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update PMIS_TdayTran set" & _
                       " trandate = " & RRTRANDATE & "," & _
                       " trantype = " & RRTRANTYPE & "," & _
                       " tranno = " & RRTRANNO & "," & _
                       " itemno = " & RRITEMNO & "," & _
                       " STOCK_ORD = " & RRSTOCK_ORD & "," & _
                       " STOCK_SUP = " & RRSTOCK_SUP & "," & _
                       " tranqty = " & RRTRANQTY & "," & _
                       " tranucost = " & RRTRANUCOST & "," & _
                       " MAC= " & RRNEWMAC & "," & _
                       " traninvamt = " & RRTRANINVAMT & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " status = " & RRSTATUS & "," & _
                       " in_out = " & RRIN_OUT & "," & _
                       " TREMARKS = " & VTXTTREMARKS & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", txtRRNo, "RR", labDetID
        
        ShowSuccessFullyUpdated
    End If
    
    cleargrid grdDetails
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    Else
        RR_TOTVAT = 0
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " ttlrramt = " & RR_TOTUCOST & "," & _
                       " netrramt = " & RR_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & RR_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    End If
    Call NEW_LogAudit("E", "MATERIAL RECEIVING", SQL_STATEMENT, labID, "", "RR NO: " & txtRRNo & " ADD/EDIT DETAILS", "", "")
    
    gconDMIS.Execute ("Update pmis_stockmas set srp = '" & Trim(txtNewSRP.Text) & "', stockdesc = '" & cboTranDescription & "' where stockno = '" & cboTranPartNo & "' and [type] = 'M'")
    'cmdOkUpdate_Click
    rsRefresh
    On Error Resume Next
    rsRR_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True
    If ADDOREDIT = "ADD" And Picture1.Visible = True Then
    Call addTran
    Picture1.Enabled = False
    fra_Search.Enabled = False
    Screen.MousePointer = 0
    Exit Sub
    End If
ERRORCODE:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub

End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "MATERIALS RECEIVING") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        Dim tmpOnHand                   As Integer
        Dim rsTranPartNo                As ADODB.Recordset
        Set rsTdayTran = New ADODB.Recordset
        rsTdayTran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt,status from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
        If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
            rsTdayTran.MoveFirst
            
            '=================================
            'updating code:     jaa -9092008
            If NegativeValuesExist = True Then
               Exit Sub
            End If
            '=================================
          
            Do While Not rsTdayTran.EOF
                Set RSPARTMAS = New ADODB.Recordset
                RSPARTMAS.Open "Select STOCKNO,onhand,trecqty,onorder,served,receipts from CSMS_MATMAS where TYPE = " & MODULE_STOCK_TYPE & " AND STOCKNO = " & N2Str2Null(rsTdayTran!STOCK_ORD), gconDMIS
                If Not RSPARTMAS.EOF And Not RSPARTMAS.EOF Then
                    If Null2String(rsTdayTran!STATUS) = "P" Then
'                        '=================================
'                        'updating code:     jaa - 09092008
'                        tmpOnHand = N2Str2Zero(rsPartMas!ONHAND) - NumericVal(rsTdayTran!tranqty)
'                        If tmpOnHand < 0 Then
'                            'If MsgBox("Unposting this transaction will cause for negative stock of Part Number: " & N2Str2Null(rsPartMas!STOCKNO) & "" & vbCrLf & "Proceed Anyway?", vbYesNo + vbQuestion) = vbYes Then
'                            MsgBox "Issuance for Material Number: " & N2Str2Null(rsPartMas!STOCKNO) & " was already made. " & vbCrLf & "Unposting this Transaction will cause for Negative Stock of this Material Number."
'                            picPost.Visible = True
'                            Label39.Caption = N2Str2Null(rsTdayTran!stock_ord)
'                            Set rsTranPartNo = New ADODB.Recordset
'                            Set rsTranPartNo = gconDMIS.Execute("Select tranno,trantype,ID from PMIS_TDAYTRAN WHERE TYPE = 'M' AND TRANTYPE IN ('CSH','CHG','DR','RIV') AND STOCK_ORD = " & N2Str2Null(rsTdayTran!stock_ord) & " AND (STATUS = 'P' or STATUS = 'B') GROUP BY trantype,TRANNO,ID")
'                            If Not rsTranPartNo.EOF And Not rsTranPartNo.BOF Then
'                                Dim lblCtr As Integer
'                                lblCtr = 0
'                                picPost.Visible = True
'                                Do While Not rsTranPartNo.EOF
'                                    If lblCtr = 36 Then Exit Sub
'                                    Label36(lblCtr).Visible = True
'                                    Label36(lblCtr) = Null2String(rsTranPartNo!TranType) & ": " & Null2String(rsTranPartNo!TRANNO)
'                                    lblCtr = lblCtr + 1
'                                    rsTranPartNo.MoveNext
'                                Loop
'                            End If
'                            Exit Sub
'                        End If
                        '=================================
                        
                        
                        SQL_STATEMENT = "update CSMS_MATMAS set onhand =" & tmpOnHand & ", " & _
                                       " trecqty = " & N2Str2Zero(RSPARTMAS!trecqty) - NumericVal(rsTdayTran!tranqty) & ", " & _
                                       " onorder = " & N2Str2Zero(RSPARTMAS!onorder) + NumericVal(rsTdayTran!tranqty) & ", " & _
                                       " SERVED = " & N2Str2Zero(RSPARTMAS!Served) - NumericVal(rsTdayTran!tranqty) & ", " & _
                                       " receipts = " & N2Str2Zero(RSPARTMAS!receipts) - NumericVal(rsTdayTran!tranqty) & ", " & _
                                       " last_recq = " & 0 & ", " & _
                                       " last_recd = NULL, " & _
                                       " supcode = NULL" & _
                                       " where STOCKNO = " & N2Str2Null(RSPARTMAS!STOCKNO)
                        gconDMIS.Execute SQL_STATEMENT
                        Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSPARTMAS!STOCKNO), "STOCKNO", "CSMS_MATMAS"), "", "RR NO: " & txtRRNo & " UNPOSTED", "", "")
                    End If
                    
                    SQL_STATEMENT = "update PMIS_TdayTran set" & _
                                   " status = 'N'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdayTran!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "UU", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
                    
                End If
                rsTdayTran.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_RR_Hd set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "U", "MATERIALS RECEIVING", SQL_STATEMENT, labID, "Materials", "RR NO: " & txtRRNo, "RR", ""
        
        rsRefresh
        On Error Resume Next
        rsRR_HD.Find "id =" & labID.Caption
        StoreMemvars
        Call GetPrevMacAndDNP
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub
Private Sub GetPrevMacAndDNP()
'I Derived For This Formula to Get The PrevMac
'POH = NewOH - TQ
'PM = (NM[(TQ + POH)] - [(TC * TQ)])/ (POH)
'Additonal procedure created by NVB

On Error GoTo ERRORCODE

Dim rsGetPOH As New ADODB.Recordset
Dim rsGetBacker As New ADODB.Recordset
Dim rsGetMe As New ADODB.Recordset
Dim sqlGetData As String
Dim xstockno As String
    
'declaration of variable in formula
Dim TQ As Integer
Dim TC As Double
Dim POH As Integer
Dim NM As Double
Dim PM As Double
Dim xLASTM_MAC As Double
Dim xLASTM_OH As Integer
Dim old_dnp As Double
Dim recieve As Integer
Dim sqltxt As String

'this is MAC when ohand <> 0
Set rsGetBacker = New ADODB.Recordset
rsGetBacker.Open ("Select tranqty,tranucost,type,stock_ord from pmis_tdaytran where tranno = '" & txtRRNo & "' and [type] = 'M' and trantype = 'RR'"), gconDMIS
If Not (rsGetBacker.BOF And rsGetBacker.EOF) Then
End If

'sqlGetData = "select stockno from pmis_stockmas where stockno "
'sqlGetData = sqlGetData & "IN(Select stock_ord from pmis_tdaytran where tranno = '" & Trim(txtRRNo.Text) & "'"
'sqlGetData = sqlGetData & "and [type] = 'P' and trantype = 'RR')"

'Set rsGetPOH = gconDMIS.Execute(sqlGetData)
'Set rsGetMe = New ADODB.Recordset

PM = 0: old_dnp = 0:
With rsGetBacker
    .MoveFirst
    Do While Not .EOF
        xstockno = Trim(rsGetBacker!STOCK_ORD)
        TQ = Trim(rsGetBacker!tranqty)
        TC = Trim(rsGetBacker!TRANUCOST)

        rsGetMe.Open ("Select onhand,mac,dnp,lastm_mac,lastm_oh,receipts from PMIS_STOCKMAS where stockno = '" & xstockno & "' AND [TYPE] = 'M'"), gconDMIS
        If Not (rsGetMe.BOF And rsGetMe.EOF) Then
            DoEvents
            POH = Null2String(rsGetMe!ONHAND)
            NM = Null2String(rsGetMe!Mac)
            xLASTM_OH = N2Str2IntZero(rsGetMe!lastm_oh)
            xLASTM_MAC = N2Str2IntZero(rsGetMe!lastm_mac)
            recieve = N2Str2IntZero(rsGetMe!receipts)
        End If
            'if previous onhand is zero temporary quantity is given.
           If POH = 0 And xLASTM_MAC = 0 And xLASTM_OH = 0 And recieve = 0 Then 'New ITEM
           
                    'Find out if the Trancost is the same lang to its old mac
                    If TC <> NM Then
                        'Computation to Get The Previous MAC
                        'PM = (NM[(TQ + POH)] - [(TC * TQ)])/ (POH)
                        PM = Round(((NM * (TQ + POH) - (TC * TQ)) / (POH)), 2)
                        
                        'To Get Prev DNp
                        If ISNONVAT = True Then
                            old_dnp = ToDoubleNumber(PM)
                        Else
                            old_dnp = ToDoubleNumber(PM * 1.12)
                        End If
                        
                        sqltxt = "Update pmis_stockmas set mac = '" & PM & "',dnp = '" & N2Str2Zero(old_dnp) & "'"
                        sqltxt = sqltxt & " where stockno = '" & xstockno & "' and [type] = 'M'"
                
                        gconDMIS.Execute (sqltxt)
                    Else
                        sqltxt = "Update pmis_stockmas set mac = '" & PM & "',dnp = '" & N2Str2Zero(old_dnp) & "'"
                        sqltxt = sqltxt & " where stockno = '" & xstockno & "' and [type] = 'M'"
                
                        gconDMIS.Execute (sqltxt)
                        'do nothing
                    End If
            Else 'THIS OLD ITEM
            
                    If TC <> NM Then
                        'Computation to Get The Previous MAC
                        'PM = (NM[(TQ + POH)] - [(TC * TQ)])/ (POH)
                        PM = Round((((NM * (TQ + POH)) - (TC * TQ)) / (POH)), 2)
                        
                        'To Get Prev DNp
                        If ISNONVAT = True Then
                            old_dnp = ToDoubleNumber(PM)
                        Else
                            old_dnp = ToDoubleNumber(PM * 1.12)
                        End If
                        
                        sqltxt = "Update pmis_stockmas set mac = '" & PM & "',dnp = '" & Trim(old_dnp) & "'"
                        sqltxt = sqltxt & " where stockno = '" & xstockno & "' and [type] = 'M'"
                
                        gconDMIS.Execute (sqltxt)
                    Else
                        'do nothing
                    End If
            End If
    .MoveNext
    rsGetMe.Close
    Loop
End With


Set rsGetMe = Nothing
Set rsGetPOH = Nothing
Set rsGetBacker = Nothing

ERRORCODE:
    Exit Sub
End Sub


Private Sub addTran()

     fra_Search.Enabled = False
    If Picture1.Visible = True Then
        SendToBack
        cmdAddTran.ZOrder 0
        fraAddTran.ZOrder 0
        cmdTranDelete.Visible = False
        fraAddTran.Enabled = True
        ADDOREDIT = "ADD"
        InitParts
        On Error Resume Next
        cboTranPartNo.SetFocus
        Send2FrontConfirm
    End If

End Sub


Function NegativeValuesExist() As Boolean
    NegativeValuesExist = False
    
    Dim rsTranPartNo                            As ADODB.Recordset
    Dim rsParts                                 As ADODB.Recordset
    Dim rsRRno                                  As ADODB.Recordset
    Dim tmpOnHand                               As Integer
    Dim lstTrans                                As XtremeReportControl.ReportRecord
    lstRefTransNo.Records.DeleteAll
    Set rsRRno = New ADODB.Recordset
    rsRRno.Open "select tranno,STOCK_ORD,tranqty,status from PMIS_TdayTran where [TYPE] = 'M' AND trantype = 'RR' and tranno = " & N2Str2Null(rsRR_HD!RRNO) & " order by itemno asc", gconDMIS
    If Not rsRRno.EOF And Not rsRRno.BOF Then
            rsRRno.MoveFirst
            Do While Not rsRRno.EOF
                Set rsParts = New ADODB.Recordset
                rsParts.Open "Select STOCKNO,onhand from CSMS_MATMAS where STOCKNO = " & N2Str2Null(rsRRno!STOCK_ORD), gconDMIS
                If Not rsParts.EOF And Not rsParts.EOF Then
                     tmpOnHand = N2Str2Zero(rsParts!ONHAND) - NumericVal(rsRRno!tranqty)
                     If tmpOnHand < 0 Then
                        Set rsTranPartNo = New ADODB.Recordset
                        Set rsTranPartNo = gconDMIS.Execute("Select tranno,trantype,ID,stock_ord,tranqty,trandate from PMIS_TDAYTRAN WHERE TYPE = 'M' AND TRANTYPE IN ('CSH','CHG','DR','RIV') AND STOCK_ORD = " & N2Str2Null(rsRRno!STOCK_ORD) & " AND (STATUS = 'P' or STATUS = 'B') GROUP BY trantype,TRANNO,ID,stock_ord,tranqty,trandate ORDER BY TRANDATE DESC")
                        If Not rsTranPartNo.EOF And Not rsTranPartNo.BOF Then
                            rsTranPartNo.MoveFirst
                            'lstRefTransNo.Visible = True
                            Picture7.Visible = True
                            Do While Not rsTranPartNo.EOF
                                Set lstTrans = lstRefTransNo.Records.Add
                                With lstTrans
                                    .AddItem Space(2) & Null2String(rsTranPartNo!STOCK_ORD) & Space(6) & "OnHand: " & N2Str2Zero(rsParts!ONHAND) & Space(10) & "RR Qty.: " & N2Str2Zero(rsRRno!tranqty)
                                    .AddItem Null2String(rsTranPartNo!trandate)
                                    .AddItem Null2String(rsTranPartNo!TRANNO)
                                    .AddItem Null2String(rsTranPartNo!TranType)
                                    .AddItem N2Str2Zero(rsTranPartNo!tranqty)
                                End With
                                rsTranPartNo.MoveNext
                            Loop
                        End If
                        NegativeValuesExist = True
                     End If
                End If
                rsRRno.MoveNext
            Loop
            lstRefTransNo.Populate
    End If
    
End Function



Private Sub Command1_Click()
    Dim RRUNITCOST                                                    As Double
    Dim rsPartMasClone                                                As ADODB.Recordset
    Set rsPartMasClone = New ADODB.Recordset
    rsPartMasClone.Open "select STOCKNO,tpoqty,onorder,mac,dnp,srp,onhand,NON_HARI from CSMS_MATMAS where TYPE = 'M' AND STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS

    If cboTranDescription.Text = "" Then
        MsgSpeechBox "Description must have a value"
        On Error Resume Next
        cboTranDescription.Enabled = True
        cboTranDescription.SetFocus
        Exit Sub
    End If

    If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
    
        'updating code:      jaa - 09102008     - Exclude VAT if the supplier is a Non-VAT Supplier
'        RRUNITCOST = NumericVal(txtUnitCost.Text)
'        If Null2String(rsRR_HD!classcode) = "PCG" Or Null2String(rsRR_HD!classcode) = "PCS" Then
'            If CheckIfNonVatSup(Trim(txtRecvd_Code.Text)) = False Then
'               RRUNITCOST = RRUNITCOST / ConvertToBIRDecimalFormat(VAT_RATE)
'            End If
'        End If
        
        PrevPmasMAC = Format(NumericVal(rsPartMasClone!Mac), MAXIMUM_DIGIT)
        PrevPmasDNP = Format(NumericVal(rsPartMasClone!dnp), MAXIMUM_DIGIT)
        PrevPmasSRP = Format(NumericVal(rsPartMasClone!SRP), MAXIMUM_DIGIT)

        PrevPmasOnHand = N2Str2Zero(rsPartMasClone!ONHAND)
        NewPmasOnHand = NumericVal(txtTranQty.Text)
        If Null2String(rsPartMasClone!NON_HARI) = "Y" Then
            chkHARI_PARTS.Value = 0
        Else
            chkHARI_PARTS.Value = 1
        End If
        NewPmasDNP = NumericVal(txtTranINVAmt.Text)
        If PrevPmasOnHand <= 0 Then
            NewPmasMAC = (NumericVal(txtUnitCost.Text) * NumericVal(txtTranQty.Text)) / NewPmasOnHand
            'NewPmasMAC = Round((RRUNITCOST * NumericVal(txtTranQty.Text)) / NewPmasOnHand, 2)
        Else
            NewPmasMAC = ((PrevPmasMAC * PrevPmasOnHand) + (NumericVal(txtUnitCost.Text) * NumericVal(txtTranQty.Text))) / (NewPmasOnHand + PrevPmasOnHand)
            'NewPmasMAC = Round(((PrevPmasMAC * PrevPmasOnHand) + (RRUNITCOST * NumericVal(txtTranQty.Text))) / (NewPmasOnHand + PrevPmasOnHand), 2)
        End If
        NewPmasSRP = PrevPmasSRP
        txtOldMAC.Text = Format(PrevPmasMAC, MAXIMUM_DIGIT)
        txtOldDNP.Text = Format(PrevPmasDNP, MAXIMUM_DIGIT)
        txtOldSRP.Text = Format(PrevPmasSRP, MAXIMUM_DIGIT)
        txtOldOH.Text = Format(PrevPmasOnHand, DIGIT_FORMAT)
        txtNewMAC.Text = Format(NewPmasMAC, MAXIMUM_DIGIT)
        txtNewDNP.Text = Format(NewPmasDNP, MAXIMUM_DIGIT)
        txtNewSRP.Text = Format(NewPmasSRP, MAXIMUM_DIGIT)
        txtNewOH.Text = Format(NewPmasOnHand, DIGIT_FORMAT)
        Screen.MousePointer = 0
    Else
        PrevPmasMAC = "0.00": PrevPmasDNP = "0.00": PrevPmasSRP = "0.00": PrevPmasOnHand = "0"
        NewPmasOnHand = NumericVal(txtTranQty.Text)
        NewPmasSRP = NumericVal(txtNewSRP.Text)
        If NumericVal(txtDS1.Text) <= 0 Then
            NewPmasDNP = NumericVal(txtUnitCost.Text)
            'NewPmasDNP = RRUNITCOST
        Else
            NewPmasDNP = NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE)
            'NewPmasDNP = RRUNITCOST * ConvertToBIRDecimalFormat(VAT_RATE)
        End If
        If txtRecvd_Code.Text = VPAMCOR Then
            NewPmasMAC = (NumericVal(txtUnitCost.Text) * NumericVal(txtTranQty.Text)) / NewPmasOnHand
            'NewPmasMAC = Round((RRUNITCOST * NumericVal(txtTranQty.Text)) / NewPmasOnHand, 2)
            'NewPmasSRP = "0.00"
        Else
            NewPmasMAC = (NumericVal(txtUnitCost.Text) * NumericVal(txtTranQty.Text)) / NewPmasOnHand
            'NewPmasMAC = Round((RRUNITCOST * NumericVal(txtTranQty.Text)) / NewPmasOnHand, 2)
            'NewPmasSRP = "0.00"
        End If
        Send2FrontConfirm
        txtOldMAC.Text = Format(PrevPmasMAC, MAXIMUM_DIGIT)
        txtOldDNP.Text = Format(PrevPmasDNP, MAXIMUM_DIGIT)
        txtOldSRP.Text = Format(PrevPmasSRP, MAXIMUM_DIGIT)
        txtOldOH.Text = Format(PrevPmasOnHand, DIGIT_FORMAT)
        txtNewMAC.Text = Format(NewPmasMAC, MAXIMUM_DIGIT)
        txtNewDNP.Text = Format(NewPmasDNP, MAXIMUM_DIGIT)
        txtNewSRP.Text = Format(NewPmasSRP, MAXIMUM_DIGIT)
        txtNewOH.Text = Format(NewPmasOnHand, DIGIT_FORMAT)
        If Trim(cboTranPartNo.Text) <> "" Then
            gconDMIS.Execute "insert into CSMS_MATMAS " & _
                             "(TYPE,STOCKNO,STOCKDESC,MAC,DNP,SRP,date_entered,ACTIVE)" & _
                           " values ('M'," & N2Str2Null(cboTranPartNo.Text) & "," & N2Str2Null(Mid(cboTranDescription.Text, 1, 50)) & "," & NewPmasMAC & "," & NewPmasDNP & "," & NewPmasSRP & ", '" & LOGDATE & "','Y')"
        End If
        chkHARI_PARTS.Value = 0
        Screen.MousePointer = 0
    End If
    cmdTranSave.Enabled = True
    'End If

End Sub

Private Sub Command2_Click()
    lstRefTransNo.Visible = False
End Sub

Private Sub Command3_Click()
    picPost.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim fild                                                          As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    
    If Shift = 2 Then
        If KeyCode = vbKeyF1 Then
            'If picDetails.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Purchase Receiving and Storing)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "MATERIALS RECEIVING")
        End If
    End If
    
    
    Select Case KeyCode
        Case vbKeyEscape
            'picPost.Visible = False
            'lstRefTransNo.Visible = False
            Picture7.Visible = False
            If Picture1.Visible = True Then
                SendToBack
                StoreMemvars
            End If
            fra_Search.Enabled = True
            Picture1.Enabled = True
            fraDetails.Enabled = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsRR_HD!STATUS) = "P" Then
                    MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
                ElseIf Null2String(rsRR_HD!STATUS) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                Else
                    cmdAddTran_Click
                    cmdTranSave.Enabled = False
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsRR_HD!STATUS) <> "P" And Null2String(rsRR_HD!STATUS) <> "C" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If
            
           
            
        Case vbKeyF5
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsRR_HD!STATUS) <> "P" And Null2String(rsRR_HD!STATUS) <> "C" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    MODULE_STOCK_TYPE = "'M'"
    rsRefresh
    'EAP:021709 enabled search list
    textSearch.Text = "" ': SendToBack
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False
    txtPartID.Text = "": initMemvars: InitCboPayTerm
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then rsRR_HD.MoveLast
    StoreMemvars
    chkUpdateMAC.Enabled = False: chkUpdateDNP.Enabled = False
    txtNewMAC.Enabled = False: txtNewDNP.Enabled = False
    'picPost.Visible = False
    'Picture1.Visible = True
    InitGridRefTransNo
    'lstRefTransNo.Visible = False
    Picture7.ZOrder 0
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISTrans_Receiving2 = Nothing
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                                          As String
    If Null2String(rsRR_HD!STATUS) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf Null2String(rsRR_HD!STATUS) = "C" Then
        MsgSpeechBox "Item(s) are Already Cancelled and cannot be edited"
    Else
        fra_Search.Enabled = False
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            ADDOREDIT = "EDIT"
            BringToFront
            cmdTranDelete.Enabled = True
            cmdTranSave.Enabled = False
            fraAddTran.Caption = "Edit Materials"
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Materials"
            Exit Sub
        End If
    End If
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lstRR_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstRR_HD
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstRR_HD_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstRR_HD_GotFocus()
    StoreMemvars
End Sub

Private Sub lstRR_HD_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    If optRRNo.Value = True Then rsRR_HD.Bookmark = rsFind(rsRR_HD.Clone, "rrno", ITEM).Bookmark Else rsRR_HD.Bookmark = rsFind(rsRR_HD.Clone, "ID", lstRR_HD.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstRR_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub optRONo_Click()
    lstRR_HD.ColumnHeaders(1).Text = "Sup. Name": lstRR_HD.ColumnHeaders(1).Width = 4000
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optRRNo_Click()
    lstRR_HD.ColumnHeaders(1).Text = "Tran. No.": lstRR_HD.ColumnHeaders(1).Width = 2150
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optRRNo.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    Else
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstRR_HD.SetFocus
End Sub

Private Sub Timer1_Timer()
    If labRRsted.Caption <> "" Then
        If labRRsted.Visible = True Then
            labRRsted.Visible = False
        Else
            labRRsted.Visible = True
        End If
    End If
End Sub

Private Sub txtDS1_LostFocus()
    txtDS1.Text = Format(txtDS1.Text, "##0")
End Sub

Private Sub txtNewOH_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtNewSRP_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtPONo_GotFocus()
    If txtPONo.Text = "" And ADDOREDIT = "ADD" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where [TYPE] = 'M' AND modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtPONo.Text = Format(N2Str2Zero(RSCUNTER!nextnumber) - 1, "000000")
        End If
    End If
End Sub

Private Sub txtPONo_LostFocus()
    If cboClasscode.Text = "PURCHASED CHARGE" Then
        If txtPONo.Text <> "" And ADDOREDIT = "ADD" And Len(txtPONo.Text) > 0 Then
            Dim rsRR_HDDup                                            As ADODB.Recordset
            Set rsRR_HDDup = New ADODB.Recordset
            rsRR_HDDup.Open "select pono from PMIS_vw_RR_Trans where [TYPE] = 'M' AND pono = '" & txtPONo.Text & "' and status = 'P'", gconDMIS
            If Not rsRR_HDDup.EOF And Not rsRR_HDDup.BOF Then
                MsgBox "PO Number already Received.", vbInformation, "Invalid PO Number"
                Exit Sub
            End If
            RECEIVED_FROM_PO = "YES"
            Set RSPO_HD = New ADODB.Recordset
            RSPO_HD.Open "select pono,supcode,podate from PMIS_vw_PO_Trans where [TYPE] = 'M' AND pono = '" & txtPONo.Text & "'", gconDMIS
            If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
                txtRecvd_Code.Text = Null2String(RSPO_HD!SupCode): txtPODate.Text = Null2String(RSPO_HD!PODATE): cboTerms.Text = SetSupTerms(Null2String(RSPO_HD!SupCode))
                PCNT = 0: RR_TOTUCOST = 0: RR_TOTINVAMT = 0: RR_TOTVAT = 0: RR_QTY_REC = 0
                Dim rsTdaytranDup                                     As ADODB.Recordset
                Set rsTdaytranDup = New ADODB.Recordset
                rsTdaytranDup.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_AllDayTran where TYPE = 'M' and trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                    Screen.MousePointer = 11: rsTdaytranDup.MoveFirst: cleargrid grdDetails
                    Do While Not rsTdaytranDup.EOF
                        PCNT = PCNT + 1
                        grdDetails.AddItem rsTdaytranDup!ID & Chr(9) & Format(Null2String(rsTdaytranDup!itemno), "0000") & Chr(9) & _
                                           Null2String(rsTdaytranDup!STOCK_ORD) & Chr(9) & _
                                           SetSTOCKDESC(Null2String(rsTdaytranDup!STOCK_SUP)) & Chr(9) & _
                                           N2Str2IntZero(rsTdaytranDup!tranqty) & Chr(9) & _
                                           ToDoubleNumber(N2Str2Zero(rsTdaytranDup!TRANINVAMT)) & Chr(9) & _
                                           ToDoubleNumber(N2Str2Zero(rsTdaytranDup!TRANUCOST)) & Chr(9) & _
                                           ToDoubleNumber(N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANUCOST))
                        RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANUCOST))
                        RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANINVAMT))
                        rsTdaytranDup.MoveNext
                    Loop
                    If PCNT <> 0 Then grdDetails.RemoveItem 1
                    Screen.MousePointer = 0
                Else
                    cleargrid grdDetails
                End If
            Else
                MsgSpeechBox "Invalid Purchase Order Number!": txtPONo.Text = "": txtPODate.Text = "": If ADDOREDIT = "ADD" Then cleargrid grdDetails
                On Error Resume Next
                txtPONo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtRecvd_Code_Change()
    cboRecvd_Desc.Text = SetSupdesc(txtRecvd_Code.Text)
End Sub

Private Sub txtRemarks_GotFocus()
    MsgSpeech "Pls Type Your Message Here!": If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtRIV_Tranno_LostFocus()
    txtRIV_Tranno.Text = Format(txtRIV_Tranno.Text, "000000")
End Sub

Private Sub txtRRNo_LostFocus()
    txtRRNo = Format(txtRRNo, "000000")
End Sub

Private Sub txttranQty_Change()
    cmdTranSave.Enabled = False
    If txtTranQty.Text <> "" Then
        If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
            If Null2String(rsRR_HD!classcode) = "PCS" Or Null2String(rsRR_HD!classcode) = "PCG" Then
                If ISNONVAT = True Then txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text)) Else txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
            Else
                txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
                txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
            End If
        End If
    End If
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_LostFocus()
    If Trim(txtTranQty.Text) = "" Then txtTranQty.Text = 1
    If Not rsRR_HD.EOF And Not rsRR_HD.BOF Then
        If Null2String(rsRR_HD!classcode) = "PCS" Or Null2String(rsRR_HD!classcode) = "PCG" Then
            If ISNONVAT = True Then txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text)) Else txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        Else
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    End If
    txtTranQty.Text = Format(txtTranQty.Text, DIGIT_FORMAT)
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtUnitCost_Change()
    On Error Resume Next
    cmdTranSave.Enabled = False
    If Null2String(rsRR_HD!classcode) = "PCS" Or Null2String(rsRR_HD!classcode) = "PCG" Then
        If NumericVal(txtUnitCost.Text) <> 0 Then
            If ISNONVAT = True Then txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text)) Else txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    Else
        If NumericVal(txtUnitCost.Text) <> 0 Then
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    End If
End Sub

Private Sub txtUnitCost_GotFocus()
    If NumericVal(txtUnitCost.Text) = 0 Then txtUnitCost.Text = "" Else txtUnitCost.Text = NumericVal(txtUnitCost.Text)
End Sub

Private Sub txtUnitCost_LostFocus()
    txtUnitCost.Text = Format(txtUnitCost.Text, MAXIMUM_DIGIT)
End Sub


Function CheckIfNonVatSup(SupplierCode As String) As Boolean
    Dim rsSupplierMaster                                              As ADODB.Recordset
    Set rsSupplierMaster = New ADODB.Recordset
    rsSupplierMaster.Open "Select supcode,supname,NONVAT from PMIS_vw_Supplier where supcode = '" & SupplierCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplierMaster.EOF And Not rsSupplierMaster.BOF Then
        If Null2String(rsSupplierMaster!NONVAT) = "Y" Then CheckIfNonVatSup = True Else CheckIfNonVatSup = False
    Else
        CheckIfNonVatSup = False
    End If
End Function

Sub InitGridRefTransNo()
    lstRefTransNo.Columns.DeleteAll
    Call AddColumnHeader("Material No. ,Trans. Date,Trans. No.,Trans. Type,Issued Qty", lstRefTransNo)
    ResizeColumnHeader lstRefTransNo, "0,3,2.5,3,3"
    flex_FillReportPaintManager lstRefTransNo
    With lstRefTransNo
        .Columns(0).Visible = False
        .Columns(1).Alignment = xtpAlignmentLeft
        .Columns(2).Alignment = xtpAlignmentLeft
        .Columns(3).Alignment = xtpAlignmentCenter
        .Columns(4).Alignment = xtpAlignmentCenter
        .GroupsOrder.Add .Columns(0)
    End With
End Sub

