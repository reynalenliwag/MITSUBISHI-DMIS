VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMAT_MRISForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Requisition Slip Data Entry"
   ClientHeight    =   6405
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11430
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_MRISForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2685
      ScaleHeight     =   870
      ScaleWidth      =   8715
      TabIndex        =   92
      Top             =   5565
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
         MouseIcon       =   "MAT_MRISForm.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   103
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
         MouseIcon       =   "MAT_MRISForm.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   102
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
         MouseIcon       =   "MAT_MRISForm.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   101
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
         Left            =   5520
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_MRISForm.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   100
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
         MouseIcon       =   "MAT_MRISForm.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   97
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
         MouseIcon       =   "MAT_MRISForm.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":213D
         Style           =   1  'Graphical
         TabIndex        =   98
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
         MouseIcon       =   "MAT_MRISForm.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   99
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
         MouseIcon       =   "MAT_MRISForm.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   96
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
         MouseIcon       =   "MAT_MRISForm.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":2EF4
         Style           =   1  'Graphical
         TabIndex        =   95
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
         MouseIcon       =   "MAT_MRISForm.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   94
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
         MouseIcon       =   "MAT_MRISForm.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   2820
      Top             =   4140
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
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2670
      ScaleHeight     =   285
      ScaleWidth      =   8715
      TabIndex        =   83
      Top             =   5280
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
         TabIndex        =   88
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
         TabIndex        =   87
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label21 
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
         Left            =   2790
         TabIndex        =   85
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
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
         Left            =   1440
         TabIndex        =   86
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label19 
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
         Left            =   90
         TabIndex        =   84
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6375
      Left            =   60
      TabIndex        =   0
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
         Left            =   75
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4935
         Left            =   60
         TabIndex        =   5
         Top             =   1350
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8705
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "MAT_MRISForm.frx":3B49
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
         TabIndex        =   1
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9720
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   89
      Top             =   5535
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
         Left            =   780
         MouseIcon       =   "MAT_MRISForm.frx":3CAB
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":3DFD
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
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
         Left            =   60
         MouseIcon       =   "MAT_MRISForm.frx":413B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":428D
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   2670
      ScaleHeight     =   2775
      ScaleWidth      =   8715
      TabIndex        =   6
      Top             =   0
      Width           =   8715
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   2730
         TabIndex        =   25
         Top             =   570
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F1 - Assign MRS Number"
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
         Left            =   30
         TabIndex        =   13
         Top             =   90
         Width           =   2295
      End
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Top             =   75
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
         Left            =   5340
         TabIndex        =   21
         Text            =   "MRSGI07A001"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   82
         Width           =   1755
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
         TabIndex        =   9
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
         Height          =   915
         Left            =   4620
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         ToolTipText     =   "Type your message or remarks."
         Top             =   1740
         Width           =   3975
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
         Height          =   1365
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         ToolTipText     =   "Type complete name of customer."
         Top             =   1320
         Width           =   4365
      End
      Begin VB.TextBox txtTranDate 
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   24
         ToolTipText     =   "Type the date of transaction in mm/dd/yyyy format (e.g 7/5/2004)"
         Top             =   570
         Width           =   1545
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
         TabIndex        =   29
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
         Picture         =   "MAT_MRISForm.frx":45DD
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   7
         Top             =   -660
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
            TabIndex        =   8
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
         TabIndex        =   30
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
         Left            =   1140
         MaxLength       =   6
         TabIndex        =   15
         ToolTipText     =   "Input customer code (e.g. S01163)"
         Top             =   90
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
         Left            =   3810
         MaxLength       =   7
         TabIndex        =   26
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
         TabIndex        =   11
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
         TabIndex        =   16
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   90
         Width           =   1005
      End
      Begin VB.ComboBox cboSMName 
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
         Left            =   1080
         TabIndex        =   46
         Text            =   "cboSMName"
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2820
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
         Left            =   1080
         TabIndex        =   45
         Text            =   "cboSalesMan"
         Top             =   2820
         Width           =   765
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7110
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   36
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
         TabIndex        =   32
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
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   31
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1575
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
         TabIndex        =   14
         Top             =   90
         Width           =   1725
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MRIS No."
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
         TabIndex        =   19
         Top             =   120
         Width           =   855
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
         TabIndex        =   18
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
         TabIndex        =   33
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
         TabIndex        =   41
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Man"
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
         Left            =   90
         TabIndex        =   44
         Top             =   2850
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
         Left            =   2100
         TabIndex        =   35
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
         TabIndex        =   28
         Top             =   600
         Width           =   1635
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
         Left            =   3180
         TabIndex        =   27
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
         TabIndex        =   23
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
         TabIndex        =   10
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
         Left            =   2460
         TabIndex        =   17
         Top             =   120
         Width           =   1065
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
         TabIndex        =   42
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
         TabIndex        =   34
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
         TabIndex        =   22
         Top             =   105
         Width           =   1725
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
         Left            =   60
         TabIndex        =   12
         Top             =   90
         Width           =   1155
      End
   End
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2820
      TabIndex        =   47
      Top             =   150
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
      Location        =   $"MAT_MRISForm.frx":7319
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   2700
      ScaleHeight     =   2610
      ScaleWidth      =   8715
      TabIndex        =   81
      Top             =   2715
      Width           =   8715
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2475
         Left            =   30
         TabIndex        =   82
         Top             =   60
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   4366
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
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8100
         Top             =   120
      End
   End
   Begin VB.CommandButton cmdAddTran 
      Caption         =   "Command3"
      Height          =   3720
      Left            =   4530
      TabIndex        =   104
      Top             =   930
      Width           =   4785
   End
   Begin VB.CommandButton cmdSignatories 
      Caption         =   "Command3"
      Height          =   2535
      Left            =   4650
      TabIndex        =   105
      Top             =   2250
      Width           =   4620
   End
   Begin VB.PictureBox fraAddTran 
      Height          =   3525
      Left            =   4620
      ScaleHeight     =   3465
      ScaleWidth      =   4545
      TabIndex        =   48
      Top             =   1035
      Width           =   4605
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
         MouseIcon       =   "MAT_MRISForm.frx":73A7
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":74F9
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Delete Entry"
         Top             =   2580
         Width           =   735
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
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   58
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtTranDescription 
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
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   55
         Top             =   1050
         Width           =   4395
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
         TabIndex        =   66
         Top             =   2160
         Width           =   1665
      End
      Begin VB.TextBox txtTranUPrice 
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
         TabIndex        =   60
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
         TabIndex        =   57
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox txtTranItemNo 
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
         TabIndex        =   50
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
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
         Sorted          =   -1  'True
         TabIndex        =   53
         Text            =   "Combo1"
         ToolTipText     =   "Select Part Number from the list."
         Top             =   420
         Width           =   2955
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
         TabIndex        =   52
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
         MouseIcon       =   "MAT_MRISForm.frx":7824
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":7976
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Save Entry"
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
         MouseIcon       =   "MAT_MRISForm.frx":7CC6
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":7E18
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Cancel Entry"
         Top             =   2580
         Width           =   735
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
         Left            =   2430
         TabIndex        =   59
         Top             =   1470
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
         TabIndex        =   61
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
         TabIndex        =   62
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
         TabIndex        =   64
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
         TabIndex        =   65
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
         TabIndex        =   63
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
         TabIndex        =   56
         Top             =   1470
         Width           =   915
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
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   450
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
         Left            =   570
         TabIndex        =   49
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
         TabIndex        =   54
         Top             =   810
         Width           =   1275
      End
   End
   Begin VB.PictureBox fraSignatories 
      Height          =   2355
      Left            =   4755
      ScaleHeight     =   2295
      ScaleWidth      =   4350
      TabIndex        =   70
      Top             =   2355
      Width           =   4410
      Begin VB.CommandButton cmdPrintRIV 
         Caption         =   "&Print RIV"
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
         Left            =   1125
         MouseIcon       =   "MAT_MRISForm.frx":8156
         MousePointer    =   99  'Custom
         Picture         =   "MAT_MRISForm.frx":82A8
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1575
         Width           =   855
      End
      Begin VB.CheckBox chkPreview 
         BackColor       =   &H00DEDFDE&
         Height          =   255
         Left            =   4020
         TabIndex        =   80
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
         Left            =   1200
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   780
         Width           =   3045
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
         Left            =   1200
         TabIndex        =   78
         Text            =   "Text1"
         Top             =   1140
         Width           =   3045
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
         Left            =   1200
         TabIndex        =   74
         Text            =   "Text1"
         Top             =   420
         Width           =   3045
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
         Left            =   1200
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   60
         Width           =   3045
      End
      Begin VB.Label Label12 
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
         TabIndex        =   75
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label13 
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
         TabIndex        =   77
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label Label14 
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
         TabIndex        =   73
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         TabIndex        =   71
         Top             =   90
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmPMISMAT_MRISForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HD, RSTDAYTRAN, RSPARTMAS                    As ADODB.Recordset
Attribute RSTDAYTRAN.VB_VarUserMemId = 1073938432
Attribute RSPARTMAS.VB_VarUserMemId = 1073938432
Dim RSSALESMAN, RSCUNTER, RSPROFILE                    As ADODB.Recordset
Attribute RSSALESMAN.VB_VarUserMemId = 1073938435
Attribute RSCUNTER.VB_VarUserMemId = 1073938435
Attribute RSPROFILE.VB_VarUserMemId = 1073938435
Dim RSREPOR, rsCustomer                                As ADODB.Recordset
Attribute RSREPOR.VB_VarUserMemId = 1073938439
Attribute rsCustomer.VB_VarUserMemId = 1073938439
Dim KCNT                                               As Integer
Attribute KCNT.VB_VarUserMemId = 1073938441
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938442
Dim ORD_TOTUPRICE, ORD_TOTINVAMT, ORD_TOTVAT, ORD_TOTQTY As Double
Attribute ORD_TOTUPRICE.VB_VarUserMemId = 1073938443
Attribute ORD_TOTINVAMT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTVAT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTQTY.VB_VarUserMemId = 1073938443
Dim PREVORDTYPE, PREVORDNO                             As String
Attribute PREVORDTYPE.VB_VarUserMemId = 1073938447
Attribute PREVORDNO.VB_VarUserMemId = 1073938447

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
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
        If txtTranType.Text = "DR" Then
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
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
        If WAREHOUSETYPE = "ADB" Then
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
    If WAREHOUSETYPE = "ADB" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!DESCRIPTIO)
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
        Else
            Set RSPARTMAS = New ADODB.Recordset
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Stockmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
                SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
                If txtTranType.Text = "DR" Then
                    txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
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
        End If
    End If
End Function

Function SetSTOCKNO(pid As Variant)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_Stockmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
        If txtTranType.Text = "DR" Then
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
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
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_Stockmas WHERE [TYPE] = 'M' and STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_Stockmas WHERE [TYPE] = 'M' and ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select srp,STOCKNO,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            If txtTranType.Text = "DR" Then
                SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
            Else
                If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!dnp))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(RSPARTMAS!SRP))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(RSPARTMAS!MAC))
                End If
            End If
        End If
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice,tranucost from PMIS_vw_PRS_Tran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    End If
    If WAREHOUSETYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Function

Function CheckRoToParts() As String
    If txtRONO.Visible = True Then
        Dim rsCheck                                    As ADODB.Recordset
        Set rsCheck = gconDMIS.Execute("SELECT TRANNO  FROM PMIS_TDAYTRAN WHERE TRANTYPE='MRS' AND TYPE='M' AND STATUS='P'   AND STOCK_ORD ='" & Replace(cboTranPartNo, "'", "") & "' AND TRANNO IN(SELECT TRANNO  FROM PMIS_ORD_HD WHERE TRANTYPE='MRS' AND TYPE='M' AND STATUS='P'  AND RONO ='" & Replace(txtRONO, "'", "") & "')")
        If Not rsCheck.EOF Or Not rsCheck.BOF Then
            CheckRoToParts = Null2String(rsCheck!TRANNO)
        End If
    End If
End Function

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    RSORD_HD.Bookmark = rsFind(RSORD_HD.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    If WAREHOUSETYPE = "MRS" Then
        Me.Caption = "Materials Requistion Slip"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_vw_PRS where [TYPE] = 'M' and trantype = 'MRS' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
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
    If WAREHOUSETYPE = "MRS" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where modul = 'MRS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    txtRONO.Text = ""
    txtTerms.Text = ""
    txtTTLInvAmt.Text = "0.00"
    txtDS1.Text = "0"
    txtDS_Desc1.Text = "0.00"
    txtDS_Amt1.Text = "0.00"
    txtNetInvAmt.Text = "0.00"
    txtremarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    InitCbo
    InitGrid
    txtTranDate.Enabled = False: txtTranDate.Locked = True
    cleargrid grdDetails
    SendToBack
    InitSignatories

End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
End Sub

Sub StoreMemVars()
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then
        labid.Caption = RSORD_HD!ID
        txtTranType.Text = Null2String(RSORD_HD!TranType)
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(RSORD_HD!TRANNO)
        txtTranDate.Text = Null2String(RSORD_HD!trandate)
        txtCustCode.Text = Null2String(RSORD_HD!CUSTCODE)
        txtCustName.Text = Null2String(RSORD_HD!CUSTNAME)
        txtReferencePIS.Text = Null2String(RSORD_HD!REFPISNO)

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
        txtremarks.Text = Null2String(RSORD_HD!REMARKS)
        If Null2String(RSORD_HD!Status) = "C" Then
            labPosted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(RSORD_HD!Status) = "B" Then
            labPosted.Caption = "BILLED OUT"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
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
        .Text = "Materials Number"
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
    RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) * (NumericVal(txtDS1.Text) / 100))
            txtNetInvAmt.Text = ToDoubleNumber(NumericVal(txtTTLInvAmt.Text) - NumericVal(txtDS_Amt1.Text))
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
    RSPARTMAS.Open "select id,STOCKNO,STOCKDESC from PMIS_Stockmas WHERE TYPE = 'M' order by STOCKNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    RSREPOR.Open "select rep_or,niym,acct_no,invoice,plate_no from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS
    If Not RSREPOR.EOF And Not RSREPOR.BOF Then
        If Null2String(RSREPOR!invoice) <> "" Then
            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                 " Parts Request for this Repair Order is Critical!", vbCritical, "Critical Issue!"
            If MsgBox("Would You Like to Continue?", vbQuestion + vbYesNo, "Continue...") = vbNo Then
                On Error Resume Next
                txtRONO.SetFocus
                Exit Sub
            Else
                MsgBox "Pls. Input Your Notes/Reason in Remarks Field..."
                On Error Resume Next
                txtremarks.SetFocus
            End If
        End If
        txtCustName.Text = Null2String(RSREPOR!niym)
        txtCustCode.Text = Null2String(RSREPOR!ACCT_NO)

        Dim RSCUSTINFO                                 As ADODB.Recordset
        If Null2String(RSREPOR!PLATE_NO) <> "" Then
            Set RSCUSTINFO = New ADODB.Recordset
            Set RSCUSTINFO = gconDMIS.Execute("select * from CSMS_CUSVEH where Plate_NO=" & N2Str2Null(RSREPOR!PLATE_NO))
            If Not RSCUSTINFO.EOF Or Not RSCUSTINFO.BOF Then
                txtremarks = "MODEL: " & Null2String(RSCUSTINFO("model")) & vbCrLf & "ENGINE#:" & Null2String(RSCUSTINFO("ENGINE")) & vbCrLf & "VIN#:" & Null2String(RSCUSTINFO("vin")) & vbCrLf & "PLATE#:" & Null2String(RSCUSTINFO("plate_no"))
            End If
        End If
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""
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
    If WAREHOUSETYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Sub

Sub SendToBack()
    'lstOrd_Hd.Enabled = False
    cmdAddTran.ZOrder 1
    cmdAddTran.Visible = False
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    cmdSignatories.ZOrder 1
    cmdSignatories.Visible = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False
End Sub

Sub BringToFront()
    fraDetails.Enabled = False
    Picture1.Enabled = False
    lstOrd_Hd.Enabled = True
    cmdAddTran.ZOrder 0
    cmdAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True
End Sub

Sub SetCustomer()
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtCustName.Text = Null2String(rsCustomer!AcctName) & vbCrLf & Null2String(rsCustomer!CUSTOMERADD) & vbCrLf & Null2String(rsCustomer!City)
    End If
End Sub

Sub FillGrid()
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select top 20 tranno,tranno x from PMIS_vw_PRS where trantype = '" & Repleys(WAREHOUSETYPE) & "' order by Tranno asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set RSORD_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HD = gconDMIS.Execute("select top 20 tranno,tranno x from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & Repleys(WAREHOUSETYPE) & "' and tranno like '" & Repleys(XXX) & "%'")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select top 20 rono,tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & Repleys(WAREHOUSETYPE) & "' and rono is not null order by tranno asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set RSORD_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HD = gconDMIS.Execute("select top 20 Rono, tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & Repleys(WAREHOUSETYPE) & "' and rono like '" & Repleys(XXX) & "%' order by tranno asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
    txtTranDate.Enabled = True
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
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    Dim rschek                                         As New ADODB.Recordset
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)

        Set rschek = New ADODB.Recordset
        rschek.Open "Select * from PMIS_Stockmas where type = 'M'and stockno = " & N2Str2Null(cboTranPartNo) & "", gconDMIS, adOpenKeyset, adLockReadOnly

        If Not rschek.EOF And Not rschek.BOF Then
        Else
            MsgBox "Sorry partnumber is not in the list pls try again!", vbCritical
            cboTranPartNo = ""
            cboTranPartNo.SetFocus
        End If

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
    If Function_Access(LOGID, "Acess_CancelEntry", "MATERIALS REQUISITION SLIP") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCURTISSQTY                                As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    If Null2String(RSORD_HD!Status) = "P" Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                            SQL_STATEMENT = "update PMIS_Stockmas set" & _
                                          " ONREQUEST = " & PCURTISSQTY & "," & _
                                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                          " lastupdate = '" & LOGDATE & "'" & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                        Else
                            PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                            SQL_STATEMENT = "update PMIS_Stockmas set" & _
                                          " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                          " lastupdate = '" & LOGDATE & "'" & _
                                          " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                        Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Stockmas", "DETAILS", N2Str2Null("M"), "TYPE"), "", "TRAN NO: " & txtTranNo & " CANCEL", "", "")
                    End If
                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                  " status = 'C'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "C", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""
                End If
                RSTDAYTRANDUP.MoveNext
            Loop
        End If
        SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                      " status = 'C'," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "C", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
    Set RSTDAYTRANDUP = Nothing
    Set RSPARTMASDUP = Nothing

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPISNum_Click()

    With frmPMISMAT_MRIFormation

        If AddorEdit = "EDIT" Then
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
    frmPMISMAT_MRIFormation.Show 1
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "MATERIALS REQUISITION SLIP") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = True Then
        Dim PCURTISSQTY                                As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,tissqty,ONREQUEST,S_ONREQUEST from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        SQL_STATEMENT = "update PMIS_Stockmas set" & _
                                      " ONREQUEST = " & PCURTISSQTY & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                    Else
                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        SQL_STATEMENT = "update PMIS_Stockmas set" & _
                                      " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    Call NEW_LogAudit("E", "MATERIALS MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_Stockmas", "DETAILS", N2Str2Null("M"), "TYPE"), "", "TRAN NO: " & txtTranNo & " POSTED", "", "")

                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                  " status = 'P'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    NEW_LogAudit "PP", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""

                End If
                RSTDAYTRANDUP.MoveNext
            Loop
        End If

        SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                      " status = 'P'," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'" & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "P", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""

        rsRefresh
        On Error Resume Next
        RSORD_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
    Set RSTDAYTRANDUP = Nothing
    Set RSPARTMASDUP = Nothing

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgQuestionBox("Materials Requisition Slip will be Printed. Are you sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        If Mid((RSORD_HD!REFPISNO), 5, 1) = "W" Then
            rptCustomerOrder.WindowTitle = "Parts Requisition Slip"
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            'PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Material_OTC.rpt", "{PMIS_Ord_Hd.TYPE} = 'M' AND {PMIS_Ord_Hd.TRANTYPE} = 'MRS' and {PMIS_Ord_Hd.TRANNO} = " & N2Str2Null(txtTranno.Text), DMIS_REPORT_Connection, 1
            'PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Materials_OTC.rpt", "({PMIS_vw_PRS.TRANTYPE} = 'MRS' and {PMIS_vw_PRS_Tran.TYPE} = 'M' and {PMIS_vw_PRS_Tran.TRANTYPE} = 'MRS' and{PMIS_vw_PRS.STATUS} = 'P' and {PMIS_vw_PRS_Tran.TRANNO} = " & N2Str2Null(txtTranNo.Text) & ")", DMIS_REPORT_Connection, 1
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Materials.rpt", "({PMIS_vw_PRS.TYPE} = 'M' and {PMIS_vw_PRS.TRANTYPE} = 'MRS'  and{PMIS_vw_PRS.STATUS} = 'P' and {PMIS_vw_PRS_Tran.TRANNO} = " & N2Str2Null(txtTranNo.Text) & ")", DMIS_REPORT_Connection, 1
        Else
            rptCustomerOrder.WindowTitle = "Parts Requisition Slip"
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Materials.rpt", "({PMIS_vw_PRS.TYPE} = 'M' and {PMIS_vw_PRS.TRANTYPE} = 'MRS'  and{PMIS_vw_PRS.STATUS} = 'P' and {PMIS_vw_PRS_Tran.TRANNO} = " & N2Str2Null(txtTranNo.Text) & ")", DMIS_REPORT_Connection, 1
            'PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Materials.rpt", "({PMIS_vw_PRS.TYPE} = 'M' and {PMIS_vw_PRS.TRANTYPE} = 'MRS' and {PMIS_vw_PRS_Tran.TRANNO} = " & N2Str2Null(txtTranno.Text) & ")", DMIS_REPORT_Connection, 1
        End If
        Call NEW_LogAudit("V", "MATERIALS REQUISITION SLIP", "", labid, "", "TRAN NO: " & txtTranNo, "", "")
        Screen.MousePointer = 0
    End If

    Exit Sub
ErrorCode:
    ShowVBError


    '
    '    On Error GoTo ErrorCode:
    '
    '    If rsOrd_Hd!TRANTYPE = "PRS" Then
    '        cmdSignatories.Visible = True
    '        cmdSignatories.ZOrder 0
    '        fraSignatories.Visible = True
    '        fraSignatories.ZOrder 0
    '        Set rsSignatories = New ADODB.Recordset
    '        rsSignatories.Open "Select * from Signatories", gconDMIS
    '        If Not rsSignatories.EOF And Not rsSignatories.BOF Then
    '            txtPreparedBy.Text = Null2String(rsSignatories!PreparedBy)
    '            txtIssuedBy.Text = Null2String(rsSignatories!IssuedBy)
    '            txtRequestedBy.Text = Null2String(rsSignatories!requestedby)
    '            txtApprovedBy.Text = Null2String(rsSignatories!ApprovedBy)
    '            On Error Resume Next
    '            txtRequestedBy.SetFocus
    '        End If
    '        LogAudit "V", "MRIS", txtTranNo
    '        Set rsSignatories = Nothing
    '    End If
    '
    '    Exit Sub
    'ErrorCode:
    '    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    If RSORD_HD!TranType = "MRS" Then
    End If
    SendToBack
End Sub

Private Sub cmdTranCancel_Click()
    On Error GoTo ErrorCode:
    Picture1.Enabled = True
    fraDetails.Enabled = True
    SendToBack
    lstOrd_Hd.Enabled = True
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdTranDelete_Click()

    On Error GoTo ErrorCode:

    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Materials Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_vw_PRS_Tran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", "MAT CODE: " & cboTranPartNo, WAREHOUSETYPE, labDetID
        ShowDeletedMsg
    End If

    Dim CNT                                            As Integer
    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Set RSTDAYTRANDUP = New ADODB.Recordset
    RSTDAYTRANDUP.Open "select id,itemno from PMIS_vw_PRS_Tran where [TYPE] = 'M' and trantype = " & N2Str2Null(WAREHOUSETYPE) & " and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
    If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
        RSTDAYTRANDUP.MoveFirst
        CNT = 0
        Do While Not RSTDAYTRANDUP.EOF
            CNT = CNT + 1
            gconDMIS.Execute "update PMIS_vw_PRS_Tran set itemno = " & Format(CNT, "0000") & " where id = " & RSTDAYTRANDUP!ID
            RSTDAYTRANDUP.MoveNext
        Loop
    End If
    FillDetails
    SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                  " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                  " netinvamt = " & ORD_TOTINVAMT & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("E", "MATERIALS REQUISTION SLIP", SQL_STATEMENT, labid, "", "TRAN NO: " & cboTranPartNo, "MRS", "")

    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo ErrorCode
    If AddorEdit = "ADD" Then
        Dim StockISSUE: StockISSUE = CheckRoToParts
        If Len(StockISSUE) > 0 Then
            If MsgBox("Same Material number has been requested for this repair order." & vbCrLf & "Repair Order#  :" & txtRONO & vbCrLf & "Transaction#    :" & StockISSUE & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Accessories Requisition") = vbNo Then
                On Error Resume Next
                cboTranPartNo.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Materials Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_vw_PRS_Tran where [TYPE] = 'M' and STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(RSORD_HD!TRANNO) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Material Code already used in this transaction"
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
    Dim CRITICAL_QUESTION                              As String

    ORDTRANDATE = N2Date2Null(txtTranDate.Text)
    ORDTRANTYPE = N2Str2Null(txtTranType.Text)
    ORDTRANNO = N2Str2Null(txtTranNo.Text)
    ORDITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    ORDSTOCK_ORD = N2Str2Null(cboTranPartNo.Text)
    If txtTranType.Text = "ADB" Then ORDSTOCK_SUP = N2Str2Null(Left(txtTranDescription.Text, 100)) Else ORDSTOCK_SUP = N2Str2Null(cboTranPartNo.Text)
    ORDTRANQTY = NumericVal(txtTranQty.Text)
    ORDTRANUCOST = NumericVal(txtTranUCost.Text)
    ORDTRANINVAMT = NumericVal(txtTranUPrice.Text)
    ORDIN_OUT = "'R'"
    ORDSTATUS = "'N'"

    If ORDTRANINVAMT <= 0 Then
        MsgBox "Material number ( " & cboTranPartNo.Text & " ) has zero srp, please check material master file.", vbInformation + vbOKOnly
        Exit Sub
    End If
    If ORDTRANINVAMT < N2Str2IntZero(RSPARTMAS!dnp) Then
        If Mid(txtReferencePIS, 5, 1) = "W" Then
            'proceed
        Else
            If MsgBox("Your SRP of this Materials is below your Materials DNP. Proceed saving?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                'proceed saving
                CRITICAL_QUESTION = "Your SRP of this Materials is below your Materials DNP. Proceed saving?"
                Call NEW_LogAudit("MP", "MATERIALS REQUISITION SLIP", CRITICAL_QUESTION, labid, "", "TRAN NO: " & txtTranNo & " MAT NO: " & cboTranPartNo, "MRS", "")
                MsgBox "User Action has been Log in the Audit Trail", vbInformation, "Audit Trail Information"
            Else
                MsgBox "Pls. check your Materials DNP and SRP in your Materials Master File", vbInformation, "Confirm DNP/SRP"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_vw_PRS_Tran " & _
                        "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out)" & _
                      " values ('M'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", "MAT NO: " & cboTranPartNo, WAREHOUSETYPE, ""

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                      " trandate = " & ORDTRANDATE & "," & _
                      " trantype = " & ORDTRANTYPE & "," & _
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
        NEW_LogAudit "EE", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", "MAT NO: " & cboTranPartNo, WAREHOUSETYPE, labDetID

        ShowSuccessFullyUpdated
    End If
    cleargrid grdDetails
    FillDetails
    SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                  " totalqty = " & ORD_TOTQTY & "," & _
                  " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                  " netinvamt = " & ORD_TOTINVAMT & _
                  " where id = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("E", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "", "TRAN NO: " & txtTranNo & " ADD/EDIT DETAILS", "", "")

    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labid.Caption
    StoreMemVars
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
        Picture1.Enabled = False
        fraDetails.Enabled = False
    Else
        cmdTranCancel.Value = True
    End If

    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    initMemvars
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    txtTranDate.Enabled = False
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
    AddorEdit = "EDIT"
    PREVORDTYPE = txtTranType.Text
    PREVORDNO = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    txtTranDate.Enabled = False
    On Error Resume Next
    txtCustName.SetFocus
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
    On Error GoTo ErrorCode
    Dim NEXTCUNTER                                     As String
    Dim RSFINDDUP                                      As ADODB.Recordset
    Dim XSALES_ORIGIN, XSI_TYPE, XPAY_CLASS, XCHAR_YEAR, XCHAR_MONTH, XIS_SERIES, XTRACK_CODE As String

    Dim VCBOSALESMAN, VCBOSMNAME, VTXTTRANTYPE         As String
    Dim VTXTTRANNO, VTXTTRANDATE, VTXTCUSTCODE         As String
    Dim VTXTCUSTNAME, VTXTCHARGETO, VTXTRONO, VTXTREP_OR As String
    Dim VtxtTerms                                      As String
    Dim VTXTTTLINVAMT, VTXTDS1                         As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNETINVAMT                     As Double
    Dim VTXTRemarks, VStatus, Vusercode                As String
    Dim VLastUpdate                                    As String
    Dim VIN_PROCESS                                    As String
    Dim VTXTREFERENCEPIS                               As String

    If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
        MsgBox "Invalid Reference MRS Number!", vbCritical, "PRS Required!"
        Exit Sub
    End If

    
    If LTrim(RTrim(txtCustCode)) = "" Then
        MsgBox "Customer Information Is Required...", vbInformation, "Pls Select Customer Information..."
        Exit Sub
    End If

    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction No. must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            If checkdup("M", txtTranType.Text, txtTranNo.Text) = True Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!TRANNO))) Then
                If checkdup("M", txtTranType.Text, txtTranNo.Text) = True Then
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
    VCBOSALESMAN = "NULL"
    VCBOSMNAME = "NULL"

    NEXTCUNTER = NumericVal(txtTranNo.Text) + 1

    VTXTTRANTYPE = N2Str2Null(txtTranType.Text)
    VTXTTRANNO = N2Str2Null(txtTranNo.Text)
    VTXTTRANDATE = N2Date2Null(txtTranDate.Text)
    VTXTCUSTCODE = N2Str2Null(txtCustCode.Text)
    VTXTCUSTNAME = N2Str2Null(txtCustName.Text)
    VTXTREFERENCEPIS = N2Str2Null(txtReferencePIS.Text)

    VIN_PROCESS = "'Y'"
    VTXTCHARGETO = "'VAR'"
    VTXTRONO = N2Str2Null(txtRONO.Text)
    If Len(txtRONO.Text) = 10 Then
        VTXTREP_OR = "'" & Left(txtRONO.Text, 1) & "-" & Right(txtRONO.Text, 8) & "'"
    Else
        VTXTREP_OR = "NULL"
    End If
    VtxtTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLINVAMT = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNETINVAMT = NumericVal(txtNetInvAmt.Text)
    If txtremarks.Text = "Pls Type Your Message Here!" Then VTXTRemarks = "NULL" Else VTXTRemarks = N2Str2Null(Trim(txtremarks.Text))
    VStatus = "'N'"
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"

    XSALES_ORIGIN = N2Str2Null(Mid(txtReferencePIS, 3, 1))
    XSI_TYPE = N2Str2Null(Mid(txtReferencePIS, 4, 1))
    XPAY_CLASS = N2Str2Null(Mid(txtReferencePIS, 5, 1))
    XCHAR_YEAR = N2Str2Null(Mid(txtReferencePIS, 6, 2))
    XCHAR_MONTH = N2Str2Null(Mid(txtReferencePIS, 8, 1))
    XIS_SERIES = N2Str2Null(Mid(txtReferencePIS, 9, 3))
    XTRACK_CODE = N2Str2Null(Mid(txtReferencePIS, 12, 1))
    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_vw_PRS" & _
                      " ([TYPE],trantype,tranno,trandate,custcode,custname,chargeto,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " values ('M'," & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, FindTransactionID(txtTranNo, "tranno", "PMIS_vw_PRS", "DETAILS", N2Str2Null("M"), "TYPE"), "Materials", txtTranNo, WAREHOUSETYPE, ""
    Else

        SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                      " trantype = " & VTXTTRANTYPE & "," & _
                      " tranno = " & VTXTTRANNO & "," & _
                      " trandate = " & VTXTTRANDATE & "," & _
                      " custcode = " & VTXTCUSTCODE & "," & _
                      " custname = " & VTXTCUSTNAME & "," & _
                      " chargeto = " & VTXTCHARGETO & "," & _
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
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

        SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                      " SALES_ORIGIN = " & XSALES_ORIGIN & "," & _
                      " SI_TYPE = " & XSI_TYPE & "," & _
                      " PAY_CLASS = " & XPAY_CLASS & "," & _
                      " CHAR_YEAR = " & XCHAR_YEAR & "," & _
                      " CHAR_MONTH = " & XCHAR_MONTH & "," & _
                      " IS_SERIES = " & XIS_SERIES & "," & _
                      " TRACK_CODE = " & XTRACK_CODE & "" & _
                      " where id = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

        SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                      " trantype = " & VTXTTRANTYPE & "," & _
                      " trandate = " & VTXTTRANDATE & "," & _
                      " tranno = " & VTXTTRANNO & _
                      " where trantype = '" & PREVORDTYPE & "' and tranno = '" & PREVORDNO & "'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NEXTCUNTER & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' WHERE TYPE = 'M' and modul = " & VTXTTRANTYPE
    Else
        rsRefresh
        RSORD_HD.Find "Tranno = " & VTXTTRANNO
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                       " netinvamt = " & ORD_TOTINVAMT & _
                       " WHERE TYPE = 'M' and tranno = " & VTXTTRANNO & " and trantype = " & VTXTTRANTYPE
    End If

    If AddorEdit = "ADD" Then
        Picture1.Enabled = False
        fraDetails.Enabled = False

    Else
        Picture1.Enabled = True
        fraDetails.Enabled = True

    End If


    rsRefresh
    RSORD_HD.Find "tranno = " & VTXTTRANNO
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
        fraDetails.Enabled = False
    End If
    FillGrid
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
    frmCustomerSearch_PRS.LABTYPE = "M"
    frmCustomerSearch_PRS.Show 1
End Sub

Private Sub Command2_Click()
    cmdPISNum_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If Picture1.Visible = True Then
                Unload frmALL_AuditInquiry

                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (Materials Requisition Slip)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "MATERIALS REQUISITION SLIP")
            End If

        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            Picture1.Enabled = True
            fraDetails.Enabled = True
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
            If Picture1.Visible = True Then
                If Null2String(RSORD_HD!Status) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPOST", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
                    If MsgQuestionBox("Are you sure you want to UnPost this Transaction?", "UnPost Transaction") = True Then
                        Dim PCURTISSQTY                As Integer
                        Dim RSTDAYTRANDUP, RSPARTMASDUP As ADODB.Recordset

                        Set RSTDAYTRANDUP = New ADODB.Recordset
                        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(RSORD_HD!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
                            RSTDAYTRANDUP.MoveFirst
                            Do While Not RSTDAYTRANDUP.EOF
                                Set RSPARTMASDUP = New ADODB.Recordset
                                RSPARTMASDUP.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Stockmas where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                                       " ONREQUEST = " & PCURTISSQTY & "," & _
                                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                       " lastupdate = '" & LOGDATE & "'" & _
                                                       " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                    Else
                                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                                       " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                       " lastupdate = '" & LOGDATE & "'" & _
                                                       " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                    End If
                                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                                  " status = 'N'," & _
                                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                  " lastupdate = '" & LOGDATE & "'" & _
                                                  " where id = " & RSTDAYTRANDUP!ID
                                    gconDMIS.Execute SQL_STATEMENT
                                    NEW_LogAudit "U", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

                                End If
                                RSTDAYTRANDUP.MoveNext
                            Loop
                        End If
                        SQL_STATEMENT = "update PMIS_vw_PRS set" & _
                                      " status = 'N'," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where id = " & labid.Caption
                        gconDMIS.Execute SQL_STATEMENT
                        NEW_LogAudit "U", "MATERIALS REQUISITION SLIP", SQL_STATEMENT, labid, "Materials", txtTranNo, WAREHOUSETYPE, ""

                        rsRefresh
                        On Error Resume Next
                        RSORD_HD.Find "id =" & labid.Caption
                        StoreMemVars
                    End If
                    Set RSTDAYTRANDUP = Nothing
                    Set RSPARTMASDUP = Nothing
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
    Command1.Enabled = False
    Command1.Visible = False
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    initMemvars
    txtTranUPrice.Enabled = False
    rsRefresh
    StoreMemVars
    lstOrd_Hd.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISTrans_CustomerOrder = Nothing
    UnloadForm Me
End Sub

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
            MsgSpeechBox "No Entry on Parts!"
            Exit Sub
        End If
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

Private Sub txtReferencePIS_Change()

    If Mid((txtReferencePIS), 3, 1) = "S" Then
        Command1.Visible = False
        Command1.Enabled = False
        txtRONO.Enabled = True
        txtRONO.Visible = True
    Else
        Command1.Visible = True
        Command1.Enabled = True
        txtRONO.Enabled = False
        txtRONO.Visible = False
    End If


End Sub

Private Sub txtReferencePIS_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
    If txtremarks.Text = "Pls Type Your Message Here!" Then txtremarks.Text = ""
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtRONO_LostFocus()
    Dim RONOStr                                        As String
    RONOStr = txtRONO.Text
    If Left(RONOStr, 2) = "R-" Then
        RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
        If VALID_COMPANY_CODE_FORHAI = True Then
        
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
    End If
    txtRONO.Text = RONOStr
    SetCustInfo (RONOStr)
End Sub

Private Sub txtTranDate_LostFocus()
    txtTranDate.Text = Format(txtTranDate.Text, "SHORT DATE")
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
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

Private Sub lstOrd_Hd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optTranno.Value = True Then
        RSORD_HD.Bookmark = rsFind(RSORD_HD.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    Else
        RSORD_HD.Bookmark = rsFind(RSORD_HD.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    End If
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
    If KeyCode = vbKeyEscape Then On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        If Trim(textSearch.Text) = "" Then
            FillGrid
        Else
            FillSearchGrid (textSearch.Text)
        End If
    Else
        If Trim(textSearch.Text) = "" Then
            FillGrid2
        Else
            FillSearchGrid2 (textSearch.Text)
        End If
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOrd_Hd.ListItems.Count > 0 And lstOrd_Hd.Enabled = True Then: lstOrd_Hd.SetFocus
    End If
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

