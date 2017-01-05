VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISTrans_PrisForms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Requisition Slip Data Entry"
   ClientHeight    =   6015
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11430
   ForeColor       =   &H00DEDFDE&
   Icon            =   "PRISForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2685
      ScaleHeight     =   870
      ScaleWidth      =   8655
      TabIndex        =   91
      Top             =   5145
      Width           =   8655
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
         Left            =   7800
         MouseIcon       =   "PRISForm.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
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
         Left            =   7020
         MouseIcon       =   "PRISForm.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   95
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
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PRISForm.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":138C
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
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PRISForm.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   102
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
         Left            =   4680
         MouseIcon       =   "PRISForm.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   96
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
         Left            =   3900
         MouseIcon       =   "PRISForm.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":213D
         Style           =   1  'Graphical
         TabIndex        =   97
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
         Left            =   3120
         MouseIcon       =   "PRISForm.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   93
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
         Left            =   2340
         MouseIcon       =   "PRISForm.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   92
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
         Left            =   1560
         MouseIcon       =   "PRISForm.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":2EF4
         Style           =   1  'Graphical
         TabIndex        =   98
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
         Left            =   780
         MouseIcon       =   "PRISForm.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   99
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
         Left            =   0
         MouseIcon       =   "PRISForm.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   100
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
      Left            =   2760
      ScaleHeight     =   285
      ScaleWidth      =   8715
      TabIndex        =   76
      Top             =   4800
      Width           =   8715
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   81
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   80
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   79
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   78
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   77
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5985
      Left            =   60
      TabIndex        =   67
      Top             =   0
      Width           =   2595
      Begin VB.TextBox textSearch 
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
         Height          =   360
         Left            =   75
         MaxLength       =   35
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   960
         Width           =   2475
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
         Top             =   630
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
         Height          =   4575
         Left            =   60
         TabIndex        =   71
         Top             =   1350
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8070
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
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "PRISForm.frx":3B49
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tran. No."
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label18 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9720
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   88
      Top             =   5145
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
         MouseIcon       =   "PRISForm.frx":3CAB
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":3DFD
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Cancel"
         Top             =   0
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
         Left            =   0
         MouseIcon       =   "PRISForm.frx":413B
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":428D
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   2700
      ScaleHeight     =   2775
      ScaleWidth      =   8715
      TabIndex        =   28
      Top             =   90
      Width           =   8715
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   2460
         TabIndex        =   103
         Top             =   570
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F1 - Assign PRS Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         TabIndex        =   83
         Top             =   60
         Width           =   2355
      End
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   7110
         TabIndex        =   82
         Top             =   60
         Width           =   255
      End
      Begin VB.TextBox txtReferencePIS 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   5280
         TabIndex        =   1
         Text            =   "PRSGI07A001"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
         Width           =   1845
      End
      Begin VB.ComboBox cboChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   5550
         TabIndex        =   10
         Text            =   "cboChargeTo"
         ToolTipText     =   "Select option from list."
         Top             =   -405
         Width           =   1785
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   915
         Left            =   4620
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         ToolTipText     =   "Type your message or remarks."
         Top             =   1740
         Width           =   3975
      End
      Begin VB.TextBox txtCustName 
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
         Height          =   1275
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
         Enabled         =   0   'False
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
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         ToolTipText     =   "Type the date of transaction in mm/dd/yyyy format (e.g 7/5/2004)"
         Top             =   570
         Width           =   1275
      End
      Begin VB.TextBox txtDS1 
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
         Left            =   4800
         MaxLength       =   3
         TabIndex        =   11
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
         Picture         =   "PRISForm.frx":45DD
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   60
         Top             =   -660
         Width           =   435
         Begin VB.TextBox txtTranType 
            Appearance      =   0  'Flat
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
         Left            =   5700
         MaxLength       =   10
         TabIndex        =   12
         ToolTipText     =   "Input the type of the added amount."
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox txtCustCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
            Name            =   "Arial"
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
         Enabled         =   0   'False
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
         Height          =   330
         Left            =   1080
         TabIndex        =   9
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
            Name            =   "Arial"
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
         TabIndex        =   8
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
         TabIndex        =   59
         Top             =   510
         Width           =   1515
         Begin VB.TextBox txtNetInvAmt 
            Alignment       =   1  'Right Justify
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
            MaxLength       =   15
            TabIndex        =   63
            Top             =   440
            Width           =   1395
         End
         Begin VB.TextBox txtTTLInvAmt 
            Alignment       =   1  'Right Justify
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   74
         Top             =   960
         Width           =   2685
      End
      Begin VB.TextBox txtRONO 
         Appearance      =   0  'Flat
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
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lbluser 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         Left            =   6960
         TabIndex        =   106
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PRIS No."
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
         Height          =   285
         Left            =   4500
         TabIndex        =   75
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
         Enabled         =   0   'False
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
         Height          =   285
         Left            =   90
         TabIndex        =   41
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
         TabIndex        =   34
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Type"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   6900
         TabIndex        =   29
         Top             =   -30
         Width           =   1725
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   2010
      Left            =   2700
      ScaleHeight     =   2010
      ScaleWidth      =   8715
      TabIndex        =   42
      Top             =   2805
      Width           =   8715
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8100
         Top             =   120
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   1875
         Left            =   30
         TabIndex        =   14
         Top             =   60
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   3307
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
      Caption         =   "Command3"
      Height          =   3705
      Left            =   4530
      TabIndex        =   104
      Top             =   900
      Width           =   4785
   End
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2850
      TabIndex        =   27
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
      Location        =   ""
   End
   Begin VB.CommandButton cmdSignatories 
      Caption         =   "Command3"
      Height          =   2565
      Left            =   4650
      TabIndex        =   105
      Top             =   2250
      Width           =   4605
   End
   Begin VB.PictureBox fraSignatories 
      Height          =   2355
      Left            =   4755
      ScaleHeight     =   2295
      ScaleWidth      =   4350
      TabIndex        =   51
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
         MouseIcon       =   "PRISForm.frx":7319
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":746B
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   1575
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
            Name            =   "Arial"
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
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   780
         Width           =   3045
      End
      Begin VB.TextBox txtRequestedBy 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1140
         Width           =   3045
      End
      Begin VB.TextBox txtIssuedBy 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   420
         Width           =   3045
      End
      Begin VB.TextBox txtPreparedBy 
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
         Height          =   345
         Left            =   1200
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   60
         Width           =   3045
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Approved By"
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
         Height          =   255
         Left            =   90
         TabIndex        =   55
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Requested By"
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
         Height          =   255
         Left            =   90
         TabIndex        =   54
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Issued By"
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
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Prepared By"
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
         Height          =   255
         Left            =   90
         TabIndex        =   52
         Top             =   90
         Width           =   1065
      End
   End
   Begin VB.PictureBox fraAddTran 
      Height          =   3525
      Left            =   4620
      ScaleHeight     =   3465
      ScaleWidth      =   4545
      TabIndex        =   43
      Top             =   1005
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
         Left            =   2850
         MouseIcon       =   "PRISForm.frx":77D1
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":7923
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2490
         Width           =   705
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
         MouseIcon       =   "PRISForm.frx":7C4E
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":7DA0
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2490
         Width           =   705
      End
      Begin VB.TextBox txtTranUCost 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox txtTranDescription 
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
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1050
         Width           =   4395
      End
      Begin VB.TextBox txtTranTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
         TabIndex        =   18
         ToolTipText     =   "Type quantity purchased (e.g. 5, 4)"
         Top             =   1440
         Width           =   885
      End
      Begin VB.TextBox txtTranItemNo 
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
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   15
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   975
      End
      Begin VB.ComboBox cboTranPartNo 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   1470
         Sorted          =   -1  'True
         TabIndex        =   16
         ToolTipText     =   "Select Part Number from the list."
         Top             =   420
         Width           =   2985
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
         Left            =   1470
         MouseIcon       =   "PRISForm.frx":80DE
         MousePointer    =   99  'Custom
         Picture         =   "PRISForm.frx":8230
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2490
         Width           =   705
      End
      Begin VB.Label labTranUCost 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
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
         Height          =   225
         Left            =   2430
         TabIndex        =   73
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
         Caption         =   "Part No."
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
         Height          =   225
         Left            =   570
         TabIndex        =   47
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
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
            Name            =   "Arial"
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
End
Attribute VB_Name = "frmPMISTrans_PrisForms"
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
Dim RSPROFILE                                          As ADODB.Recordset
Dim RSREPOR                                            As ADODB.Recordset
Dim rsCustomer                                         As ADODB.Recordset
Dim KCNT                                               As Integer
Dim AddorEdit                                          As String
Dim ORD_TOTUPRICE                                      As Double
Dim ORD_TOTINVAMT                                      As Double
Dim ORD_TOTVAT                                         As Double
Dim ORD_TOTQTY                                         As Double
Dim PREVORDTYPE                                        As String
Dim PREVORDNO                                          As String
Dim BLNF3CNTR                                          As Boolean

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
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_Partmas where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Partmas where STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            RSPARTMAS.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Partmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    RSPARTMAS.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_Partmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_Partmas where STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_Partmas where ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select srp,STOCKNO,mac,dnp from PMIS_Partmas where STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
        Set rsCheck = gconDMIS.Execute("SELECT TRANNO  FROM PMIS_TDAYTRAN WHERE TRANTYPE='PRS' AND TYPE='P' AND STATUS='P'   AND STOCK_ORD ='" & Replace(cboTranPartNo, "'", "") & "' AND TRANNO IN(SELECT TRANNO  FROM PMIS_ORD_HD WHERE TRANTYPE='PRS' AND TYPE='P' AND STATUS='P'  AND RONO ='" & Replace(txtRONO, "'", "") & "')")
        If Not rsCheck.EOF Or Not rsCheck.BOF Then
            CheckRoToParts = Null2String(rsCheck!tranno)
        End If
    End If
End Function

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    RSORD_HD.Bookmark = rsFind(RSORD_HD.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    If WAREHOUSETYPE = "PRS" Then
        Me.Caption = "Parts Requistion Slip"
        Set RSORD_HD = New ADODB.Recordset
        RSORD_HD.Open "select * from PMIS_vw_PRS where trantype = 'PRS' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
End Sub

Sub InitCboChargeToWarehouse()
    With cboChargeTo
        .Clear
        .AddItem "MECHANICAL"
        .AddItem "COMPANY"
        .AddItem "WARRANTY"
        .AddItem "TINSMITH"
        .AddItem "VARIOUS"
        .AddItem "FLEET"
        .AddItem "PARTS CLAIM"
        .Text = "MECHANICAL"
    End With
End Sub

Sub InitCboChargeToCounter()
    With cboChargeTo
        .Clear
        .AddItem "MECHANICAL"
        .AddItem "COMPANY"
        .AddItem "WARRANTY"
        .AddItem "TINSMITH"
        .AddItem "VARIOUS"
        .AddItem "FLEET"
        .AddItem "PARTS CLAIM"
        .Text = "VARIOUS"
    End With
End Sub

Sub initMemvars()
    If WAREHOUSETYPE = "PRS" Then
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where modul = 'PRS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    cleargrid grdDetails
    SendToBack
    txtTranDate.Enabled = False
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
        txtTranNo.Text = Null2String(RSORD_HD!tranno)
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

Sub FillDetails()
    On Error Resume Next
    KCNT = 0
    ORD_TOTUPRICE = 0
    ORD_TOTINVAMT = 0
    ORD_TOTVAT = 0
    ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_vw_PRS_Tran where tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    RSPARTMAS.Open "select id,STOCKNO,STOCKDESC from PMIS_Partmas", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    cmdAddTran.ZOrder 1
    cmdAddTran.Visible = False
    fraAddTran.ZOrder 1
    fraAddTran.Visible = False
    fraAddTran.Enabled = False
    'cmdSignatories.ZOrder 1
    'cmdSignatories.Visible = False
    fraSignatories.ZOrder 1
    fraSignatories.Visible = False

End Sub

Sub BringToFront()
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
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select top 20 Tranno,tranno x from PMIS_vw_PRS where trantype = '" & WAREHOUSETYPE & "' order by Tranno asc")
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
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HD = gconDMIS.Execute("select top 20 tranno, tranno from PMIS_vw_PRS where trantype = '" & WAREHOUSETYPE & "' and tranno like '" & XXX & "%'")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select top 20 rono,tranno from PMIS_vw_PRS where trantype = '" & WAREHOUSETYPE & "' and rono is not null order by tranno asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HD = gconDMIS.Execute("select top 20 Rono, tranno from PMIS_vw_PRS where trantype = '" & WAREHOUSETYPE & "' and rono like '" & XXX & "%' order by tranno asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
        txtTranDescription.Text = SetSTOCKDESC2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
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
        rschek.Open "Select * from PMIS_Stockmas where type = 'P'and stockno = " & N2Str2Null(cboTranPartNo) & "", gconDMIS, adOpenKeyset, adLockReadOnly

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
    cmdTranDelete.Enabled = False
    InitParts
    On Error Resume Next
    cboTranPartNo.SetFocus
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "PARTS REQUISTION SLIP") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        If COMPANY_CODE = "HGC" Then
            With FrmCancelTransaction
                .lblTransaction_type = "PRS"
                .LblTransactionNo = txtTranNo.Text
                FrmCancelTransaction.Show 1
            End With

            If CANCEL_ANS = "NO" Then Exit Sub
        End If

        Dim PCURTISSQTY                                As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where tranno = " & N2Str2Null(RSORD_HD!tranno) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Partmas where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    If Null2String(RSORD_HD!Status) = "P" Then
                        If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                            PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                            SQL_STATEMENT = "update PMIS_Partmas set" & _
                                          " ONREQUEST = " & PCURTISSQTY & "," & _
                                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                          " lastupdate = '" & LOGDATE & "'" & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT

                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo & " CANCEL", "PRS", "")
                        Else
                            PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                            SQL_STATEMENT = "update PMIS_Partmas set" & _
                                          " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                          " lastupdate = '" & LOGDATE & "'" & _
                                          " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                            gconDMIS.Execute SQL_STATEMENT

                            Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo & " CANCEL", "PRS", "")
                        End If
                    End If
                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                  " status = 'C'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW_LogAudit "C", "PARTS REQUISTION SLIP", SQL_STATEMENT, labID, "Parts", txtTranNo, WAREHOUSETYPE, ""
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

        NEW_LogAudit "C", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""

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

Private Sub cmdEditTranDate_Click()
    If Function_Access(LOGID, "Acess_SYSTEM", "PARTS REQUISTION SLIP") = False Then Exit Sub
    txtTranDate.Enabled = True: txtTranDate.Locked = False
End Sub

Private Sub cmdPISNum_Click()
    With frmPMISPRIFormation
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
    frmPMISPRIFormation.Show 1
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "PARTS REQUISTION SLIP") = False Then Exit Sub


    On Error GoTo ErrorCode:

    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD = "" Or FILD = "No Entry" Then
        MsgBox "Posting of Transaction cannot proceed. Pls. Add Part(s).", vbCritical, "Confirm Posting "
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = True Then
        Dim PCURTISSQTY                                As Integer
        Dim RSTDAYTRANDUP, RSPARTMASDUP                As ADODB.Recordset

        Set RSTDAYTRANDUP = New ADODB.Recordset
        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where tranno = " & N2Str2Null(RSORD_HD!tranno) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
            RSTDAYTRANDUP.MoveFirst
            Do While Not RSTDAYTRANDUP.EOF
                Set RSPARTMASDUP = New ADODB.Recordset
                RSPARTMASDUP.Open "select STOCKNO,onhand,tissqty,ONREQUEST,S_ONREQUEST from PMIS_Partmas where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        SQL_STATEMENT = "update PMIS_Partmas set" & _
                                      " ONREQUEST = " & PCURTISSQTY & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT

                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, "PRS", "")
                    Else
                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) + N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                        SQL_STATEMENT = "update PMIS_Partmas set" & _
                                      " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                      " lastupdate = '" & LOGDATE & "'" & _
                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT

                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo, "PRS", "")
                    End If
                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                  " status = 'P'," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where id = " & RSTDAYTRANDUP!ID
                    gconDMIS.Execute SQL_STATEMENT
                    'NEW_LogAudit "P", "PARTS REQUISTION SLIP", SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""
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

        NEW_LogAudit "P", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""
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
    If Function_Access(LOGID, "Acess_Print", "PARTS REQUISTION SLIP") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If MsgQuestionBox("Parts Requisition Slip will be Printed. Are you sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        If Mid((RSORD_HD!REFPISNO), 3, 1) = "W" Then
            rptCustomerOrder.WindowTitle = "Parts Requisition Slip"
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Parts_OTC.rpt", "{PMIS_Ord_Hd.TYPE} = 'P' AND {PMIS_Ord_Hd.TRANTYPE} = 'PRS' and {PMIS_Ord_Hd.TRANNO} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Else
            rptCustomerOrder.WindowTitle = "Parts Requisition Slip"
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Parts.rpt", "({PMIS_vw_PRS.TYPE} = 'P' and {PMIS_vw_PRS.TRANTYPE} = 'PRS' and {PMIS_vw_PRS_Tran.TYPE} = 'P' and {PMIS_vw_PRS_Tran.TRANTYPE} = 'PRS' and {PMIS_vw_PRS_Tran.TRANNO} = " & N2Str2Null(txtTranNo.Text) & ")", DMIS_REPORT_Connection, 1
        End If

        NEW_LogAudit "V", "PARTS REQUISTION SLIP", "", labid, "", "TRAN NO: " & txtTranNo, "", ""
        Screen.MousePointer = 0
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    If RSORD_HD!TranType = "PRS" Then
    End If
    SendToBack
End Sub

Private Sub cmdTranCancel_Click()
    Picture1.Enabled = True
    fraDetails.Enabled = True
    picDetails.Enabled = True
    grdDetails.Enabled = True
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_vw_PRS_Tran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "PART NO: " & cboTranPartNo, WAREHOUSETYPE, labDetID
        ShowDeletedMsg
    End If
    Dim CNT                                            As Integer
    Dim RSTDAYTRANDUP                                  As ADODB.Recordset
    Set RSTDAYTRANDUP = New ADODB.Recordset
    RSTDAYTRANDUP.Open "select id,itemno from PMIS_vw_PRS_Tran where trantype = " & N2Str2Null(WAREHOUSETYPE) & " and tranno = " & N2Str2Null(RSORD_HD!tranno) & " order by itemno asc", gconDMIS
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

    Call NEW_LogAudit("E", "PART REQUISTION SLIP", SQL_STATEMENT, labid, "", "TRAN NO: " & txtTranNo & " - DETAILS REMOVE", "RIV", "")
    rsRefresh
    On Error Resume Next
    RSORD_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        Dim StockISSUE: StockISSUE = CheckRoToParts
        If Len(StockISSUE) > 0 Then
            If MsgBox("Same part number had been requested for this repair order." & vbCrLf & "Repair Order#  :" & txtRONO & vbCrLf & "Transaction#    :" & StockISSUE & vbCrLf & "Do you want to continue?", vbYesNo + vbInformation, "Parts Requisition") = vbNo Then
                On Error Resume Next
                cboTranPartNo.SetFocus
                Screen.MousePointer = 0
                Exit Sub
            End If
            'NEW_LogAudit "AA", "PARTS REQUISTION SLIP", "", labID, "Parts", "PART NO. " & cboTranPartNo, WAREHOUSETYPE, ""
        End If
    End If

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_vw_PRS_Tran where STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(RSORD_HD!tranno) & " order by itemno asc", gconDMIS
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
        MsgBox "Part number ( " & cboTranPartNo.Text & " ) has zero srp, please check parts master file.", vbInformation + vbOKOnly
        Exit Sub
    End If

    '===============================================================================================================
    'updating code:     jaa - 09062008      - Disallow this checking if the Transaction is Warranty or Internal Paid

    If Mid(txtReferencePIS, 5, 1) = "W" Or Mid(txtReferencePIS, 5, 1) = "I" Then
        'PROCEED
    Else
        If ORDTRANINVAMT < N2Str2IntZero(RSPARTMAS!dnp) Then
            If MsgBox("Your SRP of this Parts is below your Parts DNP. Proceed saving?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                'proceed saving
                'UPDATE BY   : MJP 09112008 02:35 AM
                'DESCRIPTION : SAVE THE LIMIT QUESTION IN THE AUDIT TRAIL
                Dim CRITICAL_QUESTION                  As String
                CRITICAL_QUESTION = "Your SRP of this Parts is below your Parts DNP. Proceed saving?"
                Call NEW_LogAudit("MP", "PARTS REQUISTION SLIP", CRITICAL_QUESTION, labid, "", "TRAN NO: " & txtTranNo & " PART NO: " & cboTranPartNo, "", "")
                'UPDATE BY   : MJP 09112008 02:35 AM
            Else
                MsgBox "Pls. check your Parts DNP and SRP in your Parts Master File", vbInformation, "Confirm DNP/SRP"
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
    End If
    '===============================================================================================================

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_vw_PRS_Tran " & _
                        "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out)" & _
                      " values ('P'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                      " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                      " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                      " " & ORDTRANUCOST & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "PART NO. " & cboTranPartNo, WAREHOUSETYPE, ""
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
        NEW_LogAudit "EE", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "PART NO: " & cboTranPartNo, WAREHOUSETYPE, labDetID
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
    NEW_LogAudit "E", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", txtTranNo, WAREHOUSETYPE, ""

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
    If Function_Access(LOGID, "Acess_Add", "PARTS REQUISTION SLIP") = False Then Exit Sub
    AddorEdit = "ADD"
    initMemvars
    PrisValidation
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "PARTS REQUISTION SLIP") = False Then Exit Sub
    AddorEdit = "EDIT"
    PREVORDTYPE = txtTranType.Text
    PREVORDNO = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
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
        MsgBox "Invalid Reference PRS Number!", vbCritical, "PRS Required!"
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
            If checkdup("P", txtTranType.Text, txtTranNo.Text) = True Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                txtTranNo.SetFocus
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!tranno))) Then
                If checkdup("P", txtTranType.Text, txtTranNo.Text) = True Then
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
                      " (trantype,tranno,trandate,custcode,custname,chargeto,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                      " values (" & VTXTTRANTYPE & ", " & VTXTTRANNO & ", " & VTXTTRANDATE & ", " & _
                      " " & VTXTCUSTCODE & ", " & VTXTCUSTNAME & ", " & VTXTCHARGETO & _
                        ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VCBOSALESMAN & ", " & VCBOSMNAME & _
                        ", " & VtxtTerms & ", " & VTXTTTLINVAMT & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNETINVAMT & ", " & VTXTRemarks & _
                        ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIN_PROCESS & "," & VTXTREFERENCEPIS & ", " & XSALES_ORIGIN & ", " & XSI_TYPE & ", " & XPAY_CLASS & ", " & XCHAR_YEAR & ", " & XCHAR_MONTH & ", " & XIS_SERIES & ", " & XTRACK_CODE & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "PARTS REQUISTION SLIP", SQL_STATEMENT, FindTransactionID(txtTranNo, "tranno", "PMIS_vw_PRS", "DETAILS", N2Str2Null("PRS"), "TRANTYPE"), "Parts", "TRAN NO: " & txtTranNo & " - " & txtReferencePIS, "PRS", ""
        ShowSuccessFullyAdded
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
        NEW_LogAudit "E", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo & " - " & txtReferencePIS, "PRS", ""

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
        'NEW_LogAudit "E", "PARTS REQUISTION SLIP", SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo & " - " & txtReferencePIS, "PRS", ""

        SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                      " trantype = " & VTXTTRANTYPE & "," & _
                      " trandate = " & VTXTTRANDATE & "," & _
                      " tranno = " & VTXTTRANNO & _
                      " where trantype = '" & PREVORDTYPE & "' and tranno = '" & PREVORDNO & "'"
        gconDMIS.Execute SQL_STATEMENT
        'NEW_LogAudit "E", "PARTS REQUISTION SLIP", SQL_STATEMENT, labID, "Parts", "TRAN NO: " & txtTranNo & " - " & txtReferencePIS, "PRS", ""
        ShowSuccessFullyUpdated
    End If


    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NEXTCUNTER & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' where modul = " & VTXTTRANTYPE
    Else
        rsRefresh
        RSORD_HD.Find "Tranno = " & VTXTTRANNO
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                       " netinvamt = " & ORD_TOTINVAMT & _
                       " where tranno = " & VTXTTRANNO & " and trantype = " & VTXTTRANTYPE
    End If

    'cmdCancel.Value = True

    If AddorEdit = "ADD" Then
        Picture1.Enabled = False
        fraDetails.Enabled = False
        grdDetails.Enabled = False
    Else
        Picture1.Enabled = True
        fraDetails.Enabled = True
    End If

    'fraDetails.Enabled = True
    rsRefresh
    RSORD_HD.Find "tranno = " & VTXTTRANNO
    'updated by mark..
    cmdCancel.Value = True

    FillGrid
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
        fraDetails.Enabled = False
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
    frmCustomerSearch_PRS.LABTYPE = "P"
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

    '    If Shift = 2 Then
    '        If KeyCode = vbKeyF1 Then
    '            If picDetails.Visible = False Then Exit Sub
    '                If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
    '                Unload frmALL_AuditInquiry
    '                frmALL_AuditInquiry.Show
    '                frmALL_AuditInquiry.ZOrder 0
    '                frmALL_AuditInquiry.Caption = "Audit Inquiry (Parts Issuance)"
    '                Call frmALL_AuditInquiry.DisplayHistory(labID, "PARTS REQUISTION SLIP")
    '        End If
    '    End If

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Picture1.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PARTS REQUISTION SLIP)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "PARTS REQUISTION SLIP")

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
                    BLNF3CNTR = True
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
            If Picture1.Visible = True Then
                If Function_Access(LOGID, "Acess_UNPOST", "PARTS REQUISTION SLIP") = False Then Exit Sub
                If Null2String(RSORD_HD!Status) = "P" Then
                    If MsgQuestionBox("Are you sure you want to UnPost this Transaction?", "UnPost Transaction") = True Then
                        Dim PCURTISSQTY                As Integer
                        Dim RSTDAYTRANDUP, RSPARTMASDUP As ADODB.Recordset

                        Set RSTDAYTRANDUP = New ADODB.Recordset
                        RSTDAYTRANDUP.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where tranno = " & N2Str2Null(RSORD_HD!tranno) & " and trantype = " & N2Str2Null(RSORD_HD!TranType), gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not RSTDAYTRANDUP.EOF And Not RSTDAYTRANDUP.BOF Then
                            RSTDAYTRANDUP.MoveFirst
                            Do While Not RSTDAYTRANDUP.EOF
                                Set RSPARTMASDUP = New ADODB.Recordset
                                RSPARTMASDUP.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Partmas where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD), gconDMIS
                                If Not RSPARTMASDUP.EOF And Not RSPARTMASDUP.BOF Then
                                    If Null2String(RSORD_HD!SALES_ORIGIN) = "W" Then
                                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                        SQL_STATEMENT = "update PMIS_Partmas set" & _
                                                      " ONREQUEST = " & PCURTISSQTY & "," & _
                                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                      " lastupdate = '" & LOGDATE & "'" & _
                                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT
                                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo & " UNPOST", "PRS", "")
                                    Else
                                        PCURTISSQTY = N2Str2IntZero(RSPARTMASDUP!S_ONREQUEST) - N2Str2Zero(RSTDAYTRANDUP!TRANQTY)
                                        SQL_STATEMENT = "update PMIS_Partmas set" & _
                                                      " S_ONREQUEST = " & PCURTISSQTY & "," & _
                                                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                      " lastupdate = '" & LOGDATE & "'" & _
                                                      " where STOCKNO = " & N2Str2Null(RSTDAYTRANDUP!STOCK_ORD)
                                        gconDMIS.Execute SQL_STATEMENT

                                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Date2Null(RSTDAYTRANDUP!STOCK_ORD), "STOCKNO", "PMIS_STOCKMAS"), "", "TRAN NO: " & txtTranNo & " UNPOST", "PRS", "")
                                    End If
                                    SQL_STATEMENT = "update PMIS_vw_PRS_Tran set" & _
                                                  " status = 'N'," & _
                                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                  " lastupdate = '" & LOGDATE & "'" & _
                                                  " where id = " & RSTDAYTRANDUP!ID
                                    gconDMIS.Execute SQL_STATEMENT
                                    'NEW_LogAudit "U", "PARTS REQUISTION SLIP", SQL_STATEMENT, labID, "Parts", txtTranNo, WAREHOUSETYPE, ""
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
                        NEW_LogAudit "U", "PARTS REQUISTION SLIP", SQL_STATEMENT, labid, "Parts", "TRAN NO: " & txtTranNo, WAREHOUSETYPE, ""

                        rsRefresh
                        On Error Resume Next
                        RSORD_HD.Find "id = " & labid.Caption
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
    rsRefresh
    On Error Resume Next
    If Not RSORD_HD.EOF And Not RSORD_HD.BOF Then RSORD_HD.MoveLast
    StoreMemVars

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
            cmdTranDelete.Enabled = True
            BringToFront
            StorePartsEntry (FILD)
            Picture1.Enabled = False
            fraDetails.Enabled = False
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
    If IsDate(txtTranDate) = False Then: txtTranDate = "": Exit Sub
    txtTranDate.Text = Format(txtTranDate.Text, "mm/dd/yyyy")
End Sub

Private Sub txtTranNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTranNo_LostFocus()
    Dim RSFINDDUP                                      As ADODB.Recordset
    If AddorEdit = "ADD" Then
        If checkdup("P", txtTranType.Text, txtTranNo.Text) = True Then
            MsgSpeechBox "Transaction No. already exist!"
            On Error Resume Next
            txtTranNo.SetFocus
            Exit Sub
        End If
    Else
        If LTrim(RTrim(txtTranNo)) <> LTrim(RTrim(Null2String(RSORD_HD!tranno))) Then
            If checkdup("P", txtTranType.Text, txtTranNo.Text) = True Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                txtTranNo.SetFocus
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
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
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


Public Sub PrisValidation()
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
End Sub
