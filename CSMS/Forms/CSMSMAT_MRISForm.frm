VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPMISMAT_MRISForms_CSMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Requisition Slip Data Entry"
   ClientHeight    =   6045
   ClientLeft      =   1110
   ClientTop       =   2520
   ClientWidth     =   11430
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CSMSMAT_MRISForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2685
      ScaleHeight     =   870
      ScaleWidth      =   8715
      TabIndex        =   93
      Top             =   5115
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   96
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   97
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   103
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   104
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   98
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":213D
         Style           =   1  'Graphical
         TabIndex        =   99
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   95
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   94
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":2EF4
         Style           =   1  'Graphical
         TabIndex        =   100
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   101
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox fraAddTran 
      Height          =   3525
      Left            =   4620
      ScaleHeight     =   3465
      ScaleWidth      =   4545
      TabIndex        =   43
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":3B49
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":3C9B
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Delete Entry"
         Top             =   2460
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":3FC6
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":4118
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Cancel Entry"
         Top             =   2460
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
         TabIndex        =   19
         ToolTipText     =   "Input price of item. Do not use comma and peso sign (e.g.300, 26)"
         Top             =   1440
         Width           =   1515
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
         TabIndex        =   17
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
         TabIndex        =   21
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
         TabIndex        =   18
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
         MaxLength       =   4
         TabIndex        =   15
         ToolTipText     =   "Type item number (e.g. 0001)"
         Top             =   60
         Width           =   615
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
         TabIndex        =   16
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":4456
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":45A8
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Save Entry"
         Top             =   2460
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
         TabIndex        =   75
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
         TabIndex        =   47
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
   Begin wizButton.cmd cmdAddTran 
      Height          =   3660
      Left            =   4560
      TabIndex        =   62
      Top             =   975
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6456
      TX              =   "cmd1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "CSMSMAT_MRISForm.frx":48F8
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
      Left            =   2700
      ScaleHeight     =   285
      ScaleWidth      =   8715
      TabIndex        =   78
      Top             =   4710
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
         Top             =   30
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5985
      Left            =   60
      TabIndex        =   69
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
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   960
         Width           =   2475
      End
      Begin VB.OptionButton optRONo 
         Caption         =   "RO Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   71
         Top             =   630
         Width           =   2385
      End
      Begin VB.OptionButton optTranno 
         Caption         =   "Transaction No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   70
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   4575
         Left            =   60
         TabIndex        =   73
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
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CSMSMAT_MRISForm.frx":4914
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
         TabIndex        =   74
         Top             =   150
         Width           =   1455
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
         TabIndex        =   76
         Top             =   960
         Width           =   2685
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F1 - Assign MRS Number"
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
         TabIndex        =   85
         Top             =   60
         Width           =   2295
      End
      Begin VB.CommandButton cmdPISNum 
         Caption         =   "..."
         Height          =   375
         Left            =   7080
         TabIndex        =   84
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
         Left            =   5340
         TabIndex        =   1
         Text            =   "MRSGI07A001"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   60
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
         TabIndex        =   10
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
         TabIndex        =   13
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
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
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
         Picture         =   "CSMSMAT_MRISForm.frx":4A76
         ScaleHeight     =   405
         ScaleWidth      =   435
         TabIndex        =   60
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
         TabIndex        =   12
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   60
            Width           =   1395
         End
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
         TabIndex        =   77
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
         TabIndex        =   68
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
         TabIndex        =   67
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
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2820
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
      Location        =   "http:///"
   End
   Begin wizButton.cmd cmdSignatories 
      Height          =   2505
      Left            =   4650
      TabIndex        =   63
      Top             =   2280
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   4419
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "CSMSMAT_MRISForm.frx":77B2
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":77CE
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":7920
         Style           =   1  'Graphical
         TabIndex        =   89
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
         TabIndex        =   24
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
         TabIndex        =   25
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
         TabIndex        =   23
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   90
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   9720
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   90
      Top             =   5055
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":7C86
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":7DD8
         Style           =   1  'Graphical
         TabIndex        =   91
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
         MouseIcon       =   "CSMSMAT_MRISForm.frx":8116
         MousePointer    =   99  'Custom
         Picture         =   "CSMSMAT_MRISForm.frx":8268
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   1890
      Left            =   2700
      ScaleHeight     =   1890
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
         Height          =   1815
         Left            =   30
         TabIndex        =   14
         Top             =   60
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   3201
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
End
Attribute VB_Name = "frmPMISMAT_MRISForms_CSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOrd_Hd, rsTdayTran, rsPartMas                    As ADODB.Recordset
Attribute rsTdayTran.VB_VarUserMemId = 1073938432
Attribute rsPartMas.VB_VarUserMemId = 1073938432
Dim rsSalesMan, rsCunter, rsProfile                    As ADODB.Recordset
Attribute rsSalesMan.VB_VarUserMemId = 1073938435
Attribute rsCunter.VB_VarUserMemId = 1073938435
Attribute rsProfile.VB_VarUserMemId = 1073938435
Dim rsSignatories                                      As ADODB.Recordset
Attribute rsSignatories.VB_VarUserMemId = 1073938438
Dim rsREPOR, rsCustomer                                As ADODB.Recordset
Attribute rsREPOR.VB_VarUserMemId = 1073938439
Attribute rsCustomer.VB_VarUserMemId = 1073938439
Dim kcnt                                               As Integer
Attribute kcnt.VB_VarUserMemId = 1073938441
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938442
Dim ORD_TOTUPRICE, ORD_TOTINVAMT, ORD_TOTVAT, ORD_TOTQTY As Double
Attribute ORD_TOTUPRICE.VB_VarUserMemId = 1073938443
Attribute ORD_TOTINVAMT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTVAT.VB_VarUserMemId = 1073938443
Attribute ORD_TOTQTY.VB_VarUserMemId = 1073938443
Dim PrevOrdType, PrevOrdNo                             As String
Attribute PrevOrdType.VB_VarUserMemId = 1073938447
Attribute PrevOrdNo.VB_VarUserMemId = 1073938447

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
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDSTOCKNO(cboTranPartNo.Text)
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
    If Function_Access(LOGID, "Acess_CancelEntry", "MATERIALS REQUISITION SLIP") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If MsgQuestionBox("Are you sure you want to Cancel this Transaction?", "Cancel Transaction") = True Then
        Dim PCurOnHand, PCurTISSQTY, PCurIssuances     As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    If Null2String(rsOrd_Hd!Status) = "P" Then
                        If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                            PCurTISSQTY = N2Str2IntZero(rsPartmasDup!ONREQUEST) - N2Str2Zero(rsTdaytranDup!tranqty)
                            gconDMIS.Execute "update PMIS_Stockmas set" & _
                                           " ONREQUEST = " & PCurTISSQTY & "," & _
                                           " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                           " lastupdate = '" & LOGDATE & "'" & _
                                           " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        Else
                            PCurTISSQTY = N2Str2IntZero(rsPartmasDup!S_ONREQUEST) - N2Str2Zero(rsTdaytranDup!tranqty)
                            gconDMIS.Execute "update PMIS_Stockmas set" & _
                                           " S_ONREQUEST = " & PCurTISSQTY & "," & _
                                           " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                           " lastupdate = '" & LOGDATE & "'" & _
                                           " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                        End If
                    End If
                    gconDMIS.Execute "update PMIS_vw_PRS_Tran set" & _
                                   " status = 'C'," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdaytranDup!ID
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        LogAudit "C", "PRIS", txtTranNo
        rsRefresh
        On Error Resume Next
        rsOrd_Hd.Find "id =" & labID.Caption
        StoreMemVars
    End If
    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPISNum_Click()
    With frmPMISMAT_MRIFormation_CSMS
        If AddorEdit = "EDIT" Then
            .txtedit = "EDIT"
        Else
            .txtedit = ""
        End If
        .lbl2 = Mid(txtReferencePIS, 3, 1)
        .lbl3 = Mid(txtReferencePIS, 4, 1)
        .lbl4 = Mid(txtReferencePIS, 5, 1)
        ' .lbl6_7 = Mid(txtReferencePIS, 6, 2)
        ' .lbl8 = Mid(txtReferencePIS, 8, 1)
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
    frmPMISMAT_MRIFormation_CSMS.Show 1
    On Error Resume Next
    txtCustName.SetFocus
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "MATERIALS REQUISITION SLIP") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    If MsgQuestionBox("Are you sure you want to Post this Transaction?", "Post Transaction") = True Then
        Dim PCurOnHand, PCurTISSQTY, PCurIssuances     As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onhand,tissqty,ONREQUEST,S_ONREQUEST from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                        PCurTISSQTY = N2Str2IntZero(rsPartmasDup!ONREQUEST) + N2Str2Zero(rsTdaytranDup!tranqty)
                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                       " ONREQUEST = " & PCurTISSQTY & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                    Else
                        PCurTISSQTY = N2Str2IntZero(rsPartmasDup!S_ONREQUEST) + N2Str2Zero(rsTdaytranDup!tranqty)
                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                       " S_ONREQUEST = " & PCurTISSQTY & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                    End If
                    gconDMIS.Execute "update PMIS_vw_PRS_Tran set" & _
                                   " status = 'P'," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdaytranDup!ID
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If

        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        LogAudit "P", "MRIS", txtTranNo
        rsRefresh
        On Error Resume Next
        rsOrd_Hd.Find "id =" & labID.Caption
        StoreMemVars
    End If
    Set rsTdaytranDup = Nothing
    Set rsPartmasDup = Nothing

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS REQUISITION SLIP") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
     oVoice.Speak "Print Materials Requisition Slip, Are You Sure?", SVSFlagsAsync
    If MsgBox("Print Materials Requisition Slip, Are You Sure?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        If Val(txtDS1) = 0 Then
            PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "RequisitionSlip_Materials.rpt", "{PMIS_vw_PRS.type} = 'M' and {PMIS_vw_PRS.trantype} = 'MRS' and {PMIS_vw_PRS.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        'Else
        '    PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "InventoryRequisition_IssuanceSlip_WithDiscount.rpt", "", DMIS_REPORT_Connection, 1
        End If
    End If
'    If rsOrd_Hd!TRANTYPE = "PRS" Then
'        cmdSignatories.Visible = True
'        cmdSignatories.ZOrder 0
'        fraSignatories.Visible = True
'        fraSignatories.ZOrder 0
'        Set rsSignatories = New ADODB.Recordset
'        rsSignatories.Open "Select * from Signatories", gconDMIS
'        If Not rsSignatories.EOF And Not rsSignatories.BOF Then
'            txtPreparedBy.Text = Null2String(rsSignatories!preparedby)
'            txtIssuedBy.Text = Null2String(rsSignatories!issuedby)
'            txtRequestedBy.Text = Null2String(rsSignatories!requestedby)
'            txtApprovedBy.Text = Null2String(rsSignatories!approvedby)
'            On Error Resume Next
'            txtRequestedBy.SetFocus
'        End If
        LogAudit "V", "MRIS", txtTranNo
        Set rsSignatories = Nothing
    'End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrintRIV_Click()
    If rsOrd_Hd!TRANTYPE = "MRS" Then
        'PRSPRINTING
    End If
    SendToBack
End Sub

Private Sub cmdTranCancel_Click()
'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    SendToBack
    StoreMemVars

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdTranDelete_Click()

'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Dim PnoOnhand, PnoTISSQTY, PnoIssuances            As Integer
    If labDetID.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Materials Entry") = True Then
        gconDMIS.Execute "delete from PMIS_vw_PRS_Tran where id = " & labDetID.Caption
        ShowDeletedMsg
    End If
    Dim cnt                                            As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "select id,itemno from PMIS_vw_PRS_Tran where [TYPE] = 'M' and trantype = " & N2Str2Null(WAREHOUSETYPE) & " and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        cnt = 0
        Do While Not rsTdaytranDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update PMIS_vw_PRS_Tran set itemno = " & Format(cnt, "0000") & " where id = " & rsTdaytranDup!ID
            rsTdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    gconDMIS.Execute "update PMIS_vw_PRS set" & _
                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                   " netinvamt = " & ORD_TOTINVAMT & _
                   " where id = " & labID.Caption
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
        MsgSpeechBox "Warning: Materials Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_vw_PRS_Tran where [TYPE] = 'M' and STOCK_ORD = '" & cboTranPartNo.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(rsOrd_Hd!Tranno) & " order by itemno asc", gconDMIS
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
    Dim ORDUNIT                                        As String
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

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into PMIS_vw_PRS_Tran " & _
                         "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,tranuprice,lastupdate,usercode,status,in_out)" & _
                       " values ('M'," & ORDTRANDATE & ", " & ORDTRANTYPE & ", " & ORDTRANNO & "," & _
                       " " & ORDITEMNO & "," & ORDSTOCK_ORD & "," & _
                       " " & ORDSTOCK_SUP & ", " & ORDTRANQTY & "," & _
                       " " & ORDTRANUCOST & "," & ORDTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & ORDSTATUS & ", " & ORDIN_OUT & ")"
    Else
        gconDMIS.Execute "update PMIS_vw_PRS_Tran set" & _
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
    End If
    cleargrid grdDetails
    FillDetails
    gconDMIS.Execute "update PMIS_vw_PRS set" & _
                   " totalqty = " & ORD_TOTQTY & "," & _
                   " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                   " netinvamt = " & ORD_TOTINVAMT & _
                   " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsOrd_Hd.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then cmdAddTran_Click Else cmdTranCancel.Value = True
    Exit Sub

Errorcode:
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
    '=================================
    'updating code:     JAA - 12052007
    'To disable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = False
    '=================================
    InitMemvars
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    '=================================
    'updating code:     JAA - 12052007
    'To enable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = True
    '=================================
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
    AddorEdit = "EDIT"
    PrevOrdType = txtTranType.Text
    PrevOrdNo = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    '=================================
    'updating code:     JAA - 12052007
    'To disable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = False
    '=================================
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

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Private Sub cmdFirst_Click()
    rsOrd_Hd.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsOrd_Hd.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsOrd_Hd.MoveNext
    If rsOrd_Hd.EOF Then
        rsOrd_Hd.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsOrd_Hd.MovePrevious
    If rsOrd_Hd.BOF Then
        rsOrd_Hd.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim NextCunter                                     As String
    Dim rsCunter, rsfindDup                            As ADODB.Recordset
    Dim xSALES_ORIGIN, xSI_TYPE, xPAY_CLASS, xCHAR_YEAR, xCHAR_MONTH, xIS_SERIES, xTRACK_CODE As String

    Dim VcboSalesMan, VcboSMName, VTXTTranType         As String
    Dim VTXTTranNo, VTXTTranDate, VtxtCustCode         As String
    Dim VtxtCustName, VTXTChargeTo, VTXTRONO, VTXTREP_OR As String
    Dim VtxtTerms                                      As String
    Dim VTXTTTLInvAmt, VTXTDS1                         As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNetInvAmt                     As Double
    Dim VTXTRemarks, VStatus, Vusercode                As String
    Dim VLastUpdate                                    As String
    Dim VIn_Process                                    As String
    Dim vtxtReferencePIS                               As String

    If Trim(txtReferencePIS.Text) = "" Or Len(txtReferencePIS.Text) < 10 Then
        MsgBox "Invalid Reference MRS Number!", vbCritical, "PRS Required!"
        Exit Sub
    End If

    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction No. must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select trantype,tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Transaction No. already exist!"
                On Error Resume Next
                Exit Sub
            End If
        End If
    End If
    If txtTranDate.Text = "" Or IsDate(txtTranDate.Text) = False Then
        MsgSpeechBox "Invalid Transaction Date!"
        On Error Resume Next
        txtTranDate.SetFocus
        Exit Sub
    End If
    VcboSalesMan = "NULL"
    VcboSMName = "NULL"

    NextCunter = NumericVal(txtTranNo.Text) + 1

    VTXTTranType = N2Str2Null(txtTranType.Text)
    VTXTTranNo = N2Str2Null(txtTranNo.Text)
    VTXTTranDate = N2Date2Null(txtTranDate.Text)
    VtxtCustCode = N2Str2Null(txtCustCode.Text)
    VtxtCustName = N2Str2Null(txtCustName.Text)
    vtxtReferencePIS = N2Str2Null(txtReferencePIS.Text)

    VIn_Process = "'Y'"
    VTXTChargeTo = "'VAR'"
    VTXTRONO = N2Str2Null(txtRONO.Text)
    If Len(txtRONO.Text) = 10 Then
        VTXTREP_OR = "'" & Left(txtRONO.Text, 1) & "-" & Right(txtRONO.Text, 8) & "'"
    Else
        VTXTREP_OR = "NULL"
    End If
    VtxtTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLInvAmt = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetInvAmt = NumericVal(txtNetInvAmt.Text)
    If txtRemarks.Text = "Pls Type Your Message Here!" Then VTXTRemarks = "NULL" Else VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    VStatus = "'N'"
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"

    xSALES_ORIGIN = N2Str2Null(Mid(txtReferencePIS, 3, 1))
    xSI_TYPE = N2Str2Null(Mid(txtReferencePIS, 4, 1))
    xPAY_CLASS = N2Str2Null(Mid(txtReferencePIS, 5, 1))
    xCHAR_YEAR = N2Str2Null(Mid(txtReferencePIS, 6, 2))
    xCHAR_MONTH = N2Str2Null(Mid(txtReferencePIS, 8, 1))
    xIS_SERIES = N2Str2Null(Mid(txtReferencePIS, 9, 3))
    xTRACK_CODE = N2Str2Null(Mid(txtReferencePIS, 12, 1))
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into PMIS_vw_PRS" & _
                       " ([TYPE],trantype,tranno,trandate,custcode,custname,chargeto,rono,rep_or,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate,In_Process,REFPISNO,SALES_ORIGIN,SI_TYPE,PAY_CLASS,CHAR_YEAR,CHAR_MONTH,IS_SERIES,TRACK_CODE)" & _
                       " values ('M'," & VTXTTranType & ", " & VTXTTranNo & ", " & VTXTTranDate & ", " & _
                       " " & VtxtCustCode & ", " & VtxtCustName & ", " & VTXTChargeTo & _
                         ", " & VTXTRONO & "," & VTXTREP_OR & ", " & VcboSalesMan & ", " & VcboSMName & _
                         ", " & VtxtTerms & ", " & VTXTTTLInvAmt & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetInvAmt & ", " & VTXTRemarks & _
                         ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & "," & VIn_Process & "," & vtxtReferencePIS & ", " & xSALES_ORIGIN & ", " & xSI_TYPE & ", " & xPAY_CLASS & ", " & xCHAR_YEAR & ", " & xCHAR_MONTH & ", " & xIS_SERIES & ", " & xTRACK_CODE & ")"
        LogAudit "A", "MRIS", txtTranNo
    Else

        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " tranno = " & VTXTTranNo & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " custcode = " & VtxtCustCode & "," & _
                       " custname = " & VtxtCustName & "," & _
                       " chargeto = " & VTXTChargeTo & "," & _
                       " rono = " & VTXTRONO & "," & _
                       " rep_or = " & VTXTREP_OR & "," & _
                       " salesman = " & VcboSalesMan & "," & _
                       " smname = " & VcboSMName & "," & _
                       " terms = " & VtxtTerms & "," & _
                       " ttlinvamt = " & VTXTTTLInvAmt & "," & _
                       " ds1 = " & VTXTDS1 & "," & _
                       " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                       " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                       " netinvamt = " & VTXTNetInvAmt & "," & _
                       " remarks = " & VTXTRemarks & ", " & _
                       " status = " & VStatus & ", " & _
                       " usercode = " & Vusercode & ", " & _
                       " In_Process = " & VIn_Process & ", " & _
                       " REFPISNO = " & vtxtReferencePIS & ", " & _
                       " lastupdate = " & VLastUpdate & _
                       " where id = " & labID.Caption

        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " SALES_ORIGIN = " & xSALES_ORIGIN & "," & _
                       " SI_TYPE = " & xSI_TYPE & "," & _
                       " PAY_CLASS = " & xPAY_CLASS & "," & _
                       " CHAR_YEAR = " & xCHAR_YEAR & "," & _
                       " CHAR_MONTH = " & xCHAR_MONTH & "," & _
                       " IS_SERIES = " & xIS_SERIES & "," & _
                       " TRACK_CODE = " & xTRACK_CODE & "" & _
                       " where id = " & labID.Caption

        gconDMIS.Execute "update PMIS_vw_PRS_Tran set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " tranno = " & VTXTTranNo & _
                       " where trantype = '" & PrevOrdType & "' and tranno = '" & PrevOrdNo & "'"
        LogAudit "E", "MRIS", txtTranNo
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NextCunter & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' WHERE TYPE = 'M' and modul = " & VTXTTranType
    Else
        rsRefresh
        rsOrd_Hd.Find "Tranno = " & VTXTTranNo
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                       " ttlinvamt = " & ORD_TOTUPRICE & "," & _
                       " netinvamt = " & ORD_TOTINVAMT & _
                       " WHERE TYPE = 'M' and tranno = " & VTXTTranNo & " and trantype = " & VTXTTranType
    End If
    '=================================
    'updating code:     JAA - 12052007
    'To enable lstOrd_Hd Listview for Adding and Editing Transaction
    fraDetails.Enabled = True
    '=================================
    rsRefresh
    rsOrd_Hd.Find "tranno = " & VTXTTranNo
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
'frmAllCustomer.Show
    frmPMISMAT_CustomerSearch.Show 1
End Sub

Private Sub Command2_Click()
    cmdPISNum_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim fild                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    fild = grdDetails.Text
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
        Case vbKeyF1
            If Picture1.Visible = False Then Command2.Value = True
        Case vbKeyF2
            If Command1.Visible = True And Command1.Enabled = True Then Command1.Value = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsOrd_Hd!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!Status) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change..."
                Else
                    cmdAddTran_Click
                End If
            End If
        Case vbKeyF4
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!Status) <> "P" And Null2String(rsOrd_Hd!Status) <> "C" And Null2String(rsOrd_Hd!Status) <> "B" Then
                        grdDetails_DblClick
                    End If
                End If
            End If
        Case vbKeyF5
            If fild <> "" And fild <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(rsOrd_Hd!Status) <> "P" And Null2String(rsOrd_Hd!Status) <> "C" And Null2String(rsOrd_Hd!Status) <> "B" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF8
            If cmdPost.Enabled = True Then cmdPost.Value = True
        Case vbKeyF12
            If Picture1.Visible = True Then
                If Null2String(rsOrd_Hd!Status) = "P" Then
                    If Function_Access(LOGID, "Acess_UNPOST", "MATERIALS REQUISITION SLIP") = False Then Exit Sub
                    If MsgQuestionBox("Are you sure you want to UnPost this Transaction?", "UnPost Transaction") = True Then
                        Dim PCurOnHand, PCurTISSQTY, PCurIssuances As Integer
                        Dim rsTdaytranDup, rsPartmasDup As ADODB.Recordset

                        Set rsTdaytranDup = New ADODB.Recordset
                        rsTdaytranDup.Open "select id,trantype,tranno,STOCK_ORD,tranqty from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(rsOrd_Hd!Tranno) & " and trantype = " & N2Str2Null(rsOrd_Hd!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
                        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                            rsTdaytranDup.MoveFirst
                            Do While Not rsTdaytranDup.EOF
                                Set rsPartmasDup = New ADODB.Recordset
                                rsPartmasDup.Open "select STOCKNO,onhand,ONREQUEST,S_ONREQUEST from PMIS_Stockmas where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD), gconDMIS
                                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                                    If Null2String(rsOrd_Hd!SALES_ORIGIN) = "W" Then
                                        PCurTISSQTY = N2Str2IntZero(rsPartmasDup!ONREQUEST) - N2Str2Zero(rsTdaytranDup!tranqty)
                                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                                       " ONREQUEST = " & PCurTISSQTY & "," & _
                                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                       " lastupdate = '" & LOGDATE & "'" & _
                                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    Else
                                        PCurTISSQTY = N2Str2IntZero(rsPartmasDup!S_ONREQUEST) - N2Str2Zero(rsTdaytranDup!tranqty)
                                        gconDMIS.Execute "update PMIS_Stockmas set" & _
                                                       " S_ONREQUEST = " & PCurTISSQTY & "," & _
                                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                       " lastupdate = '" & LOGDATE & "'" & _
                                                       " where STOCKNO = " & N2Str2Null(rsTdaytranDup!STOCK_ORD)
                                    End If
                                    gconDMIS.Execute "update PMIS_vw_PRS_Tran set" & _
                                                   " status = 'N'," & _
                                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                                   " lastupdate = '" & LOGDATE & "'" & _
                                                   " where id = " & rsTdaytranDup!ID
                                End If
                                rsTdaytranDup.MoveNext
                            Loop
                        End If
                        gconDMIS.Execute "update PMIS_vw_PRS set" & _
                                       " status = 'N'," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where id = " & labID.Caption
                        LogAudit "U", "PRIS", txtTranNo
                        rsRefresh
                        On Error Resume Next
                        rsOrd_Hd.Find "id =" & labID.Caption
                        StoreMemVars
                    End If
                    Set rsTdaytranDup = Nothing
                    Set rsPartmasDup = Nothing
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1: PMIS_ORDER_SHOW = True
    textSearch.Text = "":    'Picture5.ZOrder 0
    Command1.Enabled = False
    Command1.Visible = False
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    InitMemvars
    txtTranUPrice.Enabled = False
    rsRefresh
    'If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then rsOrd_Hd.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    If WAREHOUSETYPE = "MRS" Then
        Me.Caption = "Materials Requistion Slip"
        Set rsOrd_Hd = New ADODB.Recordset
        rsOrd_Hd.Open "select * from PMIS_vw_PRS where [TYPE] = 'M' and trantype = 'MRS' order by tranno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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

Sub InitMemvars()
    If WAREHOUSETYPE = "MRS" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where modul = 'MRS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtRONO.Enabled = True
        txtTerms.Enabled = False
        'cboSalesMan.Enabled = False
        'cboSMName.Enabled = False
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
    txtRemarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    InitCbo
    InitGrid
    cleargrid grdDetails
    SendToBack
    'cboChargeTo.Enabled = True
    InitSignatories
End Sub

Sub InitSignatories()
    txtPreparedBy.Text = ""
    txtIssuedBy.Text = ""
    txtRequestedBy.Text = ""
    txtApprovedBy.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
        labID.Caption = rsOrd_Hd!ID
        txtTranType.Text = Null2String(rsOrd_Hd!TRANTYPE)
        cboSMName.Enabled = True
        txtTranNo.Text = Null2String(rsOrd_Hd!Tranno)
        txtTranDate.Text = Null2String(rsOrd_Hd!trandate)
        txtCustCode.Text = Null2String(rsOrd_Hd!custcode)
        txtCustName.Text = Null2String(rsOrd_Hd!custname)
        txtReferencePIS.Text = Null2String(rsOrd_Hd!refpisno)

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
        txtRONO.Text = Null2String(rsOrd_Hd!rono)
        cboSMName.Text = FillSalesMan(Null2String(rsOrd_Hd!salesman))
        txtTerms.Text = Null2String(rsOrd_Hd!Terms)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(rsOrd_Hd!ds1)
        txtDS_Desc1.Text = Null2String(rsOrd_Hd!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!ds_amt1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsOrd_Hd!netinvamt))
        txtRemarks.Text = Null2String(rsOrd_Hd!remarks)
        If Null2String(rsOrd_Hd!Status) = "C" Then
            labPosted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(rsOrd_Hd!Status) = "B" Then
            labPosted.Caption = "BILLED OUT"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
        ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
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
    kcnt = 0
    ORD_TOTUPRICE = 0
    ORD_TOTINVAMT = 0
    ORD_TOTVAT = 0
    ORD_TOTQTY = 0
    Dim STOCKDESCription                               As String
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_vw_PRS_Tran where [TYPE] = 'M' and tranno = " & N2Str2Null(txtTranNo.Text) & " and trantype = " & N2Str2Null(txtTranType.Text) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        cboChargeTo.Enabled = False
        Screen.MousePointer = 11
        rsTdayTran.MoveFirst
        Do While Not rsTdayTran.EOF
            kcnt = kcnt + 1
            If txtTranType.Text = "ADB" Then
                STOCKDESCription = Null2String(rsTdayTran!STOCK_SUP)
            Else
                STOCKDESCription = SetSTOCKDESC(Null2String(rsTdayTran!STOCK_SUP))
            End If
            grdDetails.AddItem rsTdayTran!ID & Chr(9) & Null2String(rsTdayTran!itemno) & Chr(9) & _
                               Null2String(rsTdayTran!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(rsTdayTran!tranqty) & Chr(9) & _
                               Format(N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(rsTdayTran!tranqty)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
            rsTdayTran.MoveNext
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
        If kcnt <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cboChargeTo.Enabled = True
        cleargrid grdDetails
    End If
End Sub

Function FillSalesMan(XXX As String) As String
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & XXX & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        FillSalesMan = Null2String(rsSalesMan!signname)
        cboSalesMan.Text = Null2String(rsSalesMan!empno)
    Else
        cboSalesMan.Text = ""
    End If
End Function

Sub InitCbo()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select id,STOCKNO,STOCKDESC from PMIS_Stockmas WHERE TYPE = 'M' order by STOCKNO ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboTranPartNo.Clear
        Do While Not rsPartMas.EOF
            cboTranPartNo.AddItem Null2String(rsPartMas!STOCKNO)
            rsPartMas.MoveNext
        Loop
    End If
    FillCboSalesMan
End Sub

Sub FillCboSalesMan()
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan order by signname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        rsSalesMan.MoveFirst: cboSalesMan.Clear: cboSMName.Clear
        Do While Not rsSalesMan.EOF
            cboSalesMan.AddItem Null2String(rsSalesMan!empno)
            cboSMName.AddItem Null2String(rsSalesMan!signname)
            rsSalesMan.MoveNext
        Loop
    Else
        cboSalesMan.Clear: cboSMName.Clear
    End If
End Sub

Sub SetCustInfo(rep As String)
    rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set rsREPOR = New ADODB.Recordset
    'rsREPOR.Open "select rep_or,niym,acct_no,invoice from CSMS_repor where rep_or = '" & rep & "'", gconDMIS
    rsREPOR.Open "select rep_or,niym,acct_no,invoice from CSMS_repor where rep_or = '" & txtRONO.Text & "'", gconDMIS
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        If Null2String(rsREPOR!invoice) <> "" Then
            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                 " Parts Request for this Repair Order is Critical!", vbCritical, "Critical Issue!"
            If MsgBox("Would You Like to Continue?", vbQuestion + vbYesNo, "Continue...") = vbNo Then
                On Error Resume Next
                txtRONO.SetFocus
                Exit Sub
            Else
                MsgBox "Pls. Input Your Notes/Reason in Remarks Field..."
                On Error Resume Next
                txtRemarks.SetFocus
            End If
        End If
        txtCustName.Text = Null2String(rsREPOR!niym)
        txtCustCode.Text = Null2String(rsREPOR!ACCT_NO)
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""
    End If
End Sub

Function SetSTOCKDESC(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select STOCKNO,STOCKDESC,srp,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKDESC = Null2String(rsPartMas!STOCKDESC)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                'If cboChargeTo.Text = "WARRANTY" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                'ElseIf cboChargeTo.Text = "COMPANY" Then
                '   txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        End If
    Else
        If WAREHOUSETYPE = "ADB" Then
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC = Null2String(rsPartMas!DESCRIPTIO)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
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
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select STOCKNUMBER,descriptio,srp from PMIS_DNPP where STOCKNUMBER= " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetSTOCKDESC2 = Null2String(rsPartMas!DESCRIPTIO)
            txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
            txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
        Else
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = " & N2Str2Null(cboTranPartNo.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = 0
                txtTranUCost.Text = 0
            End If
        End If
    Else
        If pid <> "" Then
            Set rsPartMas = New ADODB.Recordset
            rsPartMas.Open "Select id,STOCKDESC,srp,mac,dnp from PMIS_Stockmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPartMas.EOF And Not rsPartMas.BOF Then
                SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)
                If txtTranType.Text = "DR" Then
                    If cboChargeTo.Text = "PARTS CLAIM" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    Else
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    End If
                Else
                    If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                        'If cboChargeTo.Text = "WARRANTY" Then
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                        'ElseIf cboChargeTo.Text = "COMPANY" Then
                        '   txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                    Else
                        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
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
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO,srp,dnp,mac from PMIS_Stockmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKNO = Null2String(rsPartMas!STOCKNO)
        If txtTranType.Text = "DR" Then
            If cboChargeTo.Text = "PARTS CLAIM" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        Else
            If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            Else
                txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
            End If
        End If
    Else
        txtTranUPrice.Text = "0.00"
        txtTranUCost.Text = 0
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = '" & DDD & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDSTOCKNO = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKDESC from PMIS_Stockmas WHERE TYPE = 'M' and ltrim(rtrim(STOCKDESC)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDDesc = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select srp,STOCKNO,mac,dnp from PMIS_Stockmas WHERE TYPE = 'M' and STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            'SetPartPrice = Format(N2Str2Zero(rsPARTMAS!SRP), MAXIMUM_DIGIT)
            If txtTranType.Text = "DR" Then
                If cboChargeTo.Text = "PARTS CLAIM" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                End If
            Else
                If Mid(Trim(txtReferencePIS.Text), 5, 1) = "W" Or Mid(Trim(txtReferencePIS.Text), 5, 1) = "I" Then
                    'If cboChargeTo.Text = "WARRANTY" Then
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!DNP))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                    'ElseIf cboChargeTo.Text = "COMPANY" Then
                    '   SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPARTMAS!Mac) * ConvertToBIRDecimalFormat(VAT_RATE))
                Else
                    SetPartPrice = ToDoubleNumber(N2Str2Zero(rsPartMas!SRP))
                    txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsPartMas!Mac))
                End If
            End If
        End If
    End If
End Function

Sub InitParts()
    txtTranItemNo.Text = Format(kcnt + 1, "0000")
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

Function StorePartsEntry(ByVal ID As Variant)
    Set rsTdayTran = New ADODB.Recordset
    rsTdayTran.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice,tranucost from PMIS_vw_PRS_Tran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTdayTran.EOF And Not rsTdayTran.BOF Then
        labDetID.Caption = rsTdayTran!ID
        labPartNo.Caption = Null2String(rsTdayTran!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(rsTdayTran!tranqty)
        txtTranItemNo.Text = Null2String(rsTdayTran!itemno)
        cboTranPartNo.Text = Null2String(rsTdayTran!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(rsTdayTran!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(rsTdayTran!tranqty)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUPRICE))
        txtTranUCost.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(rsTdayTran!tranqty) * N2Str2Zero(rsTdayTran!TRANUPRICE))
    End If
    If WAREHOUSETYPE = "ADB" Then
        labTranUCost.Visible = True: txtTranUCost.Visible = True
    Else
        labTranUCost.Visible = False: txtTranUCost.Visible = False
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    PMIS_ORDER_SHOW = False: Set frmPMISCustomerOrder_CSMS = Nothing
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                           As String
    If Null2String(rsOrd_Hd!Status) = "C" Then
        MsgSpeech "Transactions are Already Cancelled and cannot be Change"
        MsgBoxXP "Transactions are Already Cancelled" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(rsOrd_Hd!Status) = "B" Then
        MsgSpeech "Transactions are Already Billed-Out and cannot be Change"
        MsgBoxXP "Transactions are Already Billed-Out" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    ElseIf Null2String(rsOrd_Hd!Status) = "P" Then
        MsgSpeech "Transactions are Already Posted and cannot be Change"
        MsgBoxXP "Transactions are Already Posted" & vbCrLf & _
                 "and cannot be Change", "Edit Not Allowed!", XP_OKOnly, msg_Information
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Visible = True
            BringToFront
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Parts!"
            Exit Sub
        End If
    End If
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
End Sub

Sub BringToFront()
    cmdAddTran.ZOrder 0
    cmdAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Visible = True
    fraAddTran.Enabled = True
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

Sub SetCustomer()
    Dim rsCustomer                                     As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCustCode.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        txtCustName.Text = Null2String(rsCustomer!acctname) & vbCrLf & Null2String(rsCustomer!customeradd) & vbCrLf & Null2String(rsCustomer!City)
    End If
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

Private Sub txtRONO_LostFocus()
    Dim RONOStr                    As String
    RONOStr = txtRONO.Text
    If Left(RONOStr, 2) = "R-" Then
       RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
    Else
       RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
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

Private Sub txtUnitPrice_LostFocus()
    If txtTranUPrice.Text = "" Then txtTranUPrice.Text = 0
    txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
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

'SEARCH MODULE
Private Sub lstOrd_Hd_GotFocus()
    rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstOrd_Hd_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optTranno.Value = True Then
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    Else
        rsOrd_Hd.Bookmark = rsFind(rsOrd_Hd.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
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

Sub FillGrid()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select Tranno,tranno from PMIS_vw_PRS where trantype = '" & WAREHOUSETYPE & "' order by Tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
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
    Set rsOrd_Hd = gconDMIS.Execute("select tranno, tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & WAREHOUSETYPE & "' and tranno like '" & XXX & "%'")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set rsOrd_Hd = New ADODB.Recordset
    Set rsOrd_Hd = gconDMIS.Execute("select rono,tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & WAREHOUSETYPE & "' and rono is not null order by tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd
        lstOrd_Hd.Refresh
        lstOrd_Hd.Enabled = True
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsOrd_Hd                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set rsOrd_Hd = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsOrd_Hd = gconDMIS.Execute("select Rono, tranno from PMIS_vw_PRS WHERE TYPE = 'M' and trantype = '" & WAREHOUSETYPE & "' and rono like '" & XXX & "%' order by tranno asc")
    If Not (rsOrd_Hd.EOF And rsOrd_Hd.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, rsOrd_Hd
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
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
