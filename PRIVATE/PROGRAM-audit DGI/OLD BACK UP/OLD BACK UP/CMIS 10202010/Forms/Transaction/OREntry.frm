VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{205EA659-0BC9-4F44-85D9-FBC10C8940C1}#1.0#0"; "wizDigit.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCMISOREntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Official Receipt Data Entry [With VAT]"
   ClientHeight    =   8325
   ClientLeft      =   810
   ClientTop       =   3285
   ClientWidth     =   12420
   ForeColor       =   &H00F5F5F5&
   Icon            =   "OREntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   12420
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   4620
      ScaleHeight     =   870
      ScaleWidth      =   7695
      TabIndex        =   83
      Top             =   7380
      Width           =   7695
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
         Left            =   6945
         MouseIcon       =   "OREntry.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   94
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
         Left            =   6255
         MouseIcon       =   "OREntry.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Option"
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
         Left            =   5565
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "View Options"
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
         Left            =   4875
         MouseIcon       =   "OREntry.frx":16F2
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":1844
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
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
         Left            =   4185
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":1BA0
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":1CF2
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Post this Transaction"
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
         Left            =   3495
         MouseIcon       =   "OREntry.frx":2017
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":2169
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
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
         Left            =   2805
         MouseIcon       =   "OREntry.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":25CE
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Move to Last Record"
         Top             =   30
         Width           =   705
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
         Left            =   2115
         MouseIcon       =   "OREntry.frx":291E
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":2A70
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Move to First Record"
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
         Left            =   1425
         MouseIcon       =   "OREntry.frx":2DCE
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":2F20
         Style           =   1  'Graphical
         TabIndex        =   86
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
         Left            =   735
         MouseIcon       =   "OREntry.frx":321A
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":336C
         Style           =   1  'Graphical
         TabIndex        =   85
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
         Left            =   45
         MouseIcon       =   "OREntry.frx":36C4
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":3816
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   3390
      ScaleHeight     =   1215
      ScaleWidth      =   8925
      TabIndex        =   20
      Top             =   90
      Width           =   8955
      Begin wizDigits.wizDigit wizDigit1 
         Height          =   1215
         Left            =   -480
         TabIndex        =   21
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6375
      Left            =   30
      TabIndex        =   48
      Top             =   0
      Width           =   3285
      Begin VB.OptionButton optORNo 
         Caption         =   "OR No."
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
         TabIndex        =   51
         Top             =   390
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton optCustName 
         Caption         =   "Customer Name"
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
         TabIndex        =   50
         Top             =   630
         Width           =   1905
      End
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   960
         Width           =   3135
      End
      Begin MSComctlLib.ListView lstOFF_HD 
         Height          =   4935
         Left            =   30
         TabIndex        =   52
         Top             =   1350
         Width           =   3165
         _ExtentX        =   5583
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "OREntry.frx":3B75
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR Number."
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label20 
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
         TabIndex        =   53
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox picDetail 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   3225
      TabIndex        =   70
      Top             =   6420
      Width           =   3225
      Begin VB.CommandButton cmdInvoices 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Customers &Invoices"
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
         Left            =   30
         TabIndex        =   134
         ToolTipText     =   "Invoice App. Detail"
         Top             =   585
         Width           =   3150
      End
      Begin VB.CommandButton cmdInvoiceDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice App. Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   72
         ToolTipText     =   "Invoice App. Detail"
         Top             =   285
         Width           =   3150
      End
      Begin VB.CommandButton cmdORDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "O.R. Payment Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   71
         ToolTipText     =   "O.R. Payment Detail"
         Top             =   0
         Width           =   3150
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   60
      ScaleHeight     =   945
      ScaleWidth      =   3195
      TabIndex        =   77
      Top             =   7290
      Width           =   3225
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   60
         TabIndex        =   82
         Top             =   540
         Width           =   2055
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F11 - Calculator"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   60
         TabIndex        =   81
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Del OR Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   80
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Edit OR Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   79
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - Add OR Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   78
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10800
      ScaleHeight     =   855
      ScaleWidth      =   1500
      TabIndex        =   95
      Top             =   7395
      Width           =   1500
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
         Left            =   750
         MouseIcon       =   "OREntry.frx":3CD7
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":3E29
         Style           =   1  'Graphical
         TabIndex        =   96
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
         Left            =   60
         MouseIcon       =   "OREntry.frx":4167
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":42B9
         Style           =   1  'Graphical
         TabIndex        =   97
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   885
      Left            =   3390
      ScaleHeight     =   885
      ScaleWidth      =   8955
      TabIndex        =   37
      Top             =   6405
      Width           =   8955
      Begin VB.TextBox txtBranch 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6540
         TabIndex        =   47
         Top             =   60
         Width           =   2325
      End
      Begin VB.TextBox txtPaidFor 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1320
         TabIndex        =   46
         Top             =   60
         Width           =   3855
      End
      Begin VB.TextBox txtPaymentAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7470
         TabIndex        =   40
         Text            =   "0.00"
         Top             =   450
         Width           =   1400
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1320
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   450
         Width           =   1400
      End
      Begin VB.TextBox txtTaxAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   4260
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   450
         Width           =   1400
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Amt :"
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
         Height          =   285
         Left            =   6120
         TabIndex        =   45
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code :"
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
         Height          =   255
         Left            =   5250
         TabIndex        =   44
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amt :"
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
         Height          =   255
         Left            =   3090
         TabIndex        =   43
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amt :"
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
         Height          =   255
         Left            =   60
         TabIndex        =   42
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for :"
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
         Height          =   255
         Left            =   90
         TabIndex        =   41
         Top             =   60
         Width           =   1185
      End
   End
   Begin VB.PictureBox picOR 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3390
      ScaleHeight     =   1455
      ScaleWidth      =   8955
      TabIndex        =   22
      Top             =   1410
      Width           =   8955
      Begin VB.TextBox cboCUSNAME 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   133
         Text            =   "Text1"
         Top             =   1020
         Width           =   4215
      End
      Begin MSMask.MaskEdBox txtOR_NUM 
         Height          =   525
         Left            =   2100
         TabIndex        =   0
         Top             =   60
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   15857397
         ForeColor       =   0
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "F7-Select Customer"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7425
         TabIndex        =   100
         Top             =   1035
         Width           =   1500
      End
      Begin VB.CommandButton cmdVarious 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2850
         TabIndex        =   75
         Top             =   1050
         Width           =   285
      End
      Begin Crystal.CrystalReport rptChat 
         Left            =   5340
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
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
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8460
         Top             =   570
      End
      Begin VB.TextBox txtOR_DATE 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   2
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox txtCUSCDE 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "BBBBBB"
         Top             =   1020
         Width           =   705
      End
      Begin MSMask.MaskEdBox txtVNF 
         Height          =   525
         Left            =   4140
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   15857397
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label labStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*** Cancelled OR ***"
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
         Left            =   5400
         TabIndex        =   63
         Top             =   180
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Official Receipt Date:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code-Name :"
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
         Height          =   255
         Left            =   -30
         TabIndex        =   25
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Official Receipt No. :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -210
         TabIndex        =   24
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3960
         TabIndex        =   23
         Top             =   90
         Visible         =   0   'False
         Width           =   195
      End
   End
   Begin VB.PictureBox picORDetails 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3465
      Left            =   3390
      ScaleHeight     =   3465
      ScaleWidth      =   8985
      TabIndex        =   27
      Top             =   2910
      Width           =   8985
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   3255
         Left            =   60
         TabIndex        =   4
         Top             =   90
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Type    |    Ref. #       |    Application                                |   AR                  | Balance           ||||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "OREntry.frx":4609
      End
   End
   Begin VB.PictureBox picCreditCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   3420
      ScaleHeight     =   4335
      ScaleWidth      =   8895
      TabIndex        =   107
      Top             =   3000
      Visible         =   0   'False
      Width           =   8925
      Begin VB.CommandButton cmdCardCancel 
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
         Left            =   8070
         MouseIcon       =   "OREntry.frx":4923
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":4A75
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   360
         Width           =   705
      End
      Begin VB.CommandButton cmdCardSave 
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
         Left            =   7380
         MouseIcon       =   "OREntry.frx":4DB3
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":4F05
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   360
         Width           =   705
      End
      Begin VB.CheckBox chkSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   150
         TabIndex        =   141
         Top             =   3900
         Width           =   195
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Payment Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   112
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OR Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   111
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select / Deselect All "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   4440
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   108
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvPayments 
         Height          =   2655
         Left            =   120
         TabIndex        =   110
         ToolTipText     =   "Double click to select customer"
         Top             =   1230
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR No."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Code"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer Name"
            Object.Width           =   6598
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reference No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OR Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox picCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   90
         ScaleHeight     =   495
         ScaleWidth      =   6585
         TabIndex        =   130
         Top             =   720
         Visible         =   0   'False
         Width           =   6585
         Begin VB.TextBox txtCustomer 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2580
            MaxLength       =   50
            TabIndex        =   131
            Top             =   30
            Width           =   3465
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
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
            Height          =   195
            Left            =   930
            TabIndex        =   132
            Top             =   120
            Width           =   1470
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   8145
         TabIndex        =   113
         Top             =   720
         Visible         =   0   'False
         Width           =   8145
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   30
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   405
            Left            =   870
            TabIndex        =   115
            Top             =   15
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20643841
            CurrentDate     =   38216
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   405
            Left            =   3090
            TabIndex        =   116
            Top             =   15
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   20643841
            CurrentDate     =   38216
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "From :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   2610
            TabIndex        =   117
            Top             =   60
            Width           =   435
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   8475
         TabIndex        =   119
         Top             =   720
         Width           =   8475
         Begin VB.TextBox txtReference2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2595
            MaxLength       =   6
            TabIndex        =   120
            Top             =   30
            Width           =   1815
         End
         Begin VB.Label lblReference2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter OR Number"
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
            Height          =   195
            Left            =   900
            TabIndex        =   121
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         Height          =   195
         Left            =   390
         TabIndex        =   142
         Top             =   3930
         Width           =   660
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6240
         TabIndex        =   125
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6990
         TabIndex        =   124
         Top             =   3930
         Width           =   1695
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Bank Payment        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   8421504
         GradientColorDark=   8421504
      End
      Begin VB.Label lblRefNo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1680
         TabIndex        =   122
         Top             =   4080
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   3420
      ScaleHeight     =   3945
      ScaleWidth      =   8895
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   8925
      Begin VB.TextBox txtReference1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3930
         MaxLength       =   8
         TabIndex        =   126
         Top             =   4080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboInvoiceType 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   390
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   960
         Width           =   1365
      End
      Begin VB.CheckBox chkCreditCardTrans 
         BackColor       =   &H00C0C0C0&
         Caption         =   "This is a Credit Card Transaction"
         Height          =   285
         Left            =   5880
         TabIndex        =   102
         Top             =   330
         Visible         =   0   'False
         Width           =   2745
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
         Left            =   90
         MouseIcon       =   "OREntry.frx":5255
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":53A7
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2910
         Width           =   705
      End
      Begin VB.CommandButton cmdCardPayment 
         Caption         =   "Compute for Card Tax and Discount"
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
         Height          =   465
         Left            =   5880
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":56D2
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Save changes"
         Top             =   630
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   795
         Left            =   2670
         ScaleHeight     =   795
         ScaleWidth      =   3075
         TabIndex        =   55
         Top             =   960
         Width           =   3075
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1590
            TabIndex        =   58
            Text            =   "0.00"
            Top             =   420
            Width           =   1395
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1590
            TabIndex        =   56
            Text            =   "0.00"
            Top             =   30
            Width           =   1395
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Balance :"
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
            Height          =   285
            Left            =   30
            TabIndex        =   59
            Top             =   450
            Width           =   1545
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Amount :"
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
            Height          =   285
            Left            =   30
            TabIndex        =   57
            Top             =   60
            Width           =   1545
         End
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         TabIndex        =   54
         Top             =   540
         Width           =   285
      End
      Begin VB.ComboBox cboTranType 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   390
         ItemData        =   "OREntry.frx":59DC
         Left            =   1200
         List            =   "OREntry.frx":59DE
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   510
         Width           =   4245
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   8
         Top             =   1380
         Width           =   1365
      End
      Begin VB.TextBox txtDescript 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   615
         Left            =   1200
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2190
         Width           =   4545
      End
      Begin VB.ComboBox cboBranch 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   2235
      End
      Begin VB.ComboBox cboPaidFor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Text            =   "cboPaidFor"
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   435
         Left            =   6540
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2370
         Width           =   2085
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   1140
         Width           =   1400
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1560
         Width           =   1400
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
         Left            =   7920
         MouseIcon       =   "OREntry.frx":59E0
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":5B32
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2910
         Width           =   705
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
         Left            =   7230
         MouseIcon       =   "OREntry.frx":5E70
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":5FC2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2910
         Width           =   705
      End
      Begin VB.CommandButton cmdDetails 
         Caption         =   "Command2"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7380
         TabIndex        =   101
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         TabIndex        =   73
         Top             =   510
         Width           =   285
      End
      Begin VB.Label lblDetID 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1110
         TabIndex        =   129
         Top             =   3660
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblReference1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Reference No."
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
         Height          =   195
         Left            =   1170
         TabIndex        =   128
         Top             =   4110
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - View Bank Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1170
         TabIndex        =   127
         Top             =   3390
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   225
         Left            =   90
         TabIndex        =   106
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice"
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
         Height          =   225
         Left            =   90
         TabIndex        =   104
         Top             =   960
         Width           =   1005
      End
      Begin XtremeShortcutBar.ShortcutCaption labStatusMode 
         Height          =   285
         Left            =   0
         TabIndex        =   103
         Top             =   0
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "System is Adding/Editing OR Detail"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   8421504
         GradientColorDark=   16711680
      End
      Begin VB.Label labReference 
         Caption         =   "Label10"
         Height          =   285
         Left            =   3630
         TabIndex        =   74
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label labCUSCODE 
         Caption         =   "Label21"
         Height          =   195
         Left            =   2370
         TabIndex        =   62
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label labDetID 
         Caption         =   "Label21"
         Height          =   135
         Left            =   2040
         TabIndex        =   61
         Top             =   2370
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label labDocDate 
         BackStyle       =   0  'Transparent
         Caption         =   "[DOC DATE]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   30
         TabIndex        =   60
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Label labRef 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Height          =   255
         Left            =   90
         TabIndex        =   36
         Top             =   1470
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran Type :"
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
         Height          =   255
         Left            =   -60
         TabIndex        =   35
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Application :"
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
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Amt :"
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
         Height          =   285
         Left            =   5820
         TabIndex        =   33
         Top             =   2010
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code :"
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
         Height          =   255
         Left            =   9090
         TabIndex        =   32
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "EWT"
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
         Height          =   255
         Left            =   6030
         TabIndex        =   31
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Charges"
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
         Height          =   255
         Left            =   5820
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for :"
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
         Height          =   255
         Left            =   60
         TabIndex        =   29
         Top             =   1830
         Width           =   1125
      End
   End
   Begin VB.PictureBox picDeposits 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   3420
      ScaleHeight     =   3645
      ScaleWidth      =   6435
      TabIndex        =   135
      Top             =   3300
      Visible         =   0   'False
      Width           =   6465
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6180
         TabIndex        =   136
         Top             =   0
         Width           =   255
      End
      Begin MSComctlLib.ListView lvDeposits 
         Height          =   3240
         Left            =   45
         TabIndex        =   137
         Top             =   360
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   5715
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OR Date"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "OR No."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Applied"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Deposits"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   3315
         Left            =   30
         ScaleHeight     =   3255
         ScaleWidth      =   6345
         TabIndex        =   138
         Top             =   330
         Width           =   6405
      End
      Begin XtremeShortcutBar.ShortcutCaption sc3 
         Height          =   285
         Left            =   0
         TabIndex        =   140
         Top             =   0
         Width           =   6435
         _Version        =   655364
         _ExtentX        =   11351
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Unapplied Customer Deposit/s"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   -2147483629
         GradientColorDark=   12582912
      End
      Begin VB.Label lblDepositID 
         Height          =   195
         Left            =   30
         TabIndex        =   139
         Top             =   3720
         Width           =   1395
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   1515
      Left            =   9000
      ScaleHeight     =   1455
      ScaleWidth      =   2910
      TabIndex        =   69
      Top             =   5850
      Width           =   2970
      Begin VB.CommandButton cmdRecoverOR 
         Caption         =   "Recover Cancelled Official Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1470
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":6312
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":661C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Recover Cancelled Official Receipt"
         Top             =   300
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancelOR 
         Caption         =   "Cancel Official Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   60
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":6A74
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":6D7E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel Official Receipt"
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "OPTIONS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   -240
         TabIndex        =   99
         Top             =   0
         Width           =   3435
      End
   End
   Begin VB.PictureBox picPayment 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   6825
      TabIndex        =   64
      Top             =   1170
      Visible         =   0   'False
      Width           =   6855
      Begin VB.OptionButton optCANCEL 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   5070
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCARD 
         Caption         =   "CARD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCHECK 
         Caption         =   "CHECK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCASH 
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   540
         Width           =   1545
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   98
         Top             =   0
         Width           =   6840
         _Version        =   655364
         _ExtentX        =   12065
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "  Select Type of Payment..."
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
         VisualTheme     =   3
         GradientColorLight=   16711680
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   6150
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   6480
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCMISOREntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOFF_HD                                                     As ADODB.Recordset
Dim rsOFF_DT                                                     As ADODB.Recordset
Dim TOTAL_AR_AMOUNT                                              As Double
Dim AddorEdit, PrevOR_NUM                                        As String
Attribute PrevOR_NUM.VB_VarUserMemId = 1073938435
Dim On_Update                                                    As Boolean
Attribute On_Update.VB_VarUserMemId = 1073938437
Dim ChangeORNum                                                  As Boolean
Public LocalAcess                                                As String
Dim rsINVOICEDUp                                                 As ADODB.Recordset
Dim rsCustomerDeposit                                            As ADODB.Recordset
Dim FIRST_LOAD                                                   As Boolean
Dim vTRANTYPE                                                    As String
Dim vOR_NUM2                                                     As String
Dim tmpTotal                                                     As Double
Dim vDetails                                                     As Boolean

Function SetCustomerCode(XXX As Variant)
    Dim rsCustomer                                               As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select CusCde from ALL_CUSMAS Where CusNam = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = rsCustomer!CUSCDE
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerName(XXX As Variant)
    Dim rsCustomer                                               As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select CusNam from ALL_CUSMAS Where CusCde = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = rsCustomer!CusNam
    End If
    Set rsCustomer = Nothing
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                  As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'F' and DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!code
    End If
End Function

Function SetTranType(XXX As Variant)
    Dim rsType                                                   As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'A' and Code = '" & XXX & "'")
    If Not rsType.EOF And Not rsType.BOF Then
        SetTranType = rsType!DESCNAME
    End If
End Function

Function SetTranTypeCode(XXX As Variant)
    Dim rsType                                                   As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'A' and DescName = '" & XXX & "'")
    If Not rsType.EOF And Not rsType.BOF Then
        SetTranTypeCode = rsType!code
    End If
End Function

Function SetBranch(XXX As Variant)
    Dim rsBranch                                                 As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'C' and Code = '" & XXX & "'")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        SetBranch = rsBranch!DESCNAME
    End If
End Function

Function SetBranchCode(XXX As Variant)
    Dim rsBranch                                                 As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'C' and DescName = '" & XXX & "'")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        SetBranchCode = rsBranch!code
    End If
End Function

Function SetPaidFor(XXX As Variant)
    Dim rsPayment                                                As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'D' and Code = '" & XXX & "'")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        SetPaidFor = Null2String(rsPayment!DESCNAME)
    End If
End Function

Function SetPaidForCode(XXX As Variant)
    Dim rsPayment                                                As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'D' and DescName = '" & XXX & "'")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        SetPaidForCode = Null2String(rsPayment!code)
    End If
End Function

Sub RefreshDisplay()
    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
    StoreMemVars
End Sub

Sub Save_CASH_Payment()
    If COMPANY_CODE = "HGC" Then
        Set rsINVOICEDUp = New ADODB.Recordset
        Set rsINVOICEDUp = gconDMIS.Execute("Select * from CMIS_ORS WHERE ORNO = '" & Format(txtOR_NUM.Text, "000000") & "'")
        If Not rsINVOICEDUp.EOF And Not rsINVOICEDUp.BOF Then
            If Trim(Null2String(rsINVOICEDUp!Status)) = "P" Then
                MsgSpeechBox "OR Number Already Exist!"
                Exit Sub
            End If
            If Trim(Null2String(rsINVOICEDUp!Status)) = "C" Then
                MsgSpeechBox "OR Number Already Used and Was Cancelled!"
                Exit Sub
            End If
        End If
    End If
    gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                      " VAT = " & VAT_OR & "," & _
                      " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                      " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                      " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                      " PAIDNA = 1, STATUS='P'" & _
                      " where OR_NUM = " & N2Str2Null(txtOR_NUM.Text))
    If COMPANY_CODE = "HGC" Then
        gconDMIS.Execute "Insert into CMIS_ORS (ORNO,ORDATE,STATUS) Values ('" & txtOR_NUM.Text & "','" & CDate(txtOR_DATE.Text) & "','P')"
    End If

    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & txtOR_NUM.Text
    StoreMemVars
End Sub

Sub Save_CHECK_Payment()
    If COMPANY_CODE = "HGC" Then
        Set rsINVOICEDUp = New ADODB.Recordset
        Set rsINVOICEDUp = gconDMIS.Execute("Select * from CMIS_ORS WHERE ORNO = '" & Format(txtOR_NUM.Text, "000000") & "'")
        If Not rsINVOICEDUp.EOF And Not rsINVOICEDUp.BOF Then
            If Null2String(rsINVOICEDUp!Status) = "P" Then
                MsgSpeechBox "OR Number Already Exist!"
                Exit Sub
            End If
            If Null2String(rsINVOICEDUp!Status) = "C" Then
                MsgSpeechBox "OR Number Already Used and Was Cancelled!"
                Exit Sub
            End If
        End If
    End If
    gconDMIS.Execute ("update CMIS_Off_Hd Set" & _
                      " VAT = " & VAT_OR & "," & _
                      " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                      " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                      " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                      " PAIDNA = 1, STATUS='P'" & _
                      " where OR_NUM = " & N2Str2Null(txtOR_NUM.Text))
End Sub

Sub SetCustomer()
    Call FillCustomer
    Dim rsCustomer                                               As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select * from ALL_Customer where CusCde = '" & txtCuscde.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        cboCUSNAME.Text = Null2String(rsCustomer!AcctName)
    End If
End Sub

Sub rsRefresh()
    Set rsOFF_HD = New ADODB.Recordset
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd where VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE asc, OR_NUM asc")
    'Else
    '   Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd where VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE asc, OR_NUM asc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd where VAT = 1 order by OR_DATE asc, OR_NUM asc")
    Else
        Set rsOFF_HD = gconDMIS.Execute("Select * from CMIS_Off_hd where VAT = 0 order by OR_DATE asc, OR_NUM asc")
    End If
    If FIRST_LOAD = True Then
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then rsOFF_HD.MoveLast
    End If

End Sub

Sub StoreMemVars()
    cmdEdit.Enabled = True: cmdPOST.Enabled = True
    cmdOptions.Enabled = True: cmdPrint.Enabled = False
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        labid.Caption = rsOFF_HD!Id
        txtOR_NUM.Text = Format(Left(Null2String(rsOFF_HD!OR_NUM), 6), "000000")
        txtOR_DATE.Text = Null2String(rsOFF_HD!OR_DATE)
        txtCuscde.Text = Null2String(rsOFF_HD!CUSCDE)
        'If SetCustomerName(Null2String(rsOFF_HD!CUSCDE)) <> "" Then
        cboCUSNAME.Text = SetCustomerName(Null2String(rsOFF_HD!CUSCDE))
        'Else
        '    cboCUSNAME.ListIndex = -1
        'End If
        If N2Str2Zero(rsOFF_HD!CashAmount) > 0 Then
            MODE_OF_PAYMENT = "CASH"
        End If
        If N2Str2Zero(rsOFF_HD!CardAmount) > 0 Then
            MODE_OF_PAYMENT = "CARD"
        End If
        If N2Str2Zero(rsOFF_HD!CHKAMOUNT) > 0 Then
            MODE_OF_PAYMENT = "CHECK"
        End If
        RECEIPTS_AMOUNT = N2Str2Zero(rsOFF_HD!OR_AMT)
        If rsOFF_HD!Cancel = True Then
            cmdEdit.Enabled = False
            cmdPOST.Enabled = False
            cmdOptions.Enabled = False
            cmdPrint.Enabled = False
        End If
        If rsOFF_HD!PaidNa = True Then
            cmdEdit.Enabled = False
            cmdPOST.Enabled = False
            cmdOptions.Enabled = True
            cmdPrint.Enabled = True
        Else
            cmdOptions.Enabled = False
        End If
        StoreDetails
    Else
        'MsgBox "No Such Record!", vbInformation, "Message"
        MessagePop InfoFriend, "Message", "No Such Record"
        cmdAdd.Value = True
    End If
End Sub

Sub StoreDetails()
    Dim I                                                        As Integer
    Dim vDeposit                                                 As Double
    TOTAL_AR_AMOUNT = 0: InitGrid
    Dim TRAN_INVOICE_TYPE                                        As String
    Dim rsOFF_Payment                                            As ADODB.Recordset
    Set rsOFF_DT = New ADODB.Recordset
    Set rsOFF_DT = gconDMIS.Execute("Select * from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND OR_Num = " & N2Str2Null(rsOFF_HD!OR_NUM) & " Order by [ID] asc")
    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
        rsOFF_DT.MoveFirst: I = 0
        Do While Not rsOFF_DT.EOF
            I = I + 1
            If Null2String(rsOFF_DT!TRANTYPE) = "RO" Then
                TRAN_INVOICE_TYPE = "SI"
            Else
                TRAN_INVOICE_TYPE = Null2String(rsOFF_DT!TRANTYPE)
            End If

            Set rsOFF_Payment = gconDMIS.Execute("SELECT PAYMENT,(SELECT Sum(AMOUNT) FROM CMIS_DEPOSITS WHERE ID_DET = " & N2Str2Null(rsOFF_DT!OR_NUM) & ") AS DEPOSIT FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT ID_DET FROM CMIS_DEPOSITS WHERE ID_DET = " & N2Str2Null(rsOFF_DT!OR_NUM) & " AND APPLIED = 'Y')")

            If Not rsOFF_Payment.EOF And Not rsOFF_Payment.BOF Then
                vDeposit = N2Str2Zero(rsOFF_Payment!DEPOSIT)
            End If
            grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!INVOICENO) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!AMOUNT)) & Chr(9) & ToDoubleNumber(IIf(N2Str2Zero(rsOFF_DT!BALANCE) = 0, 0, Round(N2Str2Zero(rsOFF_DT!BALANCE) - (N2Str2Zero(rsOFF_DT!payment) + N2Str2Zero(vDeposit)), 2))) & _
                               Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!payment) & Chr(9) & rsOFF_DT!Id
            TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + N2Str2Zero(rsOFF_DT!payment)
            If I = 1 Then grdDetails.RemoveItem 1
            wizDigit1.TextValue = ToDoubleNumber(TOTAL_AR_AMOUNT)
            rsOFF_DT.MoveNext
        Loop
        grdDetails.Col = 10

        ShowGridDetails grdDetails.Text
        vDetails = True
    Else
        vDetails = False
        wizDigit1.TextValue = ZERO
        txtPaidFor.Text = "": txtBranch.Text = ""
        txtDiscountAmt.Text = "0.00": txtTaxAmt.Text = "0.00": txtPaymentAmt.Text = "0.00"
    End If
End Sub

Sub ShowGridDetails(XXX As Long)
    Dim rsOFF_Details                                            As ADODB.Recordset
    Set rsOFF_Details = New ADODB.Recordset
    Set rsOFF_Details = gconDMIS.Execute("Select * from CMIS_Off_Dt Where ID = " & XXX)
    If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
        txtPaidFor.Text = SetPaidFor(Null2String(rsOFF_Details!PAIDFOR))
        xPAIDFOR = Null2String(rsOFF_Details!PAIDFOR)
        txtBranch.Text = SetBranch(Null2String(rsOFF_Details!BRANCH))
        txtDiscountAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!DISCOUNT))
        txtTaxAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!tax))
        txtPaymentAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!payment))
        lblDetID.Caption = Null2String(rsOFF_Details!OR_NUM)
        vREFERENCENO = Null2String(rsOFF_Details!ReferenceNo)
    Else
        txtPaidFor.Text = "": txtBranch.Text = ""
        txtDiscountAmt.Text = "0.00": txtTaxAmt.Text = "0.00": txtPaymentAmt.Text = "0.00"
    End If
End Sub

Sub StoreGridDetails(XXX As Long)
    Dim rsOFF_Details                                            As ADODB.Recordset
    Set rsOFF_Details = New ADODB.Recordset
    Set rsOFF_Details = gconDMIS.Execute("Select * from CMIS_Off_Dt Where ID = " & XXX)
    If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
        AddorEdit = "EDIT"
        labStatusMode.Caption = "System is Editing OR Detail..."
        cmdRefresh.Enabled = False
        cmdTranDelete.Visible = True
        labDetID.Caption = rsOFF_Details!Id
        labCUSCODE.Caption = Null2String(rsOFF_Details!CUSCDE)
        cboTranType.Text = SetTranType(Null2String(rsOFF_Details!TRANTYPE))
        txtReference.Text = Null2String(rsOFF_Details!REFERENCE)
        labReference.Caption = Null2String(rsOFF_Details!INVOICENO)
        txtDescript.Text = Null2String(rsOFF_Details!DESCRIPT)
        lblRefNo.Caption = Null2String(rsOFF_Details!ReferenceNo)
        If Null2String(rsOFF_Details!PAIDFOR) <> "" Then
            cboPaidFor.Text = SetPaidFor(Null2String(rsOFF_Details!PAIDFOR))
        Else
            cboPaidFor.ListIndex = -1
        End If
        If Null2String(rsOFF_Details!BRANCH) <> "" Then
            cboBranch.Text = SetBranch(Null2String(rsOFF_Details!BRANCH))
        Else
            cboBranch.ListIndex = -1
        End If
        txtAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!AMOUNT))
        txtBalance.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!BALANCE))
        txtDiscount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!DISCOUNT))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!tax))
        txtPayment.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!payment))
        lblDetID.Caption = Null2String(rsOFF_Details!OR_NUM)
    End If
End Sub

Sub InitGridMemvars()
    AddorEdit = "ADD": cmdRefresh.Enabled = True: cmdTranDelete.Visible = False
    cboTranType.ListIndex = -1: cboTranType.Enabled = True
    labDocDate.Caption = "[DOC DATE]"
    labCUSCODE.Caption = "V00009"
    txtReference.Text = "": txtDescript.Text = ""
    cboPaidFor.ListIndex = -1: cboBranch.ListIndex = -1

    txtAmount.Text = "0.00": txtBalance.Text = "0.00"
    txtDiscount.Text = "0.00": txtTax.Text = "0.00"
    txtPayment.Text = "0.00"

    txtReference.Enabled = False: txtDescript.Enabled = False
    cboPaidFor.Enabled = False: cboBranch.Enabled = False
    txtDiscount.Enabled = False: txtTax.Enabled = False
    txtPayment.Enabled = False
    On Error Resume Next
    cboTranType.SetFocus
End Sub

Sub initMemvars()
    txtOR_NUM.Text = ""
    txtOR_DATE.Text = LOGDATE
    txtCuscde.Text = ""
    cboCUSNAME = ""
    txtPaymentAmt.Text = ZERO
    wizDigit1.TextValue = ZERO
    InitGrid
End Sub

Sub InitGrid()
    cleargrid grdDetails
    grdDetails.FormatString = "  Type    |    Ref. #       |    Application                                |   AR                  | Balance           "
    grdDetails.ColWidth(5) = 1: grdDetails.ColWidth(6) = 1: grdDetails.ColWidth(7) = 1: grdDetails.ColWidth(8) = 1: grdDetails.ColWidth(9) = 1: grdDetails.ColWidth(10) = 1
End Sub

Sub FillCustomer()
    Exit Sub

    'TO DO
    'TAKE OUT THIS PROCEDURE AND USE ONLY THE SELECT BUTTON
    Dim rsCustomer                                               As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("Select cusnam from ALL_CUSMAS where cusnam <> '' and cusnam is not null Order by cusnam asc")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Combo_Loadval cboCUSNAME, rsCustomer
    End If
End Sub

Sub FillType()
    Dim rsType                                                   As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'A' order by DescName asc")
    If Not rsType.EOF And Not rsType.BOF Then
        Combo_Loadval cboTranType, rsType
    End If
End Sub

Sub FillBranch()
    Dim rsBranch                                                 As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'C' order by DescName asc")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        Combo_Loadval cboBranch, rsBranch
    End If
End Sub

Sub FillPayment()
    Dim rsPayment                                                As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("Select DescName from CMIS_SBOOK Where Book = 'D' order by DescName asc")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        Combo_Loadval cboPaidFor, rsPayment
    End If
End Sub

Sub FillGrid()
    Dim rsOFF_HD2                                                As ADODB.Recordset
    lstOFF_HD.Sorted = False: lstOFF_HD.ListItems.Clear: lstOFF_HD.Enabled = False
    lstOFF_HD.Enabled = False
    Set rsOFF_HD2 = New ADODB.Recordset
    If OR_VAT_NONVAT = "VAT" Then
        'Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM,ID from CMIS_Off_hd where VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE desc, OR_NUM desc")
        Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM,ID from CMIS_Off_hd where VAT = 1 order by OR_DATE desc, OR_NUM desc")
    Else
        'Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM,ID from CMIS_Off_hd where VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE desc, OR_NUM desc")
        Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM,ID from CMIS_Off_hd where VAT = 0 order by OR_DATE desc, OR_NUM desc")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If

    Set rsOFF_HD2 = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsOFF_HD2                                                As ADODB.Recordset
    lstOFF_HD.Sorted = False: lstOFF_HD.ListItems.Clear
    lstOFF_HD.Enabled = False
    XXX = Repleys(XXX)
    Set rsOFF_HD2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM, ID from CMIS_Off_hd where VAT = 1 AND OR_NUM like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE Desc, OR_NUM desc")
    'Else
    '   Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM, ID from CMIS_Off_hd where VAT = 0 AND OR_NUM like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by OR_DATE Desc, OR_NUM desc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM, ID from CMIS_Off_hd where VAT = 1 AND OR_NUM like '" & ReplaceQuote(XXX) & "%' order by OR_DATE Desc, OR_NUM desc")
    Else
        Set rsOFF_HD2 = gconDMIS.Execute("select OR_NUM, ID from CMIS_Off_hd where VAT = 0 AND OR_NUM like '" & ReplaceQuote(XXX) & "%' order by OR_DATE Desc, OR_NUM desc")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If

    Set rsOFF_HD2 = Nothing
End Sub

Sub FillGrid2()
    Dim rsOFF_HD2                                                As ADODB.Recordset
    lstOFF_HD.Sorted = False: lstOFF_HD.ListItems.Clear: lstOFF_HD.Enabled = False
    lstOFF_HD.Enabled = False
    Set rsOFF_HD2 = New ADODB.Recordset
    If OR_VAT_NONVAT = "VAT" Then
        'Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME,ID from CMIS_Off_hd where VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by CUSNAME asc")
        Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME,ID from CMIS_Off_hd where VAT = 1 order by CUSNAME asc")
    Else
        'Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME,ID from CMIS_Off_hd where VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by CUSNAME asc")
        Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME,ID from CMIS_Off_hd where VAT = 0 order by CUSNAME asc")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If

    Set rsOFF_HD2 = Nothing
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsOFF_HD2                                                As ADODB.Recordset
    lstOFF_HD.Enabled = False
    lstOFF_HD.Sorted = False: lstOFF_HD.ListItems.Clear
    Set rsOFF_HD2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME, ID from CMIS_Off_hd where VAT = 1 and CUSNAME like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by CUSNAME asc")
    'Else
    '   Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME, ID from CMIS_Off_hd where VAT = 0 and CUSNAME like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " order by CUSNAME asc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME, ID from CMIS_Off_hd where VAT = 1 and CUSNAME like '" & XXX & "%' order by CUSNAME asc")
    Else
        Set rsOFF_HD2 = gconDMIS.Execute("select CUSNAME, ID from CMIS_Off_hd where VAT = 0 and CUSNAME like '" & XXX & "%' order by CUSNAME asc")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If

    Set rsOFF_HD2 = Nothing
End Sub

Private Sub cboBranch_GotFocus()
    VBComBoBoxDroppedDown cboBranch
End Sub

Private Sub cboPaidFor_Click()
    txtDescript.Text = cboPaidFor.Text
    If SetPaidForCode(cboPaidFor.Text) = "427" Then
        If CheckIfBank(txtCuscde.Text) = True Then
            picCreditCard.Visible = True
            txtReference2.Text = ""
            txtCustomer.Text = ""
            picCreditCard.ZOrder 0
            Option1.Value = True
            chkCreditCardTrans.Value = 0
            chkCreditCardTrans.Enabled = False
            Option1_Click
            chkSelect.Value = 0
        Else
            MsgBox "For BANK use only", vbInformation, "Payment Receive from Bank"
            cboPaidFor.ListIndex = -1
            Exit Sub
        End If
    Else
        chkCreditCardTrans.Enabled = True
    End If
End Sub

Private Sub cboPaidFor_GotFocus()
    VBComBoBoxDroppedDown cboPaidFor
End Sub

Private Sub cboPaidFor_KeyPress(KeyAscii As Integer)
    txtDescript.Text = ""
    If KeyAscii = 13 Then
        txtDescript.Text = cboPaidFor.Text
        If SetPaidForCode(cboPaidFor.Text) = "427" Then
            If CheckIfBank(txtCuscde.Text) = True Then
                picCreditCard.Visible = True
                txtReference1.Text = "NULL"
                txtReference2.Text = ""
                txtCustomer.Text = ""
                picCreditCard.ZOrder 0
                Option1.Value = True
                chkCreditCardTrans.Value = 0
                chkCreditCardTrans.Enabled = False
            Else
                MsgBox "For BANK use only", vbInformation, "Payment Receive from Bank"
                cboPaidFor.ListIndex = -1
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub cboPaidFor_LostFocus()
    If SetPaidForCode(cboPaidFor.Text) = "427" Then
        If CheckIfBank(txtCuscde.Text) = True Then
            picCreditCard.Visible = True
            txtReference2.Text = ""
            txtCustomer.Text = ""
            picCreditCard.ZOrder 0
            Option1.Value = True
            chkCreditCardTrans.Value = 0
            chkCreditCardTrans.Enabled = False
            Option1_Click
        Else
            MsgBox "For BANK use only", vbInformation, "Payment Receive from Bank"
            cboPaidFor.ListIndex = -1
            Exit Sub
        End If
    Else
        chkCreditCardTrans.Enabled = True
    End If
End Sub

Private Sub cboTranType_Click()
    txtReference.Enabled = True
    If SetTranTypeCode(cboTranType.Text) = "PI" Or SetTranTypeCode(cboTranType.Text) = "SI" Or SetTranTypeCode(cboTranType.Text) = "AI" Or SetTranTypeCode(cboTranType.Text) = "MI" Then
        cboInvoiceType.Enabled = True
        chkCreditCardTrans.Enabled = True
    Else
        cboInvoiceType.Enabled = False
        chkCreditCardTrans.Value = 0
        chkCreditCardTrans.Enabled = False
    End If
End Sub

Private Sub cboTranType_GotFocus()
    VBComBoBoxDroppedDown cboTranType
End Sub

Private Sub cboTranType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cboTranType.Text) = "" Then VBComBoBoxDroppedDown cboTranType
    End If
End Sub

Private Sub cboTranType_LostFocus()
    cboTranType.Enabled = False
    If SetTranTypeCode(cboTranType.Text) = "OTH" Then
        Dim rsJoy                                                As ADODB.Recordset
        Set rsJoy = New ADODB.Recordset
        Set rsJoy = gconDMIS.Execute("Select * from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'OTH' order by reference desc")
        If Not rsJoy.EOF And Not rsJoy.BOF Then
            txtReference.Text = Format(N2Str2Zero(rsJoy!REFERENCE) + 1, "00000000")
        Else
            txtReference.Text = "00000001"
        End If
        txtReference.Enabled = False: cmdTranSave.Enabled = True
        cboPaidFor.Enabled = True: cboBranch.Enabled = True: txtDescript.Enabled = True
        txtDiscount.Enabled = True: txtTax.Enabled = True: txtPayment.Enabled = True
        'Call txtReference_KeyDown(vbKeyReturn, 0)
        On Error Resume Next
        cboPaidFor.SetFocus
    End If
End Sub

Private Sub chkCreditCardTrans_Click()
    If chkCreditCardTrans.Value = 1 Then
        'cmdCardPayment.Enabled = True
        txtReference1.Text = GetReferenceNo
        txtReference1.Locked = True
        lblReference1.Visible = True
        txtReference1.Visible = True
    Else
        cmdCardPayment.Enabled = False
        txtReference1.Visible = False
        lblReference1.Visible = False
        txtReference1.Text = ""
        txtDiscount.Text = "0.00"
        txtTax.Text = "0.00"
    End If
End Sub

Private Sub chkSelect_Click()
    Dim iCount As Integer
    If chkSelect.Value = 1 Then
        For iCount = 1 To lvPayments.ListItems.Count
            lvPayments.ListItems.Item(iCount).Checked = True
        Next
    Else
        For iCount = 1 To lvPayments.ListItems.Count
            lvPayments.ListItems.Item(iCount).Checked = False
        Next
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    On_Update = True
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    picOR.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    cmdORDetail.Enabled = False
    cmdInvoiceDetail.Enabled = False
    cmdInvoices.Enabled = True
    initMemvars
    On Error Resume Next
    txtOR_NUM.SetFocus
End Sub

Private Sub cmdCancel_Click()
    On_Update = False
    picOR.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    cmdORDetail.Enabled = True
    cmdInvoiceDetail.Enabled = True
    cmdInvoices.Enabled = False
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancelOR_Click()
    If CheckORCutOff(txtOR_NUM) = True Then
        MsgBox "Cancel is not allowed. Cut Off has been processed.", vbInformation, "Message"
        picOptions.Visible = False
        Exit Sub
    Else
        If MsgBox("Cancel this O.R. Entries, Are you Sure?", vbQuestion + vbYesNo, "Confirm Cancelation") = vbYes Then
            If COMPANY_CODE = "HGC" Then
                gconDMIS.Execute "Update CMIS_ORS set Status='C',CANCELLEDDATE = '" & CDate(LOGDATE) & "' where ORNO='" & txtOR_NUM.Text & "'"
                'Update By BTT:06/05/2008

                SQL_STATEMENT = "Update CMIS_OFF_HD set dateCancel='" & CDate(LOGDATE) & "' where OR_NUM='" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "C", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", "OR NO: " & Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                Else
                    NEW_LogAudit "C", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", "OR NO: " & Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                End If

            End If

            SQL_STATEMENT = "update CMIS_Off_Hd Set Cancel = 1 Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "C", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            Else
                NEW_LogAudit "C", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            End If


            SQL_STATEMENT = "update CMIS_Off_Dt Set payment = 0, Cancel = 1 Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "CC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            Else
                NEW_LogAudit "CC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            End If


            If MODE_OF_PAYMENT = "CASH" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " CASH = CASH - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If MODE_OF_PAYMENT = "CHECK" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " [CHECK] = [CHECK] - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If MODE_OF_PAYMENT = "CARD" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " CARD = CARD - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If

            rsRefresh
            rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
            picOptions.Visible = False
            StoreMemVars
        End If
    End If
End Sub

Private Sub cmdCardCancel_Click()
    lvPayments.ListItems.Clear
    picCreditCard.ZOrder 1
    picCreditCard.Visible = False
    lblView.Visible = True
End Sub

Private Sub cmdCardPayment_Click()
    'txtDiscount.Text = ToDoubleNumber(NumericVal(txtAmount.Text) * 0.05)
    'txtTax.Text = ToDoubleNumber(NumericVal(txtAmount.Text) * 0.005)
End Sub

Private Sub cmdDetails_Click()
    On_Update = True
    cmdDetails.Enabled = False
    cmdDetails.ZOrder 0
    cmdDetails.Visible = True
    picDetails.ZOrder 0
    picDetails.Visible = True
    'UPDATE BY   : MJP 09032008 05:41 PM
    'DESCRIPTION : TO LIMIT THE USER ON CLICKING THE NAVIGATION BUTTON WHILE ADDING A DETAILS
    Picture1.Enabled = False
    fraDetails.Enabled = False
    chkCreditCardTrans.Value = 0
'    SetTranTypeCode(cboTranType.Text) = "Vehicle Invoice"
    'UPDATE BY   : MJP 09032008 05:41 PM

    InitGridMemvars
    If TRANTYPE <> "" Then
        cboTranType.Text = TRANTYPE
    End If
    AddorEdit = "ADD"
    labStatusMode.Caption = "System is Adding OR Detail..."
    cmdTranSave.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub
    On_Update = True
    AddorEdit = "EDIT"
    PrevOR_NUM = txtOR_NUM.Text
    grdDetails.Enabled = False
    picOR.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    cmdORDetail.Enabled = False
    cmdInvoiceDetail.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsOFF_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdInvoiceDetail_Click()
    grdDetails.Col = 0
    INVOICE_DETAIL_TYPE = grdDetails.Text
    grdDetails.Col = 1
    INVOICE_DETAIL_TRANNO = grdDetails.Text
    frmInvoiceAppDetail.Show vbModal
End Sub

Private Sub cmdLast_Click()
    rsOFF_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsOFF_HD.MoveNext
    If rsOFF_HD.EOF Then
        rsOFF_HD.MoveLast
        ShowLastRecordMsg
        ' MsgBox "Last of Record!", vbInformation, "Info"
        '   MessagePop NaviBegin, "End of Record", "Last Record", 1000
    End If
    StoreMemVars
End Sub

Private Sub cmdOK_Click()
    If Option2.Value = True Then
        Dim xList                                                As ListItem
        Dim rsCMIS_OFF_HD                                        As ADODB.Recordset
        Set rsCMIS_OFF_HD = New ADODB.Recordset
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD Where TOF = '3'and OR_Date >= '" & dtFrom & "' and OR_Date <= '" & dtTo & "' and Paidby <> 'Y' order by OR_Date")
        If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
            lvPayments.ListItems.Clear
            lblTotal = "0.00"
            Do While Not rsCMIS_OFF_HD.EOF
                Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
                xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
                xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!CUSNAME)
                xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
                xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
                xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
                tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
                lblTotal = Format(tmpTotal, "#,###,##0.00")
                rsCMIS_OFF_HD.MoveNext
            Loop
        Else
            MessagePop RecNotFound, "No record to view", "No Record"
        End If
    End If
End Sub

Private Sub cmdOptions_Click()
    If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub
    picOptions.Visible = True
    picOptions.ZOrder 0
End Sub

Private Sub cmdORDetail_Click()
    OR_NUMBER_GLOBAL = txtOR_NUM.Text
    frmORPaymentDetail.Show vbModal
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    If vDetails = 0 Then
        MsgBox "No information to be posted!", vbInformation, "Message"
    Else
        If CheckPostedOR(txtOR_NUM) = True Then
            MsgBox "Transaction has been posted.", vbInformation, "Posted"
        Else
            picPayment.ZOrder 0
            picPayment.Visible = True
            optCASH.Value = True
            optCHECK.Value = False
            optCARD.Value = False
            optCANCEL.Value = False
            If CheckIfBank(txtCuscde.Text) = True Then
                optCARD.Enabled = False
            Else
                optCARD.Enabled = True
            End If
            On Error Resume Next
            optCASH.SetFocus
        End If
    End If
    'LogAudit "P", "OFFICIAL RECEIPT", txtOR_NUM
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrevious_Click()
    rsOFF_HD.MovePrevious
    If rsOFF_HD.BOF Then
        rsOFF_HD.MoveFirst
        ShowFirstRecordMsg
        'MsgBox "First of Record!", vbInformation, "Info"
        'MessagePop NaviBegin, "Beginning of Record", "First Record", 1000

    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LocalAcess) = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Dim xlApp                                                    As Excel.Application
    Dim xlBook                                                   As Excel.Workbook
    Dim xlSheet                                                  As Excel.Worksheet
    Dim rsTMP                                                    As New ADODB.Recordset
    Dim XCNT                                                     As Integer
    'UPDATED BY: JUN
    'DATE UPDATED: 02-10-2009
    rptChat.Reset
    'UPDATED BY: JUN

    If MsgBox("Print Official Receipt now?", vbQuestion + vbYesNo) = vbYes Then
        If COMPANY_CODE = "HEI" Then
            XCNT = 3
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Open(CMIS_REPORT_PATH & "\OR.xlt")
            Set xlSheet = xlBook.Worksheets(1)

            Set rsTMP = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_NUM = " & N2Str2Null(txtOR_NUM) & "")
            If Not (rsTMP.BOF And rsTMP.EOF) Then
                Do While Not rsTMP.EOF
                    xlSheet.Cells(XCNT, "A") = Null2String(rsTMP!TRANTYPE) & Null2String(rsTMP!REFERENCE)
                    xlSheet.Cells(XCNT, "B") = Format(NumericVal(rsTMP!payment), MAXIMUM_DIGIT)

                    XCNT = XCNT + 1
                    rsTMP.MoveNext
                Loop
            End If

            XCNT = 9
            xlSheet.Cells(XCNT, "i") = txtOR_DATE

            XCNT = 10
            xlSheet.Cells(XCNT, "F") = cboCUSNAME

            XCNT = 16
            xlSheet.Cells(XCNT, "F") = NumToText(NumericVal(rsOFF_HD!OR_AMT))

            If rsOFF_HD!TOF = 1 Then XCNT = 18
            If rsOFF_HD!TOF = 2 Then XCNT = 19
            If rsOFF_HD!TOF = 3 Then XCNT = 20
            xlSheet.Cells(XCNT, "B") = Format(NumericVal(rsOFF_HD!OR_AMT), MAXIMUM_DIGIT)

            xlApp.Windows.Item(1).Caption = "Official Receipt"
            xlApp.Visible = True
            Set xlApp = Nothing
        Else
            If OR_VAT_NONVAT = "VAT" Then
                PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
            Else
                PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 0" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
            End If
        End If
    End If

    If OR_VAT_NONVAT = "VAT" Then
        NEW_LogAudit "V", "TRANSACTION O.R. WITH VAT", "", labid, "", "OR NO: " & txtOR_NUM, "VAT", ""
    Else
        NEW_LogAudit "V", "TRANSACTION O.R. WITHOUT VAT", "", labid, "", "OR NO: " & txtOR_NUM, "NON VAT", ""
    End If

    Exit Sub

Errorcode:
    ShowVBError
End Sub

Private Sub cmdRecoverOR_Click()
    If MsgBox("Recover this O.R. Entries? Are you Sure?", vbQuestion + vbYesNo, "Confirm Recovery") = vbYes Then
        If CheckIfCancel(txtOR_NUM) = True Then
            SQL_STATEMENT = "update CMIS_Off_Hd Set Cancel = 0 Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "RC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            Else
                NEW_LogAudit "RC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            End If
            '=================================================

            SQL_STATEMENT = "update CMIS_Off_Dt Set Cancel = 0 Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "RC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            Else
                NEW_LogAudit "RC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            End If

            SQL_STATEMENT = "update CMIS_Off_Dt Set payment = " & RECEIPTS_AMOUNT & " Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT

            If MODE_OF_PAYMENT = "CASH" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " CASH = CASH + " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If MODE_OF_PAYMENT = "CHECK" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " [CHECK] = [CHECK] + " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
            If MODE_OF_PAYMENT = "CARD" Then
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                  " CARD = CARD + " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        Else
            MsgBox "OR No. is not cancelled.", vbInformation, "O.R. Recovery"
        End If

        '        If OR_VAT_NONVAT = "VAT" Then
        '            LogAudit "R", "OFFICIAL RECEIPT DATA ENTRY [VAT]", "OR NO.: " & txtOR_NUM
        '        Else
        '            LogAudit "R", "OFFICIAL RECEIPT DATA ENTRY [NON VAT]", "OR NO.: " & txtOR_NUM
        '        End If

        rsRefresh
        rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
        picOptions.Visible = False
        StoreMemVars
    End If
End Sub

Private Sub cmdRef_Click()
    If labRef.Caption = "Ref. '" Then
        labRef.Caption = "Inv. '"
    Else
        labRef.Caption = "Ref. '"
    End If
    On Error Resume Next
    txtReference.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    InitGridMemvars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Screen.MousePointer = 11
    Dim varOR_NUM, varOR_DATE, varCUSCDE, varCUSNAME             As String

    If txtCuscde.Text = "" Then
        MsgBox "Customer not yet added in Customer Master File..."
        Exit Sub
    ElseIf txtOR_NUM.Text = "" Then
        MessagePop InfoFriend, "OR Number", "OR Number is blank"
        Exit Sub
    End If
    varOR_NUM = N2Str2Null(Format(txtOR_NUM.Text, "000000"))
    varOR_DATE = N2Str2Null(txtOR_DATE.Text)
    varCUSCDE = N2Str2Null(txtCuscde.Text)
    varCUSNAME = N2Str2Null(cboCUSNAME.Text)

    Dim IS_VAT                                                   As Integer
    If OR_VAT_NONVAT = "VAT" Then
        IS_VAT = 1
    Else
        IS_VAT = 0
    End If

    Dim rsCheckORNUM                                             As ADODB.Recordset

    If AddorEdit = "ADD" Then
        Set rsCheckORNUM = gconDMIS.Execute("Select OR_NUM from CMIS_Off_hd Where VAT = " & IS_VAT & " AND OR_NUM = " & varOR_NUM)
        If Not rsCheckORNUM.EOF And Not rsCheckORNUM.BOF Then
            Screen.MousePointer = 0
            MsgBox "OR Number already used! Pls. input valid OR number...", vbCritical + vbOKOnly, "Invalid OR No."
            On Error Resume Next
            txtOR_NUM.SetFocus
            txtOR_NUM.SelLength = Len(txtOR_NUM)
            Exit Sub
        End If
    Else
        If varOR_NUM <> N2Str2Null(rsOFF_HD!OR_NUM) Then
            'If PrevOR_NUM <> txtOR_NUM.Text Then
            Set rsCheckORNUM = gconDMIS.Execute("Select OR_NUM from CMIS_Off_hd Where VAT = " & IS_VAT & " AND OR_NUM = " & varOR_NUM)
            If Not rsCheckORNUM.EOF And Not rsCheckORNUM.BOF Then
                Screen.MousePointer = 0
                MsgBox "OR Number already used! Pls. input valid OR number...", vbCritical + vbOKOnly, "Invalid OR No."
                On Error Resume Next
                txtOR_NUM.SetFocus
                txtOR_NUM.SelLength = Len(txtOR_NUM)
                Exit Sub
            End If
            'End If
        End If
    End If

    If AddorEdit = "ADD" Then
        
        SQL_STATEMENT = "Insert into CMIS_Off_hd " & _
                        "(OR_NUM,OR_DATE,CUSCDE,CUSNAME,DATECREATE,TIMECREATE,VAT,STATUS)" & _
                        " values (" & varOR_NUM & "," & varOR_DATE & "," & varCUSCDE & "," & varCUSNAME & ",'" & LOGDATE & "','" & Time & "'," & IS_VAT & ",'N')"
        gconDMIS.Execute SQL_STATEMENT
        
        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "A", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        Else
            NEW_LogAudit "A", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        End If

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CMIS_Off_Dt Set" & _
                        " OR_NUM = " & N2Str2Null(varOR_NUM) & "," & _
                        " VAT = " & VAT_OR & _
                        " where OR_NUM = " & N2Str2Null(PrevOR_NUM)
        gconDMIS.Execute SQL_STATEMENT

        If OR_VAT_NONVAT = "VAT" Then
            If NumericVal(FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd")) > 0 Then NEW_LogAudit "EE", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_Dt")
        Else
            If NumericVal(FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd")) > 0 Then NEW_LogAudit "EE", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_Dt")
        End If
        SQL_STATEMENT = " update CMIS_Off_Hd Set" & _
                        " VAT = " & VAT_OR & "," & _
                        " OR_NUM = " & N2Str2Null(varOR_NUM) & "," & _
                        " OR_DATE = " & N2Str2Null(varOR_DATE) & "," & _
                        " CUSCDE = " & N2Str2Null(varCUSCDE) & "," & _
                        " CUSNAME = " & N2Str2Null(varCUSNAME) & _
                        " where OR_NUM = " & N2Str2Null(PrevOR_NUM)
        gconDMIS.Execute SQL_STATEMENT
        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "E", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        Else
            NEW_LogAudit "E", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        End If

        ShowSuccessFullyUpdated
    End If

    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & varOR_NUM
    cmdCancel.Value = True
    FillGrid
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then cmdDetails_Click
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    SelectCustomer = "Customer"
    frmCustomerSearch1.Show 1
End Sub

Private Sub cmdTranCancel_Click()
'updating code:    JAA - 07112007
    On Error GoTo Errorcode:
    On_Update = False
    picDetails.ZOrder 1: picDetails.Visible = False
    cmdDetails.ZOrder 1: cmdDetails.Visible = False
    'UPDATE BY   : MJP 09032008 05:41 PM
    'DESCRIPTION : TO LIMIT THE USER ON CLICKING THE NAVIGATION BUTTON WHILE ADDING A DETAILS
    fraDetails.Enabled = True
    Picture1.Enabled = True
    'UPDATE BY   : MJP 09032008 05:41 PM
    StoreMemVars
    picCreditCard.ZOrder 1: picCreditCard.Visible = False
    picDeposits.ZOrder 1: picDeposits.Visible = False
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo Errorcode:

    If MsgQuestionBox("Delete This Entry, Are you Sure?", "Delete OR Entry") = True Then
        Dim rsDeposits                                           As ADODB.Recordset
        Set rsDeposits = gconDMIS.Execute("SELECT OR_Num FROM CMIS_OFF_DT WHERE OR_Num IN (SELECT OR_Num FROM CMIS_DEPOSITS WHERE OR_Num='" & lblDetID.Caption & "' AND Applied='Y')")
        If Not rsDeposits.EOF And Not rsDeposits.BOF Then
            MessagePop InfoWarning, "Applied Payment", "Customer deposit cannot be deleted!"
        Else
            SQL_STATEMENT = "delete from CMIS_Off_Dt where id = " & labDetID.Caption
            gconDMIS.Execute SQL_STATEMENT

            gconDMIS.Execute ("Delete from CMIS_Deposits where OR_Num ='" & lblDetID.Caption & "'")
            gconDMIS.Execute ("Update CMIS_Deposits SET Applied ='N',ID_DET=NULL,INVOICENO=NULL where ID_Det ='" & lblDetID.Caption & "'")
            gconDMIS.Execute "update CMIS_Off_Hd Set PAIDBY = 'N' where ReferenceNo = '" & lblRefNo & "'"
            gconDMIS.Execute "update CMIS_Off_Hd set OR_AMT=NULL,BAYADAMT=NULL,CASHAMOUNT=NULL,CHKAMOUNT=NULL,TOF=NULL,ReferenceNo=NULL,Bank=NULL where OR_NUM = '" & txtOR_NUM & "'"

            '=================================================
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "XX", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            Else
                NEW_LogAudit "XX", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            End If
            '=================================================
            ShowDeletedMsg
        End If
    End If

    rsRefresh
    On Error Resume Next
    rsOFF_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    On Error GoTo Errorcode
    If cboTranType.Text = "" Then
        MsgBox "Transaction cannot be save", vbCritical, "Error!"
        Exit Sub
    ElseIf txtReference.Text = "" Then
        MsgBox "Transaction cannot be save", vbCritical, "Error!"
        Exit Sub
    ElseIf SetTranTypeCode(cboTranType.Text) = "OTH" Then
        If cboPaidFor.Text = "" Then
            MsgBox "Field cannot be empty. Please select.", vbCritical, "Error!"
            cboPaidFor.SetFocus
            Exit Sub
        End If
    End If
    Dim vOR_NUM                                                  As String
    Dim vSUB_OR_NUM                                              As String
    Dim vReference                                               As String
    Dim vINVOICENO                                               As String
    Dim vCUSCDE                                                  As String
    Dim varCUSCDE                                                As String
    Dim vDESCRIPT                                                As String
    Dim vDOCDTE                                                  As String
    Dim vORDATE                                                  As String
    Dim vPAIDFOR                                                 As String
    Dim vBRANCH                                                  As String
    Dim vBALANCE                                                 As String
    Dim vAMOUNT                                                  As String
    Dim vPAYMENT                                                 As String
    Dim vDISCOUNT                                                As String
    Dim vTAX                                                     As Double
    Dim IS_VAT                                                   As Integer
    Dim vOVER                                                    As Double

    '***************************************************************************
    'updating code:     jaa - 11202008      - save trantype for PI,SI,MI,AI only
    Dim vInvoiceType                                             As String
    If SetTranTypeCode(cboTranType.Text) = "PI" Or SetTranTypeCode(cboTranType.Text) = "AI" Or SetTranTypeCode(cboTranType.Text) = "MI" Or SetTranTypeCode(cboTranType.Text) = "SI" Or SetTranTypeCode(cboTranType.Text) = "VI" Or SetTranTypeCode(cboTranType.Text) = "UI" Then
        vInvoiceType = N2Str2Null(cboInvoiceType.Text)
    Else
        vInvoiceType = "NULL"
    End If
    '***************************************************************************

    vOR_NUM = N2Str2Null(txtOR_NUM.Text)
    vSUB_OR_NUM = N2Str2Null(txtOR_NUM.Text)
    vTRANTYPE = N2Str2Null(SetTranTypeCode(cboTranType.Text))
    vREFERENCENO = N2Str2Null(txtReference1.Text)
    varCUSCDE = N2Str2Null(txtCuscde.Text)

    If SetTranTypeCode(cboTranType.Text) <> "RO" Then
        vReference = N2Str2Null(txtReference.Text)
        vINVOICENO = N2Str2Null(txtReference.Text)
    Else
        If labRef.Caption = "Ref. '" Then
            vReference = N2Str2Null(txtReference.Text)
            vINVOICENO = N2Str2Null(labReference.Caption)
        Else
            vReference = N2Str2Null(labReference.Caption)
            vINVOICENO = N2Str2Null(txtReference.Text)
        End If
    End If

    If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
        vCUSCDE = N2Str2Null(txtCuscde.Text)
    Else
        vCUSCDE = N2Str2Null(labCUSCODE.Caption)
    End If
    vDESCRIPT = N2Str2Null(txtDescript.Text)

    vBALANCE = NumericVal(txtBalance.Text)
    vAMOUNT = NumericVal(txtAmount.Text)

    vPAYMENT = NumericVal(txtPayment.Text)

    vDISCOUNT = NumericVal(txtDiscount.Text)
    vTAX = NumericVal(txtTax.Text)
    vOVER = NumericVal(NumericVal(txtPayment.Text) - NumericVal(txtBalance.Text))
    vPAIDFOR = N2Str2Null(SetPaidForCode(cboPaidFor.Text))
    vBRANCH = N2Str2Null(SetBranchCode(cboBranch.Text))
    If vPAYMENT <= 0 Then
        MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
        txtPayment.SetFocus
        Exit Sub
    End If

    If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
        If NumericVal(vPAYMENT) > NumericVal(vBALANCE) Then
            MsgBox "The Payment Amount is Greater than balance Amount", vbInformation, "Message"
            If MsgBox("Accept Over Payment?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            Else
                Exit Sub
            End If
        End If
    End If
    If labDocDate.Caption = "[DOC DATE]" Then vDOCDTE = "NULL" Else vDOCDTE = N2Date2Null(labDocDate.Caption)
    vORDATE = N2Str2Null(txtOR_DATE.Text)
    If OR_VAT_NONVAT = "VAT" Then IS_VAT = 1 Else IS_VAT = 0
    If AddorEdit = "ADD" Then
        'Updated: ACL 05292009
        Dim rsCardCompany                                        As ADODB.Recordset
        Dim vBankCharges                                         As Double
        Dim vEWT                                                 As Double
        Dim vTotal                                               As Double
        Set rsCardCompany = New ADODB.Recordset
        rsCardCompany.Open "Select * from CMIS_CardBank where CUSCDE = '" & txtCuscde.Text & "'", gconDMIS, adOpenKeyset
        If Not rsCardCompany.EOF And Not rsCardCompany.BOF Then
            vBankCharges = NumericVal(rsCardCompany!BankCharges) / 100
            vEWT = NumericVal(rsCardCompany!EWT) / 100
            vTotal = 1 - (vBankCharges + vEWT)
        End If

        If lvPayments.ListItems.Count <> 0 Then
            If SetPaidForCode(cboPaidFor.Text) = "427" Then
                vREFERENCENO = N2Str2Null(lvPayments.SelectedItem.SubItems(4))
                vCUSCDE = N2Str2Null(lvPayments.SelectedItem.SubItems(1))
                'vAMOUNT = NumericVal(lvPayments.SelectedItem.SubItems(3)) - (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vBankCharges + (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vEWT))
                'vPAYMENT = NumericVal(lvPayments.SelectedItem.SubItems(3)) - (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vBankCharges + (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vEWT))
                'vOR_NUM2 = lvPayments.SelectedItem.Text
                vAMOUNT = Round(NumericVal(lvPayments.SelectedItem.SubItems(3) * vTotal), 2)
                vPAYMENT = Round(NumericVal(lvPayments.SelectedItem.SubItems(3) * vTotal), 2)
                vOR_NUM2 = lvPayments.SelectedItem.Text
            End If
        End If

        SQL_STATEMENT = "Insert into CMIS_Off_Dt " & _
                        "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT)" & _
                        " values (" & vOR_NUM & "," & vInvoiceType & "," & vTRANTYPE & "," & vReference & "," & vREFERENCENO & "," & vINVOICENO & "," & vCUSCDE & "," & vDESCRIPT & "," & vBALANCE & "," & vAMOUNT & "," & vDOCDTE & "," & vORDATE & "," & vPAYMENT & "," & vDISCOUNT & "," & vTAX & "," & vPAIDFOR & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & ")"
        gconDMIS.Execute SQL_STATEMENT
        lvDeposits.ListItems.Clear

        'BANK FOR CREDIT CARD TRANSACTION
        If CheckIfBank(txtCuscde.Text) = True Then
            gconDMIS.Execute "update CMIS_Off_Hd Set PAIDBY = 'Y' where OR_NUM = '" & vOR_NUM2 & "'"
        End If

        If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
            vInvoiceType = SetPaidForCode(cboPaidFor.Text)
            Select Case vInvoiceType
            Case "412P"
                vInvoiceType = "'PI'"
            Case "412S"
                vInvoiceType = "'SI'"
            Case "412V"
                vInvoiceType = "'VI'"
            End Select
            Dim rsDet_ID                                         As ADODB.Recordset
            Set rsDet_ID = gconDMIS.Execute("select * from CMIS_OFF_DT where OR_Num = " & N2Str2Null(txtOR_NUM.Text) & "")
            If Not rsDet_ID.EOF And Not rsDet_ID.BOF Then
                SQL_STATEMENT = "Insert into CMIS_Deposits " & _
                                "(CusCde,ORDate,OR_Num,Amount,Applied,PaidFor,InvoiceType)" & _
                                "values (" & varCUSCDE & "," & vORDATE & "," & vOR_NUM & ", " & vPAYMENT & ", 'N'," & vPAIDFOR & "," & vInvoiceType & ")"
                gconDMIS.Execute SQL_STATEMENT
            End If
            Set rsDet_ID = Nothing
        End If

        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "AA", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        Else
            NEW_LogAudit "AA", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        End If


        gconDMIS.Execute ("Insert into CMIS_TranList " & _
                          "(VAT,TRANTYPE,REFERENCE,BALANCE,DOCDTE)" & _
                          " values (" & VAT_OR & "," & vTRANTYPE & "," & vReference & "," & vPAYMENT - vBALANCE & "," & vDOCDTE & ")")

        ShowSuccessFullyAdded
    Else
        vREFERENCENO = Null2String(lblRefNo.Caption)
        SQL_STATEMENT = "update CMIS_Off_Dt Set " & _
                        " VAT = " & IS_VAT & "," & _
                        " INVOICETYPE = " & vInvoiceType & "," & _
                        " TRANTYPE = " & N2Str2Null(vTRANTYPE) & "," & _
                        " REFERENCE = " & N2Str2Null(vReference) & "," & _
                        " REFERENCENO = " & N2Str2Null(vREFERENCENO) & "," & _
                        " INVOICENO = " & N2Str2Null(vINVOICENO) & "," & _
                        " CUSCDE = " & N2Str2Null(vCUSCDE) & "," & _
                        " DESCRIPT = " & N2Str2Null(vDESCRIPT) & "," & _
                        " BALANCE = " & N2Str2Zero(vBALANCE) & "," & _
                        " AMOUNT = " & N2Str2Zero(vAMOUNT) & "," & _
                        " DOCDTE = " & N2Str2Null(vDOCDTE) & "," & _
                        " ORDATE = " & N2Str2Null(vORDATE) & "," & _
                        " PAYMENT = " & N2Str2Zero(vPAYMENT) & "," & _
                        " DISCOUNT = " & N2Str2Zero(vDISCOUNT) & "," & _
                        " TAX = " & N2Str2Null(vTAX) & "," & _
                        " PAIDFOR = " & N2Str2Null(vPAIDFOR) & "," & _
                        " BRANCH = " & N2Str2Null(vBRANCH) & "," & _
                        " [OVER] = " & N2Str2Null(vOVER) & _
                        " Where ID = " & labDetID.Caption

        gconDMIS.Execute SQL_STATEMENT

        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "EE", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        Else
            NEW_LogAudit "EE", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        End If
        '=================================================
        ShowSuccessFullyUpdated
    End If

    rsRefresh
    rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
    StoreMemVars
    cmdTranCancel.Value = True
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdVarious_Click()
'UPDATED BY AXP-061920071101
'frmALLCustomer.Show vbModal
''FillCustomer
'If CURRENT_CUST_CODE <> "" Then
'    txtCUSCDE.Text = CURRENT_CUST_CODE
'    cboCUSNAME.Text = SetCustomerName(txtCUSCDE.Text)
'End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdInvoices_Click()
    If txtOR_NUM.Text = "" Then
        MessagePop InfoFriend, "OR Number", "OR Number is blank"
        Exit Sub
    Else
        frmCMIS_viewInvoiceDetail.Show 1
    End If
End Sub

Private Sub Command4_Click()
    picDeposits.Visible = False
    cmdTranSave.Enabled = True
End Sub

Private Sub Form_Activate()
'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If picDetails.Visible = True Then
            If picDeposits.Visible = True Then
                picDetails.Enabled = True
                picDeposits.Visible = False
                picDeposits.ZOrder 1
            ElseIf picCreditCard.Visible = True Then
                picCreditCard.Visible = False
                picCreditCard.ZOrder 1
            Else
                picDetails.Enabled = True
                fraDetails.Enabled = True
                On_Update = False
                picDetails.ZOrder 1
                picDetails.Visible = False

                cmdDetails.ZOrder 1
                cmdDetails.Visible = False
                picDeposits.Visible = False
                picCreditCard.Visible = False
                picCreditCard.ZOrder 1

                Picture1.Enabled = True
                On Error Resume Next
                grdDetails.SetFocus
            End If
        End If
        If picOptions.Visible = True Then
            picOptions.Visible = False
            picOptions.ZOrder 1
        End If
        'If picDetail.Visible = True Then
        '   picDetail.Visible = False
        'End If
    Case vbKeyF2
        If Null2Bool(rsOFF_HD!PaidNa) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
            lblView.Visible = False
            If Picture1.Visible = True Then cmdDetails_Click
        End If
    Case vbKeyF3
        grdDetails_DblClick
        'Case vbKeyF7
        '     picDetail.ZOrder 0
        '     picDetail.Visible = True
        '     cmdORDetail.SetFocus
    Case vbKeyF5
        If SetPaidForCode(cboPaidFor.Text) = "427" Then
            If CheckIfBank(txtCuscde.Text) = True Then
                picCreditCard.Visible = True
                picCreditCard.ZOrder 0
                lblView.Visible = False
            End If
        End If
    Case vbKeyF6
        If picDetails.Visible = True Then
            picDeposits.Visible = True
            picDeposits.ZOrder 0
            cmdTranSave.Enabled = False
            Call Unapplied_Deposits(Null2String(txtCuscde.Text))
        End If
    Case vbKeyF7
        If picOR.Enabled = True Then Call cmdSelect_Click
    Case vbKeyF8
        If picDetails.Visible = False Then cmdPost_Click
    Case vbKeyF11
        Shell "calc.exe"
    Case vbKeyF12
        If CheckPostedOR(txtOR_NUM.Text) = False Then
            MsgBox "Transaction is not yet posted.", vbInformation, "Message"
            Exit Sub
        Else
            If MsgBox("Unpost this O.R. Entries, Are you Sure?", vbQuestion + vbYesNo, "Unpost Transaction") = vbYes Then
                If CheckORCutOff(txtOR_NUM) = True Then
                    MsgBox "Unposting is not allowed. Cut Off has been processed.", vbInformation, "Message"
                    Exit Sub
                ElseIf CheckIfImportedinAMIS(txtOR_NUM) = True Then
                    MsgBox "Unposting is not allowed. Already imported in accounting.", vbInformation, "Message"
                    Exit Sub
                ElseIf CheckAppliedDeposit(txtOR_NUM) = True Then
                    MessagePop InfoWarning, "Applied Deposit", "Customer deposit cannot be unposted!"
                    Exit Sub
                Else
                    'DESCRIPTION: DEDUCT FROM CMIS_CASH_POS IF TRANSACTION IS UNPOSTED
                    '             POSTED NOT YET DEPOSITED
                    If CheckDeposited(txtOR_NUM) = False Then
                        Call UnPost_CashPos
                    Else
                        'DESCRIPTION: POSTED AND DEPOSITED
                        Call UnPost_CashPos
                        'DESCRIPTION: UNPOSTING OF TRANSACTION, SET DEPOSIT=FALSE
                        If COMPANY_CODE = M_COMPANY_CODE Then
                            gconDMIS.Execute ("update CMIS_Off_Hd Set Deposit1 = 0 Where OR_NUM = '" & txtOR_NUM & "'")
                            gconDMIS.Execute ("update CMIS_Off_Hd Set Deposit2 = 0 Where OR_NUM = '" & txtOR_NUM & "'")
                            gconDMIS.Execute ("update CMIS_Off_Hd Set Deposit3 = 0 Where OR_NUM = '" & txtOR_NUM & "'")
                        Else
                            gconDMIS.Execute ("update CMIS_Off_Hd Set Deposit = 0 Where OR_NUM = '" & txtOR_NUM & "'")
                        End If
                        'DESCRIPTION: DELETE FROM BANKDEPOSIT AND CASH POSITION IF CUT OFF IS NOT YET PROCESS
                        gconDMIS.Execute ("Delete from CMIS_BankDepo where OR_NUM = " & N2Str2Null(txtOR_NUM))

                    End If
                End If
                '================================================
                'UPDATING CODE:     JAA - 08272008   11:00PM
                SQL_STATEMENT = "update CMIS_Off_Hd Set paidna = 0, STATUS='N' Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "U", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_hd"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                Else
                    NEW_LogAudit "U", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_hd"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                End If
                SQL_STATEMENT = "update CMIS_Off_Dt Set paidna = 0, STATUS='N' Where VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "UU", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                Else
                    NEW_LogAudit "UU", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                End If
                '=================================================

                rsRefresh
                rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
                StoreMemVars
            End If
        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
    If Shift = 1 Then
        If KeyCode = vbKeyF1 Then
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If OR_VAT_NONVAT = "VAT" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION O.R. WITH VAT)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "TRANSACTION O.R. WITH VAT")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION O.R. WITHOUT VAT)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "TRANSACTION O.R. WITHOUT VAT")
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Dim rsProfile                                                As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing
    CenterMe frmMain, Me, 1: picOptions.Visible = False
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    picOR.Enabled = False: FillGrid: initMemvars
    textSearch.Text = "": FillCustomer: FillType: FillBranch: FillPayment: FillInvoiceType
    On_Update = False
    If OR_VAT_NONVAT = "VAT" Then
        VAT_OR = 1
    Else
        VAT_OR = 0
    End If
    FIRST_LOAD = True: rsRefresh: FIRST_LOAD = False: StoreMemVars
    picPayment.Top = 3120
    dtFrom = LOGDATE
    dtTo = LOGDATE
    ChangeORNum = False
    Screen.MousePointer = 0
End Sub

Sub FillInvoiceType()
    cboInvoiceType.Clear
    cboInvoiceType.AddItem "CSH"
    cboInvoiceType.AddItem "CHG"
    cboInvoiceType.AddItem "DR"
    cboInvoiceType.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsOFF_HD = Nothing
    Set rsOFF_DT = Nothing
    Set frmCMISOREntry = Nothing
    LocalAcess = ""
End Sub

Private Sub grdDetails_Click()
    grdDetails.Col = 10
    If grdDetails.Text <> "" Then
        ShowGridDetails grdDetails.Text
    End If
End Sub

Private Sub grdDetails_DblClick()
    If Null2Bool(rsOFF_HD!PaidNa) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
        grdDetails.Col = 10
        If grdDetails.Text <> "" Then
            On_Update = True
            cmdDetails.Enabled = False: cmdDetails.ZOrder 0
            cmdDetails.Visible = True: picDetails.ZOrder 0
            picDetails.Visible = True: fraDetails.Enabled = False: Picture1.Enabled = False
            StoreGridDetails grdDetails.Text
            If SetPaidForCode(cboPaidFor.Text) = "427" Then
                chkCreditCardTrans.Value = 0
                chkCreditCardTrans.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lvDeposits_DblClick()
'DESCRIPTION: Apply customer deposits
    Dim vINVOICENO                                               As String
    Dim reply                                                    As String
    Dim iCtr                                                     As Integer
    If Not lvDeposits.SelectedItem Is Nothing Then
        reply = MsgBox("Are you sure you want to apply" + vbCrLf + "this customer deposit?", vbQuestion + vbYesNo, "Customer Deposit")
        If reply = vbYes Then
            txtPayment.Text = ToDoubleNumber(NumericVal(txtPayment.Text) - NumericVal(lvDeposits.SelectedItem.SubItems(3)))
            'txtBalance.Text = ToDoubleNumber(NumericVal(txtPayment.Text) - NumericVal(lvDeposits.SelectedItem.SubItems(3)))
            'picDeposits.Visible = False
            If SetTranTypeCode(cboTranType.Text) <> "RO" Then
                vINVOICENO = N2Str2Null(txtReference.Text)
            Else
                If labRef.Caption = "Ref. '" Then
                    vINVOICENO = N2Str2Null(labReference.Caption)
                Else
                    vINVOICENO = N2Str2Null(txtReference.Text)
                End If
            End If
            If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
                'gconDMIS.Execute ("Update CMIS_Deposits Set Applied = 'Y',ID_Det = '" & txtOR_NUM & "',InvoiceNo =" & vINVOICENO & " where Cuscde ='" & txtCUSCDE.Text & "' and OR_Num = '" & lvDeposits.SelectedItem.SubItems(2) & "'")
                gconDMIS.Execute ("Update CMIS_Deposits Set Applied = 'Y',ID_Det = '" & txtOR_NUM & "',InvoiceNo =" & vINVOICENO & " where ID ='" & lblDepositID.Caption & "'")
            End If
            picDetails.Enabled = True
            cmdTranSave.Enabled = True
            iCtr = lvDeposits.SelectedItem.Index
            lvDeposits.ListItems.Remove (iCtr)
            'lvDeposits.ListItems.ITEM(iCtr).ForeColor = vbRed
            If lvDeposits.ListItems.Count = 0 Then
                picDeposits.Visible = False
            End If

        Else
            Exit Sub
        End If
    Else
        MessagePop RecNotFound, "", "No Record Found"
    End If
End Sub

Private Sub lvDeposits_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblDepositID.Caption = lvDeposits.SelectedItem.SubItems(6)
End Sub

Private Sub lvPayments_DblClick()
'DESCRIPTION: Apply payments from Credit Card Company with Bank Charges and Expanded Withheld Tax
    Dim rsCardCompany                                            As ADODB.Recordset
    Dim vBankCharges                                             As Double
    Dim vEWT                                                     As Double
    Set rsCardCompany = New ADODB.Recordset
    rsCardCompany.Open "Select * from CMIS_CardBank where CUSCDE = '" & txtCuscde.Text & "'", gconDMIS, adOpenKeyset
    If Not rsCardCompany.EOF And Not rsCardCompany.BOF Then
        vBankCharges = NumericVal(rsCardCompany!BankCharges) / 100
        vEWT = NumericVal(rsCardCompany!EWT) / 100
    End If

    If Not lvPayments.SelectedItem Is Nothing Then
        txtPayment = Format(NumericVal(lvPayments.SelectedItem.SubItems(3)) - (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vBankCharges + (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vEWT)), "#,###,##0.00")
        txtDiscount.Text = ToDoubleNumber(lvPayments.SelectedItem.SubItems(3)) * vBankCharges
        txtTax.Text = ToDoubleNumber(lvPayments.SelectedItem.SubItems(3)) * vEWT
        picCreditCard.Visible = False
    Else
        MessagePop RecNotFound, "", "No Record Found"
    End If
End Sub

Private Sub optCANCEL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        picPayment.ZOrder 1: picPayment.Visible = False
    End If
End Sub

Private Sub optCARD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CARD"
        picPayment.ZOrder 1: picPayment.Visible = False

        frmCMISCARDPaymentEntry.Show vbModal
    End If
End Sub

Private Sub optCASH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CASH"
        picPayment.ZOrder 1: picPayment.Visible = False
        frmCMISCASHPaymentEntry.Show vbModal
    End If
End Sub

Private Sub optCHECK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CHECK"
        picPayment.ZOrder 1: picPayment.Visible = False
        frmCMISCHECKPaymentEntry.Show vbModal
    End If
End Sub

Private Sub Option1_Click()
    Picture3.Visible = True
    Picture5.Visible = False
    picCustomer.Visible = False
    lvPayments.ListItems.Clear
    txtReference2 = ""
    'lvPayments.Checkboxes = False
    txtReference2.SetFocus
    txtPayment = "0.00"
    CreditCardPayments
End Sub

Private Sub Option2_Click()
    Picture3.Visible = False
    Picture5.Visible = True
    picCustomer.Visible = False
    lvPayments.ListItems.Clear
    'lvPayments.Checkboxes = True
    lblTotal = "0.00"
    txtPayment = "0.00"
    CreditCardPayments
End Sub

Private Sub Option3_Click()
    Picture3.Visible = False
    Picture5.Visible = False
    picCustomer.Visible = True
    lvPayments.ListItems.Clear
    txtCustomer.Text = ""
    'lvPayments.Checkboxes = True
    lblTotal = "0.00"
    txtPayment = "0.00"
    txtCustomer.SetFocus
    CreditCardPayments
End Sub

Private Sub textSearch_Change()
    If optORNo.Value = True Then
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

Private Sub Timer1_Timer()
    If On_Update = False Then
        If rsOFF_HD!Cancel = True Then
            labStatus.Caption = "*** Cancelled OR ***"
            If labStatus.Visible = True Then
                labStatus.Visible = False
            Else
                labStatus.Visible = True
            End If
        Else
            If rsOFF_HD!PaidNa = True Then
                labStatus.Caption = "*** PAID OR ***"
                If labStatus.Visible = True Then
                    labStatus.Visible = False
                Else
                    labStatus.Visible = True
                End If
            Else
                labStatus.Visible = False
            End If
        End If
    Else
        labStatus.Visible = False
    End If
End Sub

Private Sub txtBalance_Change()
    txtPayment.Text = ToDoubleNumber(NumericVal(txtBalance.Text) - (NumericVal(txtDiscount.Text) + NumericVal(txtTax.Text)))
End Sub

Private Sub txtCustomer_Change()
    Dim xList                                                    As ListItem
    Dim rsCMIS_OFF_HD                                            As ADODB.Recordset
    Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD Where TOF = '3' and Paidby <> 'Y' and CusName like '" & txtCustomer.Text & "%' order by OR_Date")
    If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
        lvPayments.ListItems.Clear
        lblTotal = "0.00"
        Do While Not rsCMIS_OFF_HD.EOF
            Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
            xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
            xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!CUSNAME)
            xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
            xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
            xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
            tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
            lblTotal = Format(tmpTotal, "#,###,##0.00")
            rsCMIS_OFF_HD.MoveNext
        Loop
    Else
        'MsgBox "No customer record to view", vbCritical, "No Record"
        '        MessagePop RecNotFound, "No record to view", "No Record"
    End If
End Sub

Private Sub txtDescript_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDiscount_Change()
'txtPayment.Text = ToDoubleNumber(NumericVal(txtBalance.Text) - (NumericVal(txtDiscount.Text) + NumericVal(txtTax.Text)))
End Sub

Private Sub txtDiscount_GotFocus()
    txtDiscount.Text = NumericVal(txtDiscount.Text)
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = ToDoubleNumber(txtDiscount.Text)
End Sub

Private Sub txtOR_DATE_GotFocus()
    If IsDate(txtOR_DATE.Text) = True Then txtOR_DATE.Text = Format(txtOR_DATE.Text, "MM/DD/YYYY") Else txtOR_DATE.Text = ""
End Sub

Private Sub txtOR_DATE_LostFocus()
    If IsDate(txtOR_DATE.Text) = True Then txtOR_DATE.Text = Format(txtOR_DATE.Text, "DD-MMM-YYYY") Else txtOR_DATE.Text = ""
End Sub

Private Sub txtOR_NUM_KeyPress(KeyAscii As Integer)
'    KeyAscii = OnlyNumeric(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtOR_NUM_LostFocus()
    txtOR_NUM.Text = Format(txtOR_NUM.Text, "000000")
End Sub

Private Sub txtPayment_Change()
    If NumericVal(txtPayment.Text) <= 0 Then
        wizDigit1.TextValue = 0
    Else
        If AddorEdit = "EDIT" Then
            'wizDigit1.TextValue = ToDoubleNumber(NumericVal(TOTAL_AR_AMOUNT) + NumericVal(txtBalance.Text))
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtPayment.Text))
        Else
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtPayment.Text))
        End If
    End If
End Sub

Private Sub txtPayment_GotFocus()
    txtPayment.Text = NumericVal(txtPayment.Text)
End Sub

Private Sub txtPayment_LostFocus()
    txtPayment.Text = ToDoubleNumber(txtPayment.Text)
End Sub

Private Sub txtReference_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtReference.Text <> "" Then
        txtReference.Text = Format(txtReference.Text, "000000")
        cboPaidFor.ListIndex = -1
        If labRef.Caption = "Inv. #" Then txtReference.Text = Format(txtReference.Text, "000000")
        Dim rsOrd_Hd                                             As ADODB.Recordset
        Dim rsOFF_DT                                             As ADODB.Recordset
        If SetTranTypeCode(cboTranType.Text) = "SI" Then
            Dim rsREPOR                                          As ADODB.Recordset
            Set rsREPOR = New ADODB.Recordset
            If labRef.Caption = "Ref. '" Then
                Set rsREPOR = gconDMIS.Execute("Select Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt from CSMS_REPOR Where Rep_or = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCuscde.Text))
            Else
                Set rsREPOR = gconDMIS.Execute("Select Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt from CSMS_REPOR Where invoice = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCuscde.Text))
            End If
            If Not rsREPOR.EOF And Not rsREPOR.BOF Then
                If labRef.Caption = "Ref. '" Then
                    labReference.Caption = Null2String(rsREPOR!invoice)
                Else
                    labReference.Caption = Null2String(rsREPOR!REP_OR)
                End If
                txtDescript.Text = Null2String(rsREPOR!niym)
                txtAmount.Text = ToDoubleNumber(NumericVal(rsREPOR!RO_Amount) - NumericVal(rsREPOR!INSAMT))
                labDocDate.Caption = Null2Date(rsREPOR!dte_comp)
                labCUSCODE.Caption = Null2String(rsREPOR!acct_no)
                Set rsOFF_DT = New ADODB.Recordset
                '                Set rsOFF_DT = gconDMIS.Execute("Select SUM(PAYMENT) as MGA_BAYAD from CMIS_Off_Dt Where (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' and Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCUSCDE.Text) & "")
                '                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                '                    txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                '                    Call BalanceCash(cboInvoiceType, txtReference)
                '                Else
                '                    txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                '                End If
                Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' and Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCuscde.Text) & " GROUP BY REFERENCE,TRANTYPE")
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    Set rsCustomerDeposit = New ADODB.Recordset
                    rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                    If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                        Call BalanceCash(cboInvoiceType, txtReference)
                    End If
                Else
                    txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                End If
                Call Unapplied_Deposits(Null2String(txtCuscde.Text))
            Else
                If labRef.Caption = "Ref. '" Then
                    Set rsREPOR = gconDMIS.Execute("Select Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR Where Rep_or = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCuscde.Text))
                Else
                    Set rsREPOR = gconDMIS.Execute("Select Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR Where invoice = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCuscde.Text))
                End If
                If Not rsREPOR.EOF And Not rsREPOR.BOF Then
                    If labRef.Caption = "Ref. '" Then
                        labReference.Caption = Null2String(rsREPOR!invoice)
                    Else
                        labReference.Caption = Null2String(rsREPOR!REP_OR)
                    End If
                    txtDescript.Text = Null2String(rsREPOR!InsCde)
                    txtAmount.Text = ToDoubleNumber(N2Str2Zero(rsREPOR!PARTLABOR) + N2Str2Zero(rsREPOR!PARTPARTS) + N2Str2Zero(rsREPOR!PARTMATERIALS) + N2Str2Zero(rsREPOR!PARTACCESSORIES))
                    labDocDate.Caption = Null2Date(rsREPOR!dte_comp)
                    labCUSCODE.Caption = Null2String(rsREPOR!Participat)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select SUM(PAYMENT) as MGA_BAYAD from CMIS_Off_Dt Where (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' and Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCUSCDE.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCash(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' and Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCuscde.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))

                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    On Error Resume Next
                    txtReference.SetFocus
                    Exit Sub
                End If

            End If
        End If
        If SetTranTypeCode(cboTranType.Text) = "PI" Then
            Set rsOrd_Hd = New ADODB.Recordset
            '************************************************
            'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
            If cboInvoiceType = "CSH" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'P' AND trantype = 'CSH' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)

                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select Round(Sum(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCash(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    cmdDetails_Click
                    '                    txtReference.SetFocus
                    Exit Sub
                End If
            ElseIf cboInvoiceType = "CHG" Then
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'P' AND trantype = 'CHG' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select Round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' and TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' and TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            Else
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'P' AND trantype = 'DR' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'CHG' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' and TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            End If
            '************************************************
        End If
        If SetTranTypeCode(cboTranType.Text) = "AI" Then
            Set rsOrd_Hd = New ADODB.Recordset
            '************************************************
            'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
            If cboInvoiceType = "CSH" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'A' AND trantype = 'CSH' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType ='AI' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCash(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType = 'AI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            ElseIf cboInvoiceType = "CHG" Then
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'A' AND trantype = 'CHG' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'CHG' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' and TranType = 'AI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            Else
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'A' AND trantype = 'DR' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'CHG' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' and TranType = 'AI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            End If
        End If
        If SetTranTypeCode(cboTranType.Text) = "MI" Then
            Set rsOrd_Hd = New ADODB.Recordset
            '************************************************
            'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
            If cboInvoiceType = "CSH" Then
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'M' AND trantype = 'CSH' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType = 'MI' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCash(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' and TranType = 'MI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            ElseIf cboInvoiceType = "CHG" Then
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'M' AND trantype = 'CHG' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'CHG' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' and TranType = 'MI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    Call Unapplied_Deposits(Null2String(txtCuscde.Text))
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            Else
                Set rsOrd_Hd = New ADODB.Recordset
                Set rsOrd_Hd = gconDMIS.Execute("Select custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY where TYPE = 'M' AND trantype = 'DR' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCuscde.Text))
                If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                    txtDescript.Text = Null2String(rsOrd_Hd!custname)
                    txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                    labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                    labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                    Set rsOFF_DT = New ADODB.Recordset
                    '                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND trantype = 'CHG' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    '                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    '                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
                    '                        Call BalanceCharge(cboInvoiceType, txtReference)
                    '                    Else
                    '                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    '                    End If
                    Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' and TranType = 'MI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            End If
        End If

        If SetTranTypeCode(cboTranType.Text) = "VI" Then
            Dim rsPurchAgree                                     As ADODB.Recordset
            Set rsPurchAgree = New ADODB.Recordset
            Set rsPurchAgree = gconDMIS.Execute("Select SMIS_PurchAgree.code,SMIS_PurchAgree.deyt,SMIS_PurchAgree.NetSalesPrice AS Total,ALL_Customer.LastName + ALL_Customer.FirstName as CustomerName From ALL_Customer Inner Join SMIS_PurchAgree on ALL_Customer.CusCde = SMIS_PurchAgree.Code Where SMIS_PurchAgree.VI_No = " & N2Str2Null(txtReference.Text) & " AND SMIS_PurchAgree.code =" & N2Str2Null(txtCuscde.Text))
            If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                txtDescript.Text = Null2String(rsPurchAgree!CUSTOMERNAME)
                txtAmount.Text = ToDoubleNumber(rsPurchAgree!Total)
                labDocDate.Caption = Null2Date(rsPurchAgree!deyt)
                labCUSCODE.Caption = Null2String(rsPurchAgree!code)
                Set rsOFF_DT = New ADODB.Recordset
                Set rsOFF_DT = gconDMIS.Execute("Select round(SUM(PAYMENT),2) as MGA_BAYAD,TRANTYPE,REFERENCE from CMIS_Off_Dt Where trantype = 'VI' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE")
                If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                    Set rsCustomerDeposit = New ADODB.Recordset
                    rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TRANTYPE) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                    If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                        Call BalanceCash(cboInvoiceType, txtReference)
                    End If
                Else
                    txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                End If
                Call Unapplied_Deposits(Null2String(txtCuscde.Text))
            Else
                MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                txtReference.SetFocus
                Exit Sub
            End If
        End If
        
        'If COMPANY_CODE = "HMH" Then
            If SetTranTypeCode(cboTranType.Text) = "UI" Then
                Dim rsJournalHD As ADODB.Recordset
                Set rsJournalHD = New ADODB.Recordset
                rsJournalHD.Open "SELECT CUSTOMERCODE,CUSTOMERNAME,JDATE,INVOICEAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE=" & N2Str2Null(txtCuscde.Text) & "", gconDMIS, adOpenKeyset
                If Not rsJournalHD.EOF And Not rsJournalHD.BOF Then
                    txtDescript.Text = Null2String(rsJournalHD!CUSTOMERNAME)
                    txtAmount.Text = ToDoubleNumber(rsJournalHD!INVOICEAMT)
                    labDocDate.Caption = Null2Date(rsJournalHD!JDATE)
                    labCUSCODE.Caption = Null2String(rsJournalHD!CustomerCode)
                    
                    Set rsOFF_DT = gconDMIS.Execute("Select ISNULL(ROUND(SUM(PAYMENT),2),0) as MGA_BAYAD from CMIS_Off_Dt Where VAT = " & VAT_OR & " AND TranType = 'UI' and Reference = " & N2Str2Null(txtReference.Text) & "")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!Mga_Bayad)))
                        Call BalanceCash(cboInvoiceType, txtReference)
                    Else
                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                    End If
                    
                Else
                    MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                    txtReference.SetFocus
                    Exit Sub
                End If
            End If
        'End If
        
        If SetTranTypeCode(cboTranType.Text) = "SI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = True: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "PI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "AI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "MI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "VI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "UI" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = False
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "OTH" Then
            If lvDeposits.ListItems.Count <> 0 Then
                cmdTranSave.Enabled = False
            Else
                cmdTranSave.Enabled = True
            End If
            cboPaidFor.Enabled = True: cboBranch.Enabled = True: txtDescript.Enabled = True
            txtDiscount.Enabled = True: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
        If SetTranTypeCode(cboTranType.Text) = "EST" Then
            cmdTranSave.Enabled = True
            cboPaidFor.Enabled = False: cboBranch.Enabled = False: txtDescript.Enabled = True
            txtDiscount.Enabled = False: txtTax.Enabled = True: txtPayment.Enabled = True
        End If
    End If
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtReference_LostFocus()
    If labRef.Caption = "Inv. #" Then txtReference.Text = Format(txtReference.Text, "000000")
End Sub

Private Sub txtReference2_Change()
'DESCRIPTION: Search for OR No. to be linked with CITIBANK Payment
    Dim xList                                                    As ListItem
    Dim rsCMIS_OFF_HD                                            As ADODB.Recordset
    Set rsCMIS_OFF_HD = New ADODB.Recordset
    Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD Where TOF = '3' and Paidby <> 'Y' and OR_NUM like '" & txtReference2 & "%' order by OR_Date")
    If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
        lvPayments.ListItems.Clear
        lblTotal = "0.00"
        Do While Not rsCMIS_OFF_HD.EOF
            Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
            xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
            xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!CUSNAME)
            xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
            xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
            xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
            tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
            lblTotal = Format(tmpTotal, "#,###,##0.00")
            rsCMIS_OFF_HD.MoveNext
        Loop
    End If
    Set rsCMIS_OFF_HD = Nothing
End Sub

Private Sub txtTax_Change()
'    txtPayment.Text = ToDoubleNumber(NumericVal(txtBalance.Text) - (NumericVal(txtDiscount.Text) + NumericVal(txtTax.Text)))
End Sub

Private Sub txtTax_GotFocus()
    txtTax.Text = NumericVal(txtTax.Text)
End Sub

Private Sub txtTax_LostFocus()
    txtTax.Text = ToDoubleNumber(txtTax.Text)
End Sub

'SEARCH MODULE
Private Sub lstOFF_HD_GotFocus()
'On Error Resume Next
    If lstOFF_HD.Enabled = True Then
        rsOFF_HD.MoveFirst
        rsOFF_HD.Find ("ID=" & lstOFF_HD.SelectedItem.SubItems(1))
        StoreMemVars
    End If
End Sub

Private Sub lstOFF_HD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lstOFF_HD.Enabled = True Then
        If optORNo.Value = True Then
            'On Error Resume Next
            rsOFF_HD.Bookmark = rsFind(rsOFF_HD.Clone, "OR_NUM", Item).Bookmark
        Else
            'On Error Resume Next
            rsOFF_HD.Bookmark = rsFind(rsOFF_HD.Clone, "ID", lstOFF_HD.SelectedItem.SubItems(1)).Bookmark
        End If
        StoreMemVars
    End If
End Sub

Private Sub lstOFF_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOFF_HD
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

Private Sub lstOFF_HD_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstOFF_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Then
'   If optORNo.Value = True Then
'       If Trim(textSearch.Text) = "" Then
'         FillGrid
'       Else
'          FillSearchGrid (textSearch.Text)
'       End If
'   Else
'       If Trim(textSearch.Text) = "" Then
'           FillGrid2
'       Else
'           FillSearchGrid2 (textSearch.Text)
'       End If
'   End If
'End If

    If KeyCode = vbKeyDown Then

        If lstOFF_HD.ListItems.Count > 0 And lstOFF_HD.Enabled = True Then: lstOFF_HD.SetFocus
    End If
End Sub

Private Sub optCustName_Click()
    lstOFF_HD.Enabled = False
    lstOFF_HD.ColumnHeaders(1).Text = "Cust. Name"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optORNo_Click()
    lstOFF_HD.Enabled = False
    lstOFF_HD.ColumnHeaders(1).Text = "OR No."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub


'Function GETTRANTYPE(vtranno As String) As String
'    Dim rsOrdTranno As ADODB.Recordset
'    Set rsOrdTranno = New ADODB.Recordset
'    Set rsOrdTranno = gconDMIS.Execute("SELECT TRANNO,TRANTYPE FROM PMIS_ORD_HD WHERE TYPE = '" & Mid(vTRANTYPE, 2, 1) & "' AND TRANNO = " & N2Str2Null(vtranno) & "")
'    If Not rsOrdTranno.EOF And Not rsOrdTranno.BOF Then
'        GETTRANTYPE = Null2String(rsOrdTranno!TRANTYPE)
'    Else
'        GETTRANTYPE = "NULL"
'    End If
'    Set rsOrdTranno = Nothing
'End Function

'DESCRIPTION: Generate new ReferenceNo upon posting credit card transaction
Function GetReferenceNo() As String
    Dim rsCMIS_OFF_HD                                            As ADODB.Recordset
    Set rsCMIS_OFF_HD = New ADODB.Recordset
    Set rsCMIS_OFF_HD = gconDMIS.Execute("Select CAST(ReferenceNo AS int) AS MAX_REFERENCENO from CMIS_Off_HD Order by MAX_REFERENCENO desc")
    If Not rsCMIS_OFF_HD.EOF And Not rsCMIS_OFF_HD.BOF Then
        GetReferenceNo = Format(NumericVal(rsCMIS_OFF_HD!MAX_REFERENCENO) + 1, "00000000")
    Else
        GetReferenceNo = "00000001"
    End If
End Function

Function BalanceCash(xINVOICETYPE As String, xReference As String)
'DESCRIPTION: Check for Customer Balance
    Dim rsOFF_DTStat                                             As ADODB.Recordset
    Set rsOFF_DTStat = New ADODB.Recordset
    Set rsOFF_DTStat = gconDMIS.Execute("select PaidNa from CMIS_OFF_DT Where INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " and Reference = " & N2Str2Null(xReference))
    If Not rsOFF_DTStat.EOF And Not rsOFF_DTStat.BOF Then
        If txtBalance.Text <= 0 And rsOFF_DTStat!PaidNa = True Then
            cmdTranCancel.Value = True
            MessagePop Star, "Information", "Balance has been fully paid."
        ElseIf txtBalance.Text <= 0 And rsOFF_DTStat!PaidNa = False Then
            Call Unapplied_Deposits(Null2String(txtCuscde.Text))
            'MessagePop InfoWarning, "Information", "Payment has been made but not yet POSTED."
            If Null2Bool(rsOFF_HD!PaidNa) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
                If Picture1.Visible = True Then cmdDetails_Click
                chkCreditCardTrans.Value = 0
            End If
        End If
    End If
    Set rsOFF_DTStat = Nothing
End Function

Sub BalanceCharge(xINVOICETYPE As String, xReference As String)
    Dim rsOFF_DTStat                                             As ADODB.Recordset
    Set rsOFF_DTStat = New ADODB.Recordset
    Set rsOFF_DTStat = gconDMIS.Execute("select PaidNa from CMIS_OFF_DT Where INVOICETYPE = " & N2Str2Null(xINVOICETYPE) & " and Reference = " & N2Str2Null(xReference))
    If Not rsOFF_DTStat.EOF And Not rsOFF_DTStat.BOF Then
        If txtBalance.Text <= 0 And rsOFF_DTStat!PaidNa = True Then
            cmdTranCancel.Value = True
            MessagePop Star, "Information", "Balance has been fully paid."
        ElseIf txtBalance.Text <= 0 And rsOFF_HD!PaidNa = False Then
            MessagePop InfoWarning, "Information", "Payment has been made but not yet POSTED."
            If Null2Bool(rsOFF_HD!PaidNa) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
                If Picture1.Visible = True Then cmdDetails_Click
                chkCreditCardTrans.Value = 0
            End If
        End If
    End If
    Set rsOFF_DTStat = Nothing
End Sub

Sub Unapplied_Deposits(XXX As String)
'DESCRIPTION: List of Customer Deposits
    Dim xList                                                    As ListItem
    Dim rsUnapplied                                              As ADODB.Recordset
    Set rsUnapplied = New ADODB.Recordset
    rsUnapplied.Open "SELECT HD.OR_NUM,HD.STATUS,DP.* FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM WHERE HD.CUSCDE ='" & XXX & "' AND HD.PAIDNA =1 AND DP.APPLIED <> 'Y'", gconDMIS, adOpenKeyset
    lvDeposits.ListItems.Clear
    If Not rsUnapplied.EOF And Not rsUnapplied.BOF Then
        picDeposits.Visible = True
        picDeposits.ZOrder 0
        cmdTranSave.Enabled = False
        Do While Not rsUnapplied.EOF
            Set xList = lvDeposits.ListItems.Add(, , Null2String(rsUnapplied!CUSCDE))
            xList.SubItems(1) = Null2Date(Format(rsUnapplied!ORDATE, "mm/dd/yyyy"))
            xList.SubItems(2) = rsUnapplied!OR_NUM
            xList.SubItems(3) = ToDoubleNumber(rsUnapplied!AMOUNT)
            xList.SubItems(4) = N2Str2Null(rsUnapplied!Applied)
            xList.SubItems(5) = Null2String(rsUnapplied!ID_Det)
            xList.SubItems(6) = Null2String(rsUnapplied!Id)
            xList.SubItems(7) = SetPaidFor(Null2String(rsUnapplied!PAIDFOR))
            rsUnapplied.MoveNext
        Loop
    End If
    Set rsUnapplied = Nothing
End Sub

Sub CreditCardPayments()
'DESCRIPTION: List all Credit Card Payments
    Dim xList                                                    As ListItem
    Dim rsCMIS_OFF_HD                                            As ADODB.Recordset
    If COMPANY_CODE = "HGC" Then
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD Where TOF = '3' and Paidby <> 'Y' and OR_DATE >='2/1/2010'  order by OR_Date")
    Else
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD Where TOF = '3' and Paidby <> 'Y'  order by OR_Date")
    End If
    If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
        lvPayments.ListItems.Clear
        lblTotal = "0.00"
        Do While Not rsCMIS_OFF_HD.EOF
            Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
            xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
            xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!CUSNAME)
            xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
            xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
            xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
            tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
            lblTotal = Format(tmpTotal, "#,###,##0.00")
            rsCMIS_OFF_HD.MoveNext
        Loop
    Else
        MessagePop RecNotFound, "No record to view", "No Record"
    End If
    Set rsCMIS_OFF_HD = Nothing
End Sub

Function CheckIfBank(xCusCde As String) As Boolean
    Dim rsCheckCode                                              As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "Select Cuscde from All_Customer_Table where CusCde = " & N2Str2Null(xCusCde) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                                      As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "Select CusCde from CMIS_CardBank where CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Function CheckDeposited(xORNUM As String) As Boolean
    Dim rsCheckDeposited                                         As ADODB.Recordset
    Set rsCheckDeposited = New ADODB.Recordset
    rsCheckDeposited.Open "Select * from CMIS_BANKDEPO where OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckDeposited.EOF And Not rsCheckDeposited.BOF Then
        CheckDeposited = True
    End If
End Function

Function CheckORCutOff(xORNUM As String) As Boolean
    On Error Resume Next
    Dim rsCheckORCutOff                                          As ADODB.Recordset
    Set rsCheckORCutOff = New ADODB.Recordset
    rsCheckORCutOff.Open "Select * from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "' and CutDate IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsCheckORCutOff.EOF And Not rsCheckORCutOff.BOF Then
        CheckORCutOff = True
    End If
End Function

Function CheckCutOffDate(xORNUM As String) As String
    On Error Resume Next
    Dim rsCheckORCutOff                                          As ADODB.Recordset
    Set rsCheckORCutOff = New ADODB.Recordset
    rsCheckORCutOff.Open "Select * from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "' and CutDate IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsCheckORCutOff.EOF And Not rsCheckORCutOff.BOF Then
        CheckCutOffDate = CDate(rsCheckORCutOff!CUTDATE)
    End If
End Function

Function CheckPostedOR(xORNUM As String) As Boolean
    Dim rsCheckPosted                                            As ADODB.Recordset
    Set rsCheckPosted = New ADODB.Recordset
    If VAT_OR = 1 Then
        rsCheckPosted.Open "Select * from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "' and PAIDNA=1 and VAT=1", gconDMIS, adOpenKeyset
    Else
        rsCheckPosted.Open "Select * from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "' and PAIDNA=1 and VAT=0", gconDMIS, adOpenKeyset
    End If
    If Not rsCheckPosted.EOF And Not rsCheckPosted.BOF Then
        CheckPostedOR = True
    End If
End Function

Function CashAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                          As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "Select CashAmount,ChkAmount,CardAmount from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CashAmount = NumericVal(rsCheckPayments!CashAmount)
    End If
End Function

Function CheckAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                          As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "Select CashAmount,ChkAmount,CardAmount from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CheckAmount = NumericVal(rsCheckPayments!CHKAMOUNT)
    End If
End Function

Function CardAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                          As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "Select CashAmount,ChkAmount,CardAmount from CMIS_OFF_HD where OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CardAmount = NumericVal(rsCheckPayments!CardAmount)
    End If
End Function

Function CheckIfImportedinAMIS(xOR_Num As String) As Boolean
    Dim rsPostedCRJ                                              As ADODB.Recordset
    Set rsPostedCRJ = New ADODB.Recordset
    rsPostedCRJ.Open "Select * from AMIS_Journal_HD where JTYPE='CRJ' and Status <> 'C' and InvoiceNo ='" & xOR_Num & "'", gconDMIS, adOpenKeyset
    If Not rsPostedCRJ.EOF And Not rsPostedCRJ.BOF Then
        CheckIfImportedinAMIS = True
    End If
End Function

Sub UnPost_CashPos()
    If MODE_OF_PAYMENT = "CASH" Then
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                          " CASH = CASH - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    End If
    If MODE_OF_PAYMENT = "CHECK" Then
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                          " [CHECK] = [CHECK] - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    End If
    If MODE_OF_PAYMENT = "CARD" Then
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                          " CARD = CARD - " & RECEIPTS_AMOUNT & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    End If

    If CheckDeposited(txtOR_NUM) = True Then
        If MODE_OF_PAYMENT = "CASH" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                              " CASH = CASH + " & RECEIPTS_AMOUNT & "," & _
                              " CASHDEPO = CASHDEPO - " & RECEIPTS_AMOUNT & _
                              " where CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
        ElseIf MODE_OF_PAYMENT = "CHECK" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                              " [CHECK] = [CHECK] + " & RECEIPTS_AMOUNT & "," & _
                              " CHECKDEPO = CHECKDEPO - " & RECEIPTS_AMOUNT & _
                              " where CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
        ElseIf MODE_OF_PAYMENT = "CARD" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos set" & _
                              " CARD = CARD + " & RECEIPTS_AMOUNT & "," & _
                              " CARDDEPO = CARDDEPO - " & RECEIPTS_AMOUNT & _
                              " where CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
        End If
    End If
End Sub

Function CheckIfCancel(xOR_Num As String) As Boolean
    Dim rsCheckCancel                                            As ADODB.Recordset
    Set rsCheckCancel = New ADODB.Recordset
    rsCheckCancel.Open "Select * from CMIS_OFF_HD where Cancel=1 and OR_NUM = '" & xOR_Num & "'", gconDMIS, adOpenKeyset
    If Not rsCheckCancel.EOF And Not rsCheckCancel.BOF Then
        CheckIfCancel = True
    End If
End Function

Function CheckAppliedDeposit(xOR_Num As String) As Boolean
    Dim rsDeposit                                                As ADODB.Recordset
    Set rsDeposit = New ADODB.Recordset
    rsDeposit.Open "SELECT * FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT OR_NUM FROM CMIS_DEPOSITS WHERE OR_NUM = '" & xOR_Num & "' AND APPLIED = 'Y')", gconDMIS, adOpenKeyset
    If Not rsDeposit.EOF And Not rsDeposit.BOF Then
        CheckAppliedDeposit = True
    End If
End Function
