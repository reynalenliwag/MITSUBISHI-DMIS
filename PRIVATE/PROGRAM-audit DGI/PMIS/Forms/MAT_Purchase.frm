VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMAT_Purchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Purchase Order Entry"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_Purchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11775
   Begin VB.PictureBox picConfirmation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2220
      ScaleHeight     =   255
      ScaleWidth      =   9435
      TabIndex        =   85
      Top             =   7260
      Width           =   9465
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "F9 - Update PO Upon Confirmation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1290
         TabIndex        =   87
         Top             =   30
         Width           =   3435
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "F11 - View Confirmation Window"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   5340
         TabIndex        =   86
         Top             =   30
         Width           =   3105
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2220
      ScaleHeight     =   255
      ScaleWidth      =   9435
      TabIndex        =   60
      Top             =   6000
      Width           =   9465
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
         TabIndex        =   92
         Top             =   30
         Width           =   2445
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
         Left            =   5070
         TabIndex        =   64
         Top             =   30
         Width           =   1905
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
         TabIndex        =   63
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
         TabIndex        =   62
         Top             =   30
         Width           =   1455
      End
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
         TabIndex        =   61
         Top             =   30
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport rptPurchaseOrder 
      Left            =   2400
      Top             =   4470
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Height          =   7635
      Left            =   30
      TabIndex        =   54
      Top             =   0
      Width           =   2115
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
         TabIndex        =   57
         Text            =   "TEXT"
         Top             =   960
         Width           =   1995
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
         TabIndex        =   56
         Top             =   630
         Width           =   1875
      End
      Begin VB.OptionButton optPONo 
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
         TabIndex        =   55
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstPO_HD 
         Height          =   5865
         Left            =   60
         TabIndex        =   58
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10345
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
         MouseIcon       =   "MAT_Purchase.frx":08CA
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
      Begin VB.Label LblRefRRno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   60
         TabIndex        =   96
         Top             =   7260
         Width           =   1995
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
         TabIndex        =   59
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2220
      ScaleHeight     =   870
      ScaleWidth      =   9465
      TabIndex        =   65
      Top             =   6330
      Width           =   9465
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
         Left            =   8640
         MouseIcon       =   "MAT_Purchase.frx":0A2C
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":0B7E
         Style           =   1  'Graphical
         TabIndex        =   68
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
         Left            =   7860
         MouseIcon       =   "MAT_Purchase.frx":0EE4
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelPO 
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Arial"
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
         MouseIcon       =   "MAT_Purchase.frx":139C
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":14EE
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
         BeginProperty Font 
            Name            =   "Arial"
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
         MouseIcon       =   "MAT_Purchase.frx":1828
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
         BeginProperty Font 
            Name            =   "Arial"
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
         MouseIcon       =   "MAT_Purchase.frx":1CBF
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":1E11
         Style           =   1  'Graphical
         TabIndex        =   77
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
         MouseIcon       =   "MAT_Purchase.frx":2136
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":2288
         Style           =   1  'Graphical
         TabIndex        =   70
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
         MouseIcon       =   "MAT_Purchase.frx":25E4
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":2736
         Style           =   1  'Graphical
         TabIndex        =   71
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
         MouseIcon       =   "MAT_Purchase.frx":2A49
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":2B9B
         Style           =   1  'Graphical
         TabIndex        =   67
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
         MouseIcon       =   "MAT_Purchase.frx":2EEB
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":303D
         Style           =   1  'Graphical
         TabIndex        =   66
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
         MouseIcon       =   "MAT_Purchase.frx":339B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":34ED
         Style           =   1  'Graphical
         TabIndex        =   72
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
         MouseIcon       =   "MAT_Purchase.frx":37E7
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":3939
         Style           =   1  'Graphical
         TabIndex        =   73
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
         MouseIcon       =   "MAT_Purchase.frx":3C91
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":3DE3
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10170
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   78
      Top             =   6330
      Width           =   1470
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
         MouseIcon       =   "MAT_Purchase.frx":4142
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":4294
         Style           =   1  'Graphical
         TabIndex        =   79
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
         Left            =   -60
         MouseIcon       =   "MAT_Purchase.frx":45D2
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":4724
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3045
      Left            =   2220
      TabIndex        =   89
      Top             =   2910
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2805
         Left            =   60
         TabIndex        =   90
         Top             =   150
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   4948
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
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   2220
      TabIndex        =   24
      Top             =   0
      Width           =   9495
      Begin VB.ComboBox cboContactCode 
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
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   2190
         Width           =   3015
      End
      Begin VB.TextBox txtDON 
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
         Left            =   5670
         TabIndex        =   7
         Text            =   "16A070101"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdDON 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   82
         Top             =   180
         Width           =   255
      End
      Begin VB.ComboBox cboModelCode 
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
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   2580
         Width           =   3015
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   7440
         Top             =   120
      End
      Begin VB.TextBox txtDS1 
         Alignment       =   1  'Right Justify
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
         Left            =   5490
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox txtPODate 
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
         Height          =   360
         Left            =   3180
         MaxLength       =   10
         TabIndex        =   47
         ToolTipText     =   "Type the date of the purchase order in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   180
         Width           =   1365
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
         Left            =   4800
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "MAT_Purchase.frx":4A74
         ToolTipText     =   "Type your message or your remarks."
         Top             =   1890
         Width           =   4575
      End
      Begin VB.ComboBox cboPP_No 
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
         Left            =   3420
         TabIndex        =   1
         Text            =   "cboRecvd_Desc"
         ToolTipText     =   "Select PP Number from the list."
         Top             =   -600
         Width           =   1305
      End
      Begin VB.TextBox txtShipTo 
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
         Left            =   150
         MaxLength       =   40
         TabIndex        =   9
         ToolTipText     =   "Type the name of addressee (e.g. CALEB MOTOR CORPORATION)"
         Top             =   3390
         Width           =   4545
      End
      Begin VB.TextBox txtDealerCode 
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
         Left            =   1230
         MaxLength       =   5
         TabIndex        =   8
         ToolTipText     =   "Type the place where the order should be delivered (e.g. PCMC0)"
         Top             =   3030
         Width           =   1005
      End
      Begin VB.TextBox txtSupCode 
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
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Type the supplier code (e.g. 00001)"
         Top             =   660
         Width           =   1005
      End
      Begin VB.ComboBox cboSupName 
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
         TabIndex        =   3
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1050
         Width           =   4605
      End
      Begin VB.TextBox txtPONo 
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
         Height          =   360
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   180
         Width           =   1005
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   60
         ScaleHeight     =   795
         ScaleWidth      =   4635
         TabIndex        =   29
         Top             =   1410
         Width           =   4635
         Begin VB.TextBox txtSup_Addrs 
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
            Height          =   705
            Left            =   30
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   30
            Width           =   4605
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1185
         Left            =   6480
         ScaleHeight     =   1185
         ScaleWidth      =   2925
         TabIndex        =   30
         Top             =   660
         Width           =   2925
         Begin VB.TextBox txtNetPOAmt 
            Alignment       =   1  'Right Justify
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
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   51
            Top             =   780
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
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   50
            Top             =   390
            Width           =   1395
         End
         Begin VB.TextBox txtPO_Amount 
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
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   49
            Top             =   0
            Width           =   1395
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
            Left            =   30
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "Type the type of the additional amount (e.g. VAT)"
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TOT Amount"
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
            Left            =   120
            TabIndex        =   46
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Left            =   120
            TabIndex        =   45
            Top             =   30
            Width           =   1245
         End
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Person"
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
         Left            =   120
         TabIndex        =   84
         Top             =   2250
         Width           =   1965
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
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
         Left            =   4770
         TabIndex        =   83
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Model"
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
         Left            =   120
         TabIndex        =   81
         Top             =   2640
         Width           =   1965
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   4860
         X2              =   9390
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label17 
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
         Left            =   6120
         TabIndex        =   48
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   4800
         TabIndex        =   42
         Top             =   1620
         Width           =   1965
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PP Number"
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
         Height          =   315
         Left            =   2370
         TabIndex        =   44
         Top             =   -570
         Width           =   1845
      End
      Begin VB.Label labPosted 
         Alignment       =   1  'Right Justify
         Caption         =   "POSTED"
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
         Height          =   315
         Left            =   7320
         TabIndex        =   43
         Top             =   180
         Width           =   2115
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5040
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4740
         X2              =   4740
         Y1              =   90
         Y2              =   3000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ship To"
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
         TabIndex        =   41
         Top             =   3060
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PO Number"
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
         Left            =   120
         TabIndex        =   28
         Top             =   210
         Width           =   1845
      End
      Begin VB.Label Label2 
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
         Height          =   315
         Index           =   1
         Left            =   2310
         TabIndex        =   27
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1965
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
         Left            =   3960
         TabIndex        =   25
         Top             =   1050
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdAddTran 
      Caption         =   "Command1"
      Height          =   4245
      Left            =   4500
      TabIndex        =   93
      Top             =   840
      Width           =   4755
   End
   Begin VB.Frame fraAddTran 
      Caption         =   "Add/Edit Materials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4590
      TabIndex        =   31
      Top             =   900
      Width           =   4575
      Begin VB.TextBox txtwvatm 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   2670
         TabIndex        =   94
         Top             =   2070
         Width           =   1785
      End
      Begin VB.CheckBox chkUseHARIDNP 
         Caption         =   "Use HARI DNP"
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
         Left            =   2850
         TabIndex        =   91
         Top             =   1650
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtVIN 
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
         Height          =   315
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2700
         Width           =   1815
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
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2700
         Width           =   1125
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
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2340
         Width           =   1125
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
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1980
         Width           =   1125
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
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1620
         Width           =   765
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
         MaxLength       =   4
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.ComboBox cboTranDescription 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   1230
         Width           =   4335
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
         TabIndex        =   14
         Text            =   "Combo1"
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
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   600
         Width           =   585
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
         Height          =   855
         Left            =   3480
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_Purchase.frx":4A8E
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":4BE0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Delete Entry"
         Top             =   3120
         Width           =   915
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
         Height          =   855
         Left            =   2580
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_Purchase.frx":4F0B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":505D
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel Entry"
         Top             =   3120
         Width           =   915
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
         Height          =   855
         Left            =   1680
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MAT_Purchase.frx":539B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_Purchase.frx":54ED
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Save Entry"
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "* Unit Cost w/ VAT *"
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
         Height          =   255
         Index           =   1
         Left            =   2610
         TabIndex        =   95
         Top             =   1860
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "VIN"
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
         Height          =   285
         Left            =   2700
         TabIndex        =   88
         Top             =   2400
         Width           =   1275
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   150
         TabIndex        =   32
         Top             =   2730
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
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
         Index           =   0
         Left            =   420
         TabIndex        =   40
         Top             =   2370
         Width           =   1005
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
         Left            =   1710
         TabIndex        =   38
         Top             =   3300
         Width           =   285
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amt."
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
         Left            =   240
         TabIndex        =   37
         Top             =   2010
         Width           =   1185
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         Left            =   630
         TabIndex        =   36
         Top             =   1650
         Width           =   795
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Code"
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
         Left            =   180
         TabIndex        =   35
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
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
         Left            =   570
         TabIndex        =   34
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   180
         TabIndex        =   33
         Top             =   990
         Width           =   1245
      End
   End
   Begin VB.Label Label3 
      Caption         =   "- required fields"
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
      Left            =   10260
      TabIndex        =   53
      Top             =   7290
      Width           =   1425
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
      Height          =   225
      Index           =   2
      Left            =   10080
      TabIndex        =   52
      Top             =   7320
      Width           =   135
   End
End
Attribute VB_Name = "frmPMISMAT_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPO_HD, rsPMIS_PP_HD, rsPMIS_Tdaytran             As ADODB.Recordset
Attribute rsPMIS_PP_HD.VB_VarUserMemId = 1073938432
Attribute rsPMIS_Tdaytran.VB_VarUserMemId = 1073938432
Dim rsCSMS_MATMAS, rsSupplier                          As ADODB.Recordset
Attribute rsCSMS_MATMAS.VB_VarUserMemId = 1073938435
Attribute rsSupplier.VB_VarUserMemId = 1073938435
Dim rsALL_Profile, rsPMIS_Counter                      As ADODB.Recordset
Attribute rsALL_Profile.VB_VarUserMemId = 1073938437
Attribute rsPMIS_Counter.VB_VarUserMemId = 1073938437
Dim Pcnt                                               As Integer
Attribute Pcnt.VB_VarUserMemId = 1073938439
Dim PO_TOTQTY                                          As Integer
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938440
Dim PO_TOTUCOST, PO_TOTINVAMT                          As Double
Attribute PO_TOTUCOST.VB_VarUserMemId = 1073938441
Attribute PO_TOTINVAMT.VB_VarUserMemId = 1073938441
Dim PO_TOTVAT                                          As Double
Attribute PO_TOTVAT.VB_VarUserMemId = 1073938443
Dim PO_T_ONORDER                                       As Long
Attribute PO_T_ONORDER.VB_VarUserMemId = 1073938444
Dim PrevPONO                                           As String
Attribute PrevPONO.VB_VarUserMemId = 1073938445
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasSRP              As Double
Attribute PrevPmasMAC.VB_VarUserMemId = 1073938446
Attribute PrevPmasDNP.VB_VarUserMemId = 1073938446
Attribute PrevPmasSRP.VB_VarUserMemId = 1073938446
Dim NewPmasMAC, NewPmasDNP, NewPmasSRP                 As Double
Attribute NewPmasMAC.VB_VarUserMemId = 1073938450
Attribute NewPmasDNP.VB_VarUserMemId = 1073938450
Attribute NewPmasSRP.VB_VarUserMemId = 1073938450
Dim NewPmasOnHand, PrevTranQty                         As Integer
Attribute NewPmasOnHand.VB_VarUserMemId = 1073938453
Attribute PrevTranQty.VB_VarUserMemId = 1073938453
Dim ISNONVAT                                           As Boolean
Attribute ISNONVAT.VB_VarUserMemId = 1073938456
Dim DON_TYPE                                           As String
Attribute DON_TYPE.VB_VarUserMemId = 1073938434

Sub fill_LblRefRRno()
    Dim rsTMP As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_RR_HD" & vbCrLf
    SQL = SQL & "Union All" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_REC_HIST" & vbCrLf
    SQL = SQL & ") T WHERE PONO = '" & Null2String(txtPONo) & "' AND TYPE = 'M' AND STATUS = 'P'" & vbCrLf
 
    Set rsTMP = gconDMIS.Execute(SQL)
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        LblRefRRno.Caption = "REF RRNO" & "-" & Null2String(rsTMP!RRNO)
    Else
         LblRefRRno.Caption = "NO REF RRNO"
    End If
    
    Set rsTMP = Nothing
End Sub

Function SetOrderType(XXX As String)
    Dim rsOrderType                                    As ADODB.Recordset
    Set rsOrderType = New ADODB.Recordset
    Set rsOrderType = gconDMIS.Execute("Select * from PMIS_OrderType Where CODE = '" & XXX & "'")
    If Not rsOrderType.EOF And Not rsOrderType.BOF Then
        SetOrderType = Null2String(rsOrderType!Description)
    End If
    Set rsOrderType = Nothing
End Function

Function SetPartDesc(ppp As String)
    Set rsCSMS_MATMAS = New ADODB.Recordset
    rsCSMS_MATMAS.Open "Select STOCKNO,STOCKDESC from CSMS_MATMAS where STOCKNO = '" & ppp & "'", gconDMIS
    If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
        SetPartDesc = UCase(Null2String(rsCSMS_MATMAS!STOCKDESC))
    End If
End Function

Function SetPartDesc2(pid As Variant)
    If pid <> "" Then
        Set rsCSMS_MATMAS = New ADODB.Recordset
        rsCSMS_MATMAS.Open "Select id,STOCKDESC,dnp from CSMS_MATMAS where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
            SetPartDesc2 = Null2String(rsCSMS_MATMAS!STOCKDESC)
            
            txtUnitCost.Text = Round(N2Str2Zero(rsCSMS_MATMAS!dnp) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
            
            If NumericVal(txtUnitCost.Text) <> 0 Then
                If ISNONVAT = True Then
                    txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
                Else
                    txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
                End If
                txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
            Else
                txtTranINVAmt.Text = 0
                txtTranTotalAmt.Text = 0
            End If
            
            
            
            
            
            
            
            
            
            
            
            
        End If
    End If
End Function

Function SetPartNo(pid As Variant)
    If pid <> "" Then
        Set rsCSMS_MATMAS = New ADODB.Recordset
        rsCSMS_MATMAS.Open "Select id,STOCKNO from CSMS_MATMAS where id = " & pid, gconDMIS
        If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
            SetPartNo = Null2String(rsCSMS_MATMAS!STOCKNO)
        End If
    End If
End Function

Function SetPartIDPartNo(DDD As String)
    Set rsCSMS_MATMAS = New ADODB.Recordset
    rsCSMS_MATMAS.Open "Select id,STOCKNO from CSMS_MATMAS where STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS
    If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
        SetPartIDPartNo = N2Str2IntZero(rsCSMS_MATMAS!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsCSMS_MATMAS = New ADODB.Recordset
    rsCSMS_MATMAS.Open "Select id,STOCKDESC from CSMS_MATMAS where (ltrim(rtrim(STOCKDESC))) = '" & UCase(LTrim(RTrim(DDD))) & "'", gconDMIS
    If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
        SetPartIDDesc = N2Str2IntZero(rsCSMS_MATMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsCSMS_MATMAS = New ADODB.Recordset
        rsCSMS_MATMAS.Open "Select STOCKNO,mac from CSMS_MATMAS where STOCKNO = '" & ppp & "'", gconDMIS
        If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
            SetPartPrice = N2Str2Zero(rsCSMS_MATMAS!MAC)
        End If
    End If
End Function

Function SetSupdesc(ppp As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supcode,supname,sup_addrs,vat_percnt,NONVAT,CONTACT from PMIS_vw_SUPPLIER where supcode = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupdesc = Null2String(rsSupplier!supname)
        txtSup_Addrs.Text = Null2String(rsSupplier!sup_addrs)
        cboContactCode.Text = Null2String(rsSupplier!CONTACT)
        If Null2String(rsSupplier!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(rsSupplier!vat_percnt)
        End If
    Else
        cboContactCode.Text = ""
        txtSup_Addrs.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
    End If
End Function

Function SetSupCode(nnn As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supname,supcode,sup_addrs,vat_percnt,NONVAT,CONTACT from PMIS_vw_SUPPLIER where supname = '" & nnn & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupCode = Null2String(rsSupplier!SupCode)
        txtSup_Addrs.Text = Null2String(rsSupplier!sup_addrs)
        cboContactCode.Text = Null2String(rsSupplier!CONTACT)
        If Null2String(rsSupplier!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(rsSupplier!vat_percnt)
        End If
    Else
        cboContactCode.Text = ""
        txtSup_Addrs.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    PrevTranQty = 0
    Set rsPMIS_Tdaytran = New ADODB.Recordset
    rsPMIS_Tdaytran.Open "select * from PMIS_Tdaytran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
        labDetID.Caption = rsPMIS_Tdaytran!ID
        txtTranItemNo.Text = Format(Null2String(rsPMIS_Tdaytran!itemno), "0000")
        cboTranPartNo.Text = Null2String(rsPMIS_Tdaytran!STOCK_ORD)
        cboTranDescription.Text = SetPartDesc(Null2String(rsPMIS_Tdaytran!STOCK_SUP))
        txtTranQty.Text = N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY)
        PrevTranQty = N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY)
        txtTranINVAmt.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_Tdaytran!TRANINVAMT))
        txtUnitCost.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_Tdaytran!TRANQTY) * N2Str2Zero(rsPMIS_Tdaytran!TRANINVAMT))
        txtVIN.Text = Null2String(rsPMIS_Tdaytran!Vin)
    End If
End Function

Function SetModelCode(XXX As String)
    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select ModelCode from ALL_model where Model = " & N2Str2Null(XXX))
    If Not rsModel.EOF And Not rsModel.BOF Then
        SetModelCode = Null2String(rsModel!MODELCODE)
    End If
End Function

Function SetModelDesc(XXX As String)
    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select Model from ALL_model where ModelCode = " & N2Str2Null(XXX))
    If Not rsModel.EOF And Not rsModel.BOF Then
        SetModelDesc = Null2String(rsModel!Model)
    End If
End Function

Function SetContactCode(XXX As String)
    Dim rsContact                                      As ADODB.Recordset
    Set rsContact = New ADODB.Recordset
    Set rsContact = gconDMIS.Execute("Select ContactCode from ALL_Contact where ContactName = " & N2Str2Null(XXX))
    If Not rsContact.EOF And Not rsContact.BOF Then
        SetContactCode = Null2String(rsContact!contactcode)
    End If
    Set rsContact = Nothing
End Function

Function SetContactName(XXX As String)
    Dim rsContact                                      As ADODB.Recordset
    Set rsContact = New ADODB.Recordset
    Set rsContact = gconDMIS.Execute("Select ContactName from ALL_Contact where ContactCode = " & N2Str2Null(XXX))
    If Not rsContact.EOF And Not rsContact.BOF Then
        SetContactName = Null2String(rsContact!ContactName)
    End If
    Set rsContact = Nothing
End Function

Sub Send2FrontConfirm()
    Frame1.Enabled = False: Picture1.Enabled = False: fraDetails.Enabled = False: cmdAddTran.Enabled = False: fraAddTran.Enabled = False
End Sub

Sub Send2BackConfirm()
    Frame1.Enabled = False: Picture1.Enabled = True: fraDetails.Enabled = True: cmdAddTran.Enabled = True: fraAddTran.Enabled = True
End Sub

Sub SendToFrontConfirmPO()
    With frmPMISTrans_POConfirmation
        .txtPONo.Text = txtPONo.Text
        .txtPODate.Text = Format(txtPODate.Text, "DD-MMM-YY")
        DoEvents
        .txtDealerCode.Text = Left(txtDON.Text, 2)
        .txtDealerName.Text = cboSupName.Text
        .txtSOType.Text = SetOrderType(Mid(txtDON.Text, 3, 1))
        .txtSOYear.Text = Mid(txtDON.Text, 4, 2)
        .txtSOMonth.Text = The_month(Mid(txtDON.Text, 6, 2))
        .txtSONum.Text = txtDON.Text
        .FillDetails (txtPONo.Text)
        Me.KeyPreview = False
        .Show 1
        Me.KeyPreview = True
    End With
End Sub

Sub SendToBackConfirmPO()
    Unload frmPMISTrans_POConfirmation
End Sub

Sub FindDupPOno(DDD As String)
    RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select * from PMIS_PO_HD WHERE [TYPE] = 'M' order by pono asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtPONo.Text = ""
    Set rsPMIS_Counter = New ADODB.Recordset
    rsPMIS_Counter.Open "select modul,nextnumber from PMIS_Counter where [TYPE] = 'M' AND modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_Counter.EOF And Not rsPMIS_Counter.BOF Then
        txtPONo.Text = Format(N2Str2IntZero(rsPMIS_Counter!nextnumber), "000000")
    End If
    chkUseHARIDNP.Value = 0
    txtPartID.Text = ""
    cboPP_No.Text = ""
    txtPODate.Text = LOGDATE
    txtSupCode.Text = ""

    txtDON.Text = ""

    FillCboSupName
    txtSup_Addrs.Text = ""
    Filltxtshipto
    txtPO_Amount.Text = ""
    txtDS1.Text = ""
    txtDS_Desc1.Text = ""
    txtDS_Amt1.Text = ""
    txtNetPOAmt.Text = ""
    labPosted.Visible = False
    labPosted.Caption = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    cleargrid grdDetails
    InitGrid
    InitCbo
    InitParts
End Sub

Sub StoreMemVars()
    DON_TYPE = ""
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        labID.Caption = RSPO_HD!ID
        txtPONo.Text = Null2String(RSPO_HD!PONO)
        cboPP_No.Text = Null2String(RSPO_HD!ppno)
        txtPODate.Text = Null2String(RSPO_HD!PODATE)
        txtDON.Text = Null2String(RSPO_HD!DON)
        DON_TYPE = Right(Left(Null2String(RSPO_HD!DON), 3), 1)
        txtSupCode.Text = Null2String(RSPO_HD!SupCode)
        cboSupName.Text = SetSupdesc(Null2String(RSPO_HD!SupCode))
        txtSup_Addrs.Text = Null2String(RSPO_HD!sup_addrs)
        cboContactCode.Text = Null2String(RSPO_HD!contactcode)
        cboModelCode.Text = Null2String(RSPO_HD!MODELCODE)
        txtDealerCode.Text = Null2String(RSPO_HD!dealercode)
        Filltxtshipto2 (Null2String(RSPO_HD!dealercode))
        txtPO_Amount.Text = ToDoubleNumber(N2Str2Zero(RSPO_HD!po_amount))
        txtDS1.Text = N2Str2IntZero(RSPO_HD!ds1)
        txtDS_Desc1.Text = Null2String(RSPO_HD!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSPO_HD!DS_AMT1))
        txtNetPOAmt.Text = ToDoubleNumber(N2Str2Zero(RSPO_HD!netpoamt))
        txtRemarks.Text = Null2String(RSPO_HD!REMARKS)
        If Null2String(RSPO_HD!Status) = "P" Then
            labPosted.Visible = True
            labPosted.Caption = "POSTED [" & Null2String(RSPO_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            picConfirmation.Visible = False
            cmdCancelPO.Enabled = False
            cmdUnPost.Enabled = True
            cmdAddTran.Enabled = False
            cmdPrint.Enabled = True
        ElseIf Null2String(RSPO_HD!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "CANCELLED [" & Null2String(RSPO_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdCancelPO.Enabled = False
            picConfirmation.Visible = False
            cmdAddTran.Enabled = False
            cmdPrint.Enabled = False
        Else
            labPosted.Visible = False
            labPosted.Caption = ""
            cmdEdit.Enabled = True
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            picConfirmation.Visible = False
            cmdCancelPO.Enabled = True
            cmdAddTran.Enabled = True
            cmdPrint.Enabled = False
        End If
        cleargrid grdDetails
        FillDetails
        fill_LblRefRRno
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub InitGrid()
    With grdDetails
        .ColWidth(0) = 1
        .ColWidth(1) = 700
        .ColWidth(2) = 1200
        .ColAlignment(2) = 2
        .ColWidth(3) = 2500
        .ColWidth(4) = 500
        .ColWidth(5) = 800
        .ColWidth(6) = 1100
        .ColWidth(7) = 2500

        .Row = 0
        .Col = 1: .Text = "Item"
        .Col = 2: .Text = "Material Code"
        .Col = 3: .Text = "Description"
        .Col = 4: .Text = "Qty"
        .Col = 5: .Text = "Amount"
        .Col = 6: .Text = "Total Order"
        .Col = 7: .Text = "VIN"
    End With
End Sub

Sub FillDetails()
    Pcnt = 0: PO_TOTUCOST = 0: PO_TOTINVAMT = 0: PO_TOTVAT = 0: PO_T_ONORDER = 0: PO_TOTQTY = 0
    Set rsPMIS_Tdaytran = New ADODB.Recordset
    rsPMIS_Tdaytran.Open "select id,tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,VIN from PMIS_Tdaytran where [TYPE] = 'M' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
        Screen.MousePointer = 11
        rsPMIS_Tdaytran.MoveFirst
        Do While Not rsPMIS_Tdaytran.EOF
            Pcnt = Pcnt + 1
            grdDetails.AddItem rsPMIS_Tdaytran!ID & Chr(9) & Format(Null2String(rsPMIS_Tdaytran!itemno), "0000") & Chr(9) & _
                               Null2String(rsPMIS_Tdaytran!STOCK_ORD) & Chr(9) & _
                               SetPartDesc(Null2String(rsPMIS_Tdaytran!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_Tdaytran!TRANQTY) * N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Null2String(rsPMIS_Tdaytran!Vin)
            PO_TOTQTY = PO_TOTQTY + N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY)
            PO_TOTUCOST = PO_TOTUCOST + (N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY) * N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST))
            PO_TOTINVAMT = PO_TOTINVAMT + (N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY) * N2Str2Zero(rsPMIS_Tdaytran!TRANINVAMT))
            rsPMIS_Tdaytran.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        If ISNONVAT = True Then PO_TOTVAT = 0 Else PO_TOTVAT = PO_TOTINVAMT - PO_TOTUCOST '(PO_TOTINVAMT / ConvertToBIRDecimalFormat(VAT_RATE))
        PO_TOTUCOST = NumericVal(PO_TOTINVAMT - PO_TOTVAT)
        If NumericVal(PO_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtPO_Amount.Text = ToDoubleNumber(PO_TOTUCOST)
'            txtDS_Amt1.Text = ToDoubleNumber(PO_TOTVAT)
'            txtNetPOAmt.Text = ToDoubleNumber(PO_TOTINVAMT)
            txtNetPOAmt.Text = ToDoubleNumber(PO_TOTUCOST * 1.12)
            txtDS_Amt1.Text = ToDoubleNumber(NumericVal(txtNetPOAmt.Text) - NumericVal(txtPO_Amount.Text))
        Else
            txtDS1.Text = ""
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = ZERO
            txtPO_Amount.Text = ToDoubleNumber(PO_TOTUCOST)
            txtNetPOAmt.Text = ToDoubleNumber(PO_TOTUCOST)
        End If
        Screen.MousePointer = 0
    Else
        cleargrid grdDetails
    End If
End Sub

Sub FillCboSupName()
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select supname from PMIS_vw_SUPPLIER order by supname asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        cboSupName.Clear
        Do While Not rsSupplier.EOF
            cboSupName.AddItem Null2String(rsSupplier!supname)
            rsSupplier.MoveNext
        Loop
    End If
End Sub

Sub Filltxtshipto()
    Set rsALL_Profile = New ADODB.Recordset
    rsALL_Profile.Open "select * from ALL_Profile WHERE MODULENAME = 'PMIS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsALL_Profile.EOF And Not rsALL_Profile.BOF Then
        txtDealerCode.Text = Null2String(rsALL_Profile!COMPANYCODE)
        txtShipTo.Text = Null2String(rsALL_Profile!CompanyName)
    End If
End Sub

Sub Filltxtshipto2(param As String)
    Set rsALL_Profile = New ADODB.Recordset
    rsALL_Profile.Open "select * from ALL_Profile where MODULENAME = 'PMIS' AND companycode = '" & param & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsALL_Profile.EOF And Not rsALL_Profile.BOF Then
        txtDealerCode.Text = Null2String(rsALL_Profile!COMPANYCODE)
        txtShipTo.Text = Null2String(rsALL_Profile!CompanyName)
    End If
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(Pcnt + 1, "0000")
    cboTranPartNo.Text = ""
    cboTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranINVAmt.Text = 0#
    txtUnitCost.Text = 0#
    txtTranTotalAmt.Text = 0#
    txtVIN.Text = ""
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1: cmdAddTran.Enabled = False
    fraAddTran.ZOrder 1
    fraAddTran.Enabled = False
    Send2BackConfirm
End Sub

Sub BringToFront()
    cmdAddTran.ZOrder 0
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    Picture1.Enabled = False
End Sub

Sub InitCbo()
    Set rsCSMS_MATMAS = New ADODB.Recordset
    rsCSMS_MATMAS.Open "select STOCKNO,STOCKDESC from CSMS_MATMAS ORDER BY STOCKDESC ASC", gconDMIS
    If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
        rsCSMS_MATMAS.MoveFirst
        cboTranPartNo.Clear
        cboTranDescription.Clear
        Do While Not rsCSMS_MATMAS.EOF
            cboTranPartNo.AddItem Null2String(rsCSMS_MATMAS!STOCKNO)
            cboTranDescription.AddItem Null2String(rsCSMS_MATMAS!STOCKDESC)
            rsCSMS_MATMAS.MoveNext
        Loop
    End If
    FillCboContact
    FillCboModel
End Sub

Sub FillGrid()
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Enabled = False
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    Set RSPO_HD = New ADODB.Recordset
    Set RSPO_HD = gconDMIS.Execute("select pono from PMIS_PO_HD WHERE [TYPE] = 'M' order by pono asc")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        lstPO_HD.Enabled = True: Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    lstPO_HD.Enabled = False
    Set RSPO_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSPO_HD = gconDMIS.Execute("select pono, pono from PMIS_PO_HD where [TYPE] = 'M' AND pono like'" & XXX & "%'")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        lstPO_HD.Enabled = True: Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    lstPO_HD.Enabled = False
    Set RSPO_HD = New ADODB.Recordset
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HD WHERE [TYPE] = 'M' order by pono asc")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        lstPO_HD.Enabled = True: Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Enabled = False
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    Set RSPO_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HD where [TYPE] = 'M' AND supname like '" & XXX & "%' order by pono asc")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh: lstPO_HD.Enabled = True
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillCboModel()
    Dim rsModel                                        As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    Set rsModel = gconDMIS.Execute("Select DISTINCT Description from ALL_ModelCode order by Description asc")
    If Not rsModel.EOF And Not rsModel.BOF Then
        rsModel.MoveFirst: cboModelCode.Clear
        Do While Not rsModel.EOF
            cboModelCode.AddItem Null2String(rsModel!Description)
            rsModel.MoveNext
        Loop
    End If
    Set rsModel = Nothing
End Sub

Sub FillCboContact()
    Dim rsContact                                      As ADODB.Recordset
    Set rsContact = New ADODB.Recordset
    Set rsContact = gconDMIS.Execute("Select ContactName from ALL_Contact order by ContactName asc")
    If Not rsContact.EOF And Not rsContact.BOF Then
        rsContact.MoveFirst: cboContactCode.Clear
        Do While Not rsContact.EOF
            cboContactCode.AddItem Null2String(rsContact!ContactName)
            rsContact.MoveNext
        Loop
    End If
    Set rsContact = Nothing
End Sub

Private Sub cboSupName_Click()
    txtSupCode.Text = SetSupCode(cboSupName.Text)
End Sub

Private Sub cboSupName_GotFocus()
    VBComBoBoxDroppedDown cboSupName
End Sub

Private Sub cboSupName_LostFocus()
    txtSupCode.Text = SetSupCode(cboSupName.Text)
End Sub

Private Sub cboTranDescription_Click()
    If cboTranDescription.Text <> "" Then
        txtPartID.Text = SetPartIDDesc(cboTranDescription.Text)
        cboTranPartNo.Text = SetPartNo(txtPartID.Text)
        cboTranDescription.Text = SetPartDesc2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranPartNo_Change()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDPartNo(cboTranPartNo.Text)
        cboTranDescription.Text = SetPartDesc2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranPartNo_Click()
    If cboTranPartNo.Text <> "" Then
        txtPartID.Text = SetPartIDPartNo(cboTranPartNo.Text)
        cboTranDescription.Text = SetPartDesc2(txtPartID.Text)
    End If
End Sub

Private Sub cboTranPartNo_LostFocus()
    cboTranPartNo.Text = UCase(cboTranPartNo.Text)
End Sub

Private Sub cboTranPartNo_GotFocus()
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub chkUseHARIDNP_Click()
    Dim rsDNPP                                         As ADODB.Recordset
    If chkUseHARIDNP.Value = 1 Then
        Set rsDNPP = New ADODB.Recordset
        Set rsDNPP = gconDMIS.Execute("Select * from PMIS_Dnpp Where PARTNUMBER = '" & cboTranPartNo.Text & "'")
        If Not rsDNPP.EOF And Not rsDNPP.BOF Then
            If DON_TYPE = "V" Then txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
            If DON_TYPE = "S" Then txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
            If DON_TYPE = "R" Then txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
            If DON_TYPE = "A" Then txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP)
            If DON_TYPE = "E" Then txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP3)
            If ISNONVAT = True Then
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtUnitCost.Text))
        End If
    Else
        txtTranINVAmt.Text = "0.00"
        txtUnitCost.Text = "0.00"
        txtTranTotalAmt.Text = "0.00"
    End If
End Sub

Private Sub cmdAddTran_Click()
    Frame2.Enabled = False
    If Picture1.Visible = True And cmdAddTran.Enabled = True Then
        SendToBack
        cmdAddTran.ZOrder 0: cmdAddTran.Enabled = True
        fraAddTran.ZOrder 0
        fraAddTran.Enabled = True
        cmdTranDelete.Enabled = False
        fraAddTran.Enabled = True
        AddorEdit = "ADD"
        InitParts
        On Error Resume Next
        cboTranPartNo.SetFocus
    End If
End Sub

Private Sub cmdCancelPO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "MATERIALS PURCHASE ORDER") = False Then Exit Sub

    On Error GoTo ErrorCode:

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
       'updated by: IEBV 11172011
       'description: to rollback transaction if error occured
        gconDMIS.BeginTrans
        If Cancel = False Then
            str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
            str_MSG = str_MSG & "Description: "
            str_MSG = str_MSG & " " & error_msg
            str_MSG = str_MSG & " " & vbCrLf
            str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
            str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Cancellation of Transaction")
            MsgBox str_MSG, vbCritical, "Cancellation Error"
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
        gconDMIS.CommitTrans
        rsRefresh
        On Error Resume Next
        RSPO_HD.Find "id =" & labID.Caption
        StoreMemVars
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Function Cancel() As Boolean
On Error GoTo errordaa

    Dim rsPMIS_TdaytranDup, rsCSMS_MATMASDup           As ADODB.Recordset
    Dim PCurOnOrder, PCurTpoQty                        As Integer
    SQL_STATEMENT = "update PMIS_PO_HD set" & _
                  " status = 'C'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "C", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "", ""

    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        Do While Not rsPMIS_TdaytranDup.EOF
            Set rsCSMS_MATMASDup = New ADODB.Recordset
            rsCSMS_MATMASDup.Open "select STOCKNO,onorder,tpoqty,ordered,emergency_po from CSMS_MATMAS where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), gconDMIS
            If Not rsCSMS_MATMASDup.EOF And Not rsCSMS_MATMASDup.BOF Then
                PCurOnOrder = N2Str2IntZero(rsCSMS_MATMASDup!ONORDER) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                PCurTpoQty = N2Str2IntZero(rsCSMS_MATMASDup!tpoqty) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                If Null2String(rsPMIS_TdaytranDup!Status) = "P" Then
                    gconDMIS.Execute "update CSMS_MATMAS set" & _
                                   " purchases = " & N2Str2Zero(rsCSMS_MATMASDup!purchases) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                   " onorder = " & PCurOnOrder & "," & _
                                   " ORDERED = " & N2Str2IntZero(rsCSMS_MATMASDup!Ordered) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                   " tpoqty = " & PCurTpoQty & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    If Mid(txtDON.Text, 3, 1) = "E" Then
                        gconDMIS.Execute "update CSMS_MATMAS set" & _
                                       " EMERGENCY_PO = " & N2Str2IntZero(rsCSMS_MATMASDup!emergency_po) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & _
                                       " where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    End If
                End If
            End If
            rsPMIS_TdaytranDup.MoveNext
        Loop
    End If
    SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                  " status = '" & "C" & "'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where [TYPE] = 'M' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO'"
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "C", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", ""

    Set rsPMIS_TdaytranDup = Nothing
    Set rsCSMS_MATMASDup = Nothing

     Cancel = True
     Exit Function
errordaa:
    error_msg = error
    Cancel = False
End Function

Private Sub cmdDON_Click()
    With frmPMISMAT_DONFormation
        If AddorEdit = "EDIT" Then
            .txtedit = "EDIT"
            .lbl1 = Mid(txtDON, 1, 2)
            .lbl2 = Mid(txtDON, 3, 1)
            .lbl3 = Mid(txtDON, 4, 2)
            .lbl4 = Mid(txtDON, 6, 2)
            .lbl5 = Mid(txtDON, 8, 2)
            .dtTranDate = CDate(txtPODate.Text)
        Else
            .txtedit = ""
        End If
    End With
    frmPMISMAT_DONFormation.Show 1
    On Error Resume Next
    cboSupName.SetFocus
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "MATERIALS PURCHASE ORDER") = False Then Exit Sub

    On Error GoTo ErrorCode:

    'updating code: JAA - 06272008     'Do not allow posting of transaction without issuance of Parts
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD = "" Or FILD = "No Entry" Then
        MsgBox "Posting of Transaction cannot proceed. Pls. Add Material(s).", vbCritical, "Confirm Posting"
        Exit Sub
    End If
    '====================================================================================================


    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        Set rsPMIS_Tdaytran = New ADODB.Recordset
        rsPMIS_Tdaytran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
        If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
            rsPMIS_Tdaytran.MoveFirst
            Do While Not rsPMIS_Tdaytran.EOF
                If N2Str2Zero(rsPMIS_Tdaytran!TRANINVAMT) <= 0 Then
                    MessagePop InfoWait, "Action Void", "Warning: Transaction with Invoice Amount equal to Zero Encountered!"
                    Exit Sub
                End If
                rsPMIS_Tdaytran.MoveNext
            Loop
        End If
       'updated by: IEBV 11172011
       'description: to rollback transaction if error occured
        gconDMIS.BeginTrans
        If post = False Then
        
            str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
            str_MSG = str_MSG & "Description: "
            str_MSG = str_MSG & " " & error_msg
            str_MSG = str_MSG & " " & vbCrLf
            str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
            str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Posting of Transaction")
            MsgBox str_MSG, vbCritical, "Posting Error"
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
        gconDMIS.CommitTrans
        
        rsRefresh
        On Error Resume Next
        RSPO_HD.Find "id =" & labID.Caption
        StoreMemVars
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Function post() As Boolean
On Error GoTo errordaa

    Set rsPMIS_Tdaytran = New ADODB.Recordset
    rsPMIS_Tdaytran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
    If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
        rsPMIS_Tdaytran.MoveFirst
        Do While Not rsPMIS_Tdaytran.EOF
            Set rsCSMS_MATMAS = New ADODB.Recordset
            rsCSMS_MATMAS.Open "Select STOCKNO,onhand,tpoqty,onorder,ordered,emergency_po,purchases from CSMS_MATMAS where STOCKNO = " & N2Str2Null(rsPMIS_Tdaytran!STOCK_ORD), gconDMIS
            If Not rsCSMS_MATMAS.EOF And Not rsCSMS_MATMAS.BOF Then
                gconDMIS.Execute "update CSMS_MATMAS set " & _
                               " purchases = " & N2Str2Zero(rsCSMS_MATMAS!purchases) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                               " tpoqty = " & N2Str2Zero(rsCSMS_MATMAS!tpoqty) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                               " ONORDER = " & N2Str2Zero(rsCSMS_MATMAS!ONORDER) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                               " ORDERED = " & N2Str2Zero(rsCSMS_MATMAS!Ordered) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & _
                               " where STOCKNO = " & N2Str2Null(rsCSMS_MATMAS!STOCKNO)
                SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                              " status = 'P'" & "," & _
                              " usercode = " & N2Str2Null(LOGCODE) & "," & _
                              " lastupdate = '" & LOGDATE & "'" & _
                              " where id = " & rsPMIS_Tdaytran!ID
                gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "P", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", ""

                If Mid(txtDON.Text, 3, 1) = "E" Then
                    gconDMIS.Execute "update CSMS_MATMAS set" & _
                                   " EMERGENCY_PO = " & N2Str2Zero(rsCSMS_MATMAS!emergency_po) + N2Str2Zero(rsPMIS_Tdaytran!TRANQTY) & _
                                   " where STOCKNO = " & N2Str2Null(rsPMIS_Tdaytran!STOCK_ORD)
                End If
            End If
            rsPMIS_Tdaytran.MoveNext
        Loop
    End If
    SQL_STATEMENT = "update PMIS_PO_HD set" & _
                  " status = 'P'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "P", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "", ""

    Set rsPMIS_Tdaytran = Nothing
    Set rsCSMS_MATMAS = Nothing

    post = True
    Exit Function
errordaa:
    error_msg = error
    post = False

End Function

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS PURCHASE ORDER") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgQuestionBox("PO Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        If NumericVal(txtDS1.Text) > 0 Then
            rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptPurchaseOrder.Formulas(2) = "PREPAREDBY = '" & (gconDMIS.Execute("Select PreparedBy from all_profile where modulename = 'PMIS'").Fields(0).Value) & "'"
            rptPurchaseOrder.Formulas(3) = "APPROVEDBY= '" & (gconDMIS.Execute("Select ApprovedBY from all_profile where modulename = 'PMIS'").Fields(0).Value) & "'"
              
              If COMPANY_CODE = "HPI" Then
                    rptPurchaseOrder.Formulas(3) = "PREPAREDBY = '" & GetSetting("PMIS", "SIGNATORIES", "PO-PREPBY", "") & "'"
                    rptPurchaseOrder.Formulas(4) = "APPROVEDBY= '" & GetSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", "") & "'"
                    rptPurchaseOrder.Formulas(5) = "CHECKEDBY= '" & GetSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", "") & "'"
                    rptPurchaseOrder.Formulas(6) = "OWNER= '" & GetSetting("PMIS", "SIGNATORIES", "PO-OWNER", "") & "'"
              End If
                
             If COMPANY_CODE = "HMH" Then
                    rptPurchaseOrder.Formulas(3) = "PreparedBy = '" & GetSignitories("PreparedBy", "PMIS") & "'"
                    rptPurchaseOrder.Formulas(4) = "CheckedBy = '" & GetSignitories("CheckedBy", "PMIS") & "'"
                    rptPurchaseOrder.Formulas(5) = "ApprovedBy = '" & GetSignitories("ApprovedBy", "PMIS") & "'"
             End If
                        
            PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO.RPT", "{Po_hd.type} = 'M' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
        Else
    
            rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            rptPurchaseOrder.Formulas(2) = "PREPAREDBY = '" & (gconDMIS.Execute("Select PreparedBy from all_profile where modulename = 'PMIS'").Fields(0).Value) & "'"
            rptPurchaseOrder.Formulas(3) = "APPROVEDBY= '" & (gconDMIS.Execute("Select ApprovedBY from all_profile where modulename = 'PMIS'").Fields(0).Value) & "'"
              If COMPANY_CODE = "HPI" Then
                    rptPurchaseOrder.Formulas(3) = "PREPAREDBY = '" & GetSetting("PMIS", "SIGNATORIES", "PO-PREPBY", "") & "'"
                    rptPurchaseOrder.Formulas(4) = "APPROVEDBY= '" & GetSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", "") & "'"
                    rptPurchaseOrder.Formulas(5) = "CHECKEDBY= '" & GetSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", "") & "'"
                    rptPurchaseOrder.Formulas(6) = "OWNER= '" & GetSetting("PMIS", "SIGNATORIES", "PO-OWNER", "") & "'"
             End If
                
             If COMPANY_CODE = "HMH" Then
                    rptPurchaseOrder.Formulas(3) = "PreparedBy = '" & GetSignitories("PreparedBy", "PMIS") & "'"
                    rptPurchaseOrder.Formulas(4) = "CheckedBy = '" & GetSignitories("CheckedBy", "PMIS") & "'"
                    rptPurchaseOrder.Formulas(5) = "ApprovedBy = '" & GetSignitories("ApprovedBy", "PMIS") & "'"
             End If
                
            PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_nonvat.RPT", "{Po_hd.type} = 'M' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
        End If
        Screen.MousePointer = 0
    End If
    NEW_LogAudit "V", "MATERIALS PURCHASE ORDER", "", labID, "Materials", "", "", ""
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Function GetSignitories(xFields As String, xMODULENAME As String) As String
    Dim SQLTXT As String
    Dim rsTMP As New ADODB.Recordset
        
    SQLTXT = "SELECT " & xFields & " AS UFIELD FROM ALL_PROFILE WHERE ModuleName = '" & xMODULENAME & "'"
    Set rsTMP = gconDMIS.Execute(SQLTXT)
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
          GetSignitories = Null2String(rsTMP!UFIELD)
    Else
          GetSignitories = ""
    End If

    Set rsTMP = Nothing
End Function

Private Sub cmdTranCancel_Click()
    Picture1.Enabled = True
    fraDetails.Enabled = True
    Frame2.Enabled = True
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()

    On Error GoTo ErrorCode:

    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_Tdaytran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", labDetID
    End If
    Dim CNT                                            As Integer
    Dim rsPMIS_TdaytranDup                             As ADODB.Recordset
    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select id,itemno from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        CNT = 0
        Do While Not rsPMIS_TdaytranDup.EOF
            CNT = CNT + 1
            gconDMIS.Execute "update PMIS_Tdaytran set itemno = " & Format(CNT, "0000") & " where id = " & rsPMIS_TdaytranDup!ID
            rsPMIS_TdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labID.Caption
    Else
        PO_TOTVAT = NumericVal(txtDS_Amt1.Text)
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labID.Caption
    End If
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    Dim CTR                                 As Integer
    Dim sqlcommand                          As String
    Dim Msg                                 As String
    
    Screen.MousePointer = 11
    On Error GoTo ErrorCode

    If cboTranPartNo.Text = "" Then
        MsgBox "Warning: Part Number must have a value", vbInformation + vbOKOnly
        On Error Resume Next
        cboTranPartNo.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If cboTranDescription.Text = "" Then
        MsgBox "Warning: Material's Description must have a value ", vbInformation + vbOKOnly
        On Error Resume Next
        cboTranDescription.SetFocus
        Exit Sub
    End If

'updated by: IEBV 02012011_0400pm
'description:  saves the part number that dont exist on Master file
'----------------------------------------------------------------------------------------------------------------
    CTR = 0
    CTR = (gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))) & " AND [TYPE] = 'M' ").Fields(0).Value)
    If CTR > 0 Then
        'do nothing
    Else
        Msg = "Part Number Doesn't Exist On Material Master file." & vbCrLf
        Msg = Msg + "Do You Want To Add It To Master File?"
        If MsgBox(Msg, vbQuestion + vbYesNo) = vbYes Then
            sqlcommand = "Insert into PMIS_stockmas ([TYPE], STOCKNO,STOCKDESC,STOCKTYPE,USERCODE,LASTUPDATE,ACTIVE) "
            sqlcommand = sqlcommand + " VALUES ('M'," & N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))) & ",'" & UCase(cboTranDescription.Text) & "', "
            sqlcommand = sqlcommand + " 'GJ', '" & Null2String(RSPO_HD!USERCODE) & "','" & LOGDATE & "', 'N') "
            gconDMIS.Execute sqlcommand
        Else
            Exit Sub
        End If
    End If
'----------------------------------------------------------------------------------------------------------------
    
    If AddorEdit = "ADD" Then
        Dim rsPMIS_TdaytranClone                       As ADODB.Recordset
        Set rsPMIS_TdaytranClone = New ADODB.Recordset
        rsPMIS_TdaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_Tdaytran where [TYPE] = 'M' AND STOCK_ORD = " & UCase(N2Str2Null(LTrim(RTrim(cboTranPartNo.Text)))) & " and trantype = 'PO' and tranno =" & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
        If Not rsPMIS_TdaytranClone.EOF And Not rsPMIS_TdaytranClone.BOF Then
            MsgSpeechBox "Warning: Material Code already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    Dim POTRANDATE, POTRANNO, POTRANTYPE               As String
    Dim POITEMNO, POSTOCK_ORD, POSTOCK_SUP             As String
    Dim POTRANQTY                                      As Integer
    Dim POTRANUCOST                                    As Double
    Dim POSTATUS                                       As String
    Dim POTRANINVAMT                                   As Double
    Dim POTRANVIN                                      As String

    POTRANDATE = N2Date2Null(txtPODate.Text)
    POTRANTYPE = "'" & "PO" & "'"
    POTRANNO = N2Str2Null(txtPONo.Text)
    POITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    POSTOCK_ORD = UCase(N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))))
    POSTOCK_SUP = UCase(N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))))
    POTRANQTY = NumericVal(txtTranQty.Text)
    POTRANINVAMT = NumericVal(txtTranINVAmt.Text)
    POTRANUCOST = NumericVal(txtUnitCost.Text)
    POSTATUS = "'N'"
    POTRANVIN = N2Str2Null(txtVIN.Text)

    If POTRANINVAMT <= 0 Then
        If MsgBox("Warning: Invoice Amount Is zero! Do You Want to Continue", vbInformation + vbYesNo) = vbNo Then
            On Error Resume Next
            txtTranINVAmt.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_Tdaytran " & _
                        "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt,VIN,lastupdate,usercode,status)" & _
                      " values ('M'," & POTRANDATE & ", " & POTRANTYPE & ", " & POTRANNO & "," & _
                      " " & POITEMNO & "," & POSTOCK_ORD & "," & _
                      " " & POSTOCK_SUP & ", " & POTRANQTY & "," & _
                      " " & POTRANUCOST & ", " & POTRANINVAMT & "," & POTRANVIN & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & POSTATUS & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", ""
    Else
        SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                      " trandate = " & POTRANDATE & "," & _
                      " trantype = " & POTRANTYPE & "," & _
                      " tranno = " & POTRANNO & "," & _
                      " itemno = " & POITEMNO & "," & _
                      " STOCK_ORD = " & POSTOCK_ORD & "," & _
                      " STOCK_SUP = " & POSTOCK_SUP & "," & _
                      " tranqty = " & POTRANQTY & "," & _
                      " tranucost = " & POTRANUCOST & "," & _
                      " traninvamt = " & POTRANINVAMT & "," & _
                      " VIN = " & POTRANVIN & "," & _
                      " lastupdate = '" & LOGDATE & "'," & _
                      " status = " & POSTATUS & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "" & _
                      " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", labDetID
    End If

    Dim rsCSMS_MATMASClone                             As ADODB.Recordset
    Set rsCSMS_MATMASClone = New ADODB.Recordset
    rsCSMS_MATMASClone.Open "select STOCKNO,tpoqty,onorder,mac,dnp,srp,onhand from CSMS_MATMAS where STOCKNO = " & POSTOCK_ORD, gconDMIS
    If Not rsCSMS_MATMASClone.EOF And Not rsCSMS_MATMASClone.BOF Then
    Else
        If txtSupCode.Text = VPAMCOR Then
            MsgBox "Material Code doesn't exist this will automatically add to Master file!", vbInformation + vbOKOnly, "Invalid Material code"
            gconDMIS.Execute "insert into CSMS_MATMAS " & _
                             "(TYPE,STOCKNO,STOCKDESC,date_entered)" & _
                           " values ('M'," & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
        Else
            MsgBox "Material Code doesn't exist this will automatically add to Master file!", vbInformation + vbOKOnly, "Invalid Material code"
            gconDMIS.Execute "insert into CSMS_MATMAS " & _
                             "(TYPE,STOCKNO,STOCKDESC,date_entered)" & _
                           " values ('M'," & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
        End If
    End If
    cleargrid grdDetails
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " totalqty = " & PO_TOTQTY & "," & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labID.Caption
    Else
        PO_TOTVAT = NumericVal(txtDS_Amt1.Text)
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " totalqty = " & PO_TOTQTY & "," & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labID.Caption
    End If
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labID.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
        Picture1.Enabled = False
        fraDetails.Enabled = False
    Else
        cmdTranCancel.Value = True
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Function PO_EXISTS(PO_NO As String) As Boolean
    Dim rsTMP As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_RR_HD" & vbCrLf
    SQL = SQL & "Union All" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_REC_HIST" & vbCrLf
    SQL = SQL & ") T WHERE PONO = '" & PO_NO & "' AND TYPE = 'M' AND STATUS = 'P'" & vbCrLf

    Set rsTMP = gconDMIS.Execute(SQL)
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        PO_EXISTS = True
    Else
        PO_EXISTS = False
    End If
    
    Set rsTMP = Nothing
End Function

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "MATERIALS PURCHASE ORDER") = False Then Exit Sub

    On Error GoTo ErrorCode:
    
    'Updated By: IEBV
    'description:   TO avoin uposting of PO if PO is already receive but not yet posted
    If chkfnotyet_posted(txtPONo.Text, "M") = True Then MessagePop InfoFriend, "Action Void", "You cannot Unpost this transaction, Its already received but not yet posted!": Exit Sub

    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        
        If PO_EXISTS(txtPONo) = True Then
            MessagePop InfoFriend, "Action Void", "You cannot Unpost this transaction, Its already Received!"
            Exit Sub
        End If
       
        'updated by: IEBV 11172011
        'description: to rollback transaction if error occured
        gconDMIS.BeginTrans
        If UNpost = False Then
        
            str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
            str_MSG = str_MSG & "Description: "
            str_MSG = str_MSG & " " & error_msg
            str_MSG = str_MSG & " " & vbCrLf
            str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
            str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
            
            str_MSG = Replace(str_MSG, "@UTX83912839123", "Unposting of Transaction")
            MsgBox str_MSG, vbCritical, "Unposting Error"
            gconDMIS.RollbackTrans
            Screen.MousePointer = 0
            Exit Sub
        End If
        gconDMIS.CommitTrans
        
        rsRefresh
        RSPO_HD.Find "id =" & labID.Caption
        StoreMemVars
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Function UNpost() As Boolean
On Error GoTo errordaa

    SQL_STATEMENT = "update PMIS_PO_HD set" & _
                  " status = 'N'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "U", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "", ""

    Dim rsPMIS_TdaytranDup, rsCSMS_MATMASDup       As ADODB.Recordset
    Dim PCurOnOrder, PCurTpoQty                    As Integer
    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select ID,Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_Tdaytran where [TYPE] = 'M' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        Do While Not rsPMIS_TdaytranDup.EOF
            Set rsCSMS_MATMASDup = New ADODB.Recordset
            rsCSMS_MATMASDup.Open "select STOCKNO,onorder,tpoqty,ordered,emergency_po,purchases from CSMS_MATMAS where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), gconDMIS
            If Not rsCSMS_MATMASDup.EOF And Not rsCSMS_MATMASDup.BOF Then
                PCurOnOrder = N2Str2IntZero(rsCSMS_MATMASDup!ONORDER) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                PCurTpoQty = N2Str2IntZero(rsCSMS_MATMASDup!tpoqty) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                If Null2String(rsPMIS_TdaytranDup!Status) = "P" Then
                    gconDMIS.Execute "update CSMS_MATMAS set" & _
                                   " purchases = " & N2Str2Zero(rsCSMS_MATMASDup!purchases) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                   " onorder = " & PCurOnOrder & "," & _
                                   " tpoqty = " & PCurTpoQty & "," & _
                                   " ORDERED = " & N2Str2IntZero(rsCSMS_MATMASDup!Ordered) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    If Mid(txtDON.Text, 3, 1) = "E" Then
                        gconDMIS.Execute "update CSMS_MATMAS set" & _
                                       " EMERGENCY_PO = " & N2Str2IntZero(rsCSMS_MATMASDup!emergency_po) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & _
                                       " where STOCKNO = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    End If
                End If
            End If
            SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                          " status = 'N'," & _
                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                          " lastupdate = '" & LOGDATE & "'" & _
                          " where ID = " & rsPMIS_TdaytranDup!ID
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "U", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "", ""
            rsPMIS_TdaytranDup.MoveNext
        Loop
    End If

    Set rsPMIS_TdaytranDup = Nothing
    Set rsCSMS_MATMASDup = Nothing


    UNpost = True
    Exit Function
errordaa:
    error_msg = error
    UNpost = False
    
End Function

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "MATERIALS PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next
    txtPONo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Frame2.Enabled = True
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "EDIT"
    PrevPONO = Format(txtPONo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next

    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSPO_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSPO_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSPO_HD.MoveNext
    If RSPO_HD.EOF Then
        RSPO_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSPO_HD.MovePrevious
    If RSPO_HD.BOF Then
        RSPO_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsPO_HDDup                                     As ADODB.Recordset
    'AXP02232008
    If Len(Trim(RTrim(txtPONo))) <> 6 Then
        MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
        On Error Resume Next
        txtPONo.SetFocus
        Exit Sub
    End If
    If txtSupCode.Text = "" Then
        MsgSpeechBox "Warning: Supplier Code must not be empty!"
        On Error Resume Next
        txtSupCode.SetFocus
        Exit Sub
    End If
    If txtPODate.Text = "" Or IsDate(txtPODate.Text) = False Then
        MsgSpeechBox "Invalid Date!"
        On Error Resume Next
        txtPODate.SetFocus
        Exit Sub
    End If

    If cboModelCode.Text = "" Then
        MsgBox "Vehicle model must not be empty!", vbCritical, "Purchase Order"
        cboModelCode.SetFocus
        Exit Sub
    End If

    If cboSupName.Text = "" Then
        MsgBox "Supplier name cannot be blank!", vbCritical + vbOKOnly
        On Error Resume Next
        cboSupName.SetFocus
        Exit Sub
        
    End If
    'VALIDATION FOR TRANSACTION NUMBER
    If IsNull(txtPONo.Text) = True Then
        MsgSpeechBox "Warning: Purchase Order Number must not be empty"
        On Error Resume Next
        txtPONo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            If checkdup_PO("M", txtPONo.Text) = True Then
                MsgSpeechBox "Purchase Order Number already exist!"
                On Error Resume Next
                txtPONo.SetFocus
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
                If checkdup_PO("M", txtPONo.Text) = True Then
                    MsgSpeechBox "Purchase Order Number already exist!"
                    On Error Resume Next
                    txtPONo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    'updated by: IEBV 11172011
    'description: to rollback transaction if error occured
     gconDMIS.BeginTrans
     If save = False Then
     
         str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
         str_MSG = str_MSG & "Description: "
         str_MSG = str_MSG & " " & error_msg
         str_MSG = str_MSG & " " & vbCrLf
         str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
         str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
         str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
         
         str_MSG = Replace(str_MSG, "@UTX83912839123", "Saving of Transaction")
         MsgBox str_MSG, vbCritical, "Saving Error"
         gconDMIS.RollbackTrans
         Screen.MousePointer = 0
         Exit Sub
     End If
     gconDMIS.CommitTrans

    Exit Sub



ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Function save() As Boolean
On Error GoTo errordaa

    Dim NewPOPMIS_Counter                              As String
    Dim VTXTPONo, VTXTPPNo, VTXTPODate                 As String
    Dim VcboSupName, VTXTSup_Addrs, VTXTDealerCode     As String
    Dim VTXTShipTo, VTXTPO_Amount                      As String
    Dim VTXTDS1, VTXTDS_Desc1, VTXTDS_Amt1             As String
    Dim VTXTNetPOAmt, VTXTRemarks, VTXTSupCode         As String

    Dim VTXTDON, VTXTORDERTYPE, VTXTORDER_SERIES       As String
    Dim VCBOContactCode, VCBOModelCode                 As String

    NewPOPMIS_Counter = NumericVal(txtPONo.Text) + 1

    VTXTSupCode = N2Str2Null(txtSupCode.Text)
    VTXTPONo = N2Str2Null(txtPONo.Text)
    VTXTPPNo = N2Str2Null(cboPP_No.Text)
    VTXTPODate = N2Date2Null(txtPODate.Text)

    VTXTORDERTYPE = N2Str2Null(Mid(txtDON.Text, 3, 1))
    VTXTORDER_SERIES = N2Str2Null(Mid(txtDON.Text, 8, 2))
    VTXTDON = N2Str2Null(txtDON.Text)

    VcboSupName = N2Str2Null(cboSupName.Text)
    VTXTSup_Addrs = N2Str2Null(Trim(txtSup_Addrs.Text))
    VTXTDealerCode = N2Str2Null(txtDealerCode.Text)

    VCBOContactCode = N2Str2Null(cboContactCode.Text)
    VCBOModelCode = N2Str2Null(cboModelCode.Text)

    VTXTShipTo = N2Str2Null(txtShipTo.Text)
    VTXTPO_Amount = NumericVal(txtPO_Amount.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetPOAmt = NumericVal(txtNetPOAmt.Text)
    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "Insert into PMIS_PO_HD" & _
                      " (TYPE,pono,ppno,podate,DON,ORDERTYPE,ORDER_SERIES,supcode,supname,sup_addrs,dealercode,ContactCode,ModelCode,po_amount,ds1,ds_desc1,ds_amt1,netpoamt,usercode,lastupdate,remarks)" & _
                      " values ('M'," & VTXTPONo & ", " & VTXTPPNo & ", " & VTXTPODate & "," & VTXTDON & ", " & VTXTORDERTYPE & "," & VTXTORDER_SERIES & _
                        ", " & VTXTSupCode & ", " & VcboSupName & _
                        ", " & VTXTSup_Addrs & ", " & VTXTDealerCode & "," & VCBOContactCode & "," & VCBOModelCode & _
                        ", " & VTXTPO_Amount & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNetPOAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, FindTransactionID(txtPONo, "pono", "PMIS_PO_HD", "DETAILS", N2Str2Null("M"), "TYPE"), "Materials", txtPONo, "PO", ""

    Else
        SQL_STATEMENT = "update PMIS_PO_HD set" & _
                      " pono = " & VTXTPONo & "," & _
                      " ppno = " & VTXTPPNo & "," & _
                      " podate = " & VTXTPODate & "," & _
                      " DON = " & VTXTDON & "," & _
                      " ORDERTYPE = " & VTXTORDERTYPE & "," & _
                      " ORDER_SERIES = " & VTXTORDER_SERIES & "," & _
                      " supcode = " & VTXTSupCode & "," & _
                      " supname = " & VcboSupName & "," & _
                      " sup_addrs = " & VTXTSup_Addrs & "," & _
                      " dealercode = " & VTXTDealerCode & "," & _
                      " Contactcode = " & VCBOContactCode & "," & _
                      " Modelcode = " & VCBOModelCode & "," & _
                      " po_amount = " & VTXTPO_Amount & "," & _
                      " ds1 = " & VTXTDS1 & "," & _
                      " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                      " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                      " netpoamt = " & VTXTNetPOAmt & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "," & _
                      " lastupdate = '" & LOGDATE & "'," & _
                      " remarks = " & VTXTRemarks & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "", ""

        SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                      " trandate = " & VTXTPODate & "," & _
                      " tranno = " & VTXTPONo & _
                      " where [TYPE] = 'M' AND tranno = '" & Null2String(RSPO_HD!PONO) & "' and trantype = 'PO'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS PURCHASE ORDER", SQL_STATEMENT, labID, "Materials", txtPONo, "PO", ""

    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NewPOPMIS_Counter & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where [TYPE] = 'M' AND modul = 'PO'"
        Call FillGrid
    End If
    rsRefresh
    RSPO_HD.Find "pono = " & VTXTPONo
    cmdCancel.Value = True
    DoEvents
    cleargrid grdDetails
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " totalqty = " & PO_TOTQTY & "," & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labID.Caption
    Else
        PO_TOTVAT = NumericVal(txtDS_Amt1.Text)
        gconDMIS.Execute "update PMIS_PO_HD set" & _
                       " totalqty = " & PO_TOTQTY & "," & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labID.Caption
    End If

    If AddorEdit = "ADD" Then
        Picture1.Enabled = False
        fraDetails.Enabled = False

    Else
        Picture1.Enabled = True
        fraDetails.Enabled = True

    End If

    rsRefresh
    RSPO_HD.Find "id = " & labID.Caption
    StoreMemVars
    'If AddorEdit = "ADD" Then cmdAddTran_Click
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
        Picture1.Enabled = False
    End If

    save = True
    Exit Function
errordaa:
    error_msg = error
    save = False
End Function


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text

    If Shift = 2 Then
        If KeyCode = vbKeyF1 Then
            'If picDetails.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MATERIALS PURCHASE ORDER)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "MATERIALS PURCHASE ORDER")
        End If
    End If

    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            Frame2.Enabled = True
            SendToBackConfirmPO
            Picture1.Enabled = True
            fraDetails.Enabled = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSPO_HD!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change"
                ElseIf Null2String(RSPO_HD!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                Else
                    cmdAddTran_Click
                    Picture1.Enabled = False
                    fraDetails.Enabled = False
                End If
            End If
        Case vbKeyF4
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSPO_HD!Status) <> "P" And Null2String(RSPO_HD!Status) <> "C" Then
                        grdDetails_DblClick
                    End If
                End If
            End If
        Case vbKeyF5
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If Null2String(RSPO_HD!Status) <> "P" And Null2String(RSPO_HD!Status) <> "C" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If

            Picture1.Enabled = False
            fraDetails.Enabled = False

        Case vbKeyF9
            If picConfirmation.Visible = True Then
                SendToFrontConfirmPO
            End If
        Case vbKeyF11
            If picConfirmation.Visible = True Then
                SendToFrontConfirmPO
            End If
        Case vbKeyF12
            If cmdUnPost.Enabled = True Then cmdUnPost.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    If COMPANY_CODE = "HCI" Then
        txtwvatm.Visible = True
    Else
        txtwvatm.Visible = False
    End If
    
    Frame1.Enabled = False: SendToBackConfirmPO: cmdAddTran.Enabled = False: picConfirmation.Visible = False
    Picture1.Visible = True: SendToBack
    Picture2.Visible = False: textSearch.Text = "": initMemvars: rsRefresh
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then RSPO_HD.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISTrans_Purchase = Nothing
    UnloadForm Me
End Sub



Private Sub grdDetails_DblClick()
    If Null2String(RSPO_HD!Status) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf Null2String(RSPO_HD!Status) = "C" Then
        MsgSpeechBox "Item(s) are Already Cancelled and cannot be edited"
    Else
        Frame2.Enabled = False
        Dim FILD                                       As String
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        FILD = grdDetails.Text
        If FILD <> "" And FILD <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Enabled = True
            BringToFront
            fraAddTran.Caption = "Edit Materials"
            StorePartsEntry (FILD)
        Else
            MsgSpeechBox "No Entry on Materials"
            Exit Sub
        End If
    End If
End Sub

Private Sub cboTranDescription_LostFocus()
    cboTranDescription.Text = UCase(cboTranDescription.Text)
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
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

Private Sub txtPODate_LostFocus()
    txtPODate.Text = Format(txtPODate.Text, "SHORT DATE")
End Sub

Private Sub txtPONo_LostFocus()
    txtPONo.Text = Format(txtPONo.Text, "000000")
'    If Frame1.Enabled = True Then
'        If Len(txtPONo.Text) >= 3 Then
'            Dim rsPO_HDDup                             As ADODB.Recordset
'            If AddorEdit = "ADD" Then
'                Set rsPO_HDDup = New ADODB.Recordset
'                rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'M' AND pono = '" & txtPONo.Text & "'", gconDMIS
'                If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                    MsgSpeechBox "PO Number Already Exist!"
'                    On Error Resume Next
'                    txtPONo.SetFocus
'                End If
'            ElseIf AddorEdit = "EDIT" Then
'                If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
'                    Set rsPO_HDDup = New ADODB.Recordset
'                    rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'M' AND pono = '" & txtPONo.Text & "'", gconDMIS
'                    If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                        MsgSpeechBox "PO Number Already Exist!"
'                        On Error Resume Next
'                        txtPONo.SetFocus
'                    End If
'                End If
'            End If
'        End If
'    End If
'---------------------------------------------------------------------------------
'updated by: IEBV 12172010_1155Am
'description: to Check if po no. already exist
'    If Frame1.Enabled = True Then
'        If Len(txtPONo.Text) >= 3 Then
'            Dim rsPO_HDDup                             As ADODB.Recordset
'            If AddorEdit = "ADD" Then
'                If checkdup_PO("M", txtPONo.Text) = True Then
'                    MsgSpeechBox "PO Number Already Exist!"
'                    On Error Resume Next
'                    txtPONo.SetFocus
'                End If
'            ElseIf AddorEdit = "EDIT" Then
'                If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
'                    If checkdup_PO("M", txtPONo.Text) = True Then
'                        MsgSpeechBox "PO Number Already Exist!"
'                        On Error Resume Next
'                        txtPONo.SetFocus
'                    End If
'                End If
'            End If
'        End If
'    End If
'---------------------------------------------------------------------------------
End Sub

Private Sub txtRemarks_GotFocus()
    MsgSpeechBox "Pls Type Your Message Here!"
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSup_Addrs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSupCode_Change()
    cboSupName.Text = SetSupdesc(txtSupCode.Text)
End Sub

Private Sub txtTranINVAmt_GotFocus()
    If NumericVal(txtTranINVAmt.Text) = 0 Then txtTranINVAmt.Text = ""
End Sub

Private Sub txtTranINVAmt_LostFocus()
    If txtTranINVAmt.Text = "" Then txtTranINVAmt.Text = 0
    txtTranINVAmt.Text = Format(txtTranINVAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        If ISNONVAT = True Then
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
        Else
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
        End If
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
    End If
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_LostFocus()
    If Trim(txtTranQty.Text) = "" Then txtTranQty.Text = 1
    If ISNONVAT = True Then
        txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
    Else
        txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
    End If
    txtTranQty.Text = Format(txtTranQty.Text, DIGIT_FORMAT)
    txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
End Sub

Private Sub txtTranINVAmt_Change()
    txtTranINVAmt.Text = ToDoubleNumber(txtTranINVAmt.Text)
End Sub

Private Sub txtTranTotalAmt_LostFocus()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtUnitCost_Change()
    On Error Resume Next
    If NumericVal(txtUnitCost.Text) <> 0 Then
        If ISNONVAT = True Then
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text))
        Else
            txtTranINVAmt.Text = ToDoubleNumber(NumericVal(txtUnitCost.Text) * ConvertToBIRDecimalFormat(VAT_RATE))
        End If
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
    End If
End Sub

Private Sub txtUnitCost_GotFocus()
    If NumericVal(txtUnitCost.Text) > 0 Then
        txtUnitCost.Text = NumericVal(txtUnitCost.Text)
    Else
        txtUnitCost.Text = ""
    End If
End Sub

Private Sub txtUnitCost_LostFocus()
    txtUnitCost.Text = Format(txtUnitCost.Text, MAXIMUM_DIGIT)
End Sub

Private Sub lstPO_HD_GotFocus()
    RSPO_HD.Find = "PONO=" & lstPO_HD.SelectedItem.Text
    'RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstPO_HD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optPONo.Value = True Then
        RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", Item).Bookmark
    Else
        RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstPO_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPO_HD
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstPO_HD_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstPO_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optPONo.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    Else
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstPO_HD.ListItems.Count > 0 And lstPO_HD.Enabled = True Then: lstPO_HD.SetFocus
    End If
End Sub

Private Sub optRONo_Click()
    lstPO_HD.ColumnHeaders(1).Text = "Sup. Name"
    lstPO_HD.ColumnHeaders(1).Width = 4000
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optpono_Click()
    lstPO_HD.ColumnHeaders(1).Text = "Tran. No."
    lstPO_HD.ColumnHeaders(1).Width = 2150
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub txtwvatm_Change()
    If COMPANY_CODE = "HCI" Then
        If NumericVal(txtwvatm) <> 0 Then
            txtUnitCost.Text = Format(NumericVal(txtwvatm) / 1.12, "#,###,##0.00")
        End If
End If
End Sub

Sub click()
    RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

