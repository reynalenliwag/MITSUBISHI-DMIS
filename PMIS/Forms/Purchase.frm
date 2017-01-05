VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISTrans_Purchase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Entry"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Purchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11745
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2280
      ScaleHeight     =   870
      ScaleWidth      =   9405
      TabIndex        =   70
      Top             =   6330
      Width           =   9405
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
         Left            =   8580
         MouseIcon       =   "Purchase.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   73
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
         Left            =   7800
         MouseIcon       =   "Purchase.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   74
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
         Left            =   7020
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Purchase.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   80
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
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Purchase.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   81
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
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "Purchase.frx":1B5D
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":1CAF
         Style           =   1  'Graphical
         TabIndex        =   82
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
         MouseIcon       =   "Purchase.frx":1FD4
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   75
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
         MouseIcon       =   "Purchase.frx":2482
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":25D4
         Style           =   1  'Graphical
         TabIndex        =   76
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
         MouseIcon       =   "Purchase.frx":28E7
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":2A39
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Move to First Record"
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
         MouseIcon       =   "Purchase.frx":2D89
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":2EDB
         Style           =   1  'Graphical
         TabIndex        =   71
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
         MouseIcon       =   "Purchase.frx":3239
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":338B
         Style           =   1  'Graphical
         TabIndex        =   77
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
         MouseIcon       =   "Purchase.frx":3685
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":37D7
         Style           =   1  'Graphical
         TabIndex        =   78
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
         MouseIcon       =   "Purchase.frx":3B2F
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":3C81
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox picConfirmation 
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
      Height          =   315
      Left            =   2220
      ScaleHeight     =   285
      ScaleWidth      =   9435
      TabIndex        =   90
      Top             =   7200
      Width           =   9465
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F9 - View/Update PO Upon Confirmation"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   30
         MouseIcon       =   "Purchase.frx":3FE0
         MousePointer    =   99  'Custom
         TabIndex        =   91
         Top             =   30
         Width           =   9285
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
      ScaleWidth      =   9435
      TabIndex        =   65
      Top             =   6000
      Width           =   9465
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7110
         TabIndex        =   96
         Top             =   30
         Width           =   2445
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5070
         TabIndex        =   69
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3360
         TabIndex        =   68
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1740
         TabIndex        =   67
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   120
         TabIndex        =   66
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7530
      Left            =   60
      TabIndex        =   59
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstPO_HD 
         Height          =   5730
         Left            =   60
         TabIndex        =   63
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10107
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
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
         MouseIcon       =   "Purchase.frx":4132
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
         TabIndex        =   110
         Top             =   7150
         Width           =   1965
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
         TabIndex        =   64
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10050
      ScaleHeight     =   855
      ScaleWidth      =   1680
      TabIndex        =   83
      Top             =   6330
      Width           =   1680
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
         MouseIcon       =   "Purchase.frx":4294
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":43E6
         Style           =   1  'Graphical
         TabIndex        =   84
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
         Left            =   30
         MouseIcon       =   "Purchase.frx":4724
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":4876
         Style           =   1  'Graphical
         TabIndex        =   85
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
      Height          =   2985
      Left            =   2220
      TabIndex        =   31
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
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
         Height          =   345
         Left            =   4800
         TabIndex        =   98
         Top             =   180
         Width           =   255
      End
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
         ItemData        =   "Purchase.frx":4BC6
         Left            =   1680
         List            =   "Purchase.frx":4BC8
         TabIndex        =   10
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   2190
         Width           =   2925
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
         Left            =   6240
         TabIndex        =   12
         Text            =   "16A070101"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdDON 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   87
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
         TabIndex        =   11
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   2580
         Width           =   2925
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   9240
         Top             =   240
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
         TabIndex        =   15
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
         Left            =   3520
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   54
         ToolTipText     =   "Type the date of the purchase order in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   180
         Width           =   1275
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
         Left            =   4710
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "Purchase.frx":4BCA
         ToolTipText     =   "Type your message or your remarks."
         Top             =   1890
         Width           =   4665
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
         TabIndex        =   6
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   7
         ToolTipText     =   "Type the supplier code (e.g. 00001)"
         Top             =   630
         Width           =   1005
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   5
         ToolTipText     =   "Type the purchase order number (e.g. 002774)"
         Top             =   180
         Width           =   1245
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
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
         Height          =   795
         Left            =   60
         ScaleHeight     =   795
         ScaleWidth      =   4575
         TabIndex        =   36
         Top             =   1410
         Width           =   4575
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
            TabIndex        =   9
            Top             =   30
            Width           =   4515
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
         Height          =   1185
         Left            =   6480
         ScaleHeight     =   1185
         ScaleWidth      =   2925
         TabIndex        =   37
         Top             =   660
         Width           =   2925
         Begin VB.TextBox txtNetPOAmt 
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
            TabIndex        =   58
            Top             =   780
            Width           =   1485
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
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   57
            Top             =   390
            Width           =   1485
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
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   56
            Top             =   0
            Width           =   1485
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
            TabIndex        =   16
            ToolTipText     =   "Type the type of the additional amount (e.g. VAT)"
            Top             =   390
            Width           =   1335
         End
         Begin VB.Label Label9 
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
            Left            =   90
            TabIndex        =   52
            Top             =   30
            Width           =   1245
         End
         Begin VB.Label Label10 
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
            Left            =   90
            TabIndex        =   53
            Top             =   810
            Width           =   1245
         End
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
         TabIndex        =   8
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1050
         Width           =   4515
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   118
         Top             =   240
         Width           =   135
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   115
         Top             =   2640
         Width           =   135
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   114
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   113
         Top             =   660
         Width           =   135
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   2580
         TabIndex        =   112
         Top             =   195
         Width           =   135
      End
      Begin VB.Label lblimportant 
         BackStyle       =   0  'Transparent
         Caption         =   "* "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   111
         Top             =   195
         Width           =   135
      End
      Begin VB.Label LBL_T_S 
         Caption         =   "Label12"
         Height          =   315
         Left            =   5010
         TabIndex        =   106
         Top             =   780
         Visible         =   0   'False
         Width           =   1635
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
         TabIndex        =   89
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
         Left            =   5280
         TabIndex        =   88
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
         Left            =   240
         TabIndex        =   86
         Top             =   2640
         Width           =   1965
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
         TabIndex        =   55
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
         Left            =   4710
         TabIndex        =   49
         Top             =   1590
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
         TabIndex        =   51
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
         Left            =   7290
         TabIndex        =   50
         Top             =   180
         Width           =   2175
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
         TabIndex        =   48
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
         Left            =   240
         TabIndex        =   35
         Top             =   210
         Width           =   1125
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
         Left            =   2710
         TabIndex        =   34
         Top             =   210
         Width           =   825
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
         Left            =   480
         TabIndex        =   33
         Top             =   660
         Width           =   765
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
         TabIndex        =   32
         Top             =   1050
         Visible         =   0   'False
         Width           =   525
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
      Height          =   3045
      Left            =   2220
      TabIndex        =   93
      Top             =   2940
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2835
         Left            =   70
         TabIndex        =   94
         Top             =   120
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   5001
         _Version        =   393216
         Cols            =   10
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
      Caption         =   "Command2"
      Height          =   4635
      Left            =   4170
      TabIndex        =   107
      Top             =   810
      Width           =   5505
   End
   Begin VB.PictureBox picPrintPOExcel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      Height          =   4305
      Left            =   5280
      ScaleHeight     =   4275
      ScaleWidth      =   3825
      TabIndex        =   99
      Top             =   1590
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtowner 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   105
         Top             =   2520
         Width           =   3405
      End
      Begin VB.TextBox txtSIG_NotedbyDesign 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   2010
         Width           =   3405
      End
      Begin VB.TextBox txtSIG_PreparedBy 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   240
         TabIndex        =   0
         Top             =   630
         Width           =   3405
      End
      Begin VB.TextBox txtSIG_Notedby 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   1
         Top             =   1320
         Width           =   3405
      End
      Begin VB.CommandButton Command4 
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
         Left            =   2730
         MouseIcon       =   "Purchase.frx":4BE4
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":4D36
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Exit Window"
         Top             =   3300
         Width           =   795
      End
      Begin VB.CommandButton cmdSaveSig 
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
         Left            =   1950
         MouseIcon       =   "Purchase.frx":509C
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":51EE
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Save this Record"
         Top             =   3300
         Width           =   795
      End
      Begin VB.CommandButton Command3 
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
         Left            =   1170
         MouseIcon       =   "Purchase.frx":553E
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":5690
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print this Record"
         Top             =   3300
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SR. MNGR. OPERATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   103
         Top             =   1770
         Width           =   1935
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   102
         Top             =   0
         Width           =   4305
         _Version        =   655364
         _ExtentX        =   7594
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Print PO In Excel  Format"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   101
         Top             =   390
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   100
         Top             =   1080
         Width           =   900
      End
   End
   Begin VB.Frame fraAddTran 
      Caption         =   "Add/Edit Parts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   4260
      TabIndex        =   38
      Top             =   840
      Width           =   5325
      Begin VB.TextBox txtwvat 
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
         Left            =   3390
         TabIndex        =   108
         Top             =   2250
         Width           =   1815
      End
      Begin VB.OptionButton optKILL 
         Caption         =   "Kill"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   26
         Top             =   2520
         Width           =   1485
      End
      Begin VB.OptionButton optFILL 
         Caption         =   "Fill"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2730
         TabIndex        =   25
         Top             =   2280
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   97
         Top             =   630
         Width           =   285
      End
      Begin VB.CheckBox chkUseHARIDNP 
         Caption         =   "Use HARI DNP"
         Height          =   195
         Left            =   2850
         TabIndex        =   95
         Top             =   1650
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
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3090
         Width           =   2925
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   18
         Top             =   240
         Width           =   915
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   46
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
         MouseIcon       =   "Purchase.frx":59F6
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":5B48
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Delete Entry"
         Top             =   3510
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
         MouseIcon       =   "Purchase.frx":5E73
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":5FC5
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel Entry"
         Top             =   3510
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
         MouseIcon       =   "Purchase.frx":6303
         MousePointer    =   99  'Custom
         Picture         =   "Purchase.frx":6455
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save Entry"
         Top             =   3510
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "*Unit Cost w/ VAT *"
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
         Left            =   3150
         TabIndex        =   109
         Top             =   1980
         Width           =   2025
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
         Left            =   900
         TabIndex        =   92
         Top             =   3120
         Width           =   1275
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   150
         TabIndex        =   39
         Top             =   2730
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   420
         TabIndex        =   47
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
         Left            =   1980
         TabIndex        =   45
         Top             =   3660
         Width           =   285
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amt."
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   240
         TabIndex        =   44
         Top             =   2010
         Width           =   1185
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   630
         TabIndex        =   43
         Top             =   1650
         Width           =   795
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   600
         TabIndex        =   42
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Item No."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   570
         TabIndex        =   41
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   40
         Top             =   990
         Width           =   1245
      End
   End
   Begin VB.Label lblimportant 
      BackStyle       =   0  'Transparent
      Caption         =   "* "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   6
      Left            =   10320
      TabIndex        =   117
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "Required Field's"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   10440
      TabIndex        =   116
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Menu cmdmenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menuhist 
         Caption         =   "See Transaction History..."
      End
      Begin VB.Menu menumaster 
         Caption         =   "See Master File..."
      End
   End
End
Attribute VB_Name = "frmPMISTrans_Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPO_HD As ADODB.Recordset, rsPMIS_PP_HD As ADODB.Recordset, rsPMIS_Tdaytran As ADODB.Recordset
Attribute rsPMIS_PP_HD.VB_VarUserMemId = 1073938432
Attribute rsPMIS_Tdaytran.VB_VarUserMemId = 1073938432
Dim rsPMIS_Partmas As ADODB.Recordset, rsSupplier      As ADODB.Recordset
Attribute rsPMIS_Partmas.VB_VarUserMemId = 1073938435
Attribute rsSupplier.VB_VarUserMemId = 1073938435
Dim rsALL_Profile As ADODB.Recordset, rsPMIS_Counter   As ADODB.Recordset
Attribute rsALL_Profile.VB_VarUserMemId = 1073938437
Attribute rsPMIS_Counter.VB_VarUserMemId = 1073938437
Dim RSDAYTRAN                                          As New ADODB.Recordset

Dim PO_TOTQTY                                          As Integer
Dim Pcnt                                               As Integer
Attribute Pcnt.VB_VarUserMemId = 1073938439
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938440
Dim PO_TOTUCOST As Double, PO_TOTINVAMT                As Double
Attribute PO_TOTUCOST.VB_VarUserMemId = 1073938441
Attribute PO_TOTINVAMT.VB_VarUserMemId = 1073938441
Dim PO_TOTVAT                                          As Double
Attribute PO_TOTVAT.VB_VarUserMemId = 1073938443
Dim PO_T_ONORDER                                       As Long
Attribute PO_T_ONORDER.VB_VarUserMemId = 1073938444
Dim PrevPONO                                           As String
Attribute PrevPONO.VB_VarUserMemId = 1073938445
Dim PrevPmasMAC As Double, PrevPmasDNP As Double, PrevPmasSRP As Double
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
Attribute DON_TYPE.VB_VarUserMemId = 1073938435

Dim xlApp                                              As Excel.Application
Attribute xlApp.VB_VarUserMemId = 1073938436
Dim xlBook                                             As Excel.Workbook
Attribute xlBook.VB_VarUserMemId = 1073938437
Dim xlSheet                                            As Excel.Worksheet
Attribute xlSheet.VB_VarUserMemId = 1073938438

Sub fill_LblRefRRno()
    Dim rsTMP As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_RR_HD" & vbCrLf
    SQL = SQL & "Union All" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_REC_HIST" & vbCrLf
    SQL = SQL & ") T WHERE PONO = '" & Null2String(txtPONo) & "' AND TYPE = 'P' AND STATUS = 'P'" & vbCrLf
 
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
    Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "Select partno,partdesc from PMIS_Partmas where partno = '" & ppp & "'", gconDMIS
    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
        SetPartDesc = UCase(Null2String(rsPMIS_Partmas!PARTDESC))
    End If
    ''EAP:032309 MODIFY CODE IF PARTNUMBER DOES NOT EXIST IN PARTMASTERFILE
    'Set rsPMIS_Partmas = New ADODB.Recordset
    '
    '    rsPMIS_Partmas.Open "Select partno,partdesc from PMIS_Partmas where partno = '" & ppp & "'", gconDMIS
    '    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
    '            SetPartDesc = UCase(Null2String(rsPMIS_Partmas!PARTDESC))
    '    Else
    '        rsPMIS_Partmas.Close
    '        Set rsPMIS_Partmas = Nothing
    '        Set rsPMIS_Partmas = New ADODB.Recordset
    '
    '        rsPMIS_Partmas.Open "Select partnumber,descriptio from PMIS_DNPP where partnumber = '" & ppp & "'", gconDMIS
    '
    '            If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
    '                SetPartDesc = UCase(Null2String(rsPMIS_Partmas!DESCRIPTIO))
    '            End If
    '    End If

End Function

Function SetPartDesc2(pid As Variant)
    If pid <> "" Then

        Set rsPMIS_Partmas = New ADODB.Recordset
        rsPMIS_Partmas.Open "Select id,partdesc,NON_HARI,DNP from PMIS_Partmas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
            SetPartDesc2 = Null2String(rsPMIS_Partmas!PARTDESC)

            If chkUseHARIDNP.Value = 1 Then
                chkUseHARIDNP_Click
            Else
                txtUnitCost.Text = Round(N2Str2Zero(rsPMIS_Partmas!dnp) / ConvertToBIRDecimalFormat(VAT_RATE), 2)
            End If
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
        Set rsPMIS_Partmas = New ADODB.Recordset
        rsPMIS_Partmas.Open "Select id,partno from PMIS_Partmas where id = " & pid, gconDMIS
        If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
            SetPartNo = Null2String(rsPMIS_Partmas!PARTNO)
        End If
    End If
End Function

Function SetPartIDPartNo(DDD As String)
    Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "Select id,partno from PMIS_Partmas where partno = " & N2Str2Null(DDD) & "", gconDMIS
    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
        SetPartIDPartNo = N2Str2IntZero(rsPMIS_Partmas!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "Select id,partdesc from PMIS_Partmas where (ltrim(rtrim(partdesc))) = '" & UCase(LTrim(RTrim(DDD))) & "'", gconDMIS
    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
        SetPartIDDesc = N2Str2IntZero(rsPMIS_Partmas!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsPMIS_Partmas = New ADODB.Recordset
        rsPMIS_Partmas.Open "Select partno,mac from PMIS_Partmas where partno = '" & ppp & "'", gconDMIS
        If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
            SetPartPrice = N2Str2Zero(rsPMIS_Partmas!MAC)
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
        If Null2String(rsSupplier!SupCode) = "H00001" Then
            txtDON.Enabled = True
            cmdDON.Enabled = True
        Else
            txtDON.Enabled = False
            txtDON.Text = ""
            cmdDON.Enabled = False
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
        optFILL = Null2Bool(rsPMIS_Tdaytran!PO_FILL)
        optKILL = Null2Bool(rsPMIS_Tdaytran!PO_KILL)
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

Sub PrintPOExcel(XXX As String)
    Screen.MousePointer = 11
    If Len(Dir(App.Path & "\MMPC_PO.xlt")) <= 0 Then
        If EXTRACT_FILES(106, "MMPC_PO.xlt") = False Then
            MsgBox "PO Excel file cannot be located. Please add PO Template file in DMIS 2.0 Program Folder.", vbInformation, "PMIS"
            Exit Sub
        End If
    End If

    Dim vPOCONTACT                                     As String
    Dim vPOSUPPLIER_ADDRESS                            As String

    Dim vPOOrder_Date                                  As String
    Dim vPOORDER_NO                                    As String
    Dim vPOVEHICLE                                     As String

    Dim vPOORDER_TYPE                                  As String
    Dim vPODEALER_CODE                                 As String
    Dim vPODEALER_NAME                                 As String
    Dim vPODEALER_ADDRESS                              As String

    Dim vPOLINE                                        As String
    Dim vPOPART                                        As String
    Dim vPOPART_NAME                                   As String
    Dim vPOQTY                                         As String
    Dim vPOAMOUNT                                      As String
    Dim vPOTOTAL_ORDER                                 As String
    Dim vPOVIN                                         As String

    Dim vPOCounter                                     As Integer
    Dim rsPO                                           As ADODB.Recordset
    Dim rsPODetail                                     As ADODB.Recordset
    Dim TOTAL_AMT                                      As Double
    Dim TOTAL_QTY                                      As Integer
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\MMPC_PO.XLT")
    Set xlSheet = xlBook.Worksheets(1)


    Set rsPO = New ADODB.Recordset
    Set rsPO = gconDMIS.Execute("Select * from PMIS_PO_HD where TYPE = 'P' AND PONO = '" & XXX & "'")
    If Not rsPO.EOF And Not rsPO.BOF Then
        vPOCONTACT = Null2String(rsPO!contactcode)
        vPOSUPPLIER_ADDRESS = Null2String(rsPO!sup_addrs)

        vPOOrder_Date = Null2String(rsPO!PODATE)
        vPOORDER_NO = Null2String(rsPO!DON)
        vPOVEHICLE = Null2String(rsPO!MODELCODE)

        vPOORDER_TYPE = Null2String(rsPO!ORDERTYPE)
        vPODEALER_CODE = DEALER_CODE
        vPODEALER_NAME = COMPANY_NAME
        vPODEALER_ADDRESS = COMPANY_ADDRESS

        If vPOORDER_TYPE = "A" Then
            vPOORDER_TYPE = "Advance Purchase Order"
        ElseIf vPOORDER_TYPE = "R" Then
            vPOORDER_TYPE = "Regular Purchase Order"
        ElseIf vPOORDER_TYPE = "V" Then
            vPOORDER_TYPE = "Vehicle Off-Road Purchase Order"
        ElseIf vPOORDER_TYPE = "E" Then
            vPOORDER_TYPE = "Emergency Purchase Order"
        ElseIf vPOORDER_TYPE = "S" Then
            vPOORDER_TYPE = "Special Purchase Order"
        Else
            vPOORDER_TYPE = "Warranty Purchase Order"
        End If

'        xlSheet.Cells(4, "B") = vPOCONTACT
'        xlSheet.Cells(5, "B") = vPOSUPPLIER_ADDRESS
'        xlSheet.Cells(4, "F") = vPOOrder_Date
'        xlSheet.Cells(5, "F") = vPOORDER_NO
'        'xlSheet.Cells(6, "E") = "Tran. No."
'        'xlSheet.Cells(6, "F") = "***" & XXX & "***"
'        xlSheet.Cells(5, "H") = vPOVEHICLE
'        xlSheet.Cells(8, "A") = vPOORDER_TYPE
'        xlSheet.Cells(10, "B") = vPODEALER_CODE
'        xlSheet.Cells(11, "B") = vPODEALER_NAME
'        xlSheet.Cells(12, "B") = vPODEALER_ADDRESS
        Set rsPODetail = New ADODB.Recordset
        Set rsPODetail = gconDMIS.Execute("Select  * from PMIS_Tdaytran where TYPE = 'P' AND trantype = 'PO' and tranno = '" & XXX & "' order by itemno asc")
        If Not rsPODetail.EOF And Not rsPODetail.BOF Then
            rsPODetail.MoveFirst: vPOCounter = 0
            Do While Not rsPODetail.EOF

                vPOLINE = Format(Null2String(rsPODetail!itemno), "0000")
                vPOPART = Null2String(rsPODetail!STOCK_ORD)
                vPOPART_NAME = SetPartDesc(Null2String(rsPODetail!STOCK_ORD))
                vPOQTY = N2Str2Zero(rsPODetail!TRANQTY)
                vPOAMOUNT = N2Str2Zero(rsPODetail!TRANINVAMT)
                vPOTOTAL_ORDER = ToDoubleNumber(N2Str2Zero(rsPODetail!TRANQTY) * N2Str2Zero(rsPODetail!TRANINVAMT))
                vPOVIN = Null2String(rsPODetail!Vin)
'                xlSheet.Cells(16 + vPOCounter, "A") = vPOLINE
                xlSheet.Cells(12 + vPOCounter, "B") = vPOPART
                xlSheet.Cells(12 + vPOCounter, "E") = vPOPART_NAME
                xlSheet.Cells(12 + vPOCounter, "D") = vPOQTY
'                xlSheet.Cells(16 + vPOCounter, "E") = vPOAMOUNT
'                xlSheet.Cells(16 + vPOCounter, "F") = vPOTOTAL_ORDER
'                xlSheet.Cells(16 + vPOCounter, "G") = "F"
'                xlSheet.Cells(16 + vPOCounter, "H") = vPOVIN
                vPOCounter = vPOCounter + 1
                rsPODetail.MoveNext
                TOTAL_QTY = TOTAL_QTY + vPOQTY
                TOTAL_AMT = TOTAL_AMT + vPOTOTAL_ORDER
            Loop
           
        End If
         xlSheet.Cells(43, "E") = txtSIG_PreparedBy
         xlSheet.Cells(43, "f") = vPOOrder_Date
'        xlSheet.Cells(16 + vPOCounter, "C") = "TOTAL"
'        xlSheet.Cells(16 + vPOCounter, "C").Font.Bold = True
'        xlSheet.Cells(16 + vPOCounter, "C").HorizontalAlignment = Excel.Constants.xlRight
'        xlSheet.Cells(16 + vPOCounter, "D") = TOTAL_QTY
'        xlSheet.Cells(16 + vPOCounter, "D").Font.Bold = True
'        xlSheet.Cells(16 + vPOCounter, "F") = TOTAL_AMT
'        xlSheet.Cells(16 + vPOCounter, "F").Font.Bold = True
'        xlSheet.Cells(16 + vPOCounter, "F").Font.Underline = xlUnderlineStyleDouble

'        If COMPANY_CODE = "HPI" Then
'            'HPI dont want to have a signitories upon printing excel
'        Else
'            xlSheet.Range("A" & 16 + vPOCounter + 3, "B" & 16 + vPOCounter + 3).MergeCells = True
'            xlSheet.Cells(16 + vPOCounter + 3, "A") = txtSIG_PreparedBy
'            xlSheet.Cells(16 + vPOCounter + 3, "A").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Cells(16 + vPOCounter + 4, "A") = "PREPARED BY"
'            xlSheet.Cells(16 + vPOCounter + 4, "A").Font.Bold = True
'            xlSheet.Cells(16 + vPOCounter + 4, "A").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Range("A" & 16 + vPOCounter + 4, "B" & 16 + vPOCounter + 4).MergeCells = True
'
'            xlSheet.Range("F" & 16 + vPOCounter + 3, "G" & 16 + vPOCounter + 3).MergeCells = True
'            xlSheet.Cells(16 + vPOCounter + 3, "F") = txtSIG_NotedbyDesign
'            xlSheet.Cells(16 + vPOCounter + 3, "F").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Cells(16 + vPOCounter + 4, "F") = "SR. MNGR. OPERATION"
'            xlSheet.Cells(16 + vPOCounter + 4, "F").Font.Bold = True
'            xlSheet.Cells(16 + vPOCounter + 4, "F").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Range("F" & 16 + vPOCounter + 4, "G" & 16 + vPOCounter + 4).MergeCells = True
'
'            xlSheet.Range("C" & 16 + vPOCounter + 3, "D" & 16 + vPOCounter + 3).MergeCells = True
'            xlSheet.Cells(16 + vPOCounter + 3, "C") = txtSIG_Notedby
'            xlSheet.Cells(16 + vPOCounter + 3, "C").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Cells(16 + vPOCounter + 4, "C") = "NOTED BY"
'            xlSheet.Cells(16 + vPOCounter + 4, "C").Font.Bold = True
'            xlSheet.Cells(16 + vPOCounter + 4, "C").HorizontalAlignment = Excel.Constants.xlCenter
'            xlSheet.Range("C" & 16 + vPOCounter + 4, "D" & 16 + vPOCounter + 4).MergeCells = True
'
'        End If
        xlApp.Windows.Item(1).Caption = vPOORDER_NO


        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBYDESG", txtSIG_NotedbyDesign)

        xlApp.Visible = True
        DoEvents
        Set xlApp = Nothing
    End If
    Screen.MousePointer = 0
    Exit Sub
End Sub

Sub Send2FrontConfirm()
    Frame1.Enabled = False: Picture1.Enabled = False: fraDetails.Enabled = False: cmdAddTran.Enabled = False: fraAddTran.Enabled = False
End Sub

Sub Send2BackConfirm()
    Frame1.Enabled = False: Picture1.Enabled = True: fraDetails.Enabled = True: cmdAddTran.Enabled = False: fraAddTran.Enabled = True
End Sub

Sub SendToFrontConfirmPO()
    With frmPMISTrans_POConfirmation
        Screen.MousePointer = 11
        .txtPONo.Text = txtPONo.Text
        .txtPODate.Text = Format(txtPODate.Text, "DD-MMM-YY")
        DoEvents
        .txtDealerCode.Text = Left(txtDON.Text, 2)
        .txtConfirmDate.Text = Format(LOGDATE, "DD-MMM-YY")
        .txtSEQ_NO.Text = "00"
        .txtDealerName.Text = cboSupName.Text
        .txtSOType.Text = SetOrderType(Mid(txtDON.Text, 3, 1))
        .txtSOYear.Text = Mid(txtDON.Text, 4, 2)
        .txtSOMonth.Text = The_month(Mid(txtDON.Text, 6, 2))
        .txtSONum.Text = txtDON.Text
        .FillDetails (txtPONo.Text)
        Me.KeyPreview = False
        Screen.MousePointer = 0
        .Show 1
        Me.KeyPreview = True
    End With

End Sub

Sub SendToBackConfirmPO()
    Unload frmPMISTrans_POConfirmation
End Sub

Sub FindDupPOno(DDD As String)
    On Error Resume Next
    RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", Format(DDD, "000000")).Bookmark
    StoreMemVars

End Sub

Sub rsRefresh()
    Dim rsPo_det                                       As New ADODB.Recordset
    Dim qtybackOrder                                   As Integer
    Dim SQL                                            As String


    Set RSPO_HD = New ADODB.Recordset

    RSPO_HD.Open "select * from PMIS_PO_HD WHERE [TYPE] = 'P' order by pono asc", gconDMIS, adOpenForwardOnly, adLockReadOnly

    '        SQL = "(select 'CURRENT' AS T_S,* from PMIS_PO_HD WHERE [TYPE] = 'P'  " & vbCrLf
    '        SQL = SQL & "union ALL" & vbCrLf
    '        SQL = SQL & "SELECT 'HISTORY' AS T_S ,* FROM PMIS_PO_HIST WHERE [TYPE] = 'P' and supcode = 'H00001') order by pono asc"
    '
    '        RSPO_HD.Open SQL, gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtPONo.Text = ""
    Set rsPMIS_Counter = New ADODB.Recordset
    rsPMIS_Counter.Open "select modul,nextnumber from PMIS_Counter where [TYPE] = 'P' AND modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_Counter.EOF And Not rsPMIS_Counter.BOF Then
        txtPONo.Text = Format(N2Str2IntZero(rsPMIS_Counter!nextnumber), "000000")
    Else
        txtPONo.Text = "000001"
    End If
    
    'JJE Prefixes
'    If COMPANY_CODE = "DJM" Then       ** FOR APPROVAL **
'        txtPONo.Text = "PP" + txtPONo.Text
'    End If
    'JJE
    If COMPANY_CODE = "DJM" Then
        txtPONo.Locked = True
    End If
    chkUseHARIDNP.Value = 0
    txtPartID.Text = ""
    cboPP_No.Text = ""
    txtPODate.Text = LOGDATE
    txtSupCode.Text = ""
    txtPODate.Locked = True
    txtDON.Text = ""
    cmdDON.Enabled = False
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
        'LBL_T_S = RSPO_HD("T_S")
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
            cmdPrint.Enabled = True
            cmdCancelPO.Enabled = False
            If Trim(txtDON.Text) <> "" Then picConfirmation.Visible = True Else picConfirmation.Visible = False
            cmdUnPost.Enabled = True
        ElseIf Null2String(RSPO_HD!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "CANCELLED [" & Null2String(RSPO_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
            cmdCancelPO.Enabled = False
            cmdPrint.Enabled = False
            If Trim(txtDON.Text) = "" Then picConfirmation.Visible = False
        Else
            cmdCancelPO.Enabled = True
            cmdPrint.Enabled = False
            labPosted.Visible = False
            labPosted.Caption = ""
            cmdEdit.Enabled = True
            cmdPost.Enabled = True
            cmdUnPost.Enabled = False
            If Trim(txtDON.Text) = "" Then picConfirmation.Visible = False
            cmdCancelPO.Enabled = True
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
        .ColWidth(7) = 400
        .ColWidth(8) = 2100

        .Row = 0
        .Col = 1: .Text = "Item"
        .Col = 2: .Text = "Part Number"
        .Col = 3: .Text = "Description"
        .Col = 4: .Text = "Qty"
        .Col = 5: .Text = "Amount"
        .Col = 6: .Text = "Total Order"
        .Col = 7: .Text = "F/K"
        .Col = 8: .Text = "VIN"
    End With
End Sub

Sub FillDetails()
    Pcnt = 0: PO_TOTUCOST = 0: PO_TOTINVAMT = 0: PO_TOTVAT = 0: PO_T_ONORDER = 0: PO_TOTQTY = 0
    Dim rsPo_det                                       As New ADODB.Recordset

    Dim Fill_Kill                                      As String
    Dim SQL                                            As String
    Dim back_ord                                       As Integer
    Dim SQLTXT                                         As String



    Set rsPMIS_Tdaytran = New ADODB.Recordset
    rsPMIS_Tdaytran.Open "select id,tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,PO_FILL,PO_KILL,VIN from PMIS_Tdaytran where [TYPE] = 'P' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'SQL = "select id,tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,PO_FILL,PO_KILL,VIN from PMIS_Tdaytran where [TYPE] = 'P' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO'" & vbCrLf
    'SQL = SQL & "union ALL" & vbCrLf
    'SQL = SQL & "select id,tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,PO_FILL,PO_KILL,VIN from PMIS_daytran where [TYPE] = 'P' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO' and non_hari = 'N' order by itemno asc"
    'rsPMIS_Tdaytran.Open SQL, gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
        Screen.MousePointer = 11
        rsPMIS_Tdaytran.MoveFirst

        Do While Not rsPMIS_Tdaytran.EOF
            Pcnt = Pcnt + 1
            If Null2Bool(rsPMIS_Tdaytran!PO_FILL) = True Then
                Fill_Kill = "F"
            Else
                Fill_Kill = "K"
            End If

            grdDetails.AddItem rsPMIS_Tdaytran!ID & Chr(9) & Format(Null2String(rsPMIS_Tdaytran!itemno), "0000") & Chr(9) & _
                               Null2String(rsPMIS_Tdaytran!STOCK_ORD) & Chr(9) & _
                               SetPartDesc(Null2String(rsPMIS_Tdaytran!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(rsPMIS_Tdaytran!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_Tdaytran!TRANQTY) * N2Str2Zero(rsPMIS_Tdaytran!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Fill_Kill & Chr(9) & Null2String(rsPMIS_Tdaytran!Vin)
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
    optFILL.Value = True
    optKILL.Value = False
    chkUseHARIDNP.Value = 0
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1
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
    Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "select partno,partdesc from PMIS_Partmas where active = 'Y' ORDER BY PARTDESC ASC", gconDMIS
    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
        rsPMIS_Partmas.MoveFirst
        cboTranPartNo.Clear
        cboTranDescription.Clear
        Do While Not rsPMIS_Partmas.EOF
            cboTranPartNo.AddItem Null2String(rsPMIS_Partmas!PARTNO)
            cboTranDescription.AddItem Null2String(rsPMIS_Partmas!PARTDESC)
            rsPMIS_Partmas.MoveNext
        Loop
    End If
    FillCboContact
    FillCboModel
End Sub

Sub RefreshPartsCbo()
    Screen.MousePointer = 11
    Set rsPMIS_Partmas = New ADODB.Recordset
    rsPMIS_Partmas.Open "select partno,partdesc from PMIS_Partmas ORDER BY PARTDESC ASC", gconDMIS
    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
        rsPMIS_Partmas.MoveFirst
        cboTranPartNo.Clear
        cboTranDescription.Clear
        Do While Not rsPMIS_Partmas.EOF
            cboTranPartNo.AddItem Null2String(rsPMIS_Partmas!PARTNO)
            cboTranDescription.AddItem Null2String(rsPMIS_Partmas!PARTDESC)
            rsPMIS_Partmas.MoveNext
        Loop
    End If
    Screen.MousePointer = 0

    '  Screen.MousePointer = 11
    '    Set rsPMIS_Partmas = New ADODB.Recordset
    '    rsPMIS_Partmas.Open "select partnumber,descriptio from PMIS_DNPP ORDER BY descriptio ASC", gconDMIS
    '    If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
    '        rsPMIS_Partmas.MoveFirst
    '        cboTranPartNo.Clear
    '        cboTranDescription.Clear
    '        Do While Not rsPMIS_Partmas.EOF
    '            cboTranPartNo.AddItem Null2String(rsPMIS_Partmas!partnumber)
    '            cboTranDescription.AddItem Null2String(rsPMIS_Partmas!descriptio)
    '            rsPMIS_Partmas.MoveNext
    '        Loop
    '    End If
    '    Screen.MousePointer = 0

End Sub

Sub FillGrid()
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    lstPO_HD.Enabled = False
    Set RSPO_HD = New ADODB.Recordset
    Set RSPO_HD = gconDMIS.Execute("select pono,pono x from PMIS_PO_HD WHERE [TYPE] = 'P' order by pono asc")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        lstPO_HD.Enabled = True: Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Enabled = False
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    Set RSPO_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSPO_HD = gconDMIS.Execute("select ID, pono from PMIS_PO_HD where [TYPE] = 'P' AND pono like'" & XXX & "%'")
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
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HD WHERE [TYPE] = 'P' order by pono asc")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        lstPO_HD.Enabled = True: Listview_Loadval Me.lstPO_HD.ListItems, RSPO_HD: lstPO_HD.Refresh
    Else
        lstPO_HD.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    lstPO_HD.Enabled = False
    Set RSPO_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HD where [TYPE] = 'P' AND supname like '" & XXX & "%' order by pono asc")
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

    'Dim RSSUPER As ADODB.Recordset
    'Set RSSUPER = gconDMIS.Execute("SELECT * FROM PMIS_DNPP WHERE PARTNUMBER=" & N2Str2Null(cboTranPartNo))

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
            If DON_TYPE = "V" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
                'EAP:032309 add description to all DON_TYPE condition when hari dnp has been checked
                'cboTranDescription = rsDNPP!DESCRIPTIO
                '
            End If
            If DON_TYPE = "S" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
                'cboTranDescription = rsDNPP!DESCRIPTIO
            End If
            If DON_TYPE = "R" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
                'cboTranDescription = rsDNPP!DESCRIPTIO
            End If
            '*******************************************************************
            'updating code:     jaa - 11172008      - include Warranty
            If DON_TYPE = "W" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP2)
                'cboTranDescription = rsDNPP!DESCRIPTIO
            End If
            '*******************************************************************
            If DON_TYPE = "A" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP)
                'cboTranDescription = rsDNPP!DESCRIPTIO
            End If
            If DON_TYPE = "E" Then
                txtTranINVAmt.Text = N2Str2Zero(rsDNPP!DNPP3)
                'cboTranDescription = rsDNPP!DESCRIPTIO
            End If

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
    If Picture1.Visible = True Then
        SendToBack
        cmdAddTran.ZOrder 0
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
On Error GoTo ErrorCode
    If Function_Access(LOGID, "Acess_CancelEntry", "PURCHASE ORDER") = False Then Exit Sub

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MessagePop InfoVoid, "Action void", "Transaction already posted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
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
        RSPO_HD.Find "id =" & labID.Caption
        StoreMemVars
    End If
    Exit Sub
ErrorCode:
    ShowVBError
End Sub


Function Cancel() As Boolean
On Error GoTo errordaa

    Dim rsPMIS_TdaytranDup, rsPMIS_PartmasDup          As ADODB.Recordset
    Dim PCurOnOrder, PCurTpoQty                        As Integer

    SQL_STATEMENT = "update PMIS_PO_HD set" & _
                  " status = 'C'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "C", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PO NO: " & txtPONo, "", ""

    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_Tdaytran where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        Do While Not rsPMIS_TdaytranDup.EOF
            Set rsPMIS_PartmasDup = New ADODB.Recordset
            rsPMIS_PartmasDup.Open "select partno,onorder,tpoqty,ordered,emergency_po from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), gconDMIS
            If Not rsPMIS_PartmasDup.EOF And Not rsPMIS_PartmasDup.BOF Then
                PCurOnOrder = N2Str2IntZero(rsPMIS_PartmasDup!ONORDER) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                PCurTpoQty = N2Str2IntZero(rsPMIS_PartmasDup!tpoqty) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                If Null2String(rsPMIS_TdaytranDup!Status) = "P" Then
                    SQL_STATEMENT = "update PMIS_Partmas set" & _
                                  " purchases = " & N2Str2Zero(rsPMIS_PartmasDup!purchases) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                  " onorder = " & PCurOnOrder & "," & _
                                  " ORDERED = " & N2Str2IntZero(rsPMIS_PartmasDup!Ordered) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                  " tpoqty = " & PCurTpoQty & "," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    gconDMIS.Execute SQL_STATEMENT

                    Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "PO NO: " & txtPONo & " CANCEL", "", "")

                    If Mid(txtDON.Text, 3, 1) = "E" Then
                        gconDMIS.Execute "update PMIS_Partmas set" & _
                                       " EMERGENCY_PO = " & N2Str2IntZero(rsPMIS_PartmasDup!emergency_po) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & _
                                       " where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT

                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "PO NO: " & txtPONo & " CANCEL EMERGENCY", "", "")
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
                  " where tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO' and [TYPE] = 'P'"
    gconDMIS.Execute SQL_STATEMENT
    Call NEW_LogAudit("C", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PO NO: " & txtPONo, "PO", "")

    Set rsPMIS_TdaytranDup = Nothing
    Set rsPMIS_PartmasDup = Nothing
    
    Cancel = True
    Exit Function
errordaa:
    error_msg = error
    Cancel = False
End Function

Private Sub cmdDON_Click()
    With frmPMISDONFormation
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
    frmPMISDONFormation.Show 1
    On Error Resume Next
    cboModelCode.SetFocus
End Sub

'Private Sub cmdEditTranDate_Click()
'
'If Function_Access(LOGID, "Acess_SYSTEM", "PURCHASE ORDER") = False Then Exit Sub
'        txtPODate.Enabled = True
'        txtPODate.Locked = False
'End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "PURCHASE ORDER") = False Then Exit Sub

    On Error GoTo ErrorCode:
    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MessagePop InfoVoid, "Action void", "Transaction already posted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
        Exit Sub
    End If
        
    'updating code: JAA - 06272008     'Do not allow posting of transaction without issuance of Parts
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD = "" Or FILD = "No Entry" Then
        MsgBox "Posting of Transaction cannot proceed. Pls. Add Part(s).", vbCritical, "Confirm Posting "
        Exit Sub
    End If
    '====================================================================================================
    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        'updated by: IEBV 11172011
        'description: to rollback transaction if error occured
        gconDMIS.BeginTrans
        If POST = False Then
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

Function POST() As Boolean
On Error GoTo errordaa

    Set rsPMIS_Tdaytran = New ADODB.Recordset
    rsPMIS_Tdaytran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_Tdaytran where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
    If Not rsPMIS_Tdaytran.EOF And Not rsPMIS_Tdaytran.BOF Then
        rsPMIS_Tdaytran.MoveFirst
        Do While Not rsPMIS_Tdaytran.EOF
            Set rsPMIS_Partmas = New ADODB.Recordset
            rsPMIS_Partmas.Open "Select partno,onhand,tpoqty,onorder,ordered,emergency_po,purchases from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_Tdaytran!STOCK_ORD), gconDMIS
            If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.BOF Then
                SQL_STATEMENT = "update PMIS_Partmas set " & _
                              " purchases = " & N2Str2Zero(rsPMIS_Partmas!purchases) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                              " tpoqty = " & N2Str2Zero(rsPMIS_Partmas!tpoqty) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                              " ONORDER = " & N2Str2Zero(rsPMIS_Partmas!ONORDER) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & "," & _
                              " ORDERED = " & N2Str2Zero(rsPMIS_Partmas!Ordered) + NumericVal(rsPMIS_Tdaytran!TRANQTY) & _
                              " where partno = " & N2Str2Null(rsPMIS_Partmas!PARTNO)
                gconDMIS.Execute SQL_STATEMENT
                Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_Partmas!PARTNO), "STOCKNO", "PMIS_PARTMAS"), "P", "PO NO: " & txtPONo & " POSTED", "", "")

                SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                              " status = 'P'" & "," & _
                              " usercode = " & N2Str2Null(LOGCODE) & "," & _
                              " lastupdate = '" & LOGDATE & "'" & _
                              " where id = " & rsPMIS_Tdaytran!ID
                gconDMIS.Execute SQL_STATEMENT
                Call NEW_LogAudit("PP", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "RR NO: " & txtPONo, "PO", "")

                If Mid(txtDON.Text, 3, 1) = "E" Then
                    SQL_STATEMENT = "update PMIS_Partmas set" & _
                                  " EMERGENCY_PO = " & N2Str2Zero(rsPMIS_Partmas!emergency_po) + N2Str2Zero(rsPMIS_Tdaytran!TRANQTY) & _
                                  " where partno = " & N2Str2Null(rsPMIS_Tdaytran!STOCK_ORD)
                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_Partmas!PARTNO), "STOCKNO", "PMIS_PARTMAS"), "", "PO NO: " & txtPONo & " EMERGENCY POSTED", "", "")
                End If
                'updated by: IEBV 092820121200pm
                'description: to set partnumber as active if not active
                Call TAGIFINACTIVE(Null2String(rsPMIS_Partmas!PARTNO), "P")
                '-----------------------------------------------------
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
    Call NEW_LogAudit("P", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PO NO: " & txtPONo, "PO", "")

    
    Set rsPMIS_Tdaytran = Nothing
    Set rsPMIS_Partmas = Nothing
    
    POST = True
    Exit Function
errordaa:
    error_msg = error
    POST = False
End Function

Function PO_EXISTS(PO_NO As String) As Boolean
    Dim rsTMP As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT * FROM" & vbCrLf
    SQL = SQL & "(" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_RR_HD" & vbCrLf
    SQL = SQL & "Union All" & vbCrLf
    SQL = SQL & "SELECT RRNO,[TYPE], ISNULL(PONO,'000000') AS PONO,STATUS FROM PMIS_REC_HIST" & vbCrLf
    SQL = SQL & ") T WHERE PONO = '" & PO_NO & "' AND TYPE = 'P' AND STATUS = 'P'" & vbCrLf

    Set rsTMP = gconDMIS.Execute(SQL)
    
    If Not (rsTMP.EOF And rsTMP.BOF) Then
        PO_EXISTS = True
    Else
        PO_EXISTS = False
    End If
    
    Set rsTMP = Nothing
End Function

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "PURCHASE ORDER") = False Then Exit Sub
    If MsgQuestionBox("PO Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then

        Screen.MousePointer = 11
'        If txtDON.Text = "" Then
        If txtSupCode <> "M00001" Then
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
                
'                If COMPANY_CODE = "HCI" Then
'                    rptPurchaseOrder.Formulas(2) = "ContactPerson = '" & cboContactCode & "'"
'                End If

                PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "' and {Tdaytran.trantype} = 'PO'", DMIS_REPORT_Connection, 1
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
                
'                If COMPANY_CODE = "HCI" Then
'                    rptPurchaseOrder.Formulas(2) = "ContactPerson = '" & cboContactCode & "'"
'                End If
                
                PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_nonvat.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
            End If
        Else
'            If MsgBox("Current Purchase Order is For HARI, Would you like to View PO in Excel Template?", vbQuestion + vbYesNo, "Select PO Format") = vbYes Then
'                If COMPANY_CODE = "HPI" Then
'                    Call PrintPOExcel(txtPONo.Text)
'                Else
'                    txtSIG_PreparedBy = GetSetting("PMIS", "SIGNATORIES", "PO-PREPBY", "")
'                    txtSIG_Notedby = GetSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", "")
'                    txtSIG_NotedbyDesign = GetSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", "")
'                    cmdSaveSig.Enabled = False
'                    picPrintPOExcel.Visible = True
'                    picPrintPOExcel.ZOrder 0
'                    Command3.Enabled = True
'                    txtowner.Visible = False
'                End If
            'Else
                If NumericVal(txtDS1.Text) > 0 Then
                    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    
                     If COMPANY_CODE = "HMH" Then
                        rptPurchaseOrder.Formulas(3) = "PreparedBy = '" & GetSignitories("PreparedBy", "PMIS") & "'"
                        rptPurchaseOrder.Formulas(4) = "CheckedBy = '" & GetSignitories("CheckedBy", "PMIS") & "'"
                        rptPurchaseOrder.Formulas(5) = "ApprovedBy = '" & GetSignitories("ApprovedBy", "PMIS") & "'"
                    
                    End If
                    
                    If COMPANY_CODE = "HCI" Then
                        rptPurchaseOrder.Formulas(2) = "ContactPerson = '" & cboContactCode & "'"
                    End If
                    
                    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                Else
                    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    
                    If COMPANY_CODE = "HCI" Then
                        rptPurchaseOrder.Formulas(2) = "ContactPerson = '" & cboContactCode & "'"
                    End If
                    
                     If COMPANY_CODE = "HMH" Then
                        rptPurchaseOrder.Formulas(3) = "PreparedBy = '" & GetSignitories("PreparedBy", "PMIS") & "'"
                        rptPurchaseOrder.Formulas(4) = "CheckedBy = '" & GetSignitories("CheckedBy", "PMIS") & "'"
                        rptPurchaseOrder.Formulas(5) = "ApprovedBy = '" & GetSignitories("ApprovedBy", "PMIS") & "'"
                    
                    End If
                    
                    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "POnonvat.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                End If
            'End If
        End If
        Screen.MousePointer = 0
        Call NEW_LogAudit("V", "PURCHASE ORDER", "", labID, "Parts", "PO NO: " & txtPONo, "", "")
    End If
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
Private Sub cmdSaveSig_Click()

    Call SaveSetting("PMIS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", txtSIG_NotedbyDesign)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-OWNER", txtowner)

    picPrintPOExcel.Visible = False

End Sub

Private Sub cmdTranCancel_Click()
    Picture1.Enabled = True
    fraDetails.Enabled = True
    SendToBack
    StoreMemVars
    Frame2.Enabled = True
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo ErrorCode:

    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MessagePop InfoVoid, "Action void", "Transaction already posted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
        Exit Sub
    End If

    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        SQL_STATEMENT = "delete from PMIS_Tdaytran where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "XX", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PART NO: " & cboTranPartNo, "PO", labDetID
    End If

    Dim CNT                                            As Integer
    Dim rsPMIS_TdaytranDup                             As ADODB.Recordset
    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select id,itemno from PMIS_Tdaytran where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        CNT = 0
        Do While Not rsPMIS_TdaytranDup.EOF
            CNT = CNT + 1
            gconDMIS.Execute "update PMIS_Tdaytran set itemno = '" & Format(CNT, "0000") & "' where id = " & rsPMIS_TdaytranDup!ID
            rsPMIS_TdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        SQL_STATEMENT = "update PMIS_PO_HD set" & _
                      " po_amount = " & PO_TOTUCOST & "," & _
                      " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                      " ds_desc1 = NULL," & _
                      " ds_amt1 = 0" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    Else
        PO_TOTVAT = NumericVal(txtDS_Amt1.Text)
        SQL_STATEMENT = "update PMIS_PO_HD set" & _
                      " po_amount = " & PO_TOTUCOST & "," & _
                      " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                      " ds_desc1 = '" & "VAT" & "'," & _
                      " ds_amt1 = " & PO_TOTVAT & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    End If
    Call NEW_LogAudit("E", "PURCHASE ORDER", SQL_STATEMENT, labID, "", "PO NO: " & txtPONo, "", "")

    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labID.Caption
    cmdTranCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo ErrorCode
    Dim CRITICAL_QUESTION                              As String
    Dim CTR                                            As Integer
    Dim rsactive                                       As ADODB.Recordset
    Dim Msg                                            As String
    Dim sqlcommand                                     As String
    Dim POTRANDATE, POTRANNO, POTRANTYPE               As String
    Dim POITEMNO, POSTOCK_ORD, POSTOCK_SUP             As String
    Dim POTRANQTY                                      As Integer
    Dim POTRANUCOST                                    As Double
    Dim POSTATUS                                       As String
    Dim POTRANINVAMT                                   As Double
    Dim POTRANVIN                                      As String
    Dim PO_FILL, PO_KILL                               As String
    
    If cboTranDescription.Text = "" Then
        MsgBox "Part number must have description!", vbCritical + vbOKOnly
        On Error Resume Next
        cboTranDescription.SetFocus
        Exit Sub
    End If
    
    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MessagePop InfoVoid, "Action void", "Transaction already posted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
        Exit Sub
    End If
'updated by: IEBV 02012011_0400pm
'description:  saves the part number that dont exist on Master file
'----------------------------------------------------------------------------------------------------------------
    CTR = 0
    CTR = (gconDMIS.Execute("SELECT COUNT(*) FROM PMIS_STOCKMAS WHERE STOCKNO = " & N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))) & " AND [TYPE] = 'P' ").Fields(0).Value)
    If CTR > 0 Then
        'do nothing
    Else
        Msg = "Part Number Doesn't Exist On Parts Master file." & vbCrLf
        Msg = Msg + "This Will Automatically Add To Parts Master File"
        If MsgBox(Msg, vbQuestion + vbYesNo) = vbYes Then
            sqlcommand = "Insert into PMIS_stockmas ([TYPE], STOCKNO,STOCKDESC,STOCKTYPE,USERCODE,LASTUPDATE,ACTIVE) "
            sqlcommand = sqlcommand + " VALUES ('P'," & N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))) & ",'" & UCase(cboTranDescription.Text) & "', "
            sqlcommand = sqlcommand + " 'GJ', '" & Null2String(RSPO_HD!USERCODE) & "','" & LOGDATE & "', 'N') "
            gconDMIS.Execute sqlcommand
        Else
            Exit Sub
        End If
    End If
'----------------------------------------------------------------------------------------------------------------
    
    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsPMIS_TdaytranClone                       As ADODB.Recordset
        Set rsPMIS_TdaytranClone = New ADODB.Recordset
        rsPMIS_TdaytranClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_Tdaytran where [TYPE] = 'P' AND STOCK_ORD = " & N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))) & " and trantype = 'PO' and tranno =" & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
        If Not rsPMIS_TdaytranClone.EOF And Not rsPMIS_TdaytranClone.BOF Then
            MsgSpeechBox "Warning: Part Number already used in this transaction"
            On Error Resume Next
            cboTranPartNo.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If


    POSTOCK_ORD = UCase(N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))))
    POSTOCK_SUP = UCase(N2Str2Null(LTrim(RTrim(cboTranPartNo.Text))))
    POTRANDATE = N2Date2Null(txtPODate.Text)
    POTRANTYPE = "'" & "PO" & "'"
    POTRANNO = N2Str2Null(txtPONo.Text)
    POITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    POTRANQTY = NumericVal(txtTranQty.Text)
    POTRANINVAMT = NumericVal(txtTranINVAmt.Text)
    POTRANUCOST = NumericVal(txtUnitCost.Text)
    POSTATUS = "'N'"
    POTRANVIN = N2Str2Null(txtVIN.Text)
    If optFILL.Value = True Then
        PO_FILL = 1
    Else
        PO_FILL = 0
    End If
    If optKILL.Value = True Then
        PO_KILL = 1
    Else
        PO_KILL = 0
    End If

    If POTRANINVAMT <= 0 Then
        If MsgBox("Warning: Invoice Amount Is zero! Do You Want to Continue", vbInformation + vbYesNo) = vbNo Then
            On Error Resume Next
            txtTranINVAmt.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        CRITICAL_QUESTION = "Warning: Invoice Amount Is zero! Do You Want to Continue"
        Call NEW_LogAudit("MP", "PURCHASE ORDER", CRITICAL_QUESTION, labID, "", "PO NO: " & txtPONo & " PART NO: " & cboTranPartNo & " " & CRITICAL_QUESTION, "", "")
        MsgBox "User action has been log in the Audit Trail", vbInformation, "Audit Trail Information"
    End If

    If AddorEdit = "ADD" Then
        SQL_STATEMENT = "insert into PMIS_Tdaytran " & _
                        "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt,PO_FILL,PO_KILL,VIN,lastupdate,usercode,status)" & _
                      " values ('P'," & POTRANDATE & ", " & POTRANTYPE & ", " & POTRANNO & "," & _
                      " " & POITEMNO & "," & POSTOCK_ORD & "," & _
                      " " & POSTOCK_SUP & ", " & POTRANQTY & "," & _
                      " " & POTRANUCOST & ", " & POTRANINVAMT & "," & PO_FILL & "," & PO_KILL & "," & POTRANVIN & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & POSTATUS & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "AA", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PART NO: " & cboTranPartNo, "PO", ""
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
                      " PO_FILL = " & PO_FILL & "," & _
                      " PO_KILL = " & PO_KILL & "," & _
                      " VIN = " & POTRANVIN & "," & _
                      " lastupdate = '" & LOGDATE & "'," & _
                      " status = " & POSTATUS & "," & _
                      " usercode = " & N2Str2Null(LOGCODE) & "" & _
                      " where id = " & labDetID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "EE", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PART NO: " & cboTranPartNo, "PO", labDetID
    End If

    Dim rsPMIS_PartmasClone                            As ADODB.Recordset
    Set rsPMIS_PartmasClone = New ADODB.Recordset
    rsPMIS_PartmasClone.Open "select partno,tpoqty,onorder,mac,dnp,srp,onhand from PMIS_Partmas where partno = " & POSTOCK_ORD, gconDMIS
    If Not rsPMIS_PartmasClone.EOF And Not rsPMIS_PartmasClone.BOF Then
    Else
        If Len(POSTOCK_ORD) > 11 Then
            If txtSupCode.Text = VPAMCOR Then
                SQL_STATEMENT = "insert into PMIS_Partmas " & _
                                "(partno,partdesc,date_entered)" & _
                              " values (" & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
                gconDMIS.Execute SQL_STATEMENT
            Else
           
                SQL_STATEMENT = "insert into PMIS_Partmas " & _
                                "(partno,partdesc,date_entered)" & _
                              " values (" & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
                gconDMIS.Execute SQL_STATEMENT
            End If
            Call NEW_LogAudit("A", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(cboTranPartNo), "PARTNO", "PMIS_PARTMAS"), "", "PO NO: " & txtPONo, "", "")
        End If
    End If
    cleargrid grdDetails
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        SQL_STATEMENT = "update PMIS_PO_HD set" & _
                      " totalqty = " & PO_TOTQTY & "," & _
                      " po_amount = " & PO_TOTUCOST & "," & _
                      " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                      " ds_desc1 = NULL," & _
                      " ds_amt1 = 0" & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    Else
        PO_TOTVAT = NumericVal(txtDS_Amt1.Text)
        SQL_STATEMENT = "update PMIS_PO_HD set" & _
                      " totalqty = " & PO_TOTQTY & "," & _
                      " po_amount = " & PO_TOTUCOST & "," & _
                      " netpoamt = " & NumericVal(txtNetPOAmt.Text) & "," & _
                      " ds_desc1 = '" & "VAT" & "'," & _
                      " ds_amt1 = " & PO_TOTVAT & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
    End If
    Call NEW_LogAudit("E", "PURCHASE ORDER", SQL_STATEMENT, labID, "", "PO NO: " & txtPONo, "", "")


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

Private Sub cmdUnPost_Click()
    
    If Function_Access(LOGID, "Acess_UnPost", "PURCHASE ORDER") = False Then Exit Sub

    On Error GoTo ErrorCode:
    
    'Updated By: IEBV
    'description:   TO avoin uposting of PO if PO is already receive but not yet posted
    If chkfnotyet_posted(txtPONo.Text, "P") = True Then MessagePop InfoFriend, "Action Void", "You cannot Unpost this transaction, Its already received but not yet posted!": Exit Sub

    If chkstatus(txtPONo.Text, "P", "PO") = "N" Then
        MessagePop InfoVoid, "Action void", "Transaction already unposted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
        Exit Sub
    End If

    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        
        If PO_EXISTS(txtPONo) = True Then
            MessagePop InfoFriend, "Action Void", "You cannot Unpost this transaction, Its already Received!"
            Exit Sub
        End If
       'updated by: IEBV 11172011
       'description: to rollback transaction if error occured
        gconDMIS.BeginTrans
        If UNPOST = False Then
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

Function UNPOST() As Boolean
On Error GoTo errordaa
    SQL_STATEMENT = "update PMIS_PO_HD set" & _
                  " status = 'N'," & _
                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                  " lastupdate = '" & LOGDATE & "'" & _
                  " where id = " & labID.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "U", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PO NO: " & txtPONo, "PO", ""

    Dim rsPMIS_TdaytranDup, rsPMIS_PartmasDup      As ADODB.Recordset
    Dim PCurOnOrder, PCurTpoQty                    As Integer
    Set rsPMIS_TdaytranDup = New ADODB.Recordset
    rsPMIS_TdaytranDup.Open "select ID,Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_Tdaytran where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
    If Not rsPMIS_TdaytranDup.EOF And Not rsPMIS_TdaytranDup.BOF Then
        rsPMIS_TdaytranDup.MoveFirst
        Do While Not rsPMIS_TdaytranDup.EOF
            Set rsPMIS_PartmasDup = New ADODB.Recordset
            rsPMIS_PartmasDup.Open "select partno,onorder,tpoqty,ordered,emergency_po,purchases from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), gconDMIS
            If Not rsPMIS_PartmasDup.EOF And Not rsPMIS_PartmasDup.BOF Then
                PCurOnOrder = N2Str2IntZero(rsPMIS_PartmasDup!ONORDER) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                PCurTpoQty = N2Str2IntZero(rsPMIS_PartmasDup!tpoqty) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY)
                If Null2String(rsPMIS_TdaytranDup!Status) = "P" Then
                    SQL_STATEMENT = "update PMIS_Partmas set" & _
                                  " purchases = " & N2Str2Zero(rsPMIS_PartmasDup!purchases) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                  " onorder = " & PCurOnOrder & "," & _
                                  " tpoqty = " & PCurTpoQty & "," & _
                                  " ORDERED = " & N2Str2IntZero(rsPMIS_PartmasDup!Ordered) - NumericVal(rsPMIS_TdaytranDup!TRANQTY) & "," & _
                                  " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                  " lastupdate = '" & LOGDATE & "'" & _
                                  " where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                    gconDMIS.Execute SQL_STATEMENT
                    Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "PO NO: " & txtPONo & " UNPOSTED", "", "")

                    If Mid(txtDON.Text, 3, 1) = "E" Then
                        SQL_STATEMENT = "update PMIS_Partmas set" & _
                                      " EMERGENCY_PO = " & N2Str2IntZero(rsPMIS_PartmasDup!emergency_po) - N2Str2Zero(rsPMIS_TdaytranDup!TRANQTY) & _
                                      " where partno = " & N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD)
                        gconDMIS.Execute SQL_STATEMENT
                        Call NEW_LogAudit("E", "PART MASTER FILE", SQL_STATEMENT, FindTransactionID(N2Str2Null(rsPMIS_TdaytranDup!STOCK_ORD), "STOCKNO", "PMIS_PARTMAS"), "P", "PO NO: " & txtPONo & " UNPOSTED EMERGENCY", "", "")
                    End If
                End If
            End If
            SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                          " status = 'N'," & _
                          " usercode = " & N2Str2Null(LOGCODE) & "," & _
                          " lastupdate = '" & LOGDATE & "'" & _
                          " where ID = " & rsPMIS_TdaytranDup!ID
            gconDMIS.Execute SQL_STATEMENT
            Call NEW_LogAudit("UU", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", "PO NO: " & txtPONo, "PO", "")

            rsPMIS_TdaytranDup.MoveNext
        Loop
    End If

    Set rsPMIS_TdaytranDup = Nothing
    Set rsPMIS_PartmasDup = Nothing
    UNPOST = True
    Exit Function
errordaa:
    error_msg = error
    UNPOST = False

End Function

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "ADD"
    PoValidation
    initMemvars
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    Frame2.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "PURCHASE ORDER") = False Then Exit Sub
    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MessagePop InfoVoid, "Action void", "Transaction already posted!"
        Exit Sub
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
        MessagePop InfoVoid, "Action void", "Transaction already cancelled!"
        Exit Sub
    End If
    AddorEdit = "EDIT"
    PrevPONO = Format(txtPONo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Frame2.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    cmdDON.Enabled = True
    txtDON.Enabled = True
    txtPODate.Locked = True
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

    'axp02232008
    'JJE
'    If COMPANY_CODE <> "DJM" Then  ** FOR APPROVAL **
        If Len(Trim(RTrim(txtPONo))) <> 6 Then
            MsgBox "Invalid Transaction Number. Should Be Six Digit In Length!", vbCritical, "Transaction Number!"
            On Error Resume Next
            txtPONo.SetFocus
            Exit Sub
        End If
'    End If
    'JJE

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

    If Null2String(rsSupplier!SupCode) = "H00001" Then
        If txtDON.Text = "" Then
            MsgSpeechBox "Invalid Order Number!"
            Exit Sub
        End If
    End If

    If cboModelCode.Text = "" Then
        MsgBox "Vehicle model must not be empty!", vbCritical, "Purchase Order"
        cboModelCode.SetFocus
        Exit Sub
    End If

    'VALIDATION FOR TRANSACTION NUMBER
    If IsNull(txtPONo.Text) = True Then
        MsgSpeechBox "Warning: Purchase Order Number must not be empty"
        On Error Resume Next
        txtPONo.SetFocus
        Exit Sub
    Else
'updated bugs:  6032010 :
'****************************************************************************

        If AddorEdit = "ADD" Then
'            Dim rsPO_HDDup                                     As ADODB.Recordset
'            Dim SQL As String
'
'
'            Set rsPO_HDDup = New ADODB.Recordset
'            'rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
'
'            SQL = "SELECT * FROM" & vbCrLf
'            SQL = SQL & "(" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hd WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & "Union All" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hist WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & ") T WHERE PONO = '" & txtPONo.Text & "' AND TYPE = 'P'" & vbCrLf

'            rsPO_HDDup.Open (SQL), gconDMIS, adOpenForwardOnly, adLockReadOnly
            If checkdup_PO("P", txtPONo.Text) = True Then
                If MsgBox("Purchase Order Number already exist! Do you want to generate new PO number?", vbQuestion + vbYesNo) = vbYes Then
                    txtPONo.Text = getnextISSPORR("P", "PO")
                Else
                    On Error Resume Next
                    txtPONo.SetFocus
                    Exit Sub
                End If
            End If
            
         Else
            If Null2String(RSPO_HD!PONO) <> txtPONo.Text Then
                If checkdup_PO("P", txtPONo.Text) = True Then
                    MsgSpeechBox "Purchase Order Number already exist!"
                    On Error Resume Next
                    txtPONo.SetFocus
                    Exit Sub
                End If
            End If
'updated ivan:  6032010 :
'****************************************************************************

'        If ADDOREDIT = "ADD" Then
'            Set rsPO_HDDup = New ADODB.Recordset
'            rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
'            If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                MsgSpeechBox "Purchase Order Number already exist!"
'                On Error Resume Next
'                txtPONo.SetFocus
'                Exit Sub
'            End If
'        Else
'            If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
'                Set rsPO_HDDup = New ADODB.Recordset
'                rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
'                If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                    MsgSpeechBox "Purchase Order Number already exist!"
'                    On Error Resume Next
'                    txtPONo.SetFocus
'                    Exit Sub
'                End If
'            End If
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
                      " values ('P'," & VTXTPONo & ", " & VTXTPPNo & ", " & VTXTPODate & "," & VTXTDON & ", " & VTXTORDERTYPE & "," & VTXTORDER_SERIES & _
                        ", " & VTXTSupCode & ", " & VcboSupName & _
                        ", " & VTXTSup_Addrs & ", " & VTXTDealerCode & "," & VCBOContactCode & "," & VCBOModelCode & _
                        ", " & VTXTPO_Amount & _
                        ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                        ", " & VTXTNetPOAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "PURCHASE ORDER", SQL_STATEMENT, FindTransactionID(txtPONo, "pono", "PMIS_PO_HD", "DETAILS", N2Str2Null("P"), "TYPE"), "Parts", txtPONo & " - " & txtDON, "PO", ""

        NewPOPMIS_Counter = NumericVal(Right(txtPONo.Text, 6)) + 1
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
        NEW_LogAudit "E", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", txtPONo & " - " & txtDON, "", ""
        SQL_STATEMENT = "update PMIS_Tdaytran set" & _
                      " trandate = " & VTXTPODate & "," & _
                      " tranno = " & VTXTPONo & _
                      " where tranno = '" & Null2String(RSPO_HD!PONO) & "' and [TYPE] = 'P' and trantype = 'PO'"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "PURCHASE ORDER", SQL_STATEMENT, labID, "Parts", txtPONo & " - " & txtDON, "", ""
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NewPOPMIS_Counter & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where TYPE = 'P' and modul = 'PO'"
        Call FillGrid
    End If
    
    If AddorEdit = "EDIT" Then
        Dim SQLTXT As String
        
        SQLTXT = "UPDATE PMIS_PO_DETAILS SET SONum = '" & Null2String(txtDON) & "'"
        SQLTXT = SQLTXT & "WHERE PO_NO = '" & Null2String(txtPONo) & "'"
        Call gconDMIS.Execute(SQLTXT)
    
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

    Frame2.Enabled = True
    rsRefresh
    RSPO_HD.Find "id = " & labID.Caption
    cmdCancel.Value = True
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

Private Sub Command1_Click()
    Call RefreshPartsCbo
    VBComBoBoxDroppedDown cboTranPartNo
End Sub


Private Sub Command3_Click()
    Call PrintPOExcel(txtPONo.Text)

    Call SaveSetting("PMIS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", txtSIG_NotedbyDesign)
    Call SaveSetting("PMIS", "SIGNATORIES", "PO-OWNER", txtowner)

End Sub

Private Sub Command4_Click()
    picPrintPOExcel.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text

    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub

            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PURCHASE ORDER)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "PURCHASE ORDER", "")

        Case vbKeyF2

            If COMPANY_CODE = "HPI" Then
                txtSIG_PreparedBy = GetSetting("PMIS", "SIGNATORIES", "PO-PREPBY", "")
                txtSIG_Notedby = GetSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", "")
                txtSIG_NotedbyDesign = GetSetting("PMIS", "SIGNATORIES", "PO-APPROVEDBY", "")
                txtowner = GetSetting("PMIS", "SIGNATORIES", "PO-OWNER", "")

                Command3.Enabled = False
                picPrintPOExcel.Visible = True
                picPrintPOExcel.ZOrder 0
                Label11.Caption = "APPROVED BY:"
            End If

        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            SendToBackConfirmPO
            Frame2.Enabled = True
            Picture1.Enabled = True
            fraDetails.Enabled = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change"
                ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
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
                    If chkstatus(txtPONo.Text, "P", "PO") <> "P" And chkstatus(txtPONo.Text, "P", "PO") <> "C" Then
                        grdDetails_DblClick
                        Picture1.Enabled = False
                        fraDetails.Enabled = False
                    End If
                End If
            End If

        Case vbKeyF5
            If FILD <> "" And FILD <> "No Entry" Then
                If Picture1.Visible = True Then
                    If chkstatus(txtPONo.Text, "P", "PO") <> "P" And chkstatus(txtPONo.Text, "P", "PO") <> "C" Then
                        grdDetails_DblClick
                        cmdTranDelete_Click
                    End If
                End If
            End If
        Case vbKeyF9
            If Trim(txtDON.Text) = "" Then Exit Sub
            If picConfirmation.Visible = True Then SendToFrontConfirmPO
        Case vbKeyF12
            If cmdUnPost.Enabled = True Then cmdUnPost.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    
    If COMPANY_CODE = "HCI" Then
        txtwvat.Visible = True
    Else
        txtwvat.Visible = False
    End If
    'JJE Invoice disable for editing
'    If COMPANY_CODE = "DJM" Then       ** FOR APPROVAL **
'        txtPONo.MaxLength = 8
'        txtPONo.Enabled = False
'    End If
    'JJE
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBackConfirmPO: cmdAddTran.Enabled = False: picConfirmation.Visible = False
    Picture1.Visible = True: SendToBack
    Picture2.Visible = False: textSearch.Text = "": initMemvars: rsRefresh
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then RSPO_HD.MoveLast
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISTrans_Purchase = Nothing
    UnloadForm Me
End Sub



Private Sub grdDetails_DblClick()
    If chkstatus(txtPONo.Text, "P", "PO") = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf chkstatus(txtPONo.Text, "P", "PO") = "C" Then
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
            fraAddTran.Caption = "Edit Parts"
            StorePartsEntry (FILD)
        Else
            MsgSpeechBox "No Entry on Parts"
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

Private Sub grdDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD <> "" And FILD <> "No Entry" Then

    If Button = vbRightButton Then
        menuhist.Visible = True
        menumaster.Visible = True
        PopupMenu cmdmenu
    End If
    End If
End Sub

Private Sub grdDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = False
End Sub

Private Sub Label36_Click()
    If Trim(txtDON.Text) = "" Then Exit Sub
    If picConfirmation.Visible = True Then SendToFrontConfirmPO
End Sub

Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = True
End Sub

Private Sub Label36_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = True
End Sub

Private Sub menuhist_Click()
   If Module_Access(LOGID, "PARTS COMPUTERIZED STOCKCARDS", "INQUIRY") = False Then Exit Sub
    Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 2
    FILD = grdDetails.Text

    Unload frmPMISInquiry_Query
    PARTSQUERY = 1

    frmPMISInquiry_Query.SetTYPE ("P")
    fromParts = True
    FormExistsShow frmPMISInquiry_Query
    frmPMISInquiry_Query.txt_Ledger_Search.Text = FILD
    frmPMISInquiry_Query.frommaster_SHOWLEDGER (FILD)
End Sub

Private Sub menumaster_Click()
    If Module_Access(LOGID, "PARTS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 2
    FILD = grdDetails.Text
    
    frmMasterFile_Parts.SETSTOCKTYPE ("P")
    FormExistsShow frmMasterFile_Parts
    frmMasterFile_Parts.textSearch.Text = FILD
    Call frmMasterFile_Parts.SearchStock(FILD, "P")
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = False
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
'updated bugs:  6032010 :
'****************************************************************************
'If AddorEdit = "ADD" Then
'            Dim rsPO_HDDup                                     As ADODB.Recordset
'            Dim SQL As String
'
'
'            Set rsPO_HDDup = New ADODB.Recordset
'            'rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
'
'            SQL = "SELECT * FROM" & vbCrLf
'            SQL = SQL & "(" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hd WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & "Union All" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hist WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & ") T WHERE PONO = '" & txtPONo.Text & "' AND TYPE = 'P'" & vbCrLf
'
'            rsPO_HDDup.Open (SQL), gconDMIS, adOpenForwardOnly, adLockReadOnly
'            If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'
'                MsgSpeechBox "Purchase Order Number already exist!"
'                On Error Resume Next
'                txtPONo.SetFocus
'                Exit Sub
'            End If
'
'Else
'            Set rsPO_HDDup = New ADODB.Recordset
'            'rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
'
'            SQL = "SELECT * FROM" & vbCrLf
'            SQL = SQL & "(" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hd WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & "Union All" & vbCrLf
'            SQL = SQL & "select PONO,[TYPE] from pmis_po_hist WHERE TYPE = 'P'" & vbCrLf
'            SQL = SQL & ") T WHERE PONO = '" & txtPONo.Text & "' AND TYPE = 'P'" & vbCrLf
'
'            rsPO_HDDup.Open (SQL), gconDMIS, adOpenForwardOnly, adLockReadOnly
'            If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'
'                MsgSpeechBox "Purchase Order Number already exist!"
'                On Error Resume Next
'                txtPONo.SetFocus
'                Exit Sub
'            End If
'End If
'****************************************************************************
'updated bugs:  6032010 :

'    If Frame1.Enabled = True Then
'        If Len(txtPONo.Text) >= 3 Then
'            Dim rsPO_HDDup                             As ADODB.Recordset
'            If ADDOREDIT = "ADD" Then
'                Set rsPO_HDDup = New ADODB.Recordset
'                rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS
'                If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                    MsgSpeechBox "PO Number Already Exist!"
'                    On Error Resume Next
'                    txtPONo.SetFocus
'                End If
'            ElseIf ADDOREDIT = "EDIT" Then
'                If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
'                    Set rsPO_HDDup = New ADODB.Recordset
'                    rsPO_HDDup.Open "select pono from PMIS_PO_HD where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS
'                    If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
'                        MsgSpeechBox "PO Number Already Exist!"
'                        '  On Error Resume Next
'                        '  txtPONo.SetFocus
'                    End If
'                End If
'            End If
'        End If
'    End If
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
    RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
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

Public Sub PoValidation()
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    Frame2.Enabled = False
End Sub

Private Sub txtwvat_Change()
     If COMPANY_CODE = "HCI" Then
        If NumericVal(txtwvat) <> 0 Then
            txtUnitCost.Text = Format(NumericVal(txtwvat) / 1.12, "#,###,##0.00")
        End If
    End If
End Sub

Sub click()
    RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub
