VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISInquiry_Purchase_Hist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Order Entry"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   DrawWidth       =   10
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Purchase_Hist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   11775
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2280
      ScaleHeight     =   870
      ScaleWidth      =   9405
      TabIndex        =   68
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
         MouseIcon       =   "Purchase_Hist.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   71
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
         MouseIcon       =   "Purchase_Hist.frx":07C2
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":0914
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelPO 
         Caption         =   "Cancel Transaction"
         Enabled         =   0   'False
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
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":0C7A
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
         Enabled         =   0   'False
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
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":0FB4
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Transaction"
         Enabled         =   0   'False
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
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":12F9
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   4680
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":161E
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         Left            =   3900
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":197A
         Style           =   1  'Graphical
         TabIndex        =   74
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
         MouseIcon       =   "Purchase_Hist.frx":1C8D
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":1DDF
         Style           =   1  'Graphical
         TabIndex        =   70
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
         MouseIcon       =   "Purchase_Hist.frx":212F
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":2281
         Style           =   1  'Graphical
         TabIndex        =   69
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
         MouseIcon       =   "Purchase_Hist.frx":25DF
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":2731
         Style           =   1  'Graphical
         TabIndex        =   75
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
         MouseIcon       =   "Purchase_Hist.frx":2A2B
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":2B7D
         Style           =   1  'Graphical
         TabIndex        =   76
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
         MouseIcon       =   "Purchase_Hist.frx":2ED5
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":3027
         Style           =   1  'Graphical
         TabIndex        =   77
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
      TabIndex        =   81
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
         Left            =   690
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":3386
         Style           =   1  'Graphical
         TabIndex        =   82
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
         Left            =   -90
         MousePointer    =   99  'Custom
         Picture         =   "Purchase_Hist.frx":36C4
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox picConfirmation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2220
      ScaleHeight     =   285
      ScaleWidth      =   9435
      TabIndex        =   88
      Top             =   7230
      Width           =   9465
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "F9 - View/Update PO Upon Confirmation"
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
         Left            =   0
         MousePointer    =   99  'Custom
         TabIndex        =   89
         Top             =   30
         Visible         =   0   'False
         Width           =   9285
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
      TabIndex        =   63
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
         TabIndex        =   94
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
         TabIndex        =   67
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label25 
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
         Left            =   3360
         TabIndex        =   66
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label26 
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
         Left            =   1740
         TabIndex        =   65
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label27 
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
         Left            =   120
         TabIndex        =   64
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
      Height          =   7545
      Left            =   60
      TabIndex        =   57
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstPO_HD 
         Height          =   6105
         Left            =   60
         TabIndex        =   61
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
         MouseIcon       =   "Purchase_Hist.frx":3A14
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
         TabIndex        =   62
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2985
      Left            =   2220
      TabIndex        =   26
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton cmdEditTranDate 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   4350
         TabIndex        =   54
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
         ItemData        =   "Purchase_Hist.frx":3B76
         Left            =   1680
         List            =   "Purchase_Hist.frx":3B78
         TabIndex        =   5
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
         Left            =   5640
         TabIndex        =   7
         Text            =   "16A070101"
         ToolTipText     =   "Type Reference PIS Number"
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmdDON 
         Caption         =   "..."
         Height          =   375
         Left            =   7110
         TabIndex        =   85
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
         Width           =   2925
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   7770
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
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   49
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
         TabIndex        =   12
         Text            =   "Purchase_Hist.frx":3B7A
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
         Left            =   930
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Type the supplier code (e.g. 00001)"
         Top             =   660
         Width           =   1155
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
         Width           =   4515
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
         Left            =   1200
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
         ScaleWidth      =   4575
         TabIndex        =   31
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
            TabIndex        =   4
            Top             =   30
            Width           =   4515
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1185
         Left            =   6480
         ScaleHeight     =   1185
         ScaleWidth      =   2925
         TabIndex        =   32
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
            TabIndex        =   53
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
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   52
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
            Left            =   1500
            MaxLength       =   15
            TabIndex        =   51
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
            Left            =   240
            TabIndex        =   48
            Top             =   810
            Width           =   1245
         End
         Begin VB.Label Label9 
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
            Left            =   240
            TabIndex        =   47
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
         TabIndex        =   87
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
         Left            =   4710
         TabIndex        =   86
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
         TabIndex        =   84
         Top             =   2640
         Width           =   1965
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   60
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
         TabIndex        =   50
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
         TabIndex        =   44
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
         TabIndex        =   46
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
         Height          =   315
         Left            =   7290
         TabIndex        =   45
         Top             =   180
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4650
         X2              =   4650
         Y1              =   150
         Y2              =   3060
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
         TabIndex        =   43
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
         Left            =   90
         TabIndex        =   30
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
         Left            =   2250
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   1050
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3045
      Left            =   2220
      TabIndex        =   91
      Top             =   2910
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2805
         Left            =   60
         TabIndex        =   92
         Top             =   150
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   4948
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
      Height          =   4485
      Left            =   3780
      TabIndex        =   33
      Top             =   1440
      Width           =   4575
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
         Left            =   2850
         TabIndex        =   21
         Top             =   2310
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
         Left            =   2850
         TabIndex        =   20
         Top             =   1980
         Value           =   -1  'True
         Width           =   1485
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   3780
         TabIndex        =   95
         Top             =   630
         Width           =   285
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
         TabIndex        =   93
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
         TabIndex        =   22
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
         TabIndex        =   41
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
         Picture         =   "Purchase_Hist.frx":3B94
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Save Entry"
         Top             =   3510
         Width           =   915
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
         TabIndex        =   90
         Top             =   3120
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
         TabIndex        =   34
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
         Left            =   420
         TabIndex        =   42
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
         TabIndex        =   40
         Top             =   3660
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   1650
         Width           =   795
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Left            =   600
         TabIndex        =   37
         Top             =   630
         Width           =   825
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
         TabIndex        =   36
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
         TabIndex        =   35
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
      TabIndex        =   56
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
      TabIndex        =   55
      Top             =   7320
      Width           =   135
   End
End
Attribute VB_Name = "frmPMISInquiry_Purchase_Hist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSPO_HD, rsPMIS_PP_HD, rsPMIS_DAYTRAN              As ADODB.Recordset
Attribute rsPMIS_PP_HD.VB_VarUserMemId = 1073938432
Attribute rsPMIS_DAYTRAN.VB_VarUserMemId = 1073938432
Dim rsPMIS_Partmas, rsSupplier                         As ADODB.Recordset
Attribute rsPMIS_Partmas.VB_VarUserMemId = 1073938435
Attribute rsSupplier.VB_VarUserMemId = 1073938435
Dim rsALL_Profile, rsPMIS_Counter                      As ADODB.Recordset
Attribute rsALL_Profile.VB_VarUserMemId = 1073938437
Attribute rsPMIS_Counter.VB_VarUserMemId = 1073938437
Dim Pcnt                                               As Integer
Attribute Pcnt.VB_VarUserMemId = 1073938439
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
Attribute DON_TYPE.VB_VarUserMemId = 1073938435

Dim xlApp                                              As Excel.Application
Attribute xlApp.VB_VarUserMemId = 1073938436
Dim xlBook                                             As Excel.Workbook
Attribute xlBook.VB_VarUserMemId = 1073938437
Dim xlSheet                                            As Excel.Worksheet
Attribute xlSheet.VB_VarUserMemId = 1073938438

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
    Set rsPMIS_DAYTRAN = New ADODB.Recordset
    rsPMIS_DAYTRAN.Open "select * from PMIS_DAYTRAN where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_DAYTRAN.EOF And Not rsPMIS_DAYTRAN.BOF Then
        labDetID.Caption = rsPMIS_DAYTRAN!ID
        txtTranItemNo.Text = Format(Null2String(rsPMIS_DAYTRAN!itemno), "0000")
        cboTranPartNo.Text = Null2String(rsPMIS_DAYTRAN!STOCK_ORD)
        cboTranDescription.Text = SetPartDesc(Null2String(rsPMIS_DAYTRAN!STOCK_SUP))
        txtTranQty.Text = N2Str2IntZero(rsPMIS_DAYTRAN!TRANQTY)
        PrevTranQty = N2Str2IntZero(rsPMIS_DAYTRAN!TRANQTY)
        txtTranINVAmt.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_DAYTRAN!TRANINVAMT))
        txtUnitCost.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_DAYTRAN!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(rsPMIS_DAYTRAN!TRANQTY) * N2Str2Zero(rsPMIS_DAYTRAN!TRANINVAMT))
        optFILL = Null2Bool(rsPMIS_DAYTRAN!PO_FILL)
        optKILL = Null2Bool(rsPMIS_DAYTRAN!PO_KILL)
        txtVIN.Text = Null2String(rsPMIS_DAYTRAN!Vin)
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

    If Len(Dir(App.Path & "\PO.xlt")) <= 0 Then
        If EXTRACT_FILES(106, "PO.xlt") = False Then
            MsgBox "Please Put PO.xlt on " & vbCrLf & App.Path, vbInformation
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
    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\PO.xlt")
    'Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\NEW PO.XLS")
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

        xlSheet.Cells(4, "B") = vPOCONTACT
        xlSheet.Cells(5, "B") = vPOSUPPLIER_ADDRESS
        xlSheet.Cells(4, "F") = vPOOrder_Date
        xlSheet.Cells(5, "F") = vPOORDER_NO
        'xlSheet.Cells(6, "E") = "Tran. No."
        'xlSheet.Cells(6, "F") = "***" & XXX & "***"
        xlSheet.Cells(5, "H") = vPOVEHICLE
        xlSheet.Cells(8, "A") = vPOORDER_TYPE
        xlSheet.Cells(10, "B") = vPODEALER_CODE
        xlSheet.Cells(11, "B") = vPODEALER_NAME
        xlSheet.Cells(12, "B") = vPODEALER_ADDRESS
        Set rsPODetail = New ADODB.Recordset
        Set rsPODetail = gconDMIS.Execute("Select  * from PMIS_Tdaytran where TYPE = 'P' AND trantype = 'PO' and tranno = '" & XXX & "' order by itemno asc")
        Dim CNT                                        As Integer
        CNT = 16

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

                xlSheet.Cells(16 + vPOCounter, "A") = vPOLINE
                xlSheet.Cells(16 + vPOCounter, "B") = vPOPART
                xlSheet.Cells(16 + vPOCounter, "C") = vPOPART_NAME
                xlSheet.Cells(16 + vPOCounter, "D") = vPOQTY
                xlSheet.Cells(16 + vPOCounter, "E") = vPOAMOUNT
                xlSheet.Cells(16 + vPOCounter, "F") = vPOTOTAL_ORDER
                xlSheet.Cells(16 + vPOCounter, "G") = "F"
                xlSheet.Cells(16 + vPOCounter, "H") = vPOVIN
                vPOCounter = vPOCounter + 1
                rsPODetail.MoveNext
                TOTAL_QTY = TOTAL_QTY + vPOQTY
                TOTAL_AMT = TOTAL_AMT + vPOTOTAL_ORDER

                CNT = CNT + 1
                If CNT > 31 Then
                    vPOCounter = 0
                    xlApp.Visible = True
                    Set xlApp = Nothing

                    Set xlApp = CreateObject("Excel.Application")
                    Set xlBook = xlApp.Workbooks.Open(PMIS_REPORT_PATH & "\PO.xlt")
                    Set xlSheet = xlBook.Worksheets(1)

                    xlSheet.Cells(4, "B") = vPOCONTACT
                    xlSheet.Cells(5, "B") = vPOSUPPLIER_ADDRESS
                    xlSheet.Cells(4, "F") = vPOOrder_Date
                    xlSheet.Cells(5, "F") = vPOORDER_NO
                    'xlSheet.Cells(6, "E") = "Tran. No."
                    'xlSheet.Cells(6, "F") = "***" & XXX & "***"
                    xlSheet.Cells(5, "H") = vPOVEHICLE
                    xlSheet.Cells(8, "B") = vPOORDER_TYPE
                    xlSheet.Cells(10, "B") = vPODEALER_CODE
                    xlSheet.Cells(11, "B") = vPODEALER_NAME
                    xlSheet.Cells(12, "B") = vPODEALER_ADDRESS
                    CNT = 16
                End If
            Loop
        End If

        'xlSheet.Cells(17 + vPOCounter, "F") = TOTAL_AMT
        'xlSheet.Cells(17 + vPOCounter, "D") = TOTAL_QTY
        'xlSheet.Cells(20 + vPOCounter, "A") = "PREPARED BY"

        'xlSheet.Cells(36, "A") = txtSIG_PreparedBy
        'xlSheet.Cells(36, "C") = txtSIG_Notedby
        'xlSheet.Cells(21 + vPOCounter, "E") = txtSIG_Notedby
        'xlSheet.Cells(36, "E") = txtSIG_NotedbyDesign

        xlApp.Windows.Item(1).Caption = vPOORDER_NO


        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
        'Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBYDESG", txtSIG_NotedbyDesign)

        xlApp.Visible = True
        Set xlApp = Nothing
    End If



    Exit Sub
    '    Set rsPO = New ADODB.Recordset
    '    Set rsPO = gconDMIS.Execute("Select * from PMIS_PO_HD where TYPE = 'P' AND PONO = '" & XXX & "'")
    '    If Not rsPO.EOF And Not rsPO.BOF Then
    '        vPOCONTACT = Null2String(rsPO!contactcode)
    '        vPOSUPPLIER_ADDRESS = Null2String(rsPO!sup_addrs)
    '
    '        vPOOrder_Date = Null2String(rsPO!PODATE)
    '        vPOORDER_NO = Null2String(rsPO!DON)
    '        vPOVEHICLE = Null2String(rsPO!modelcode)
    '
    '        vPOORDER_TYPE = Null2String(rsPO!ORDERTYPE)
    '        vPODEALER_CODE = DEALER_CODE
    '        vPODEALER_NAME = COMPANY_NAME
    '        vPODEALER_ADDRESS = COMPANY_ADDRESS
    '
    '        If vPOORDER_TYPE = "A" Then
    '            vPOORDER_TYPE = "Advance Purchase Order"
    '        ElseIf vPOORDER_TYPE = "R" Then
    '            vPOORDER_TYPE = "Regular Purchase Order"
    '        ElseIf vPOORDER_TYPE = "V" Then
    '            vPOORDER_TYPE = "Vehicle Off-Road Purchase Order"
    '        ElseIf vPOORDER_TYPE = "E" Then
    '            vPOORDER_TYPE = "Emergency Purchase Order"
    '        ElseIf vPOORDER_TYPE = "S" Then
    '            vPOORDER_TYPE = "Special Purchase Order"
    '        Else
    '            vPOORDER_TYPE = "Warranty Purchase Order"
    '        End If
    '
    '        xlSheet.Cells(4, 2) = vPOCONTACT
    '        xlSheet.Cells(5, 2) = vPOSUPPLIER_ADDRESS
    '        xlSheet.Cells(4, 6) = vPOOrder_Date
    '        xlSheet.Cells(5, 6) = vPOORDER_NO
    '        xlSheet.Cells(6, 5) = "Tran. No."
    '        xlSheet.Cells(6, 6) = "***" & XXX & "***"
    '        xlSheet.Cells(5, 7) = vPOVEHICLE
    '        xlSheet.Cells(8, 2) = vPOORDER_TYPE
    '        xlSheet.Cells(10, 2) = vPODEALER_CODE
    '        xlSheet.Cells(11, 2) = vPODEALER_NAME
    '        xlSheet.Cells(12, 2) = vPODEALER_ADDRESS
    '        Set rsPODetail = New ADODB.Recordset
    '        Set rsPODetail = gconDMIS.Execute("Select  * from PMIS_Tdaytran where TYPE = 'P' AND trantype = 'PO' and tranno = '" & XXX & "' order by itemno asc")
    '
    '
    '        If Not rsPODetail.EOF And Not rsPODetail.BOF Then
    '            rsPODetail.MoveFirst: vPOCounter = 0
    '            Dim iExcel                  As Integer
    '            For iExcel = 16 To 44
    '                xlSheet.Cells(iExcel, 1) = ""
    '                xlSheet.Cells(iExcel, 2) = ""
    '                xlSheet.Cells(iExcel, 3) = ""
    '                xlSheet.Cells(iExcel, 4) = ""
    '                xlSheet.Cells(iExcel, 5) = ""
    '                xlSheet.Cells(iExcel, 6) = ""
    '                xlSheet.Cells(iExcel, 7) = ""
    '                xlSheet.Cells(iExcel, 8) = ""
    '            Next
    '            Do While Not rsPODetail.EOF
    '                vPOLINE = Null2String(rsPODetail!itemno)
    '                vPOPART = Null2String(rsPODetail!STOCK_ORD)
    '                vPOPART_NAME = SetPartDesc(Null2String(rsPODetail!STOCK_ORD))
    '                vPOQTY = N2Str2Zero(rsPODetail!tranqty)
    '                vPOAMOUNT = N2Str2Zero(rsPODetail!TRANINVAMT)
    '                vPOTOTAL_ORDER = ToDoubleNumber(N2Str2Zero(rsPODetail!tranqty) * N2Str2Zero(rsPODetail!TRANINVAMT))
    '                vPOVIN = Null2String(rsPODetail!vin)
    '
    '                xlSheet.Cells(16 + vPOCounter, 1) = vPOLINE
    '                xlSheet.Cells(16 + vPOCounter, 2) = vPOPART
    '                xlSheet.Cells(16 + vPOCounter, 3) = vPOPART_NAME
    '                xlSheet.Cells(16 + vPOCounter, 4) = vPOQTY
    '                xlSheet.Cells(16 + vPOCounter, 5) = vPOAMOUNT
    '                xlSheet.Cells(16 + vPOCounter, 6) = vPOTOTAL_ORDER
    '                xlSheet.Cells(16 + vPOCounter, 7) = "F"
    '                xlSheet.Cells(16 + vPOCounter, 8) = vPOVIN
    '                vPOCounter = vPOCounter + 1
    '                rsPODetail.MoveNext
    '                TOTAL_QTY = TOTAL_QTY + vPOQTY
    '                TOTAL_AMT = TOTAL_AMT + vPOTOTAL_ORDER
    '            Loop
    '        End If
    '
    '        xlSheet.Cells(17 + vPOCounter, "F") = TOTAL_AMT
    '        xlSheet.Cells(17 + vPOCounter, "D") = TOTAL_QTY
    '        xlSheet.Cells(20 + vPOCounter, "A") = "PREPARED BY"
    '
    '        xlSheet.Cells(22 + vPOCounter, "A") = txtSIG_PreparedBy
    '
    '        xlSheet.Cells(20 + vPOCounter, "D") = "NOTED BY"
    '
    '        xlSheet.Cells(21 + vPOCounter, "E") = txtSIG_Notedby
    '
    '        xlSheet.Cells(22 + vPOCounter, "E") = txtSIG_NotedbyDesign
    '
    '        xlApp.Windows.Item(1).Caption = vPOORDER_NO
    '
    '
    '    Call SaveSetting("PMIS", "SIGNATORIES", "PO-PREPBY", txtSIG_PreparedBy)
    '    Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBY", txtSIG_Notedby)
    '    Call SaveSetting("PMIS", "SIGNATORIES", "PO-NOTEDBYDESG", txtSIG_NotedbyDesign)
    '
    '    xlApp.Visible = True
    '        Set xlApp = Nothing
    '    End If
End Sub


Sub Send2FrontConfirm()
    Frame1.Enabled = False: Picture1.Enabled = False: fraDetails.Enabled = False: fraAddTran.Enabled = False
End Sub

Sub Send2BackConfirm()
    Frame1.Enabled = False: Picture1.Enabled = True: fraDetails.Enabled = True: fraAddTran.Enabled = True
End Sub

Sub SendToFrontConfirmPO()
    With frmPMISTrans_POConfirmation
        Screen.MousePointer = 11
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
    Set RSPO_HD = New ADODB.Recordset
    RSPO_HD.Open "select * from PMIS_PO_HIST WHERE [TYPE] = 'P' order by pono asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtPONo.Text = ""
    Set rsPMIS_Counter = New ADODB.Recordset
    rsPMIS_Counter.Open "select modul,nextnumber from PMIS_Counter where [TYPE] = 'P' AND modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_Counter.EOF And Not rsPMIS_Counter.BOF Then
        txtPONo.Text = Format(N2Str2IntZero(rsPMIS_Counter!nextnumber), "000000")
    End If
    chkUseHARIDNP.Value = 0
    txtPartID.Text = ""
    cboPP_No.Text = ""
    txtPODate.Text = LOGDATE
    txtSupCode.Text = ""

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
    txtremarks.Text = "Pls Type Your Message Here!"
    cleargrid grdDetails
    InitGrid
    InitCbo
    InitParts
End Sub

Sub StoreMemVars()
    DON_TYPE = ""
    If Not RSPO_HD.EOF And Not RSPO_HD.BOF Then
        labid.Caption = RSPO_HD!ID
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
        txtremarks.Text = Null2String(RSPO_HD!REMARKS)
        If Null2String(RSPO_HD!Status) = "P" Then
            labPosted.Visible = True
            labPosted.Caption = "POSTED [" & Null2String(RSPO_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
            cmdCancelPO.Enabled = False
            If Trim(txtDON.Text) <> "" Then picConfirmation.Visible = True Else picConfirmation.Visible = False
            cmdUnpost.Enabled = False
        ElseIf Null2String(RSPO_HD!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "CAN2CELLED [" & Null2String(RSPO_HD!USERCODE) & "]"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnpost.Enabled = False
            cmdCancelPO.Enabled = False
            cmdPrint.Enabled = False
            If Trim(txtDON.Text) = "" Then picConfirmation.Visible = False
        Else
            cmdCancelPO.Enabled = False
            cmdPrint.Enabled = False
            labPosted.Visible = False
            labPosted.Caption = ""
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnpost.Enabled = False
            If Trim(txtDON.Text) = "" Then picConfirmation.Visible = False
            cmdCancelPO.Enabled = False
        End If
        cleargrid grdDetails
        FillDetails
    Else
        ShowNoRecord
        cmdAdd.Value = False
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
    Pcnt = 0: PO_TOTUCOST = 0: PO_TOTINVAMT = 0: PO_TOTVAT = 0: PO_T_ONORDER = 0

    Dim Fill_Kill                                      As String
    Set rsPMIS_DAYTRAN = New ADODB.Recordset
    rsPMIS_DAYTRAN.Open "select id,tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost,PO_FILL,PO_KILL,VIN from PMIS_DAYTRAN where [TYPE] = 'P' AND tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPMIS_DAYTRAN.EOF And Not rsPMIS_DAYTRAN.BOF Then
        Screen.MousePointer = 11
        rsPMIS_DAYTRAN.MoveFirst
        Do While Not rsPMIS_DAYTRAN.EOF
            Pcnt = Pcnt + 1
            If Null2Bool(rsPMIS_DAYTRAN!PO_FILL) = True Then
                Fill_Kill = "F"
            Else
                Fill_Kill = "K"
            End If
            grdDetails.AddItem rsPMIS_DAYTRAN!ID & Chr(9) & Format(Null2String(rsPMIS_DAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(rsPMIS_DAYTRAN!STOCK_ORD) & Chr(9) & _
                               SetPartDesc(Null2String(rsPMIS_DAYTRAN!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(rsPMIS_DAYTRAN!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_DAYTRAN!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2Zero(rsPMIS_DAYTRAN!TRANQTY) * N2Str2Zero(rsPMIS_DAYTRAN!TRANUCOST), MAXIMUM_DIGIT) & Chr(9) & _
                               Fill_Kill & Chr(9) & Null2String(rsPMIS_DAYTRAN!Vin)
            PO_TOTUCOST = PO_TOTUCOST + (N2Str2IntZero(rsPMIS_DAYTRAN!TRANQTY) * N2Str2Zero(rsPMIS_DAYTRAN!TRANUCOST))
            PO_TOTINVAMT = PO_TOTINVAMT + (N2Str2IntZero(rsPMIS_DAYTRAN!TRANQTY) * N2Str2Zero(rsPMIS_DAYTRAN!TRANINVAMT))
            rsPMIS_DAYTRAN.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        If ISNONVAT = True Then PO_TOTVAT = 0 Else PO_TOTVAT = PO_TOTINVAMT - (PO_TOTINVAMT / ConvertToBIRDecimalFormat(VAT_RATE))
        PO_TOTUCOST = NumericVal(PO_TOTINVAMT - PO_TOTVAT)
        If NumericVal(PO_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtDS_Amt1.Text = ToDoubleNumber(PO_TOTVAT)
            txtNetPOAmt.Text = ToDoubleNumber(PO_TOTINVAMT)
        Else
            txtDS1.Text = ""
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = ZERO
            txtNetPOAmt.Text = ToDoubleNumber(NumericVal(PO_TOTINVAMT))
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
    fraAddTran.ZOrder 1
    fraAddTran.Enabled = False
    Send2BackConfirm
End Sub

Sub BringToFront()
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
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
End Sub

Sub FillGrid()
    Dim RSPO_HD                                        As ADODB.Recordset
    lstPO_HD.Sorted = False: lstPO_HD.ListItems.Clear
    lstPO_HD.Enabled = False
    Set RSPO_HD = New ADODB.Recordset
    Set RSPO_HD = gconDMIS.Execute("select pono from PMIS_PO_HIST WHERE [TYPE] = 'P' order by pono asc")
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
    Set RSPO_HD = gconDMIS.Execute("select pono, pono from PMIS_PO_HIST where [TYPE] = 'P' AND pono like'" & XXX & "%'")
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
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HIST WHERE [TYPE] = 'P' order by pono asc")
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
    Set RSPO_HD = gconDMIS.Execute("select supname, pono from PMIS_PO_HIST where [TYPE] = 'P' AND supname like '" & XXX & "%' order by pono asc")
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

    '  Dim RSSUPER As ADODB.Recordset
    ' Set RSSUPER = gconDMIS.Execute("SELECT * FROM PMIS_DNPP WHERE PARTNUMBER=" & N2Str2Null(cboTranPartNo))

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
    If Picture1.Visible = True Then
        SendToBack
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
    If Function_Access(LOGID, "Acess_CancelEntry", "PURCHASE ORDER") = False Then Exit Sub

    If LOGLEVEL <> "ADM" Then
        MsgBox "Warning: Your account is not allowed to cancel this transaction!", vbCritical, "Error"
        Exit Sub
    End If
    Dim rsPMIS_DAYTRANDup, rsPMIS_PartmasDup           As ADODB.Recordset
    Dim PCurOnOrder, PCurTpoQty                        As Integer
    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        Set rsPMIS_DAYTRANDup = New ADODB.Recordset
        rsPMIS_DAYTRANDup.Open "select Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_DAYTRAN where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
        If Not rsPMIS_DAYTRANDup.EOF And Not rsPMIS_DAYTRANDup.BOF Then
            rsPMIS_DAYTRANDup.MoveFirst
            Do While Not rsPMIS_DAYTRANDup.EOF
                Set rsPMIS_PartmasDup = New ADODB.Recordset
                rsPMIS_PartmasDup.Open "select partno,onorder,tpoqty,ordered,emergency_po from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD), gconDMIS
                If Not rsPMIS_PartmasDup.EOF And Not rsPMIS_PartmasDup.BOF Then
                    PCurOnOrder = N2Str2IntZero(rsPMIS_PartmasDup!ONORDER) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY)
                    PCurTpoQty = N2Str2IntZero(rsPMIS_PartmasDup!tpoqty) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY)
                    If Null2String(rsPMIS_DAYTRANDup!Status) = "P" Then
                        gconDMIS.Execute "update PMIS_Partmas set" & _
                                       " purchases = " & N2Str2Zero(rsPMIS_PartmasDup!purchases) - NumericVal(rsPMIS_DAYTRANDup!TRANQTY) & "," & _
                                       " onorder = " & PCurOnOrder & "," & _
                                       " ORDERED = " & N2Str2IntZero(rsPMIS_PartmasDup!Ordered) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY) & "," & _
                                       " tpoqty = " & PCurTpoQty & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD)
                        If Mid(txtDON.Text, 3, 1) = "E" Then
                            gconDMIS.Execute "update PMIS_Partmas set" & _
                                           " EMERGENCY_PO = " & N2Str2IntZero(rsPMIS_PartmasDup!emergency_po) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY) & _
                                           " where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD)
                        End If
                    End If
                End If
                rsPMIS_DAYTRANDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_DAYTRAN set" & _
                       " status = '" & "C" & "'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where tranno = " & N2Str2Null(RSPO_HD!PONO) & " and trantype = 'PO'"
        LogAudit "C", "PARTS PURCHASE ORDER", txtPONo
        rsRefresh
        On Error Resume Next
        RSPO_HD.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

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
        Set rsPMIS_DAYTRAN = New ADODB.Recordset
        rsPMIS_DAYTRAN.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_DAYTRAN where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
        If Not rsPMIS_DAYTRAN.EOF And Not rsPMIS_DAYTRAN.BOF Then
            'rsPMIS_DAYTRAN.MoveFirst
            '            Do While Not rsPMIS_DAYTRAN.EOF
            '                If N2Str2Zero(rsPMIS_DAYTRAN!TRANINVAMT) <= 0 Then
            '                    MsgSpeechBox "Warning: Transaction with Invoice Amount equal to Zero Encountered!"
            '                    Exit Sub
            '                End If
            '                rsPMIS_DAYTRAN.MoveNext
            '            Loop
            rsPMIS_DAYTRAN.MoveFirst
            Do While Not rsPMIS_DAYTRAN.EOF
                Set rsPMIS_Partmas = New ADODB.Recordset
                rsPMIS_Partmas.Open "Select partno,onhand,tpoqty,onorder,ordered,emergency_po,purchases from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_DAYTRAN!STOCK_ORD), gconDMIS
                If Not rsPMIS_Partmas.EOF And Not rsPMIS_Partmas.EOF Then
                    gconDMIS.Execute "update PMIS_Partmas set " & _
                                   " purchases = " & N2Str2Zero(rsPMIS_Partmas!purchases) + NumericVal(rsPMIS_DAYTRAN!TRANQTY) & "," & _
                                   " tpoqty = " & N2Str2Zero(rsPMIS_Partmas!tpoqty) + NumericVal(rsPMIS_DAYTRAN!TRANQTY) & "," & _
                                   " ONORDER = " & N2Str2Zero(rsPMIS_Partmas!ONORDER) + NumericVal(rsPMIS_DAYTRAN!TRANQTY) & "," & _
                                   " ORDERED = " & N2Str2Zero(rsPMIS_Partmas!Ordered) + NumericVal(rsPMIS_DAYTRAN!TRANQTY) & _
                                   " where partno = " & N2Str2Null(rsPMIS_Partmas!PARTNO)
                    gconDMIS.Execute "update PMIS_DAYTRAN set" & _
                                   " status = 'P'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsPMIS_DAYTRAN!ID
                    If Mid(txtDON.Text, 3, 1) = "E" Then
                        gconDMIS.Execute "update PMIS_Partmas set" & _
                                       " EMERGENCY_PO = " & N2Str2Zero(rsPMIS_Partmas!emergency_po) + N2Str2Zero(rsPMIS_DAYTRAN!TRANQTY) & _
                                       " where partno = " & N2Str2Null(rsPMIS_DAYTRAN!STOCK_ORD)
                    End If
                End If
                rsPMIS_DAYTRAN.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        rsRefresh
        LogAudit "P", "PARTS PURCHASE ORDER", txtPONo
        On Error Resume Next
        RSPO_HD.Find "id =" & labid.Caption
        StoreMemVars


    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()
    'If Function_Access(LOGID, "Acess_Print", "PURCHASE ORDER") = False Then Exit Sub

    If MsgQuestionBox("PO Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        If txtDON.Text = "" Then
            If NumericVal(txtDS1.Text) > 0 Then
                rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                
                 If COMPANY_CODE = "HCI" Then
                    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_Hist_Printing.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                 Else
                    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_Hist.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                 End If
            Else
                rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "POnonvat.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
            End If
        Else

            If MsgBox("Current Purchase Order is For HARI, Would you like to View PO in Excel Template?", vbQuestion + vbYesNo, "Select PO Format") = vbYes Then
                Call PrintPOExcel(txtPONo.Text)
            Else
                If NumericVal(txtDS1.Text) > 0 Then
                    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    
                    If COMPANY_CODE = "HCI" Then
                        PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_Hist_Printing.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "PO_Hist.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                    End If
                Else
                    rptPurchaseOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
                    rptPurchaseOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
                    PrintSQLReport rptPurchaseOrder, PMIS_REPORT_PATH & "POnonvat.RPT", "{Po_hd.type} = 'P' AND {Po_hd.pono} = '" & txtPONo.Text & "'", DMIS_REPORT_Connection, 1
                End If
            End If
        End If
        Screen.MousePointer = 0
        NEW_LogAudit "V", "PURCHASE ORDER - HISTORY", "", "", "Parts", txtPONo, "Purchase Order", ""
    End If
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemVars
    Frame2.Enabled = True
End Sub

Private Sub cmdTranDelete_Click()

    On Error GoTo ErrorCode:

    If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
        gconDMIS.Execute "delete from PMIS_DAYTRAN where id = " & labDetID.Caption
        LogAudit "X", "Purchase Order-Detail", cboTranDescription
    End If
    Dim CNT                                            As Integer
    Dim rsPMIS_DAYTRANDup                              As ADODB.Recordset
    Set rsPMIS_DAYTRANDup = New ADODB.Recordset
    rsPMIS_DAYTRANDup.Open "select id,itemno from PMIS_DAYTRAN where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
    If Not rsPMIS_DAYTRANDup.EOF And Not rsPMIS_DAYTRANDup.BOF Then
        rsPMIS_DAYTRANDup.MoveFirst
        CNT = 0
        Do While Not rsPMIS_DAYTRANDup.EOF
            CNT = CNT + 1
            gconDMIS.Execute "update PMIS_DAYTRAN set itemno = " & Format(CNT, "0000") & " where id = " & rsPMIS_DAYTRANDup!ID
            rsPMIS_DAYTRANDup.MoveNext
        Loop
    End If
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labid.Caption
    Else
        PO_TOTVAT = PO_TOTINVAMT - PO_TOTUCOST
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo ErrorCode

    If cboTranPartNo.Text = "" Then
        MsgSpeechBox "Warning: Part Number must have a value"
        On Error Resume Next
        cboTranPartNo.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsPMIS_DAYTRANClone                        As ADODB.Recordset
        Set rsPMIS_DAYTRANClone = New ADODB.Recordset
        rsPMIS_DAYTRANClone.Open "select trantype,tranno,itemno,STOCK_ORD from PMIS_DAYTRAN where [TYPE] = 'P' AND STOCK_ORD = '" & UCase(cboTranPartNo.Text) & "' and trantype = 'PO' and tranno =" & N2Str2Null(RSPO_HD!PONO) & " order by itemno asc", gconDMIS
        If Not rsPMIS_DAYTRANClone.EOF And Not rsPMIS_DAYTRANClone.BOF Then
            MsgSpeechBox "Warning: Part Number already used in this transaction"
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

    Dim PO_FILL, PO_KILL                               As String

    POTRANDATE = N2Date2Null(txtPODate.Text)
    POTRANTYPE = "'" & "PO" & "'"
    POTRANNO = N2Str2Null(txtPONo.Text)
    POITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    POSTOCK_ORD = UCase(N2Str2Null(cboTranPartNo.Text))
    POSTOCK_SUP = UCase(N2Str2Null(cboTranPartNo.Text))
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
    End If

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into PMIS_DAYTRAN " & _
                         "(TYPE,trandate,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt,PO_FILL,PO_KILL,VIN,lastupdate,usercode,status)" & _
                       " values ('P'," & POTRANDATE & ", " & POTRANTYPE & ", " & POTRANNO & "," & _
                       " " & POITEMNO & "," & POSTOCK_ORD & "," & _
                       " " & POSTOCK_SUP & ", " & POTRANQTY & "," & _
                       " " & POTRANUCOST & ", " & POTRANINVAMT & "," & PO_FILL & "," & PO_KILL & "," & POTRANVIN & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & POSTATUS & ")"
    Else
        gconDMIS.Execute "update PMIS_DAYTRAN set" & _
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
    End If

    Dim rsPMIS_PartmasClone                            As ADODB.Recordset
    Set rsPMIS_PartmasClone = New ADODB.Recordset
    rsPMIS_PartmasClone.Open "select partno,tpoqty,onorder,mac,dnp,srp,onhand from PMIS_Partmas where partno = " & POSTOCK_ORD, gconDMIS
    If Not rsPMIS_PartmasClone.EOF And Not rsPMIS_PartmasClone.BOF Then
    Else
        If Len(POSTOCK_ORD) > 11 Then
            If txtSupCode.Text = VPAMCOR Then
                gconDMIS.Execute "insert into PMIS_Partmas " & _
                                 "(partno,partdesc,date_entered)" & _
                               " values (" & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
            Else
                gconDMIS.Execute "insert into PMIS_Partmas " & _
                                 "(partno,partdesc,date_entered)" & _
                               " values (" & POSTOCK_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
            End If
        End If
    End If
    cleargrid grdDetails
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labid.Caption
    Else
        PO_TOTVAT = PO_TOTINVAMT - PO_TOTUCOST
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labid.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then
        cmdAddTran_Click
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

    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        Dim rsPMIS_DAYTRANDup, rsPMIS_PartmasDup       As ADODB.Recordset
        Dim PCurOnOrder, PCurTpoQty                    As Integer
        Set rsPMIS_DAYTRANDup = New ADODB.Recordset
        rsPMIS_DAYTRANDup.Open "select ID,Tranqty,STOCK_ORD,trantype,tranno,STATUS from PMIS_DAYTRAN where [TYPE] = 'P' AND trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HD!PONO), gconDMIS
        If Not rsPMIS_DAYTRANDup.EOF And Not rsPMIS_DAYTRANDup.BOF Then
            rsPMIS_DAYTRANDup.MoveFirst
            Do While Not rsPMIS_DAYTRANDup.EOF
                Set rsPMIS_PartmasDup = New ADODB.Recordset
                rsPMIS_PartmasDup.Open "select partno,onorder,tpoqty,ordered,emergency_po,purchases from PMIS_Partmas where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD), gconDMIS
                If Not rsPMIS_PartmasDup.EOF And Not rsPMIS_PartmasDup.BOF Then
                    PCurOnOrder = N2Str2IntZero(rsPMIS_PartmasDup!ONORDER) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY)
                    PCurTpoQty = N2Str2IntZero(rsPMIS_PartmasDup!tpoqty) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY)
                    If Null2String(rsPMIS_DAYTRANDup!Status) = "P" Then
                        gconDMIS.Execute "update PMIS_Partmas set" & _
                                       " purchases = " & N2Str2Zero(rsPMIS_PartmasDup!purchases) - NumericVal(rsPMIS_DAYTRANDup!TRANQTY) & "," & _
                                       " onorder = " & PCurOnOrder & "," & _
                                       " tpoqty = " & PCurTpoQty & "," & _
                                       " ORDERED = " & N2Str2IntZero(rsPMIS_PartmasDup!Ordered) - NumericVal(rsPMIS_DAYTRANDup!TRANQTY) & "," & _
                                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                       " lastupdate = '" & LOGDATE & "'" & _
                                       " where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD)
                        If Mid(txtDON.Text, 3, 1) = "E" Then
                            gconDMIS.Execute "update PMIS_Partmas set" & _
                                           " EMERGENCY_PO = " & N2Str2IntZero(rsPMIS_PartmasDup!emergency_po) - N2Str2Zero(rsPMIS_DAYTRANDup!TRANQTY) & _
                                           " where partno = " & N2Str2Null(rsPMIS_DAYTRANDup!STOCK_ORD)
                        End If
                    End If
                End If
                gconDMIS.Execute "update PMIS_DAYTRAN set" & _
                               " status = 'N'," & _
                               " usercode = " & N2Str2Null(LOGCODE) & "," & _
                               " lastupdate = '" & LOGDATE & "'" & _
                               " where ID = " & rsPMIS_DAYTRANDup!ID
                rsPMIS_DAYTRANDup.MoveNext
            Loop
        End If
        LogAudit "U", "PARTS PURCHASE ORDER", txtPONo
        rsRefresh
        RSPO_HD.Find "id =" & labid.Caption
        StoreMemVars


    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "PURCHASE ORDER") = False Then Exit Sub
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    Frame2.Enabled = False
    initMemvars
    On Error Resume Next
    txtPONo.SetFocus
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

    'axp02232008
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

    If Null2String(rsSupplier!SupCode) = "H00001" Then
        If txtDON.Text = "" Then
            MsgSpeechBox "Invalid Order Number!"
            Exit Sub
        End If
    End If

    'VALIDATION FOR TRANSACTION NUMBER
    If IsNull(txtPONo.Text) = True Then
        MsgSpeechBox "Warning: Purchase Order Number must not be empty"
        On Error Resume Next
        txtPONo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsPO_HDDup = New ADODB.Recordset
            rsPO_HDDup.Open "select pono from PMIS_PO_HIST where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
                MsgSpeechBox "Purchase Order Number already exist!"
                On Error Resume Next
                txtPONo.SetFocus
                Exit Sub
            End If
        Else
            If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
                Set rsPO_HDDup = New ADODB.Recordset
                rsPO_HDDup.Open "select pono from PMIS_PO_HIST where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
                    MsgSpeechBox "Purchase Order Number already exist!"
                    On Error Resume Next
                    txtPONo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

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
    If txtremarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = N2Str2Null(Trim(txtremarks.Text))
    End If

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into PMIS_PO_HIST" & _
                       " (TYPE,pono,ppno,podate,DON,ORDERTYPE,ORDER_SERIES,supcode,supname,sup_addrs,dealercode,ContactCode,ModelCode,po_amount,ds1,ds_desc1,ds_amt1,netpoamt,usercode,lastupdate,remarks)" & _
                       " values ('P'," & VTXTPONo & ", " & VTXTPPNo & ", " & VTXTPODate & "," & VTXTDON & ", " & VTXTORDERTYPE & "," & VTXTORDER_SERIES & _
                         ", " & VTXTSupCode & ", " & VcboSupName & _
                         ", " & VTXTSup_Addrs & ", " & VTXTDealerCode & "," & VCBOContactCode & "," & VCBOModelCode & _
                         ", " & VTXTPO_Amount & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetPOAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
        LogAudit "A", "PARTS PURCHASE ORDER", txtPONo
        NewPOPMIS_Counter = NumericVal(txtPONo.Text) + 1
    Else
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
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
                       " where id = " & labid.Caption

        gconDMIS.Execute "update PMIS_DAYTRAN set" & _
                       " trandate = " & VTXTPODate & "," & _
                       " tranno = " & VTXTPONo & _
                       " where tranno = '" & Null2String(RSPO_HD!PONO) & "'"
        LogAudit "E", "PARTS PURCHASE ORDER", PrevPONO
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NewPOPMIS_Counter & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where TYPE = 'P' and modul = 'PO'"
        Call FillGrid
    End If
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "pono = " & VTXTPONo
    cmdCancel.Value = True
    DoEvents
    On Error GoTo ErrorCode
    cleargrid grdDetails
    FillDetails
    If ISNONVAT = True Then
        PO_TOTVAT = 0
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = 0" & _
                       " where id = " & labid.Caption
    Else
        PO_TOTVAT = PO_TOTINVAMT - PO_TOTUCOST
        gconDMIS.Execute "update PMIS_PO_HIST set" & _
                       " po_amount = " & PO_TOTUCOST & "," & _
                       " netpoamt = " & PO_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & PO_TOTVAT & _
                       " where id = " & labid.Caption
    End If
    Frame2.Enabled = True
    rsRefresh
    On Error Resume Next
    RSPO_HD.Find "id = " & labid.Caption
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddTran_Click
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Command1_Click()
    Call RefreshPartsCbo
    VBComBoBoxDroppedDown cboTranPartNo
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Picture1.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (PURCHASE ORDER)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "PURCHASE ORDER")
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim FILD                                           As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
            SendToBackConfirmPO
            Frame2.Enabled = True
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSPO_HD!Status) = "P" Then
                    MsgSpeechBox "Transactions are Already Posted and cannot be Change"
                ElseIf Null2String(RSPO_HD!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                Else
                    cmdAddTran_Click
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
        Case vbKeyF9
            If Trim(txtDON.Text) = "" Then Exit Sub
            If picConfirmation.Visible = True Then SendToFrontConfirmPO
        Case vbKeyF12
            If cmdUnpost.Enabled = True Then cmdUnpost.Value = True
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBackConfirmPO: picConfirmation.Visible = False
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

Private Sub grdDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = False
End Sub

Private Sub Label36_Click()
    '    If Trim(txtDON.Text) = "" Then Exit Sub
    '    If picConfirmation.Visible = True Then SendToFrontConfirmPO
End Sub

Private Sub Label36_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = True
End Sub

Private Sub Label36_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label36.FontUnderline = True
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
    If Frame1.Enabled = True Then
        If Len(txtPONo.Text) >= 3 Then
            Dim rsPO_HDDup                             As ADODB.Recordset
            If AddorEdit = "ADD" Then
                Set rsPO_HDDup = New ADODB.Recordset
                rsPO_HDDup.Open "select pono from PMIS_PO_HIST where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS
                If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
                    MsgSpeechBox "PO Number Already Exist!"
                    On Error Resume Next
                    txtPONo.SetFocus
                End If
            ElseIf AddorEdit = "EDIT" Then
                If LTrim(RTrim(txtPONo)) <> Null2String(RSPO_HD!PONO) Then
                    Set rsPO_HDDup = New ADODB.Recordset
                    rsPO_HDDup.Open "select pono from PMIS_PO_HIST where [TYPE] = 'P' AND pono = '" & txtPONo.Text & "'", gconDMIS
                    If Not rsPO_HDDup.EOF And Not rsPO_HDDup.BOF Then
                        MsgSpeechBox "PO Number Already Exist!"
                        '  On Error Resume Next
                        '  txtPONo.SetFocus
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtRemarks_GotFocus()
    MsgSpeechBox "Pls Type Your Message Here!"
    If txtremarks.Text = "Pls Type Your Message Here!" Then txtremarks.Text = ""
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
    If optPONo.Value = True Then
        RSPO_HD.Find = "pono=" & lstPO_HD.SelectedItem.Text
    Else
        RSPO_HD.Bookmark = rsFind(RSPO_HD.Clone, "pono", lstPO_HD.SelectedItem.SubItems(1)).Bookmark
    End If
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

