VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCSMS_ACL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accumulated Claim List"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11430
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_ACL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   11430
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   2565
      TabIndex        =   44
      Top             =   4350
      Width           =   2595
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   78
         Top             =   0
         Width           =   2625
         _Version        =   655364
         _ExtentX        =   4630
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Shorcut"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - DISAPPROVED ACL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   20
         Left            =   60
         TabIndex        =   71
         Top             =   840
         Width           =   1890
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SHIFT F1 - VIEW AUDIT TRAIL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   14
         Left            =   60
         TabIndex        =   57
         Top             =   1110
         Width           =   2325
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F9 - APPROVE ACL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   12
         Left            =   60
         TabIndex        =   52
         Top             =   570
         Width           =   1425
      End
      Begin VB.Label lblCAP 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - ADD QIR CLAIMS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   45
         Top             =   330
         Width           =   1725
      End
   End
   Begin VB.PictureBox picDET 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   2640
      ScaleHeight     =   2745
      ScaleWidth      =   8715
      TabIndex        =   28
      Top             =   3030
      Width           =   8745
      Begin MSComctlLib.ListView lsvDET 
         Height          =   2235
         Left            =   30
         TabIndex        =   30
         Top             =   270
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3942
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "QIR no."
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Claim no."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RO No."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Vin No."
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblCAP 
         BackStyle       =   0  'Transparent
         Caption         =   "DOUBLE CLICK TO REMOVE DETAILS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   390
         Index           =   15
         Left            =   5910
         TabIndex        =   58
         Top             =   2520
         Width           =   2835
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   8775
         _Version        =   655364
         _ExtentX        =   15478
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ACL DETAILS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picSEARCH 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   30
      ScaleHeight     =   4275
      ScaleWidth      =   2565
      TabIndex        =   6
      Top             =   60
      Width           =   2595
      Begin MSComctlLib.ListView lsvACL 
         Height          =   3525
         Left            =   60
         TabIndex        =   8
         Top             =   690
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   6218
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ACL no."
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtSEARCH 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   330
         Width           =   2475
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Index           =   0
         Left            =   -30
         TabIndex        =   77
         Top             =   0
         Width           =   2625
         _Version        =   655364
         _ExtentX        =   4630
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Search ACL"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   -540
      ScaleHeight     =   960
      ScaleWidth      =   12015
      TabIndex        =   9
      Top             =   5790
      Width           =   12015
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   780
         Top             =   480
      End
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
         Left            =   11220
         MouseIcon       =   "frmCSMS_ACL.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Exit Window"
         Top             =   75
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
         Left            =   10530
         MouseIcon       =   "frmCSMS_ACL.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print this Record"
         Top             =   75
         Width           =   705
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost"
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
         Left            =   9840
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_ACL.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Unpost this Transaction"
         Top             =   75
         Width           =   705
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post"
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
         Left            =   9150
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmCSMS_ACL.frx":1E89
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":1FDB
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Post this Transaction"
         Top             =   75
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
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
         Left            =   8460
         MouseIcon       =   "frmCSMS_ACL.frx":2300
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":2452
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Delete Selected Record"
         Top             =   75
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
         Left            =   7770
         MouseIcon       =   "frmCSMS_ACL.frx":277D
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":28CF
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Edit Selected Record"
         Top             =   75
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
         Left            =   7080
         MouseIcon       =   "frmCSMS_ACL.frx":2C2B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":2D7D
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Add Record"
         Top             =   75
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
         Left            =   6360
         MouseIcon       =   "frmCSMS_ACL.frx":3090
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":31E2
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Move to Last Record"
         Top             =   75
         Width           =   735
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
         Left            =   5640
         MouseIcon       =   "frmCSMS_ACL.frx":3532
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":3684
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Move to First Record"
         Top             =   75
         Width           =   735
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
         Left            =   4950
         MouseIcon       =   "frmCSMS_ACL.frx":39E2
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":3B34
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Find a Record"
         Top             =   75
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
         Left            =   4260
         MouseIcon       =   "frmCSMS_ACL.frx":3E2E
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":3F80
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Move to Next Record"
         Top             =   75
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
         Left            =   3570
         MouseIcon       =   "frmCSMS_ACL.frx":42D8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":442A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Move to Previous Record"
         Top             =   75
         Width           =   705
      End
      Begin Crystal.CrystalReport rptACL 
         Left            =   780
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label lblSTATUS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   990
         TabIndex        =   22
         Top             =   150
         Width           =   2445
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   9990
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   23
      Top             =   5790
      Width           =   1590
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
         MouseIcon       =   "frmCSMS_ACL.frx":4789
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":48DB
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel"
         Top             =   75
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
         Left            =   0
         MouseIcon       =   "frmCSMS_ACL.frx":4C19
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":4D6B
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save this Record"
         Top             =   75
         Width           =   705
      End
   End
   Begin VB.PictureBox picHEAD 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   2640
      ScaleHeight     =   2925
      ScaleWidth      =   8715
      TabIndex        =   26
      Top             =   60
      Width           =   8745
      Begin MSComCtl2.DTPicker dtpTranDate 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   690
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20250625
         CurrentDate     =   39602
      End
      Begin VB.TextBox txtPClaim 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   330
         Width           =   1815
      End
      Begin VB.TextBox txtACCT 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         TabIndex        =   5
         Top             =   2160
         Width           =   4035
      End
      Begin VB.TextBox txtHARI 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox txtTClaim 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1470
         Width           =   1815
      End
      Begin VB.TextBox txtSClaim 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1110
         Width           =   1815
      End
      Begin VB.TextBox txtLClaim 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtNotedBy 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtPREPBY 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtACLno 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   330
         Width           =   2235
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TRAN DATE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   13
         Left            =   270
         TabIndex        =   56
         Top             =   810
         Width           =   945
      End
      Begin VB.Label lblTranno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   8040
         TabIndex        =   55
         Top             =   -30
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL PART CLAIMED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   4890
         TabIndex        =   54
         Top             =   450
         Width           =   1845
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIVED BY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   11
         Left            =   4620
         TabIndex        =   51
         Top             =   1950
         Width           =   1110
      End
      Begin VB.Label labID 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   1650
         TabIndex        =   43
         Top             =   -30
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACL NO."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   10
         Left            =   540
         TabIndex        =   39
         Top             =   450
         Width           =   675
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACCOUNTING"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   9
         Left            =   6060
         TabIndex        =   38
         Top             =   2490
         Width           =   1155
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTED BY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   390
         TabIndex        =   37
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PREPARED BY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   36
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HARI REPRESENTATIVE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   1260
         TabIndex        =   35
         Top             =   2520
         Width           =   1950
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIVED BY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   34
         Top             =   1950
         Width           =   1110
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL AMOUNT CLAIMED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   4620
         TabIndex        =   33
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SUBLET CLAIMED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   4725
         TabIndex        =   32
         Top             =   1200
         Width           =   2040
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL LABOR CLAIMED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   4785
         TabIndex        =   31
         Top             =   810
         Width           =   1965
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   -30
         Width           =   8775
         _Version        =   655364
         _ExtentX        =   15478
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "ACL INFORMATION"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorLight=   12632256
         GradientColorDark=   4210752
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picAPROVE 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   2468
      ScaleHeight     =   4065
      ScaleWidth      =   6465
      TabIndex        =   59
      Top             =   1313
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdCancelACM 
         Caption         =   "&Cancel"
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
         Left            =   5670
         MouseIcon       =   "frmCSMS_ACL.frx":50BB
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":520D
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Cancel"
         Top             =   3150
         Width           =   705
      End
      Begin VB.TextBox txtACM_AMOUNT 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1830
         TabIndex        =   67
         Top             =   720
         Width           =   1905
      End
      Begin VB.TextBox txtREMARKS 
         Height          =   1935
         Left            =   1830
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   68
         Top             =   1110
         Width           =   4515
      End
      Begin VB.TextBox txtACMNO 
         Height          =   315
         Left            =   1830
         TabIndex        =   65
         Top             =   330
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dptACMDATE 
         Height          =   345
         Left            =   4800
         TabIndex        =   66
         Top             =   330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         Format          =   20250625
         CurrentDate     =   39602
      End
      Begin VB.CommandButton cmdAprovACM 
         Caption         =   "Approved"
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
         Left            =   4980
         MouseIcon       =   "frmCSMS_ACL.frx":554B
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_ACL.frx":569D
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Save this Record"
         Top             =   3150
         Width           =   705
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APPROVED AMOUNT"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   19
         Left            =   60
         TabIndex        =   64
         Top             =   810
         Width           =   1695
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACM NO."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   18
         Left            =   1020
         TabIndex        =   63
         Top             =   420
         Width           =   720
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ACM DATE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   17
         Left            =   3870
         TabIndex        =   62
         Top             =   420
         Width           =   870
      End
      Begin VB.Label lblCAP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   16
         Left            =   960
         TabIndex        =   61
         Top             =   1170
         Width           =   780
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   6465
         _Version        =   655364
         _ExtentX        =   11404
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "ENTER ACM INFORMATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picClaims1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   1403
      ScaleHeight     =   3585
      ScaleWidth      =   8595
      TabIndex        =   72
      Top             =   1553
      Visible         =   0   'False
      Width           =   8625
      Begin VB.TextBox TXTSEARCHI2 
         Height          =   405
         Left            =   60
         TabIndex        =   73
         Top             =   360
         Width           =   8475
      End
      Begin MSComctlLib.ListView lsvQIR 
         Height          =   2655
         Left            =   60
         TabIndex        =   74
         Top             =   840
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "QIR NO"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PWA NO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RO_NO"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CUSTOMER"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "VIN NO"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin wizButton.cmd cmd1 
         Height          =   255
         Left            =   8340
         TabIndex        =   75
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_ACL.frx":59ED
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Width           =   8775
         _Version        =   655364
         _ExtentX        =   15478
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "SEARCH A QIR"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
   End
   Begin VB.PictureBox picClaims 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   1088
      ScaleHeight     =   3705
      ScaleWidth      =   9225
      TabIndex        =   46
      Top             =   1493
      Visible         =   0   'False
      Width           =   9255
      Begin XtremeReportControl.ReportControl rptCLAIMS 
         Height          =   2925
         Left            =   90
         TabIndex        =   47
         Top             =   660
         Width           =   9015
         _Version        =   655364
         _ExtentX        =   15901
         _ExtentY        =   5159
         _StockProps     =   64
      End
      Begin wizButton.cmd cmdX 
         Height          =   285
         Left            =   8970
         TabIndex        =   50
         Top             =   -30
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "frmCSMS_ACL.frx":5A09
      End
      Begin VB.TextBox txtSEARCHI 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         TabIndex        =   49
         Top             =   300
         Width           =   9075
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption11 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   9225
         _Version        =   655364
         _ExtentX        =   16272
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "SEARCH CLAIMS"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   4210752
      End
   End
End
Attribute VB_Name = "frmCSMS_ACL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsACL                                              As ADODB.Recordset
Dim ADD_OR_EDIT                                        As String
Dim AUDIT_SQL                                          As String

Function GenerateNewTranno()
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT ACLNO,TRANNO,TRANDATE FROM CSMS_ACL_HD WHERE MONTH(TRANDATE) = " & Month(Date) & " AND YEAR(TRANDATE) = " & Year(Date) & " ORDER BY ACLNO DESC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        rstmp.MoveFirst
        GenerateNewTranno = Format(Right(rstmp!ACLNO, 2) + 1, "00")
    Else
        GenerateNewTranno = Format(1, "00")
    End If

    Set rstmp = Nothing
End Function

Function GetMissingDetail(vREPOR As String, vDET As String)
    Dim rsKUTO                                         As New ADODB.Recordset
    Set rsKUTO = gconDMIS.Execute("SELECT " & vDET & " AS DET_INFO FROM CSMS_REPOR WHERE REP_OR = '" & vREPOR & "'")
    If Not (rsKUTO.BOF And rsKUTO.EOF) Then
        If Null2String(rsKUTO!DET_INFO) = "" Then
            If vDET = "DTE_REL" Then GetMissingDetail = "RO NOT YET RELEASED"
        Else
            GetMissingDetail = Null2String(rsKUTO!DET_INFO)
        End If
    Else
        GetMissingDetail = "RO NOT FOUND"
    End If

    Set rsKUTO = Nothing
End Function

Function FindNatureDescription(NCODE As String)
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT * FROM WWTNATR WHERE NATRCODE = '" & NCODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindNatureDescription = Null2String(LTrim(RTrim(Replace(Null2String(rstmp!NATRENGL), vbCrLf, ""))))
    Else
        FindNatureDescription = ""
    End If
    Set rstmp = Nothing
End Function

Function FindCauseDescription(CCODE As String)
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT * FROM WWTCAUE WHERE CAUECODE = '" & CCODE & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        FindCauseDescription = Null2String(LTrim(RTrim(Replace(Null2String(rstmp!CAUEENGL), vbCrLf, ""))))
    Else
        FindCauseDescription = ""
    End If
    Set rstmp = Nothing
End Function

Function GenerateClaimNo() As String
    Dim rstmp                                          As New ADODB.Recordset

    Dim RSTMP1                                         As New ADODB.Recordset
    Dim COCODE                                         As String
    Set RSTMP1 = gconDMIS.Execute("SELECT COMPANYCODE FROM ALL_PROFILE WHERE MODULENAME = 'CSMS'")
    If Not (RSTMP1.BOF And RSTMP1.EOF) Then
        COCODE = Null2String(RSTMP1!COMPANYCODE)
    End If
    Set RSTMP1 = Nothing

    Set rstmp = gconDMIS.Execute("SELECT CLAIMNO FROM CSMS_CQIR WHERE STATUS = 'T' AND MONTH(DATEATTACHEDTOACL) = " & Month(Date) & " AND YEAR(DATEATTACHEDTOACL) = " & Year(Date) & " ORDER BY RIGHT(CLAIMNO,2) DESC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        GenerateClaimNo = COCODE & Right(Year(Date), 2) & Format(Month(Date), "00") & Format(Right(rstmp!CLAIMNO, 2) + 1, "00")
    Else
        GenerateClaimNo = COCODE & Right(Year(Date), 2) & Format(Month(Date), "00") & Format(1, "00")
    End If

    Set rstmp = Nothing
End Function

Public Function flex_FillReportView(RS As ADODB.Recordset, grd As XtremeReportControl.ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                            As ADODB.FIELD
    Dim j                                              As Long
    Dim REC                                            As XtremeReportControl.ReportRecord

    grd.Records.DeleteAll

    While Not RS.EOF
        j = j + 1

        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem (Trim(fld.Value))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
    Set RS = Nothing
End Function

Sub DisablePicture(COND As Boolean)
    picHEAD.Enabled = Not COND
    picDET.Enabled = COND
    picSEARCH.Enabled = COND

    'picClaims.Visible = Not COND
End Sub

Sub PrintACL()
    Dim rsCQIR                                         As New ADODB.Recordset
    Dim RSPART                                         As New ADODB.Recordset
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "ACL_FORMAT.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    Dim xPCOST As Currency: Dim xLCOST As Currency: Dim xSCOST As Currency: Dim xGCOST As Currency
    Dim DET1 As Integer: Dim DET2                      As Integer
    Dim SIGNA                                          As Integer
    Dim MAX_LINE                                       As Integer
    Dim FIRST_ITEM                                     As Integer
    Dim GRAND_LINE                                     As Integer
    Dim TOTAL_LINE                                     As Integer

    Dim Index                                          As Integer
    Dim PAGENO                                         As Integer
    Dim MAX_ITEM                                       As Integer
    Dim LINE_1                                         As Integer
    Dim LINE_2                                         As Integer
    Dim LINE_3                                         As Integer
    Dim GPARTS                                         As Currency
    Dim GSUBLETS                                       As Currency
    Dim GJOBS                                          As Currency

    Dim TGPARTS                                        As Currency
    Dim TGSUBLETS                                      As Currency
    Dim TGJOBS                                         As Currency
    Dim D_LINE                                         As Integer
    
    D_LINE = 35

    GPARTS = 0: GSUBLETS = 0: GJOBS = 0
    TGPARTS = 0: TGSUBLETS = 0: TGJOBS = 0

    LINE_1 = 9: LINE_2 = 10: LINE_3 = 11
    DET1 = 9: DET2 = 10
    PAGENO = 1: Index = 11: MAX_ITEM = 35
    SIGNA = 42: MAX_LINE = 42: FIRST_ITEM = 9
    GRAND_LINE = 39: TOTAL_LINE = 36

    xlSheet.Cells(1, "D") = COMPANY_NAME
    xlSheet.Cells(3, "E") = txtACLno.Text
    Dim ITEM_NO                                        As Integer
    Dim FLINE                                          As Integer
    Dim LLINE                                          As Integer
    Dim FL                                             As Integer
    Dim TMP_LAST_LINE                                  As Integer
    Dim tmp_line_3                                     As Integer
    Dim LEFT_F                                         As Integer
    Dim FROM_NEXT_PAGE                                 As Integer
    Dim NEXT_TMP_LINE_3                                As Integer


    Set rsCQIR = gconDMIS.Execute("SELECT * FROM CSMS_CQIR WHERE ACLNO = '" & txtACLno.Text & "' ORDER BY CLAIMNO")
    If Not (rsCQIR.BOF And rsCQIR.EOF) Then
        Do While Not rsCQIR.EOF
            FL = 0: FLINE = 0: LLINE = 0

            'CLAIM NO. COLUMN
            xlSheet.Cells(LINE_1, "A") = Null2String(rsCQIR!CLAIMNO)
            xlSheet.Cells(LINE_2, "A") = Null2String(rsCQIR!RO_NO)
            xlSheet.Cells(LINE_3, "A") = Null2String(rsCQIR!CLAIMTYPE)

            'VIN COLUMN
            xlSheet.Cells(LINE_1, "B") = Null2String(rsCQIR!VINNO)
            xlSheet.Cells(LINE_2, "B") = Null2String(rsCQIR!EngineNo)
            If COMPANY_CODE = "HCI" Then
                xlSheet.Cells(LINE_3, "B") = Null2String(rsCQIR!Customer)
            Else
                xlSheet.Cells(LINE_3, "B") = Null2String(rsCQIR!PREVACLNO)
            End If
            'DEL. DATE COLUMN
            xlSheet.Cells(LINE_1, "C") = Null2String(rsCQIR!DELDATE)
            xlSheet.Cells(LINE_2, "C") = Null2String(rsCQIR!RepairDate)
            'xlSheet.Cells(LINE_3, "C") = Null2String(rsCQIR!InspectionDate)
            xlSheet.Cells(LINE_3, "C") = GetMissingDetail(Null2String(rsCQIR!RO_NO), "DTE_COMP")

            'ODOMETER COLUMN
            xlSheet.Cells(LINE_1, "D") = Null2String(rsCQIR!MILEAGE)
            xlSheet.Cells(LINE_2, "D") = Null2String(rsCQIR!CAUSALPARTNO)
            xlSheet.Cells(LINE_3, "D") = Null2String(rsCQIR!NATURECODE) & " / " & Null2String(rsCQIR!CAUSECODE)

            'CONDITION COLUMN
            'FROM DATA BASE
            'xlSheet.Cells(LINE_1, "J") = Null2String(rsCQIR!Description)
            xlSheet.Cells(LINE_1, "J") = FindNatureDescription(Null2String(rsCQIR!NATURECODE))

            'FROM DATA BASE
            'xlSheet.Cells(LINE_2, "J") = Null2String(rsCQIR!ANALYSIS)
            xlSheet.Cells(LINE_2, "J") = FindCauseDescription(Null2String(rsCQIR!CAUSECODE))
            If COMPANY_CODE = "HCI" Then
                xlSheet.Cells(LINE_3, "J") = Null2String(rsCQIR!correctiveAction)
            Else
                xlSheet.Cells(LINE_3, "J") = Null2String(rsCQIR!RECOMMENDATION)
            End If
            xlSheet.Cells(LINE_3 + 1, "J") = "TOTAL"
            xlSheet.Cells(LINE_3 + 1, "J").Font.Bold = True

            'PART COST COLUMN
            xlSheet.Cells(LINE_1, "K") = Null2String(Format(rsCQIR!TOTALPARTCOST, "#,###,##0.00"))
            xlSheet.Cells(LINE_2, "K") = Null2String(Format(rsCQIR!TotalLaborCost, "#,###,##0.00"))
            xlSheet.Cells(LINE_3, "K") = Null2String(Format(rsCQIR!TotalSUBLETREPAIR, "#,###,##0.00"))
            xlSheet.Cells(LINE_3 + 1, "K") = NumericVal(Format(rsCQIR!TOTALPARTCOST + rsCQIR!TotalLaborCost + rsCQIR!TotalSUBLETREPAIR, "#,###,##0.00"))
            xlSheet.Cells(LINE_3 + 1, "K").Font.Bold = True

            GPARTS = GPARTS + CCur(rsCQIR!TOTALPARTCOST)
            GSUBLETS = GSUBLETS + CCur(rsCQIR!TotalSUBLETREPAIR)
            GJOBS = GJOBS + CCur(rsCQIR!TotalLaborCost)

            TGPARTS = TGPARTS + CCur(rsCQIR!TOTALPARTCOST)
            TGSUBLETS = TGSUBLETS + CCur(rsCQIR!TotalSUBLETREPAIR)
            TGJOBS = TGJOBS + CCur(rsCQIR!TotalLaborCost)


            'PWA TYPE COLUMN
            xlSheet.Cells(LINE_1, "L") = Null2String(rsCQIR!PWATYPE) & " / " & Null2String(rsCQIR!PWANO)
            xlSheet.Cells(LINE_2, "L") = Null2String(rsCQIR!SUBLETTYPE)
            xlSheet.Cells(LINE_3, "L") = Null2String(rsCQIR!PREVRONO)

            'HAI CUSTOMER NAME AND INVOICE NO
            'If COMPANY_CODE = "HAI" Then
            xlSheet.Cells(LINE_1, "M") = GetMissingDetail(Null2String(rsCQIR!RO_NO), "NIYM")
            xlSheet.Cells(LINE_3 + 1, "M") = "INVOICE NO: " & GetMissingDetail(Null2String(rsCQIR!RO_NO), "INVOICE")
            'End If
            'HAI CUSTOMER NAME AND INVOICE NO

            'SIGNATORIES COLUMN
            xlSheet.Cells(SIGNA, "E") = Null2String(txtPREPBY.Text)
            xlSheet.Cells(SIGNA, "G") = Null2String(txtNotedBy.Text)
            xlSheet.Cells(SIGNA, "J") = Null2String(txtHARI.Text)
            xlSheet.Cells(SIGNA, "k") = Null2String(txtACCT.Text)

            ITEM_NO = 0
            DET1 = LINE_1: DET2 = LINE_2
            LLINE = LINE_3

            Set RSPART = New ADODB.Recordset
            Call RSPART.Open("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & Null2String(rsCQIR!DLR_CQIR_REFERENCENO) & "' ORDER BY ID", gconDMIS, adOpenStatic)
            Dim lng                                    As Long

            'lng = RSPART.RecordCount
            tmp_line_3 = LINE_3
            If Not (RSPART.BOF And RSPART.EOF) Then
                Do While Not RSPART.EOF
                    'PART NO COLUMN
                    If FL = 0 Then
                        FLINE = DET1 + 3
                        FL = 1
                    End If
                    
                    
                    xlSheet.Cells(DET1, "E") = Null2String(RSPART!partno)
                    xlSheet.Cells(DET2, "E") = Null2String(RSPART!partname)

                    'QTY COLUMN
                    xlSheet.Cells(DET1, "F") = Null2String(RSPART!QTY)

                    'OPCODE COLUMN
                    xlSheet.Cells(DET1, "G") = Null2String(RSPART!OPCODE)

                    'QTY COLUMN
                    xlSheet.Cells(DET1, "H") = Null2String(RSPART!QTY)

                    'LTS COLUMN COLUMN
                    xlSheet.Cells(DET1, "I") = Null2String(RSPART!lts)

                    ITEM_NO = ITEM_NO + 1


                    RSPART.MoveNext
                    If RSPART.EOF = True Then
                        RSPART.MovePrevious
                        LLINE = DET2
                        TMP_LAST_LINE = DET2
                    Else
                        RSPART.MovePrevious
                    End If
                    
                   
                    LINE_1 = DET2 + 1
                    LINE_2 = DET2 + 2
                    LINE_3 = DET2 + 3
            
                    'Index = LINE_3
                    DET1 = DET1 + 2
                    DET2 = DET2 + 2

                    If DET2 >= MAX_ITEM Then
                        FROM_NEXT_PAGE = 1
                        LLINE = MAX_ITEM
                    
                        xlSheet.Range("A" & FLINE, "D" & LLINE).Merge
                        LEFT_F = tmp_line_3 + 2
                       
                        xlSheet.Range("J" & LEFT_F, "L" & LLINE).Merge
                        tmp_line_3 = DET1
                      
                        xlSheet.Cells(GRAND_LINE, "C") = Format(TGPARTS, "#,###,##0.00")
                        xlSheet.Cells(GRAND_LINE + 1, "C") = Format(TGJOBS, "#,###,##0.00")
                        xlSheet.Cells(GRAND_LINE + 2, "C") = Format(TGSUBLETS, "#,###,##0.00")
                        xlSheet.Cells(TOTAL_LINE, "K") = Format(TGSUBLETS + TGPARTS + TGJOBS, "#,###,##0.00")

                        TGPARTS = 0: TGSUBLETS = 0: TGJOBS = 0
                        GRAND_LINE = GRAND_LINE + 48
                        TOTAL_LINE = TOTAL_LINE + 48

                        FIRST_ITEM = MAX_ITEM + 22
                        MAX_ITEM = FIRST_ITEM + 26
                        SIGNA = SIGNA + 48
                        LINE_1 = FIRST_ITEM
                        LINE_2 = FIRST_ITEM + 1
                        LINE_3 = FIRST_ITEM + 2

                        DET1 = LINE_1
                        DET2 = LINE_2

                        FLINE = DET1
                        NEXT_TMP_LINE_3 = DET1
                    Else
                        'Stop
                    End If

                    RSPART.MoveNext
                Loop
                'Stop
                'If LLINE < TMP_LAST_LINE Then LLINE = LLINE + 2
                

                
                If LLINE < tmp_line_3 Then LLINE = LLINE + 2
                If FLINE = LLINE Then xlSheet.Range("A" & FLINE, "D" & LLINE).Merge
               

                If FROM_NEXT_PAGE = 1 Then
                    FROM_NEXT_PAGE = 0
                    If Not NEXT_TMP_LINE_3 + 1 = LLINE Then
                        LEFT_F = NEXT_TMP_LINE_3
                        'xlSheet.Range("J" & LEFT_F, "L" & LLINE).Merge
                    End If
                Else
                    If Not tmp_line_3 + 1 = LLINE Then
                        LEFT_F = tmp_line_3 + 2
                        'xlSheet.Range("J" & LEFT_F, "L" & LLINE).Merge
                    End If
                End If
            Else
                LINE_3 = LINE_1 + 4
                LINE_2 = LINE_1 + 3
                LINE_1 = LINE_1 + 2

                If LINE_3 >= MAX_ITEM Then
                    xlSheet.Cells(GRAND_LINE, "C") = Format(TGPARTS, "#,###,##0.00")
                    xlSheet.Cells(GRAND_LINE + 1, "C") = Format(TGJOBS, "#,###,##0.00")
                    xlSheet.Cells(GRAND_LINE + 2, "C") = Format(TGSUBLETS, "#,###,##0.00")
                    xlSheet.Cells(TOTAL_LINE, "K") = Format(TGSUBLETS + TGPARTS + TGJOBS, "#,###,##0.00")

                    TGPARTS = 0: TGSUBLETS = 0: TGJOBS = 0
                    GRAND_LINE = GRAND_LINE + 48
                    TOTAL_LINE = TOTAL_LINE + 48

                    FIRST_ITEM = MAX_ITEM + 22
                    MAX_ITEM = FIRST_ITEM + 26
                    SIGNA = SIGNA + 48
                    LINE_1 = FIRST_ITEM
                    LINE_2 = FIRST_ITEM + 1
                    LINE_3 = FIRST_ITEM + 2
                    'Else
                    '    Stop
                End If
            End If

            If ITEM_NO = 1 Then
                LINE_3 = LINE_1 + 4
                LINE_2 = LINE_1 + 3
                LINE_1 = LINE_1 + 2

                If LINE_3 >= MAX_ITEM Then
                    xlSheet.Cells(GRAND_LINE, "C") = Format(TGPARTS, "#,###,##0.00")
                    xlSheet.Cells(GRAND_LINE + 1, "C") = Format(TGJOBS, "#,###,##0.00")
                    xlSheet.Cells(GRAND_LINE + 2, "C") = Format(TGSUBLETS, "#,###,##0.00")
                    xlSheet.Cells(TOTAL_LINE, "K") = Format(TGSUBLETS + TGPARTS + TGJOBS, "#,###,##0.00")

                    TGPARTS = 0: TGSUBLETS = 0: TGJOBS = 0
                    GRAND_LINE = GRAND_LINE + 48
                    TOTAL_LINE = TOTAL_LINE + 48

                    FIRST_ITEM = MAX_ITEM + 22
                    MAX_ITEM = FIRST_ITEM + 26
                    SIGNA = SIGNA + 48
                    LINE_1 = FIRST_ITEM
                    LINE_2 = FIRST_ITEM + 1
                    LINE_3 = FIRST_ITEM + 2
                End If
            Else

            End If

            Set RSPART = Nothing

            rsCQIR.MoveNext
        Loop
    End If
    Set rsCQIR = Nothing

    xlSheet.Cells(TOTAL_LINE, "K") = Format(TGSUBLETS + TGPARTS + TGJOBS, "#,###,##0.00")
    xlSheet.Cells(GRAND_LINE, "C") = Format(GPARTS, "#,###,##0.00")
    xlSheet.Cells(GRAND_LINE + 1, "C") = Format(GJOBS, "#,###,##0.00")
    xlSheet.Cells(GRAND_LINE + 2, "C") = Format(GSUBLETS, "#,###,##0.00")
    xlSheet.Cells(GRAND_LINE + 3, "C") = Format(GPARTS + GJOBS + GSUBLETS, "#,###,##0.00")

    xlApp.Windows.Item(1).Caption = "ACL NO: " & txtACLno
    xlApp.Visible = True
    'xlBook.Close
    Set xlApp = Nothing: Set xlSheet = Nothing: Set xlBook = Nothing
End Sub

Sub PrintExcelACL()
    Dim rsCQIR                                         As New ADODB.Recordset
    Dim RSPART                                         As New ADODB.Recordset
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "ACL_BLANK.xlt")
    Set xlSheet = xlBook.Worksheets(1)
    Dim xPCOST                                         As Currency
    Dim xLCOST                                         As Currency
    Dim xSCOST                                         As Currency
    Dim xGCOST                                         As Currency


    Dim Index                                          As Integer
    Dim PAGENO                                         As Integer
    Dim MAX_ITEM                                       As Integer
    Dim LINE_1                                         As Integer
    Dim LINE_2                                         As Integer
    Dim LINE_3                                         As Integer

    LINE_1 = 9: LINE_2 = 10: LINE_3 = 11
    PAGENO = 1: Index = 11: MAX_ITEM = 47

    xlSheet.Cells(1, "D") = COMPANY_NAME
    xlSheet.Cells(3, "E") = txtACLno.Text

    Set rsCQIR = gconDMIS.Execute("SELECT * FROM CSMS_CQIR WHERE ACLNO = '" & txtACLno.Text & "' ORDER BY CLAIMNO")
    If Not (rsCQIR.BOF And rsCQIR.EOF) Then
        Do While Not rsCQIR.EOF
            'CLAIM NO. COLUMN
            xlSheet.Cells(LINE_1, "A") = Null2String(rsCQIR!CLAIMNO)
            xlSheet.Cells(LINE_2, "A") = Null2String(rsCQIR!RO_NO)
            xlSheet.Cells(LINE_3, "A") = Null2String(rsCQIR!CLAIMTYPE)

            'VIN COLUMN
            xlSheet.Cells(LINE_1, "B") = Null2String(rsCQIR!VINNO)
            xlSheet.Cells(LINE_2, "B") = Null2String(rsCQIR!EngineNo)
            xlSheet.Cells(LINE_3, "B") = ""

            'DEL. DATE COLUMN
            xlSheet.Cells(LINE_1, "C") = Null2String(rsCQIR!DELDATE)
            xlSheet.Cells(LINE_2, "C") = Null2String(rsCQIR!RepairDate)
            xlSheet.Cells(LINE_3, "C") = Null2String(rsCQIR!InspectionDate)

            'ODOMETER COLUMN
            xlSheet.Cells(LINE_1, "D") = Null2String(rsCQIR!MILEAGE)
            xlSheet.Cells(LINE_2, "D") = Null2String(rsCQIR!CAUSALPARTNO)
            xlSheet.Cells(LINE_3, "D") = Null2String(rsCQIR!NATURECODE) & " / " & Null2String(rsCQIR!CAUSECODE)

            'CONDITION COLUMN
            xlSheet.Cells(LINE_1, "J") = Null2String(rsCQIR!Description)
            xlSheet.Cells(LINE_2, "J") = Null2String(rsCQIR!ANALYSIS)
            xlSheet.Cells(LINE_3, "J") = Null2String(rsCQIR!RECOMMENDATION)

            'PART CSOT COLUMN
            xlSheet.Cells(LINE_1, "K") = Null2String(Format(rsCQIR!TOTALPARTCOST, "#,###,##0.00"))
            xlSheet.Cells(LINE_2, "K") = Null2String(Format(rsCQIR!TotalLaborCost, "#,###,##0.00"))
            xlSheet.Cells(LINE_3, "K") = Null2String(Format(rsCQIR!TotalSUBLETREPAIR, "#,###,##0.00"))

            'PWA TYPE COLUMN
            xlSheet.Cells(LINE_1, "L") = Null2String(rsCQIR!PWATYPE) & " / " & Null2String(rsCQIR!PWANO)
            xlSheet.Cells(LINE_2, "L") = ""
            xlSheet.Cells(LINE_3, "L") = Null2String(rsCQIR!RO_NO)

            'SIGNATORIES COLUMN
            xlSheet.Cells(42, "E") = Null2String(txtPREPBY.Text)
            xlSheet.Cells(42, "G") = Null2String(txtNotedBy.Text)
            xlSheet.Cells(42, "J") = Null2String(txtHARI.Text)
            xlSheet.Cells(42, "k") = Null2String(txtACCT.Text)

            xPCOST = xPCOST + NumericVal(rsCQIR!TOTALPARTCOST)
            xLCOST = xLCOST + NumericVal(rsCQIR!TotalLaborCost)
            xSCOST = xSCOST + NumericVal(rsCQIR!TotalSUBLETREPAIR)
            xGCOST = xPCOST + xSCOST + xLCOST

            Set RSPART = gconDMIS.Execute("SELECT * FROM CSMS_CQIRPARTS WHERE DLR_CQIR_REFERENCENO = '" & Null2String(rsCQIR!DLR_CQIR_REFERENCENO) & "' ORDER BY ID")
            If Not (RSPART.BOF And RSPART.EOF) Then
                Do While Not RSPART.EOF
                    'PART NO COLUMN
                    xlSheet.Cells(LINE_1, "E") = Null2String(RSPART!partno)
                    xlSheet.Cells(LINE_2, "E") = Null2String(RSPART!partname)
                    xlSheet.Cells(LINE_3, "E") = ""

                    'QTY COLUMN
                    xlSheet.Cells(LINE_1, "F") = Null2String(RSPART!QTY)
                    xlSheet.Cells(LINE_2, "F") = ""
                    xlSheet.Cells(LINE_3, "F") = ""

                    'OPCODE COLUMN
                    xlSheet.Cells(LINE_1, "G") = Null2String(RSPART!OPCODE)
                    xlSheet.Cells(LINE_2, "G") = ""
                    xlSheet.Cells(LINE_3, "G") = ""

                    'QTY COLUMN
                    xlSheet.Cells(LINE_1, "H") = Null2String(RSPART!QTY)
                    xlSheet.Cells(LINE_2, "H") = ""
                    xlSheet.Cells(LINE_3, "H") = ""

                    'LTS COLUMN COLUMN
                    xlSheet.Cells(LINE_1, "I") = Null2String(RSPART!lts)
                    xlSheet.Cells(LINE_2, "I") = ""
                    xlSheet.Cells(LINE_3, "I") = ""

                    LINE_1 = LINE_1 + 4
                    LINE_2 = LINE_2 + 4
                    LINE_3 = LINE_3 + 4

                    Index = Index + 4

                    If Index > 35 Then
                        Index = 11
                        xlSheet.Cells(1, "L") = PAGENO
                        LINE_1 = 9: LINE_2 = 10: LINE_3 = 11
                        PAGENO = PAGENO + 1

                        xlApp.Visible = True
                        xlBook.Close

                        Set xlApp = Nothing: Set xlSheet = Nothing: Set xlBook = Nothing

                        Set xlApp = CreateObject("Excel.Application")
                        Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "ACL_BLANK.xlt")
                        Set xlSheet = xlBook.Worksheets(1)

                        xlSheet.Cells(1, "D") = COMPANY_NAME
                        xlSheet.Cells(3, "E") = txtACLno.Text

                        xlSheet.Cells(42, "E") = Null2String(txtPREPBY.Text)
                        xlSheet.Cells(42, "G") = Null2String(txtNotedBy.Text)
                        xlSheet.Cells(42, "J") = Null2String(txtHARI.Text)
                        xlSheet.Cells(42, "k") = Null2String(txtACCT.Text)
                    Else
                        xlSheet.Cells(1, "L") = PAGENO
                    End If
                    RSPART.MoveNext
                Loop
            Else
                LINE_1 = LINE_1 + 4
                LINE_2 = LINE_2 + 4
                LINE_3 = LINE_3 + 4

                Index = Index + 4
            End If
            Set RSPART = Nothing

            'INDEX = INDEX + 4

            If Index > 35 Then
                Index = 11
                xlSheet.Cells(1, "L") = PAGENO
                LINE_1 = 9: LINE_2 = 10: LINE_3 = 11
                PAGENO = PAGENO + 1

                xlApp.Visible = True
                xlBook.Close
                Set xlApp = Nothing: Set xlSheet = Nothing: Set xlBook = Nothing

                Set xlApp = CreateObject("Excel.Application")
                Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "ACL_BLANK.xlt")
                Set xlSheet = xlBook.Worksheets(1)

                xlSheet.Cells(1, "D") = COMPANY_NAME
                xlSheet.Cells(3, "E") = txtACLno.Text

                xlSheet.Cells(42, "E") = Null2String(txtPREPBY.Text)
                xlSheet.Cells(42, "G") = Null2String(txtNotedBy.Text)
                xlSheet.Cells(42, "J") = Null2String(txtHARI.Text)
                xlSheet.Cells(42, "k") = Null2String(txtACCT.Text)
            Else
                xlSheet.Cells(1, "L") = PAGENO
            End If

            rsCQIR.MoveNext
        Loop

        xlApp.Visible = True
        xlBook.Close
    End If

    Set xlApp = Nothing: Set xlSheet = Nothing: Set xlBook = Nothing

    Set rsCQIR = Nothing
End Sub

Sub FillSearchClaims1()
    Dim rstmp                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set rstmp = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO, PWANO, RO_NO, CUSTOMER, VINNO, ID FROM CSMS_CQIR WHERE STATUS = 'A' ORDER BY ID")
    lsvQIR.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvQIR.ListItems.Add(, , Null2String(rstmp!DLR_CQIR_REFERENCENO))
            Item.SubItems(1) = Null2String(rstmp!PWANO)
            Item.SubItems(2) = Null2String(rstmp!RO_NO)
            Item.SubItems(3) = Null2String(rstmp!Customer)
            Item.SubItems(4) = Null2String(rstmp!VINNO)
            Item.SubItems(5) = rstmp!ID

            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub EnabledApprove(COND As Boolean)
    picSEARCH.Enabled = COND
    'picHEAD.Enabled = COND
    picDET.Enabled = COND
    picAdds.Enabled = COND
    picSaves.Enabled = COND
End Sub

Sub DisableForm(COND As Boolean)
    picHEAD.Enabled = COND
    picDET.Enabled = COND
    picSEARCH.Enabled = COND
    picAdds.Enabled = COND
    picSaves.Enabled = COND

    picClaims.Visible = Not COND
End Sub

Sub rsRefresh()
    Set rsACL = New ADODB.Recordset
    rsACL.Open "SELECT * FROM CSMS_ACL_HD order by ID", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not rsACL.EOF And Not rsACL.BOF Then
        labID.Caption = rsACL!ID
        txtACLno.Text = Null2String(rsACL!ACLNO)
        dtpTranDate.Value = Null2String(rsACL!TRANDATE)
        txtPREPBY.Text = Null2String(rsACL!PREPBY)
        txtNotedBy.Text = Null2String(rsACL!NotedBy)
        txtHARI.Text = Null2String(rsACL!HARI)
        txtACCT.Text = Null2String(rsACL!ACCT)
        lblTranno.Caption = Null2String(rsACL!TRANNO)

        If Null2String(rsACL!Status) = "" Then
            lblSTATUS.Caption = ""
            cmdPost.Enabled = True: cmdUnPost.Enabled = False
            cmdDelete.Enabled = True:
            cmdEdit.Enabled = True: cmdPrint.Enabled = False
        ElseIf Null2String(rsACL!Status) = "P" Then   'POSTED
            lblSTATUS.Caption = "** POSTED **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = True
            cmdDelete.Enabled = False:
            cmdEdit.Enabled = False: cmdPrint.Enabled = True
        ElseIf Null2String(rsACL!Status) = "S" Then   'SUBMITTED TO HARI
            lblSTATUS.Caption = "** SUBMITTED TO HARI **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = False
            cmdDelete.Enabled = False:
            cmdEdit.Enabled = False:                  'cmdPrint.Enabled = False
        ElseIf Null2String(rsACL!Status) = "AP" Then
            lblSTATUS.Caption = "** APPROVED **"
            cmdPost.Enabled = False: cmdUnPost.Enabled = False
            cmdDelete.Enabled = False: cmdEdit.Enabled = False
        End If

        txtACMNO.Text = Null2String(rsACL!ACM_NO)
        If Null2String(rsACL!ACM_DATE) = "" Then
            dptACMDATE.Value = Date
        Else
            dptACMDATE.Value = Null2Date(rsACL!ACM_DATE)
        End If
        txtACM_AMOUNT.Text = NumericVal(rsACL!APPROVED_AMOUNT)
        txtREMARKS.Text = Null2String(rsACL!REMARKS)

        FillACLDetails
    Else
        ShowNoRecord
        cmdAdd_Click
    End If
End Sub

Sub initMemvars()
    txtACLno.Text = ""
    dtpTranDate.Value = Date
    txtPREPBY.Text = ""
    txtNotedBy.Text = ""
    txtHARI.Text = ""
    txtACCT.Text = ""
    txtPClaim.Text = "0.00"
    txtLClaim.Text = "0.00"
    txtSClaim.Text = "0.00"
    txtTClaim.Text = "0.00"

    lsvDET.ListItems.Clear
End Sub

Sub FillSearchClaims()
    Dim rstmp                                          As New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO, PWANO, RO_NO, CUSTOMER, VINNO, ID FROM CSMS_CQIR WHERE STATUS = 'A' ORDER BY ID")
    Call ReportControlAddColumnHeader(rptCLAIMS, "QIR REF no., PWA no., RO no., Customer Name, Vin no.,  ")
    Call ReportControlPaintManager(rptCLAIMS)
    Call ResizeColumnHeader(rptCLAIMS, "18, 15, 10, 25, 13, 0")
    Call flex_FillReportView(rstmp, rptCLAIMS)
End Sub

Sub ReportControlAddColumnHeader(lst As ReportControl, StringHeaders As String)
    Dim ar()                                           As String
    Dim I                                              As Integer

    ar = Split(StringHeaders, ",")
    lst.Columns.DeleteAll
    For I = LBound(ar) To UBound(ar)
        lst.Columns.Add I, ar(I), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(lst As ReportControl)
    With lst
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub

Sub FillACLDetails()
    Dim rstmp                                          As New ADODB.Recordset
    Dim Item                                           As ListItem
    Dim vPCLAIM As Double: Dim vLCLAIM                 As Double
    Dim vSCLAIM As Double: Dim vTCLAIM                 As Double

    vPCLAIM = 0: vLCLAIM = 0
    vSCLAIM = 0: vTCLAIM = 0

    lsvDET.ListItems.Clear
    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_CQIR WHERE ACLNO = '" & txtACLno.Text & "' ORDER BY CLAIMNO")
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvDET.ListItems.Add(, , Null2String(rstmp!DLR_CQIR_REFERENCENO))
            Item.SubItems(1) = Null2String(rstmp!CLAIMNO)
            Item.SubItems(2) = Null2String(rstmp!RO_NO)
            Item.SubItems(3) = Null2String(rstmp!VINNO)
            Item.SubItems(4) = Format(N2Str2Zero(rstmp!grandtotal), "#,###,##0.00")
            Item.SubItems(5) = rstmp!ID

            vPCLAIM = vPCLAIM + N2Str2Zero(rstmp!TOTALPARTCOST)
            vSCLAIM = vSCLAIM + N2Str2Zero(rstmp!TotalSUBLETREPAIR)
            vLCLAIM = vLCLAIM + N2Str2Zero(rstmp!TotalLaborCost)
            vTCLAIM = vTCLAIM + N2Str2Zero(rstmp!grandtotal)

            rstmp.MoveNext
        Loop
    End If

    txtPClaim.Text = Format(vPCLAIM, "#,###,##0.00")
    txtLClaim.Text = Format(vLCLAIM, "#,###,##0.00")
    txtSClaim.Text = Format(vSCLAIM, "#,###,##0.00")
    txtTClaim.Text = Format(vTCLAIM, "#,###,##0.00")

    gconDMIS.Execute ("UPDATE CSMS_ACL_HD SET PARTSCLAIM = " & vPCLAIM & ", LABORCLAIM = " & vLCLAIM & ", SUBLETCLAIM = " & vSCLAIM & ", TOTALCLAIM = " & vTCLAIM & " WHERE ID = " & labID.Caption & "")

    Set rstmp = Nothing
End Sub

Sub FillGrid()
    Dim rstmp                                          As New ADODB.Recordset
    lsvACL.Enabled = False
    lsvACL.Sorted = False: lsvACL.ListItems.Clear
    Set rstmp = New ADODB.Recordset

    Set rstmp = gconDMIS.Execute("select ACLNO, ID from CSMS_ACL_HD Order by ACLNO")

    If Not (rstmp.EOF And rstmp.BOF) Then
        Listview_Loadval Me.lsvACL.ListItems, rstmp
        lsvACL.Refresh
        lsvACL.Enabled = True
    End If

    Set rstmp = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rstmp                                          As New ADODB.Recordset
    lsvACL.Sorted = False: lsvACL.ListItems.Clear
    lsvACL.Enabled = False
    Set rstmp = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    Set rstmp = gconDMIS.Execute("select ACLNO, ID from CSMS_ACL_HD where ACLNO LIKE '%" & XXX & "%' ORDER BY ACLNO")
    If Not (rstmp.EOF And rstmp.BOF) Then
        Listview_Loadval Me.lsvACL.ListItems, rstmp
        lsvACL.Refresh
        lsvACL.Enabled = True
    End If
    Set rstmp = Nothing
End Sub

Sub FillGrid1()
    Dim rstmp                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set rstmp = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO, PWANO, RO_NO, CUSTOMER, VINNO, ID FROM CSMS_CQIR WHERE STATUS = 'A' ORDER BY ID")
    lsvQIR.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvQIR.ListItems.Add(, , Null2String(rstmp!DLR_CQIR_REFERENCENO))
            Item.SubItems(1) = Null2String(rstmp!PWANO)
            Item.SubItems(2) = Null2String(rstmp!RO_NO)
            Item.SubItems(3) = Null2String(rstmp!Customer)
            Item.SubItems(4) = Null2String(rstmp!VINNO)
            Item.SubItems(5) = rstmp!ID

            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Sub FillSearchGrid1(XXX As String)
    Dim rstmp                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set rstmp = gconDMIS.Execute("SELECT DLR_CQIR_REFERENCENO, PWANO, RO_NO, CUSTOMER, VINNO, ID FROM CSMS_CQIR WHERE DLR_CQIR_REFERENCENO LIKE '%" & XXX & "%' AND STATUS = 'A' ORDER BY ID")
    lsvQIR.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvQIR.ListItems.Add(, , Null2String(rstmp!DLR_CQIR_REFERENCENO))
            Item.SubItems(1) = Null2String(rstmp!PWANO)
            Item.SubItems(2) = Null2String(rstmp!RO_NO)
            Item.SubItems(3) = Null2String(rstmp!Customer)
            Item.SubItems(4) = Null2String(rstmp!VINNO)
            Item.SubItems(5) = rstmp!ID

            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Private Sub cmd1_Click()
    DisableForm True
    picClaims1.ZOrder 1
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_add", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    On Error Resume Next
    ADD_OR_EDIT = "ADD"
    DisablePicture False
    Call initMemvars

    Dim rstmp                                          As New ADODB.Recordset
    Dim COCODE                                         As String
    Set rstmp = gconDMIS.Execute("SELECT COMPANYCODE FROM ALL_PROFILE WHERE MODULENAME = 'CSMS'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        COCODE = Null2String(rstmp!COMPANYCODE)
    End If
    Set rstmp = Nothing

    lblTranno.Caption = GenerateNewTranno
    txtACLno.Text = COCODE & Right(Year(Date), 2) & Format(Month(Date), "00") & "W" & lblTranno.Caption

    picAdds.Visible = False
    picSaves.Visible = True
    txtACLno.SetFocus
End Sub

Private Sub cmdAprovACM_Click()
    Dim vACMNO                                         As String
    Dim vACMDAte                                       As String
    Dim vACMAMOUNT                                     As Double
    Dim vACMREMARKS                                    As String

    If txtACMNO.Text = "" Then
        ShowIsRequiredMsg ("ACM no. Cannot be Blank")
        txtACMNO.SetFocus
        Exit Sub
    End If
    If txtACM_AMOUNT.Text = "" Then
        ShowIsRequiredMsg ("ACM Amount cannot be Blank")
        txtACM_AMOUNT.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtACM_AMOUNT) = False Then
        MsgBox "Invalid ACM Amount", vbExclamation, "CSMS"
        txtACM_AMOUNT.SetFocus
        Exit Sub
    End If

    vACMNO = N2Str2Null(txtACMNO.Text)
    vACMDAte = N2Str2Null(dptACMDATE.Value)
    vACMAMOUNT = NumericVal(txtACM_AMOUNT.Text)
    vACMREMARKS = N2Str2Null(txtREMARKS.Text)

    AUDIT_SQL = "UPDATE CSMS_ACL_HD SET STATUS = 'AP', ACM_NO = " & vACMNO & ",ACM_DATE = " & vACMDAte & _
                ",APPROVED_AMOUNT = " & vACMAMOUNT & ",REMARKS = " & vACMREMARKS & " WHERE ID = " & labID.Caption & ""

    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("AP", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno, "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call ShowSuccessFullyUpdated

    Call rsRefresh
    rsACL.Find "ID = " & labID & ""
    Call StoreMemVars

    Call cmdCancelACM_Click
End Sub

Private Sub cmdCancel_Click()
    DisablePicture True

    picSaves.Visible = False
    picAdds.Visible = True

    StoreMemVars
End Sub

Private Sub cmdCancelACM_Click()
    picAPROVE.Visible = False
    picAPROVE.ZOrder 1
    Call EnabledApprove(True)
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    If MsgBox("Delete This Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "DELETE FROM CSMS_ACL_HD WHERE ID = " & labID.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)
    'NEW LOG AUDIT-----------------------------------------------------------
    Call NEW_LogAudit("X", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------------


    AUDIT_SQL = "UPDATE CSMS_CQIR SET ACLNO = NULL, STATUS = 'A' WHERE ACLNO = '" & Null2String(rsACL!ACLNO) & "'"
    gconDMIS.Execute (AUDIT_SQL)
    'NEW LOG AUDIT-----------------------------------------------------------
    Dim X                                              As Integer
    Dim xID                                            As String

    If Not lsvDET.ListItems.Count = 0 Then
        For X = 1 To lsvDET.ListItems.Count
            xID = FindTransactionID(lsvDET.ListItems(X).Text, "DLR_CQIR_REFERENCENO", "CSMS_CQIR")
            Call NEW_LogAudit("DT", "QUALITY INFORMATION", AUDIT_SQL, xID, "", "ACL no: " & txtACLno, "", "")
        Next
    End If
    'NEW LOG AUDIT-----------------------------------------------------------

    txtSearch.Text = "a": txtSearch.Text = ""
    ShowDeletedMsg
    rsRefresh
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_add", "ACCUMULATED CLAIM LIST") = False Then Exit Sub
    ADD_OR_EDIT = "EDIT"

    DisablePicture False

    picAdds.Visible = False
    picSaves.Visible = True
    txtACLno.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsACL.MoveNext
    If rsACL.EOF Then
        rsACL.MoveLast
        Call ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    If MsgBox("Post This Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "UPDATE CSMS_ACL_HD SET STATUS = 'P' WHERE ID = " & labID.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT--------------------------------------------------------------------------
    Call NEW_LogAudit("P", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno, "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------

    ShowSuccessFullyUpdated
    rsRefresh
    rsACL.Find "ID = " & labID.Caption & ""

    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsACL.MovePrevious
    If rsACL.BOF Then
        rsACL.MoveFirst
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    Dim ITEM_CNT                                       As Integer
    Dim CQIR_CNT                                       As Integer
    Dim rsACL                                          As New ADODB.Recordset
    Dim RSDET                                          As New ADODB.Recordset

    ITEM_CNT = 0
    CQIR_CNT = 0
    If MsgBox("Print ACL Report", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    Screen.MousePointer = 11

    'If MsgBox("Print in Excel", vbQuestion + vbYesNo, "CSMS") = vbYes Then
    'Call PrintExcelACL
    Call PrintACL
    'Else
    '    rptACL.Formulas(0) = "COMPANY_NAME = '" & COMPANY_NAME & "'"
    '    rptACL.Formulas(1) = "COMPANY_ADDRESS = '" & COMPANY_ADDRESS & "'"
    '    rptACL.Formulas(2) = "Printedby = '" & LOGNAME & "'"
    '    rptACL.WindowTitle = "WARRANTY CLAIMS REPORT"
    '    PrintSQLReport rptACL, CSMS_REPORT_PATH & "WARRANTY_CLAIMS.rpt", "{CSMS_ACL_HD.ACLNO} = '" & txtACLno.Text & "'", CSMS_REPORT_CONNECTION, 1
    'End If


    'NEW LOG AUDIT-----------------------------------------------------
    Call NEW_LogAudit("V", "WARRANTY CLAIM", "", labID, "", "ACL no: " & txtACLno, "", "")
    'NEW LOG AUDIT-----------------------------------------------------

    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    Dim vACLNO                                         As String
    Dim vPREPBY                                        As String
    Dim vNOTEDBY                                       As String
    Dim vRECEIVED1                                     As String
    Dim vRECEIVED2                                     As String
    Dim vID                                            As String
    Dim rstmp                                          As New ADODB.Recordset


    Set rstmp = gconDMIS.Execute("SELECT * FROM CSMS_ACL_HD WHERE ACLNO = '" & txtACLno.Text & "'")
    If Not (rstmp.BOF And rstmp.EOF) Then
        If ADD_OR_EDIT = "ADD" Then
            MsgBox "ACL no Already Exist", vbExclamation, "CSMS"
            txtACLno.SetFocus
            Exit Sub
        Else
            If Not rstmp!ID = labID.Caption Then
                MsgBox "ACL no Already Exist", vbExclamation, "CSMS"
                txtACLno.SetFocus
                Exit Sub
            End If
        End If
    End If

    If txtACLno.Text = "" Then
        ShowIsRequiredMsg "ACL no. cannot be Blank"
        txtACLno.SetFocus
        Exit Sub
    End If

    If txtPREPBY.Text = "" Then
        ShowIsRequiredMsg "Prepared By cannot be Blank"
        txtPREPBY.SetFocus
        Exit Sub
    End If

    If txtNotedBy.Text = "" Then
        ShowIsRequiredMsg "Noted By Cannot be Blank"
        txtNotedBy.SetFocus
        Exit Sub
    End If

    '    If txtHARI.Text = "" Then
    '        ShowIsRequiredMsg "HARI Representative cannot be Blank"
    '        txtHARI.SetFocus
    '        Exit Sub
    '    End If

    '    If txtACCT.Text = "" Then
    '        ShowIsRequiredMsg "Accounting Representatvice cannot be blank"
    '        txtACCT.SetFocus
    '        Exit Sub
    '    End If


    vACLNO = N2Str2Null(txtACLno.Text)
    vPREPBY = N2Str2Null(txtPREPBY.Text)
    vNOTEDBY = N2Str2Null(txtNotedBy.Text)
    vRECEIVED1 = N2Str2Null(txtHARI.Text)
    vRECEIVED2 = N2Str2Null(txtACCT.Text)

    If ADD_OR_EDIT = "ADD" Then
        AUDIT_SQL = "INSERT INTO CSMS_ACL_HD (TRANNO, TRANDATE, ACLNO, PREPBY, NOTEDBY, HARI, ACCT, PARTSCLAIM, LABORCLAIM, SUBLETCLAIM) " & _
                  " VALUES('" & lblTranno.Caption & "'," & N2Str2Null(Date) & "," & vACLNO & "," & vPREPBY & "," & vNOTEDBY & "," & vRECEIVED1 & _
                    "," & vRECEIVED2 & "," & 0 & "," & 0 & "," & 0 & ")"
        gconDMIS.Execute (AUDIT_SQL)

        ShowSuccessFullyAdded

        Set rstmp = gconDMIS.Execute("SELECT ID FROM CSMS_ACL_HD WHERE ACLNO = '" & txtACLno.Text & "'")
        If Not (rstmp.BOF And rstmp.EOF) Then
            vID = rstmp!ID
        End If

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Call NEW_LogAudit("A", "WARRANTY CLAIM", AUDIT_SQL, vID, "", "ACL no: " & txtACLno, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------
    Else
        AUDIT_SQL = "UPDATE CSMS_ACL_HD set PREpBY = " & vPREPBY & _
                    ", TRANDATE = " & N2Str2Null(dtpTranDate.Value) & _
                    ", NOTEDBY = " & vNOTEDBY & _
                    ", HARI = " & vRECEIVED1 & _
                    ", ACCT = " & vRECEIVED2 & _
                    ", ACLNO = " & vACLNO & _
                  " WHERE ID = " & labID.Caption & ""
        gconDMIS.Execute (AUDIT_SQL)

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Call NEW_LogAudit("E", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno, "", "")
        'NEW LOG AUDIT----------------------------------------------------------------------------


        AUDIT_SQL = "UPDATE CSMS_CQIR SET ACLNO = " & vACLNO & _
                  " WHERE ACLNO = '" & rsACL!ACLNO & "'"
        gconDMIS.Execute (AUDIT_SQL)

        'NEW LOG AUDIT----------------------------------------------------------------------------
        Dim X                                          As Integer
        Dim xID                                        As String
        Dim xCQIR                                      As String

        If Not lsvDET.ListItems.Count = 0 Then
            For X = 1 To lsvDET.ListItems.Count
                xID = lsvDET.ListItems(X).ListSubItems(5)
                xCQIR = lsvDET.ListItems(X).Text

                Call NEW_LogAudit("EE", "QUALITY INFORMATION", AUDIT_SQL, xID, "", "ACL no: " & txtACCT & " ,DLR no: " & xCQIR, "", "")
            Next
        End If
        'NEW LOG AUDIT----------------------------------------------------------------------------

        ShowSuccessFullyUpdated
        vID = labID.Caption
    End If

    txtSearch.Text = "A": txtSearch.Text = ""
    rsRefresh
    rsACL.MoveFirst

    rsACL.Find "ID = " & vID & ""
    cmdCancel_Click
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_Unpost", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    If lblSTATUS.Caption = "** APPROVED **" Then
        MsgBox "ACL already Approved", vbExclamation, "CSMS"
        Exit Sub
    End If
    If MsgBox("Unpost This Transaction", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub

    AUDIT_SQL = "UPDATE CSMS_ACL_HD SET STATUS = NULL WHERE ID = " & labID.Caption & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT--------------------------------------------------------------------------------
    Call NEW_LogAudit("U", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno, "", "")
    'NEW LOG AUDIT--------------------------------------------------------------------------------

    ShowSuccessFullyUpdated
    Call rsRefresh
    rsACL.Find "ID = " & labID.Caption & ""

    Call StoreMemVars
End Sub

Private Sub cmdx_Click()
    DisableForm True
    picClaims.ZOrder 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If picAdds.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Warranty Claim)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "WARRANTY CLAIM")

        Case vbKeyF3:
            If Not lblSTATUS.Caption = "** POSTED **" Then
                If picAdds.Visible = False Then Exit Sub

                If Function_Access(LOGID, "Acess_add", "ACCUMULATED CLAIM LIST") = False Then Exit Sub
                Screen.MousePointer = 11

                'If COMPANY_CODE = "HBK" Then
                '    Call FillSearchClaims1
                '    DisableForm False
                '    picClaims1.ZOrder 0
                '    TXTSEARCHI2.Text = ""
                '    TXTSEARCHI2.SetFocus
                'Else
                Call FillSearchClaims
                DisableForm False
                picClaims.ZOrder 0
                txtSEARCHI.Text = ""
                txtSEARCHI.SetFocus
                'End If
                Screen.MousePointer = 0
            Else
                MsgBox "ACL is already Posted", vbInformation, "CSMS"
                Exit Sub
            End If

        Case vbKeyF9:
            If lblSTATUS.Caption = "** POSTED **" Then
                If Module_Access(LOGID, "APPROVED ACL", "SYSTEM") = False Then Exit Sub
                Call EnabledApprove(False)
                picAPROVE.Visible = True
                picAPROVE.ZOrder 0
                txtACMNO.SetFocus
            Else


            End If

        Case vbKeyF12:
            If lblSTATUS.Caption = "** APPROVED **" Then
                If Module_Access(LOGID, "DISAPPROVED ACL", "SYSTEM") = False Then Exit Sub
                If MsgBox("Disapproved this ACL", vbQuestion + vbYesNo, "CSMS") = vbYes Then
                    AUDIT_SQL = "UPDATE CSMS_ACL_HD SET STATUS = 'P',ACM_NO = NULL,ACM_DATE = NULL,REMARKS = NULL, APPROVED_AMOUNT = NULL WHERE ID = " & labID.Caption & ""
                    gconDMIS.Execute (AUDIT_SQL)

                    'NEW LOG AUDIT--------------------------------------------------------------------------------
                    Call NEW_LogAudit("DS", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no. " & txtACLno.Text & "", "", "")
                    'NEW LOG AUDIT--------------------------------------------------------------------------------

                    Call rsRefresh
                    rsACL.Find "ID = " & labID & ""
                    Call StoreMemVars
                End If
            End If
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1

    Call initMemvars
    txtSearch.Text = "A": txtSearch.Text = ""

    Call rsRefresh
    Call StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Text12_Change()
    rptCLAIMS.FilterText = txtSEARCHI.Text
    rptCLAIMS.Populate
End Sub

Private Sub lsvACL_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsvACL
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

Private Sub lsvACL_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsACL.MoveFirst
    rsACL.Find "ID = " & Item.ListSubItems(1) & ""

    StoreMemVars
End Sub

Private Sub lsvDET_DblClick()
    If Function_Access(LOGID, "Acess_Delete", "ACCUMULATED CLAIM LIST") = False Then Exit Sub

    Dim Index                                          As Integer
    Dim vID                                            As String
    Dim vTRANNO                                        As String
    Dim vDLRID                                         As String

    If rsACL!Status = "P" Then Exit Sub
    If lsvDET.ListItems.Count = 0 Then Exit Sub

    If MsgBox("Remove QIR Details", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    Index = lsvDET.SelectedItem.Index
    vID = lsvDET.ListItems(Index).ListSubItems(5)
    vTRANNO = Null2String(lsvDET.ListItems(Index).Text)

    AUDIT_SQL = "UPDATE CSMS_CQIR SET CLAIMNO = NULL, ACLNO = NULL, STATUS = 'A', DATEATTACHEDTOACL = NULL WHERE ID = " & vID & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT-----------------------------------------------------------------------------------
    vDLRID = FindTransactionID(N2Str2Null(vTRANNO), "DLR_CQIR_REFERENCENO", "CSMS_CQIR")
    Call NEW_LogAudit("XX", "WARRANTY CLAIM", AUDIT_SQL, labID, "", "ACL no: " & txtACLno & " ,DLR no: " & vTRANNO, "", vDLRID)
    Call NEW_LogAudit("DT", "QUALITY INFORMATION", AUDIT_SQL, vDLRID, "", "ACL no: " & txtACLno & ",DLR no: " & vTRANNO, "", "")
    'NEW LOG AUDIT-----------------------------------------------------------------------------------

    MessagePop InfoFriend, "ACL Information Updated", "Claim Sucessfully Detached to ACL!", 1000
    Call FillACLDetails
End Sub

Private Sub lsvQIR_DblClick()
    If lsvQIR.ListItems.Count = 0 Then: Exit Sub

    Dim Index                                          As Long
    Dim vID                                            As String
    Dim vRONO                                          As String
    Dim VCQIRNO                                        As String
    Dim vPWANO                                         As String
    Dim vACLNO                                         As String
    Dim lng                                            As Long
    Dim vCLAIMNO                                       As String
    Dim vDLRID                                         As String

    Index = lsvQIR.SelectedItem.Index

    vACLNO = N2Str2Null(txtACLno.Text)
    VCQIRNO = N2Str2Null(lsvQIR.ListItems(Index).Text)    'CQIR NO
    vPWANO = N2Str2Null(lsvQIR.ListItems(Index).ListSubItems(1))
    vID = N2Str2Null(lsvQIR.ListItems(Index).ListSubItems(5))    'ID

    vCLAIMNO = GenerateClaimNo

    AUDIT_SQL = "UPDATE CSMS_CQIR SET CLAIMNO = '" & vCLAIMNO & "', STATUS = 'T', ACLNO = " & vACLNO & ", DATEATTACHEDTOACL = '" & Date & "' WHERE ID = " & vID & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT------------------------------------------------------------------------------
    Call NEW_LogAudit("AA", "WARRANTY CLAIM", "", labID, "", "ACL no: " & txtACLno & " ,DLR no: " & VCQIRNO, "", vID)

    vDLRID = FindTransactionID(Null2String(VCQIRNO), "DLR_CQIR_REFERENCENO", "CSMS_CQIR")
    Call NEW_LogAudit("AT", "QUALITY INFORMATION", AUDIT_SQL, vDLRID, "", "ACL no: " & txtACLno & " ,DLR no: " & VCQIRNO, "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------

    Call FillACLDetails

    cmd1_Click
End Sub

Private Sub picPrint_Click()

End Sub

Private Sub rptCLAIMS_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim Index                                          As Long
    Dim vID                                            As String
    Dim vRONO                                          As String
    Dim VCQIRNO                                        As String
    Dim vPWANO                                         As String
    Dim vACLNO                                         As String
    Dim lng                                            As Long
    Dim vCLAIMNO                                       As String
    Dim vDLRID                                         As String

    If Row.Record Is Nothing Then: Exit Sub

    vACLNO = N2Str2Null(txtACLno.Text)
    VCQIRNO = Row.Record(0).Value                     'CQIR NO
    vPWANO = Row.Record(1).Value                      'PWA NO
    vID = Row.Record(5).Value                         'ID

    vCLAIMNO = GenerateClaimNo

    AUDIT_SQL = "UPDATE CSMS_CQIR SET CLAIMNO = '" & vCLAIMNO & "', STATUS = 'T', ACLNO = " & vACLNO & ", DATEATTACHEDTOACL = '" & Date & "' WHERE ID = " & vID & ""
    gconDMIS.Execute (AUDIT_SQL)

    'NEW LOG AUDIT------------------------------------------------------------------------------
    Call NEW_LogAudit("AA", "WARRANTY CLAIM", "", labID, "", "ACL no: " & txtACLno & " ,DLR no: " & VCQIRNO, "", vID)

    vDLRID = FindTransactionID(N2Str2Null(VCQIRNO), "DLR_CQIR_REFERENCENO", "CSMS_CQIR")
    Call NEW_LogAudit("AT", "QUALITY INFORMATION", AUDIT_SQL, vDLRID, "", "ACL no: " & txtACLno & " ,DLR no: " & VCQIRNO, "", "")
    'NEW LOG AUDIT------------------------------------------------------------------------------

    MessagePop InfoFriend, "ACL Information Updated", "Claim Succesfully attached", 1000
    Call FillACLDetails

    cmdx_Click
End Sub

Private Sub Timer1_Timer()
    If lblSTATUS.ForeColor = vbBlack Then
        lblSTATUS.ForeColor = vbRed
    Else
        lblSTATUS.ForeColor = vbBlack
    End If
End Sub

Private Sub txtACM_AMOUNT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890.", KeyAscii)
    End If
End Sub

Private Sub txtACM_AMOUNT_LostFocus()
    If txtACM_AMOUNT.Text = "" Then
        txtACM_AMOUNT.Text = "0"
    Else
        If IsNumeric(txtACM_AMOUNT.Text) = False Then
            MsgBox "Invalid Currency format", vbInformation, "CSMS"
            txtACM_AMOUNT.SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub txtSearch_Change()
    If txtSearch.Text = "" Then
        Call FillGrid
    Else
        Call FillSearchGrid(txtSearch.Text)
    End If
End Sub

Private Sub TXTSEARCHI2_Change()
    If TXTSEARCHI2.Text = "" Then
        Call FillGrid1
    Else
        Call FillSearchGrid1(TXTSEARCHI2.Text)
    End If
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False

    Dim ar()                                           As String
    Dim cWidth                                         As Long
    Dim I                                              As Integer
    Dim scwidth                                        As Long
    ar = Split(SizeArray, ",")
    cWidth = grd.Width

    If TypeOf grd Is ListView Then
        For I = LBound(ar) To UBound(ar)
            If I <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.ColumnHeaders(I + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For I = LBound(ar) To UBound(ar)
            If I < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(I)) / 100)
                grd.Columns(I).Width = scwidth
            End If
        Next

    End If

    Erase ar
    grd.Visible = True
End Sub

