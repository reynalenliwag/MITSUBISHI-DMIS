VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSMatIssuance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Issuance"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MATIssuance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   11475
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2730
      ScaleHeight     =   870
      ScaleWidth      =   8700
      TabIndex        =   71
      Top             =   5160
      Width           =   8700
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
         Left            =   7920
         MouseIcon       =   "MATIssuance.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
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
         Left            =   7200
         MouseIcon       =   "MATIssuance.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
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
         Left            =   6480
         MouseIcon       =   "MATIssuance.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelCO 
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
         Left            =   5760
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATIssuance.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   735
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
         Left            =   5040
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATIssuance.frx":1B43
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":1C95
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   735
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
         Left            =   4320
         MouseIcon       =   "MATIssuance.frx":1FBA
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":210C
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
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
         Left            =   3600
         MouseIcon       =   "MATIssuance.frx":2468
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":25BA
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
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
         Left            =   2880
         MouseIcon       =   "MATIssuance.frx":28CD
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":2A1F
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Move to Last Record"
         Top             =   0
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
         Left            =   2160
         MouseIcon       =   "MATIssuance.frx":2D6F
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":2EC1
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Move to First Record"
         Top             =   0
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
         Left            =   1440
         MouseIcon       =   "MATIssuance.frx":321F
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":3371
         Style           =   1  'Graphical
         TabIndex        =   79
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   735
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
         Left            =   720
         MouseIcon       =   "MATIssuance.frx":366B
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":37BD
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   735
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
         MouseIcon       =   "MATIssuance.frx":3B15
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":3C67
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   2820
      Top             =   4350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   2700
      ScaleHeight     =   2745
      ScaleWidth      =   8685
      TabIndex        =   26
      Top             =   90
      Width           =   8685
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1710
         ScaleHeight     =   405
         ScaleWidth      =   825
         TabIndex        =   61
         Top             =   30
         Width           =   825
         Begin VB.TextBox txtTranType 
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
            MaxLength       =   4
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7080
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   57
         Top             =   510
         Width           =   1515
         Begin MSMask.MaskEdBox txtNetInvAmt 
            Height          =   345
            Left            =   90
            TabIndex        =   58
            Top             =   840
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtTTLInvAmt 
            Height          =   345
            Left            =   90
            TabIndex        =   59
            Top             =   60
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtDS_Amt1 
            Height          =   345
            Left            =   0
            TabIndex        =   60
            Top             =   450
            Width           =   1485
            _ExtentX        =   2619
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
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
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   975
         Width           =   1335
      End
      Begin VB.TextBox txtCustCode 
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
         TabIndex        =   5
         Text            =   "Text1"
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
         Left            =   3540
         MaxLength       =   7
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   570
         Width           =   885
      End
      Begin VB.TextBox txtRONO 
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
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtChargeTo 
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
         Left            =   5640
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   60
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
         Left            =   3540
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   60
         Width           =   885
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
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2310
         Width           =   2535
      End
      Begin VB.ComboBox cboSalesMan 
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
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2310
         Width           =   765
      End
      Begin RichTextLib.RichTextBox txtCustName 
         Height          =   885
         Left            =   60
         TabIndex        =   6
         Top             =   1350
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   1561
         _Version        =   393217
         BackColor       =   16777215
         MaxLength       =   120
         TextRTF         =   $"MATIssuance.frx":3FC6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtDS1 
         Height          =   345
         Left            =   4860
         TabIndex        =   9
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   3
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
      Begin MSMask.MaskEdBox txtTranDate 
         Height          =   345
         Left            =   1170
         TabIndex        =   2
         Top             =   570
         Width           =   1575
         _ExtentX        =   2778
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
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   915
         Left            =   4590
         TabIndex        =   11
         Top             =   1740
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1614
         _Version        =   393217
         BackColor       =   16777215
         TextRTF         =   $"MATIssuance.frx":404E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   29
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
         TabIndex        =   39
         Top             =   2340
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
         TabIndex        =   38
         Top             =   990
         Visible         =   0   'False
         Width           =   525
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
         Left            =   5940
         TabIndex        =   37
         Top             =   600
         Width           =   1215
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
         Left            =   2280
         TabIndex        =   36
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
         Left            =   2790
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Charge To"
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
         Left            =   4620
         TabIndex        =   33
         Top             =   90
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran. #"
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
         TabIndex        =   32
         Top             =   90
         Width           =   825
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   28
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label labPosted 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELLED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   6900
         TabIndex        =   27
         Top             =   60
         Width           =   1635
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   2700
      ScaleHeight     =   2115
      ScaleWidth      =   8685
      TabIndex        =   40
      Top             =   2820
      Width           =   8685
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   1965
         Left            =   60
         TabIndex        =   12
         Top             =   60
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3466
         _Version        =   393216
         Cols            =   7
         BackColor       =   16777215
         ForeColor       =   0
         ForeColorFixed  =   0
         BackColorSel    =   -2147483633
         ForeColorSel    =   0
         BackColorBkg    =   -2147483633
         BorderStyle     =   0
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
   Begin VB.PictureBox fraAddTran 
      Height          =   3525
      Left            =   4680
      ScaleHeight     =   3465
      ScaleWidth      =   4545
      TabIndex        =   41
      Top             =   930
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
         Left            =   3000
         MouseIcon       =   "MATIssuance.frx":40E1
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":4233
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2550
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
         Left            =   1500
         MouseIcon       =   "MATIssuance.frx":455E
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":46B0
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   2550
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
         Left            =   2250
         MouseIcon       =   "MATIssuance.frx":4A00
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":4B52
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   2550
         Width           =   705
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
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   60
         Width           =   615
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
         Top             =   1050
         Width           =   4365
      End
      Begin VB.ComboBox cboTranMatCde 
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
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   420
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtTranQty 
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Top             =   1440
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTranUPrice 
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         Top             =   1800
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTranTotalAmt 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtMaterialID 
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
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   420
         Width           =   585
      End
      Begin VB.Label labMatCde 
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         Left            =   1590
         TabIndex        =   54
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label38 
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
         TabIndex        =   48
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label30 
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
         Left            =   810
         TabIndex        =   47
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label31 
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
         TabIndex        =   46
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label34 
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
         TabIndex        =   45
         Top             =   450
         Width           =   1515
      End
      Begin VB.Label Label35 
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
         TabIndex        =   44
         Top             =   90
         Width           =   855
      End
      Begin VB.Label Label33 
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
         TabIndex        =   43
         Top             =   780
         Width           =   1275
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   3675
      Left            =   4590
      TabIndex        =   63
      Top             =   840
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   6482
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
      MICON           =   "MATIssuance.frx":4E90
   End
   Begin VB.PictureBox fraSignatories 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   4830
      ScaleHeight     =   2025
      ScaleWidth      =   4305
      TabIndex        =   49
      Top             =   1650
      Width           =   4335
      Begin wizButton.cmd cmdPrintMRIS 
         Height          =   345
         Left            =   1080
         TabIndex        =   23
         Top             =   1590
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   609
         TX              =   "&Print MRIS"
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
         MICON           =   "MATIssuance.frx":4EAC
      End
      Begin VB.CheckBox chkPreview 
         BackColor       =   &H00DEDFDE&
         Height          =   255
         Left            =   4020
         TabIndex        =   24
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtApprovedBy 
         Appearance      =   0  'Flat
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
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   780
         Width           =   3045
      End
      Begin VB.TextBox txtRequestedBy 
         Appearance      =   0  'Flat
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
         Top             =   1140
         Width           =   3045
      End
      Begin VB.TextBox txtIssuedBy 
         Appearance      =   0  'Flat
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
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   420
         Width           =   3045
      End
      Begin VB.TextBox txtPreparedBy 
         Appearance      =   0  'Flat
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
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   60
         Width           =   3045
      End
      Begin VB.Label Label12 
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
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label Label13 
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
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label Label14 
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
         TabIndex        =   51
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label15 
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
         TabIndex        =   50
         Top             =   90
         Width           =   1065
      End
   End
   Begin wizButton.cmd cmdSignatories 
      Height          =   2235
      Left            =   4740
      TabIndex        =   64
      Top             =   1560
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3942
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
      MICON           =   "MATIssuance.frx":4EC8
   End
   Begin VB.Frame fraDetails 
      Height          =   5955
      Left            =   60
      TabIndex        =   66
      Top             =   0
      Width           =   2595
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
         TabIndex        =   69
         Top             =   390
         Value           =   -1  'True
         Width           =   2205
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
         TabIndex        =   68
         Top             =   630
         Width           =   2385
      End
      Begin VB.TextBox textSearch 
         Appearance      =   0  'Flat
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
         TabIndex        =   67
         Text            =   "Text1"
         Top             =   960
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstMATISS 
         Height          =   4545
         Left            =   30
         TabIndex        =   70
         Top             =   1350
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   8017
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
         MouseIcon       =   "MATIssuance.frx":4EE4
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
         TabIndex        =   65
         Top             =   150
         Width           =   1455
      End
   End
   Begin SHDocVwCtl.WebBrowser browMRIS 
      Height          =   2625
      Left            =   2760
      TabIndex        =   25
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   9930
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   84
      Top             =   5160
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
         MouseIcon       =   "MATIssuance.frx":5046
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":5198
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Cancel"
         Top             =   0
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
         Left            =   0
         MouseIcon       =   "MATIssuance.frx":54D6
         MousePointer    =   99  'Custom
         Picture         =   "MATIssuance.frx":5628
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Save Entry"
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCSMSMatIssuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMATISS                                           As ADODB.Recordset
Dim rsTDAYTRAN                                         As ADODB.Recordset
Dim rsMatMas                                           As ADODB.Recordset
Dim rsSalesMan                                         As ADODB.Recordset
Dim rsCunter                                           As ADODB.Recordset
Dim rsProfile                                          As ADODB.Recordset
Dim rsSignatories                                      As ADODB.Recordset
Dim rsREPOR                                            As ADODB.Recordset
Dim rsCustomer                                         As ADODB.Recordset
Dim kcnt                                               As Integer
Dim AddorEdit                                          As String
Dim MATISS_TOTUPRICE                                   As Double
Dim MATISS_TOTINVAMT                                   As Double
Dim MATISS_TOTVAT                                      As Double
Dim PrevMatIssType                                     As String
Dim PrevMatIssNo                                       As String

Private Sub cboTranDescription_Click()
    If cboTranDescription.Text <> "" Then
        txtMaterialID.Text = SetMatIDDesc(cboTranDescription.Text)
        cboTranMatCde.Text = Setmatcde(txtMaterialID.Text)
        cboTranDescription.Text = Setmatdsc2(txtMaterialID.Text)
    End If
End Sub

Private Sub cboTranmatcde_Change()
    If cboTranMatCde.Text <> "" Then
        txtMaterialID.Text = SetMatIDmatcde(cboTranMatCde.Text)
        cboTranDescription.Text = Setmatdsc2(txtMaterialID.Text)
    End If
End Sub

Private Sub cboTranmatcde_Click()
    If cboTranMatCde.Text <> "" Then
        txtMaterialID.Text = SetMatIDmatcde(cboTranMatCde.Text)
        cboTranDescription.Text = Setmatdsc2(txtMaterialID.Text)
    End If
End Sub

Private Sub cboTranmatcde_LostFocus()
    If cboTranMatCde.Text <> "" Then
        txtMaterialID.Text = SetMatIDmatcde(cboTranMatCde.Text)
        cboTranDescription.Text = Setmatdsc2(txtMaterialID.Text)
    End If
End Sub

Private Sub cmdAddTran_Click()
    SendToBack
    cmdAddTran.Visible = True
    cmdAddTran.ZOrder 0
    cmdTranDelete.Visible = False
    fraAddTran.Visible = True
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
    AddorEdit = "ADD"
    InitMaterials
    On Error Resume Next
    cboTranMatCde.SetFocus
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "ISSUANCE") = False Then Exit Sub

    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
        Dim PCurOnHand, PCurTISSQTY, PCurIssuances     As Integer
        Dim rsTdaytranDup, rsMatMasDup                 As ADODB.Recordset

        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select id,trantype,tranno,matcde,tranqty from CSMS_TdayTran where tranno = " & N2Str2Null(rsMATISS!Tranno) & " and trantype = " & N2Str2Null(rsMATISS!TRANTYPE), gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsMatMasDup = New ADODB.Recordset
                rsMatMasDup.Open "select matcde,onhand,tissqty,TISSQTY,issuances from CSMS_MatMas where matcde = " & N2Str2Null(rsTdaytranDup!MATCDE), gconDMIS
                If Not rsMatMasDup.EOF And Not rsMatMasDup.BOF Then
                    PCurOnHand = N2Str2IntZero(rsMatMasDup!ONHAND) + N2Str2Zero(rsTdaytranDup!tranqty)
                    PCurTISSQTY = N2Str2IntZero(rsMatMasDup!TISSQTY) - N2Str2Zero(rsTdaytranDup!tranqty)
                    PCurIssuances = N2Str2IntZero(rsMatMasDup!issuances) - N2Str2Zero(rsTdaytranDup!tranqty)
                    gconDMIS.Execute "update CSMS_MatMas set" & _
                                   " onhand = " & PCurOnHand & "," & _
                                   " tissqty = " & PCurTISSQTY & "," & _
                                   " issuances = " & PCurIssuances & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where matcde = " & N2Str2Null(rsTdaytranDup!MATCDE)
                    gconDMIS.Execute "update CSMS_TdayTran set" & _
                                   " status = 'C'," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTdaytranDup!ID
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update CSMS_MatIss set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        rsRefresh
        On Error Resume Next
        rsMATISS.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "ISSUANCE") = False Then Exit Sub

End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_POST", "ISSUANCE") = False Then Exit Sub

    MsgSpeechBox "Posting of Transaction is Automated by OR System or Billing System." & vbCrLf & _
                 "Manual Posting is made only by your System Administrator."
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "ISSUANCE") = False Then Exit Sub
    
    If rsMATISS!TRANTYPE = "MRIS" Then
        cmdSignatories.Visible = True
        cmdSignatories.ZOrder 0
        fraSignatories.Visible = True
        fraSignatories.ZOrder 0
        Set rsSignatories = New ADODB.Recordset
        rsSignatories.Open "Select * from Signatories", gconDMIS
        If Not rsSignatories.EOF And Not rsSignatories.BOF Then
            txtPreparedBy.Text = Null2String(rsSignatories!preparedby)
            txtIssuedBy.Text = Null2String(rsSignatories!issuedby)
            txtRequestedBy.Text = Null2String(rsSignatories!requestedby)
            txtApprovedBy.Text = Null2String(rsSignatories!approvedby)
            On Error Resume Next
            txtRequestedBy.SetFocus
        End If
    End If
    If rsMATISS!TRANTYPE = "CSH" Then
        If MsgQuestionBox("Cash Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CSHPRINTING
        End If
    End If
    If rsMATISS!TRANTYPE = "CHG" Then
        If MsgQuestionBox("Charge Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CHGPRINTING
        End If
    End If
    If rsMATISS!TRANTYPE = "DR" Then
        If MsgQuestionBox("DR Out Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            DRPRINTING
        End If
    End If
End Sub

Sub CHGPRINTING()
    Dim Filter                                         As String
    If MsgBox("Print it in Invoice Format?", vbQuestion + vbYesNo, "Invoice Format Option") = vbYes Then
        If NumericVal(txtDS1.Text) = 0 Then
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "MATCHG.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "MATCHGDisc.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
    Else
        If NumericVal(txtDS1.Text) = 0 Then
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "CHG.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        Else
            Screen.MousePointer = 11
            rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
            rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
            PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "CHGDisc.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
            Screen.MousePointer = 0
        End If
    End If
End Sub

Sub CSHPRINTING()
    Dim Filter                                         As String
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "CSH.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "CSHDisc.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub DRPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        rptCustomerOrder.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptCustomerOrder.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptCustomerOrder, CSMS_REPORT_PATH & "DR.RPT", "{MATIss.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub MRISPRINTING()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Dim Filter                                         As String
    Set rsProfile = New ADODB.Recordset
    rsProfile.Open "select * from ALL_Profile WHERE MODULENAME = 'CSMS'", gconDMIS
    Open CSMS_REPORT_PATH & "MRIS.HTML" For Output As #1
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select tranno,trantype,itemno,matcde,matdsc,tranqty,tranuprice from CSMS_TdayTran where tranno = " & N2Str2Null(rsMATISS!Tranno) & " and trantype = 'MRIS' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0
        If rsTDAYTRAN.RecordCount > 14 Then
            cntCOPY = 4
        Else
            cntCOPY = 2
        End If
        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                rsTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If rsTDAYTRAN.EOF Then
                    rsTDAYTRAN.MoveLast
                Else
                    rsTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & rsProfile!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=4 FACE=TIMES NEW ROMAN><strong>MATERIAL REQUISITION ISSUANCE SLIP</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & Null2String(rsMATISS!TRANTYPE) & "-" & Null2String(rsMATISS!Tranno) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(rsMATISS!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(rsMATISS!custcode) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Charge To: " & Null2String(rsMATISS!chargeto) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(rsMATISS!custname) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Ref RO# : " & Null2String(rsMATISS!rono) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=24%><FONT SIZE=2 FACE=TIMES NEW ROMAN>Material Code</FONT></td>"
            Print #1, "<td width=26%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If rsTDAYTRAN.RecordCount > 14 Then
                cnt2 = 0
            Else
                cnt2 = 14 - rsTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = 14 - (rsTDAYTRAN.RecordCount - 14)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If rsTDAYTRAN.AbsolutePosition > 14 Then
                rsTDAYTRAN.AbsolutePosition = 15
            End If
            Do While Not rsTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=24%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTDAYTRAN!MATCDE) & "</FONT></td>"
                Print #1, "<td width=26%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(rsTDAYTRAN!MatDsc) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(rsTDAYTRAN!tranqty) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(rsTDAYTRAN!tranqty)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If rsTDAYTRAN.AbsolutePosition = 14 Then Exit Do
                rsTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=24%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=26%><FONT SIZE=2>&nbsp;</FONT></td>"
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
                Print #1, "<td width=24%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=26%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=24%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=26%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL MRIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=24%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=26%><FONT SIZE=2>&nbsp;</FONT></td>"
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
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Prepared By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Issued By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Approved By</FONT></td>"
            Print #1, "<td width=25% align=center><FONT SIZE=2 FACE=TIMES NEW ROMAN>Requested By</FONT></td>"
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
        Open CSMS_REPORT_PATH & "MRIS.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
        Else
            Close #1
            browMRIS.Navigate "about:blank"
            browMRIS.Refresh
            browMRIS.Navigate CSMS_REPORT_PATH & "MRIS.HTML"
            DoEvents
            If chkPreview.Value = 1 Then
                browMRIS.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Else
                browMRIS.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
            End If
            Screen.MousePointer = 0
        End If
        Close #1
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdPrintMRIS_Click()
    MRISPRINTING
    SendToBack
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    Dim PnoOnhand, PnoTISSQTY, PnoIssuances            As Integer
    If labDetId.Caption = "" Then
        ShowNothingToDeleteMsg
        Exit Sub
    End If
    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Materials Entry") = True Then
        gconDMIS.Execute "delete from CSMS_TdayTran where id = " & labDetId.Caption
        ShowDeletedMsg
    End If
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select matcde,onhand,TISSQTY,issuances from CSMS_MatMas where matcde = '" & labMatCde.Caption & "'", gconDMIS
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        PnoOnhand = N2Str2IntZero(rsMatMas!ONHAND)
        PnoTISSQTY = N2Str2IntZero(rsMatMas!TISSQTY)
        PnoIssuances = N2Str2IntZero(rsMatMas!issuances)
        gconDMIS.Execute "update CSMS_MatMas set" & _
                       " onhand = " & PnoOnhand + NumericVal(txtTranQty.Text) & "," & _
                       " TISSQTY = " & PnoTISSQTY - NumericVal(txtTranQty.Text) & ", " & _
                       " issuances = " & PnoIssuances - NumericVal(txtTranQty.Text) & _
                       " where matcde = " & N2Str2Null(rsMatMas!MATCDE)
    End If
    FillDetails
    gconDMIS.Execute "update CSMS_MatIss set" & _
                   " ttlinvamt = " & MATISS_TOTUPRICE & "," & _
                   " netinvamt = " & MATISS_TOTINVAMT & _
                   " where id = " & labid.Caption
    rsRefresh
    On Error Resume Next
    rsMATISS.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    If cboTranMatCde.Text = "" Then
        MsgSpeechBox "Material Code must have a value"
        On Error Resume Next
        cboTranMatCde.SetFocus
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,matcde from CSMS_TdayTran where matcde = '" & cboTranMatCde.Text & "' and trantype = '" & txtTranType.Text & "' and tranno =" & N2Str2Null(rsMATISS!Tranno) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Material Code already used in this transaction"
            On Error Resume Next
            cboTranMatCde.SetFocus
            Exit Sub
        End If
    End If

    Dim MATISSTRANDATE, MATISSTRANNO, MATISSTRANTYPE   As String
    Dim MATISSITEMNO, MATISSmatcde, MATISSmatdsc       As String
    Dim MATISSTRANQTY                                  As Integer
    Dim MATISSUNIT                                     As String
    Dim MATISSTRANUCOST                                As Double
    Dim MATISSSTATUS, MATISSIN_OUT                     As String
    Dim MATISSTRANINVAMT                               As Double

    Dim CurONHAND, CurSAFESTOCK, CurTISSQTY            As Integer
    Dim curRESSERVICE, curIssuances, PrevCurOrdQty     As Integer

    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select matcde,onhand,sstock,resservice,TISSQTY,issuances from CSMS_MatMas where matcde = '" & cboTranMatCde.Text & "'", gconDMIS
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        CurONHAND = N2Str2IntZero(rsMatMas!ONHAND)
        CurSAFESTOCK = N2Str2IntZero(rsMatMas!SSTOCK)
        CurTISSQTY = N2Str2IntZero(rsMatMas!TISSQTY)
        curRESSERVICE = N2Str2IntZero(rsMatMas!RESSERVICE)
        curIssuances = N2Str2IntZero(rsMatMas!issuances)
        If AddorEdit <> "ADD" Then
            PrevCurOrdQty = NumericVal(labPrevOrdQty.Caption)
            CurONHAND = CurONHAND + PrevCurOrdQty
            CurTISSQTY = CurTISSQTY - PrevCurOrdQty
            curIssuances = curIssuances - PrevCurOrdQty
        End If
        If CurONHAND <= 0 Then
            Screen.MousePointer = 0
            MsgSpeechBox "Out of Stock!"
            Exit Sub
        End If

        If txtTranType.Text = "CSH" Or txtTranType.Text = "CHG" Then
            If CurONHAND <= curRESSERVICE Then
                Screen.MousePointer = 0
                If MsgQuestionBox("Stock is Reserved for Service... Continue Anyway?", "Stock Status Alert!") = False Then
                    Exit Sub
                End If
            End If
        End If

        If NumericVal(txtTranQty.Text) > CurONHAND Then
            Screen.MousePointer = 0
            MsgSpeechBox "Qty Ordered Exceeds Current Stock!"
            Exit Sub
        Else
            CurONHAND = CurONHAND - NumericVal(txtTranQty.Text)
        End If

        If CurONHAND < CurSAFESTOCK Then
            Screen.MousePointer = 0
            If MsgQuestionBox("Current On-hand is now below the Safety Stock Level... Proceed Anyway?", "Safety Stock Alert!") = False Then
                Exit Sub
            End If
            Screen.MousePointer = 11
        End If
    Else
        Screen.MousePointer = 0
        MsgSpeechBox "Material Code Not Found!"
        Exit Sub
    End If

    MATISSTRANDATE = N2Date2Null(txtTrandate.Text)
    MATISSTRANTYPE = N2Str2Null(txtTranType.Text)
    MATISSTRANNO = N2Str2Null(txtTranNo.Text)
    MATISSITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
    MATISSmatcde = N2Str2Null(cboTranMatCde.Text)
    MATISSmatdsc = N2Str2Null(cboTranDescription.Text)
    MATISSTRANQTY = NumericVal(txtTranQty.Text)
    MATISSTRANINVAMT = NumericVal(txtTranUPrice.Text)
    MATISSIN_OUT = "'O'"
    MATISSSTATUS = "'N'"

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into CSMS_TdayTran " & _
                         "(trandate,trantype,tranno,itemno,matcde,matdsc,tranqty,tranuprice,lastupdate,usercode,status,in_out)" & _
                       " values (" & MATISSTRANDATE & ", " & MATISSTRANTYPE & ", " & MATISSTRANNO & "," & _
                       " " & MATISSITEMNO & "," & MATISSmatcde & "," & _
                       " " & MATISSmatdsc & ", " & MATISSTRANQTY & "," & _
                       " " & MATISSTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & MATISSSTATUS & ", " & MATISSIN_OUT & ")"
    Else
        gconDMIS.Execute "update CSMS_TdayTran set" & _
                       " trandate = " & MATISSTRANDATE & "," & _
                       " trantype = " & MATISSTRANTYPE & "," & _
                       " tranno = " & MATISSTRANNO & "," & _
                       " itemno = " & MATISSITEMNO & "," & _
                       " matcde = " & MATISSmatcde & "," & _
                       " matdsc = " & MATISSmatdsc & "," & _
                       " tranqty = " & MATISSTRANQTY & "," & _
                       " tranuprice = " & MATISSTRANINVAMT & "," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " status = " & MATISSSTATUS & "," & _
                       " in_out = " & MATISSIN_OUT & "," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "" & _
                       " where id = " & labDetId.Caption
    End If
    cleargrid grdDetails
    FillDetails
    gconDMIS.Execute "update CSMS_MatIss set" & _
                   " ttlinvamt = " & MATISS_TOTUPRICE & "," & _
                   " netinvamt = " & MATISS_TOTINVAMT & _
                   " where id = " & labid.Caption
    gconDMIS.Execute "update CSMS_MatMas set" & _
                   " onhand = " & CurONHAND & "," & _
                   " TISSQTY = " & CurTISSQTY + NumericVal(txtTranQty.Text) & ", " & _
                   " issuances = " & curIssuances + NumericVal(txtTranQty.Text) & _
                   " where matcde = '" & cboTranMatCde.Text & "'"
    rsRefresh
    On Error Resume Next
    rsMATISS.Find "id = " & labid.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then cmdAddTran_Click Else cmdTranCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "ISSUANCE") = False Then Exit Sub

    AddorEdit = "ADD"
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next
    'txtCustName.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "ISSUANCE") = False Then Exit Sub

    AddorEdit = "EDIT"
    PrevMatIssType = txtTranType.Text
    PrevMatIssNo = Format(txtTranNo.Text, "000000")
    grdDetails.Enabled = False
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
    'Picture5.Visible = False
    'Dim findStr As String
    'Dim Kim, Joy As String
    'findStr = InputSpeechBox("Please Input Transaction Number or Customer Name ...", txtTranNo.Text)
    'If findStr <> "" Then
    '   On Error Resume Next
    '   rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "tranno", Format(findStr, "000000")).Bookmark
    '   If Err.Number = 3021 Then
    '      For Kim = 1 To Len(findStr)
    '          If Mid(findStr, Kim, 1) <> "-" And Mid(findStr, Kim, 1)) <> "A" Then Joy = Joy & Mid(findStr, Kim, 1)
    '      Next
    '      Joy = "A" & Format(Mid(Joy, 1, Len(Joy)), "000000")
    '      On Error Resume Next
    '      rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "rono", Joy).Bookmark
    '      If Err.Number = 3021 Then
    '         On Error GoTo ErrorCode
    '         rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "custname", findStr).Bookmark
    '      End If
    '   End If
    'End If
    'StoreMemvars
    'Exit Sub

    'ErrorCode:
    'If Err.Number = 3021 Then
    '   ShowCantFind findStr
    '   Resume Next
    'End If
End Sub

Sub FindDupTranno(DDD As String)
    rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    rsMATISS.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    rsMATISS.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsMATISS.MoveNext
    If rsMATISS.EOF Then
        rsMATISS.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsMATISS.MovePrevious
    If rsMATISS.BOF Then
        rsMATISS.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim NextCunter                                     As String
    Dim rsCunter, rsfindDup                            As ADODB.Recordset

    Dim VcboSalesMan, VcboSMName, VTXTTranType         As String
    Dim VTXTTranNo, VTXTTranDate, VtxtCustCode         As String
    Dim VtxtCustName, VTXTChargeTo, VTXTRONO           As String
    Dim VTXTTerms                                      As String
    Dim VTXTTTLInvAmt, VTXTDS1                         As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNetInvAmt                     As Double
    Dim VTXTRemarks, VStatus, Vusercode                As String
    Dim VLastUpdate                                    As String

    If IsNull(txtTranNo.Text) = True Then
        MsgSpeechBox "Transaction Number must not be empty"
        On Error Resume Next
        txtTranNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select trantype,tranno from CSMS_MatIss where trantype = '" & txtTranType.Text & "' and tranno = '" & txtTranNo.Text & "' order by trantype,tranno", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Transaction Number already exist!"
                On Error Resume Next
                txtTranNo.SetFocus
                Exit Sub
            End If
        End If
    End If
    If txtTrandate.Text = "" Or IsDate(txtTrandate.Text) = False Then
        MsgSpeechBox "Invalid Transaction Date!"
        On Error Resume Next
        txtTrandate.SetFocus
        Exit Sub
    End If

    If txtTranType.Text <> "MRIS" And txtTranType.Text <> "CSH" And txtTranType.Text <> "CHG" And txtTranType.Text <> "DR" Then
        MsgSpeechBox "Invalid Transaction Type!"
        On Error Resume Next
        txtTranType.SetFocus
        Exit Sub
    End If
    If txtTranType.Text = "MRIS" Then
        VcboSalesMan = "NULL"
        VcboSMName = "NULL"
    Else
        VcboSalesMan = N2Str2Null(cboSalesMan.Text)
        VcboSMName = N2Str2Null(cboSMName.Text)
    End If

    NextCunter = NumericVal(txtTranNo.Text) + 1

    VTXTTranType = N2Str2Null(txtTranType.Text)
    VTXTTranNo = N2Str2Null(txtTranNo.Text)
    VTXTTranDate = N2Date2Null(txtTrandate.Text)
    VtxtCustCode = N2Str2Null(txtCustCode.Text)
    VtxtCustName = N2Str2Null(txtCustName.Text)
    VTXTChargeTo = N2Str2Null(txtChargeTo.Text)
    VTXTRONO = N2Str2Null(txtROno.Text)
    VTXTTerms = N2Str2Null(txtTerms.Text)
    VTXTTTLInvAmt = NumericVal(txtTTLInvAmt.Text)
    VTXTDS1 = NumericVal(txtDS1.Text)
    VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
    VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
    VTXTNetInvAmt = NumericVal(txtNetInvAmt.Text)
    If txtRemarks.Text = "Pls Type Your Message Here!" Then
        VTXTRemarks = "NULL"
    Else
        VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
    End If
    VStatus = "'N'"
    Vusercode = "" & N2Str2Null(LOGCODE) & ""
    VLastUpdate = "'" & LOGDATE & "'"

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into CSMS_MatIss" & _
                       " (trantype,tranno,trandate,custcode,custname,chargeto,rono,salesman,smname,terms,ttlinvamt,ds1,ds_desc1,ds_amt1,netinvamt,remarks,status,usercode,lastupdate)" & _
                       " values (" & VTXTTranType & ", " & VTXTTranNo & ", " & VTXTTranDate & ", " & _
                       " " & VtxtCustCode & ", " & VtxtCustName & ", " & VTXTChargeTo & _
                         ", " & VTXTRONO & ", " & VcboSalesMan & ", " & VcboSMName & _
                         ", " & VTXTTerms & ", " & VTXTTTLInvAmt & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetInvAmt & ", " & VTXTRemarks & _
                         ", " & VStatus & ", " & Vusercode & ", " & VLastUpdate & ")"
    Else
        gconDMIS.Execute "update CSMS_MatIss set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " tranno = " & VTXTTranNo & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " custcode = " & VtxtCustCode & "," & _
                       " custname = " & VtxtCustName & "," & _
                       " chargeto = " & VTXTChargeTo & "," & _
                       " rono = " & VTXTRONO & "," & _
                       " salesman = " & VcboSalesMan & "," & _
                       " smname = " & VcboSMName & "," & _
                       " terms = " & VTXTTerms & "," & _
                       " ttlinvamt = " & VTXTTTLInvAmt & "," & _
                       " ds1 = " & VTXTDS1 & "," & _
                       " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                       " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                       " netinvamt = " & VTXTNetInvAmt & "," & _
                       " remarks = " & VTXTRemarks & ", " & _
                       " status = " & VStatus & ", " & _
                       " usercode = " & Vusercode & ", " & _
                       " lastupdate = " & VLastUpdate & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update CSMS_TdayTran set" & _
                       " trantype = " & VTXTTranType & "," & _
                       " trandate = " & VTXTTranDate & "," & _
                       " tranno = " & VTXTTranNo & _
                       " where trantype = '" & PrevMatIssType & "' and tranno = '" & PrevMatIssNo & "'"
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update CSMS_Cunter set nextnumber = '" & NextCunter & "', lastupdate = '" & LOGDATE & "', usercode = '" & "USER" & "' where modul = " & VTXTTranType
    Else
        rsRefresh
        On Error Resume Next
        rsMATISS.Find "id =" & labid.Caption
        cmdCancel.Value = True
        cleargrid grdDetails
        FillDetails
        gconDMIS.Execute "update CSMS_MatIss set" & _
                       " ttlinvamt = " & MATISS_TOTUPRICE & "," & _
                       " netinvamt = " & MATISS_TOTINVAMT & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsMATISS.Find "trantype = " & VTXTTranType & " AND tranno = " & VTXTTranNo
    cmdCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddTran_Click
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsMATISS!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                ElseIf Null2String(rsMATISS!Status) = "B" Then
                    MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change"
                Else
                    cmdAddTran_Click
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    textSearch.Text = "":    'Picture5.ZOrder 0
    If MAT_COUNTERTYPE <> "MRIS" Then optRONo.Enabled = False
    'If LOGLEVEL = "MRIS USER" Or LOGLEVEL = "RIV USER" Then
    '    Me.Caption = "Materials Requisition Issuance Slip Data Entry"
    '    Set rsMATISS = New ADODB.Recordset
    '    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'MRIS' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'Else
    'If MAT_COUNTERTYPE = "CSH" Then
    '    Me.Caption = "Cash Counter Issuance Data Entry"
    '    Set rsMATISS = New ADODB.Recordset
    '    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'End If
    'If MAT_COUNTERTYPE = "CHG" Then
    '    Me.Caption = "Charge Counter Issuance Data Entry"
    '    Set rsMATISS = New ADODB.Recordset
    '    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'End If
    'If MAT_COUNTERTYPE = "DR" Then
    '    Me.Caption = "DR Out Issuance Data Entry"
    '    Set rsMATISS = New ADODB.Recordset
    '    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    'End If
    'If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "AUTHOR" Or LOGLEVEL = "MANAGER" Or LOGLEVEL = "ADM" Then
    If MAT_COUNTERTYPE = "CSH" Then
        Me.Caption = "Cash Counter Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "CHG" Then
        Me.Caption = "Charge Counter Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "MRIS" Then
        Me.Caption = "Requisition Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'MRIS' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "DR" Then
        Me.Caption = "DR Out Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    'End If
    'End If
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    initMemvars
    StoreMemVars
    'Drawxpctl me, True, False, True, True
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
'If LOGLEVEL = "MRIS USER" Or LOGLEVEL = "RIV USER" Then
'    Me.Caption = "Materials Requisition Issuance Slip Data Entry"
'    Set rsMATISS = New ADODB.Recordset
'    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'MRIS' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'Else
'If MAT_COUNTERTYPE = "CSH" Then
'    Me.Caption = "Cash Counter Issuance Data Entry"
'    Set rsMATISS = New ADODB.Recordset
'    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'End If
'If MAT_COUNTERTYPE = "CHG" Then
'    Me.Caption = "Charge Counter Issuance Data Entry"
'    Set rsMATISS = New ADODB.Recordset
'    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'End If
'If MAT_COUNTERTYPE = "DR" Then
'    Me.Caption = "DR Out Issuance Data Entry"
'    Set rsMATISS = New ADODB.Recordset
'    rsMATISS.Open "select * from CSMS_MatIss where trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
'End If
'If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "AUTHOR" Or LOGLEVEL = "MANAGER" Then
    If MAT_COUNTERTYPE = "CSH" Then
        Me.Caption = "Cash Counter Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "CHG" Then
        Me.Caption = "Charge Counter Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "MRIS" Then
        Me.Caption = "Requisition Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'MRIS' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    If MAT_COUNTERTYPE = "DR" Then
        Me.Caption = "DR Out Issuance Data Entry"
        Set rsMATISS = New ADODB.Recordset
        rsMATISS.Open "select * from CSMS_MatIss where trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    'End If
    'End If
End Sub

Sub initMemvars()
    If MAT_COUNTERTYPE = "MRIS" Then
        'If LOGLEVEL <> "MRIS USER" And LOGLEVEL <> "RIV USER" And LOGLEVEL <> "SUPERVISOR" And LOGLEVEL <> "MANAGER" And LOGLEVEL <> "AUTHOR" And LOGLEVEL <> "ADM" Then
        '    MsgSpeechBox "Transaction Type Not Allowed."
        '    txtTranType.Text = ""
        '    Exit Sub
        'End If
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from CSMS_Cunter where modul = 'MRIS'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtROno.Enabled = True
        txtTerms.Enabled = False
        cboSalesMan.Enabled = False
        cboSMName.Enabled = False
    End If
    If MAT_COUNTERTYPE = "CSH" Then
        'If LOGLEVEL = "MRIS USER" Then
        '    MsgSpeechBox "Transaction Type Not Allowed"
        '    txtTranType.Text = ""
        '    Exit Sub
        'End If
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from CSMS_Cunter where modul = 'CSH'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtROno.Enabled = False
        txtTerms.Enabled = False
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    If MAT_COUNTERTYPE = "CHG" Then
        'If LOGLEVEL = "MRIS USER" Then
        '    MsgSpeechBox "Transaction Type Not Allowed."
        '    txtTranType.Text = ""
        '    Exit Sub
        'End If
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from CSMS_Cunter where modul = 'CHG'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtROno.Enabled = False
        txtTerms.Enabled = True
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    If MAT_COUNTERTYPE = "DR" Then
        'If LOGLEVEL = "MRIS USER" Then
        '    MsgSpeechBox "Transaction Type Not Allowed."
        '    txtTranType.Text = ""
        '    Exit Sub
        'End If
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from CSMS_Cunter where modul = 'DR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtTranNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
        Else
            txtTranNo.Text = "000001"
        End If
        txtROno.Enabled = False
        txtTerms.Enabled = True
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    txtTrandate.Text = LOGDATE
    txtCustCode.Text = "V00038"
    txtCustName.Text = ""
    txtChargeTo.Text = "VAR"
    txtROno.Text = ""
    txtTerms.Text = ""
    txtTTLInvAmt.Text = ""
    txtDS1.Text = ""
    txtDS_Desc1.Text = ""
    txtDS_Amt1.Text = ""
    txtNetInvAmt.Text = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    InitCbo
    InitGrid
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
    If Not rsMATISS.EOF And Not rsMATISS.BOF Then
        labid.Caption = rsMATISS!ID
        txtTranType.Text = Null2String(rsMATISS!TRANTYPE)
        If txtTranType.Text = "MRIS" Then
            cboSalesMan.Enabled = False
            cboSMName.Enabled = False
        Else
            cboSalesMan.Enabled = True
            cboSMName.Enabled = True
        End If
        txtTranNo.Text = Null2String(rsMATISS!Tranno)
        txtTrandate.Text = Null2String(rsMATISS!trandate)
        txtCustCode.Text = Null2String(rsMATISS!custcode)
        txtCustName.Text = Null2String(rsMATISS!custname)
        txtChargeTo.Text = Null2String(rsMATISS!chargeto)
        txtROno.Text = Null2String(rsMATISS!rono)
        FillSalesMan Null2String(rsMATISS!salesman)
        txtTerms.Text = Null2String(rsMATISS!terms)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATISS!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(rsMATISS!ds1)
        txtDS_Desc1.Text = Null2String(rsMATISS!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2IntZero(rsMATISS!ds_amt1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATISS!netinvamt))
        txtRemarks.Text = Null2String(rsMATISS!remarks)
        If Null2String(rsMATISS!Status) = "C" Then
            labPosted.Visible = True
            labPosted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
        ElseIf Null2String(rsMATISS!Status) = "B" Then
            labPosted.Visible = True
            labPosted.Caption = "BILLED OUT"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
        Else
            labPosted.Visible = False
            cmdEdit.Enabled = True
            If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = True
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
        .ColWidth(6) = 1200
        .Row = 0
        .Col = 1
        .Text = "Item"
        .Col = 2
        .Text = "Material Code"
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
    MATISS_TOTUPRICE = 0
    MATISS_TOTINVAMT = 0
    MATISS_TOTVAT = 0
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select trantype,tranno,id,itemno,matcde,matdsc,tranqty,tranuprice from CSMS_TdayTran where tranno = " & N2Str2Null(rsMATISS!Tranno) & " and trantype = " & N2Str2Null(rsMATISS!TRANTYPE) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        Screen.MousePointer = 11
        rsTDAYTRAN.MoveFirst
        Do While Not rsTDAYTRAN.EOF
            kcnt = kcnt + 1
            grdDetails.AddItem rsTDAYTRAN!ID & Chr(9) & Null2String(rsTDAYTRAN!itemno) & Chr(9) & _
                               Null2String(rsTDAYTRAN!MATCDE) & Chr(9) & _
                               Null2String(rsTDAYTRAN!MatDsc) & Chr(9) & _
                               N2Str2IntZero(rsTDAYTRAN!tranqty) & Chr(9) & _
                               N2Str2Zero(rsTDAYTRAN!TRANUPRICE) & Chr(9) & _
                               Format(N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
            MATISS_TOTUPRICE = MATISS_TOTUPRICE + (N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE))
            MATISS_TOTINVAMT = MATISS_TOTINVAMT + (N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE))
            rsTDAYTRAN.MoveNext
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
        MATISS_TOTINVAMT = MATISS_TOTINVAMT - NumericVal(txtDS_Amt1.Text)
        If kcnt <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    Else
        cleargrid grdDetails
    End If
End Sub

Function FillSalesMan(xxx As String)
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan where empno = '" & xxx & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        rsSalesMan.MoveFirst
        Do While Not rsSalesMan.EOF
            cboSalesMan.AddItem Null2String(rsSalesMan!empno)
            cboSMName.AddItem Null2String(rsSalesMan!signname)
            rsSalesMan.MoveNext
        Loop
        rsSalesMan.MoveFirst
        cboSalesMan.Text = rsSalesMan!empno
        cboSMName.Text = rsSalesMan!signname
    Else
        FillCboSalesMan
    End If
End Function

Sub InitCbo()
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select matcde from CSMS_MatMas order by matcde asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboTranMatCde.Clear
        Do While Not rsMatMas.EOF
            cboTranMatCde.AddItem Null2String(rsMatMas!MATCDE)
            rsMatMas.MoveNext
        Loop
    End If
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select matdsc from CSMS_MatMas order by matdsc asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        rsMatMas.MoveFirst
        cboTranDescription.Clear
        Do While Not rsMatMas.EOF
            cboTranDescription.AddItem Null2String(rsMatMas!MatDsc)
            rsMatMas.MoveNext
        Loop
    End If
End Sub

Sub FillCboSalesMan()
    Set rsSalesMan = New ADODB.Recordset
    rsSalesMan.Open "select empno,signname from PMIS_vw_SalesMan", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesMan.EOF And Not rsSalesMan.BOF Then
        rsSalesMan.MoveFirst
        cboSalesMan.Clear
        cboSMName.Clear
        Do While Not rsSalesMan.EOF
            cboSalesMan.AddItem Null2String(rsSalesMan!empno)
            cboSMName.AddItem Null2String(rsSalesMan!signname)
            rsSalesMan.MoveNext
        Loop
    End If
End Sub

Sub SetCustInfo(rep As String)
'rep = Left(rep, 1) & "-" & Right(rep, 6)
    Set rsREPOR = New ADODB.Recordset
    rsREPOR.Open "select rep_or,niym,acct_no,invoice from CSMS_RepOr where rep_or = '" & rep & "'", gconDMIS
    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
        If Null2String(rsREPOR!Invoice) <> "" Then
            MsgBox "Warning: Repair Order is Already Released!" & vbCrLf & _
                 " Materials Issuance for this Repair Order is not Allowed!", vbCritical, "Critical Issue!"
            On Error Resume Next
            txtROno.SetFocus
            Exit Sub
        End If
        txtCustName.Text = Null2String(rsREPOR!niym)
        txtCustCode.Text = Null2String(rsREPOR!ACCT_NO)
    Else
        txtCustName.Text = ""
        txtCustCode.Text = ""
    End If
End Sub

Private Sub cboTranDescription_LostFocus()
    If cboTranDescription.Text <> "" Then
        txtMaterialID.Text = SetMatIDDesc(cboTranDescription.Text)
        cboTranMatCde.Text = Setmatcde(txtMaterialID.Text)
        cboTranDescription.Text = Setmatdsc2(txtMaterialID.Text)
    End If
End Sub

Function Setmatdsc(ppp As String)
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select matcde,matdsc,s_price from CSMS_MatMas where matcde= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        Setmatdsc = Null2String(rsMatMas!MatDsc)
        txtTranUPrice.Text = N2Str2Zero(rsMatMas!s_price)
    Else
        Setmatdsc = ""
        txtTranUPrice.Text = 0
    End If
End Function

Function Setmatdsc2(pid As Variant)
    If pid <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "Select id,matdsc,s_price from CSMS_MatMas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            Setmatdsc2 = Null2String(rsMatMas!MatDsc)
            txtTranUPrice.Text = N2Str2Zero(rsMatMas!s_price)
        Else
            Setmatdsc2 = ""
            txtTranUPrice.Text = 0
        End If
    End If
End Function

Function Setmatcde(pid As Variant)
    If pid <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "Select id,matcde,s_price from CSMS_MatMas where id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            Setmatcde = Null2String(rsMatMas!MATCDE)
            txtTranUPrice.Text = N2Str2Zero(rsMatMas!s_price)
        Else
            Setmatcde = ""
            txtTranUPrice.Text = 0
        End If
    End If
End Function

Function SetMatIDmatcde(DDD As String)
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matcde from CSMS_MatMas where matcde = '" & DDD & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        SetMatIDmatcde = Null2String(rsMatMas!ID)
    End If
End Function

Function SetMatIDDesc(DDD As String)
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matdsc from CSMS_MatMas where matdsc= '" & DDD & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        SetMatIDDesc = Null2String(rsMatMas!ID)
    End If
End Function

Function SetMatPrice(ppp As String)
    If ppp <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "Select mac,matcde from CSMS_MatMas where matcde = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            SetMatPrice = N2Str2Zero(rsMatMas!Mac)
        End If
    End If
End Function

Sub InitMaterials()
    txtTranItemNo.Text = Format(kcnt + 1, "0000")
    cboTranMatCde.Text = ""
    cboTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUPrice.Text = 0#
    txtTranTotalAmt.Text = 0#
End Sub

Function StoreMaterialsEntry(ByVal ID As Variant)
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,matcde,matdsc,tranqty,itemno,tranuprice from CSMS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        labDetId.Caption = rsTDAYTRAN!ID
        labMatCde.Caption = Null2String(rsTDAYTRAN!MATCDE)
        labPrevOrdQty.Caption = N2Str2IntZero(rsTDAYTRAN!tranqty)
        txtTranItemNo.Text = Null2String(rsTDAYTRAN!itemno)
        cboTranMatCde.Text = Null2String(rsTDAYTRAN!MATCDE)
        cboTranDescription.Text = Null2String(rsTDAYTRAN!MatDsc)
        txtTranQty.Text = N2Str2IntZero(rsTDAYTRAN!tranqty)
        txtTranUPrice.Text = N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
        txtTranTotalAmt.Text = N2Str2Zero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUPRICE)
    End If
End Function

Private Sub grdDetails_DblClick()
    Dim Fild                                           As String
    If Null2String(rsMATISS!Status) = "C" Then
        MsgSpeechBox "Transactions are Already Cancelled, and cannot be Change..."
    ElseIf Null2String(rsMATISS!Status) = "B" Then
        MsgSpeechBox "Transactions are Already Billed-Out, and cannot be Change..."
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        Fild = grdDetails.Text
        If Fild <> "" And Fild <> "No Entry" Then
            AddorEdit = "EDIT"
            cmdTranDelete.Visible = True
            BringToFront
            StoreMaterialsEntry (Fild)
        Else
            MsgSpeechBox "No Entry on Materials!"
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

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRequestedBy_LostFocus()
    txtRequestedBy.Text = txtRequestedBy.Text
End Sub

Private Sub txtRONO_LostFocus()
    Dim RONOStr, RONOstr2, RONOstr3                    As String
    Dim k                                              As Integer
    RONOStr = txtROno.Text
    'If RONOStr <> "" Then
    '    If IsNumeric(RONOStr) = True Then
    '        RONOStr = Format(Left(RONOStr, 1), "A") & Format(Right(RONOStr, 6), "000000")
    '    Else
    '        For k = 1 To Len(RONOStr)
    '            RONOstr2 = Mid(RONOStr, k, 1)
    '            If IsNumeric(RONOstr2) = True Then RONOstr3 = RONOstr3 + RONOstr2
    '        Next
    '        RONOstr3 = Format(RONOstr3, "000000"): RONOStr = Format(Left(RONOstr3, 1), "A") & Format(Right(RONOstr3, 6), "000000")
    '    End If
    SetCustInfo (RONOStr)
    'End If
    txtROno.Text = RONOStr
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub txtTranQty_LostFocus()
    If txtTranQty.Text <> "" Then
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    Else
        txtTranQty.Text = 1
        txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranUPrice.Text))
    End If
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
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

'SEARCH MODULE
Private Sub lstMATISS_GotFocus()
    rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "ID", lstMATISS.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstMATISS_ItemClick(ByVal item As MSComctlLib.ListItem)
    If optTranno.Value = True Then
        If Trim(item) <> "" Then rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "tranno", item).Bookmark
    Else
        rsMATISS.Bookmark = rsFind(rsMATISS.Clone, "ID", lstMATISS.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

Private Sub lstMATISS_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstMATISS
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

Private Sub lstMATISS_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstMATISS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then: On Error Resume Next: textSearch.SetFocus
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
    Dim rsMATISS                                       As ADODB.Recordset
    lstMATISS.Enabled = False
    lstMATISS.Sorted = False: lstMATISS.ListItems.Clear
    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select Tranno,ID from CSMS_MatIss where trantype = '" & MAT_COUNTERTYPE & "' order by Tranno asc")
    If Not (rsMATISS.EOF And rsMATISS.BOF) Then
        lstMATISS.Enabled = True
        Listview_Loadval Me.lstMATISS.ListItems, rsMATISS
        lstMATISS.Refresh
        lstMATISS.Enabled = True
    Else
        lstMATISS.Enabled = False
        lstMATISS.Enabled = True
    End If
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsMATISS                                       As ADODB.Recordset
    lstMATISS.Enabled = False
    lstMATISS.Sorted = False: lstMATISS.ListItems.Clear
    Set rsMATISS = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsMATISS = gconDMIS.Execute("select tranno, ID from CSMS_MatIss where trantype = '" & MAT_COUNTERTYPE & "' and tranno like '" & xxx & "%'")
    If Not (rsMATISS.EOF And rsMATISS.BOF) Then
        lstMATISS.Enabled = True
        Listview_Loadval Me.lstMATISS.ListItems, rsMATISS
        lstMATISS.Refresh
        lstMATISS.Enabled = True
    Else
        lstMATISS.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsMATISS                                       As ADODB.Recordset
    lstMATISS.Sorted = False: lstMATISS.ListItems.Clear
    Set rsMATISS = New ADODB.Recordset
    Set rsMATISS = gconDMIS.Execute("select rono,ID from CSMS_MatIss where trantype = '" & MAT_COUNTERTYPE & "' and rono is not null order by tranno asc")
    If Not (rsMATISS.EOF And rsMATISS.BOF) Then
        lstMATISS.Enabled = True
        Listview_Loadval Me.lstMATISS.ListItems, rsMATISS
        lstMATISS.Refresh
         lstMATISS.Enabled = True
    Else
        lstMATISS.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsMATISS                                       As ADODB.Recordset
    lstMATISS.Sorted = False: lstMATISS.ListItems.Clear
    lstMATISS.Enabled = False
    Set rsMATISS = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsMATISS = gconDMIS.Execute("select Rono, ID from CSMS_MatIss where trantype = '" & MAT_COUNTERTYPE & "' and rono like '" & xxx & "%' order by tranno asc")
    If Not (rsMATISS.EOF And rsMATISS.BOF) Then
        lstMATISS.Enabled = True
        Listview_Loadval Me.lstMATISS.ListItems, rsMATISS
        lstMATISS.Refresh
        lstMATISS.Enabled = True
    Else
        lstMATISS.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstMATISS.Enabled = True And lstMATISS.ListItems.Count > 0 Then
            lstMATISS.SetFocus
        End If
    End If
End Sub

Private Sub optRONo_Click()
    lstMATISS.ColumnHeaders(1).Text = "RO Number"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optTranno_Click()
    lstMATISS.ColumnHeaders(1).Text = "Tran. No."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub
