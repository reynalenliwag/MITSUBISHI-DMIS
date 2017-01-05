VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISInquiry_CustomerOrderHist_AC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Order History"
   ClientHeight    =   7125
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11475
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_CustomerOrderHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11475
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   11475
      TabIndex        =   93
      Top             =   6780
      Width           =   11475
      Begin VB.Label labORNo 
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
         Left            =   870
         TabIndex        =   100
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " OR #:"
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
         Left            =   0
         TabIndex        =   99
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SJ #:"
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
         Index           =   0
         Left            =   2010
         TabIndex        =   98
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labSJ 
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
         Left            =   2880
         TabIndex        =   97
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label labinvNo 
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
         Left            =   4860
         TabIndex        =   96
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Inv #:"
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
         Index           =   1
         Left            =   4020
         TabIndex        =   95
         Top             =   0
         Width           =   825
      End
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
         Left            =   5970
         TabIndex        =   94
         Top             =   0
         Width           =   5445
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   2970
      ScaleHeight     =   870
      ScaleWidth      =   8550
      TabIndex        =   76
      Top             =   5880
      Width           =   8550
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
         Left            =   7680
         MouseIcon       =   "AC_CustomerOrderHist.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   60
         Width           =   765
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
         Left            =   6930
         MouseIcon       =   "AC_CustomerOrderHist.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton cmdCancelCO 
         Caption         =   "Cancel Transaction"
         Enabled         =   0   'False
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
         Left            =   6180
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "AC_CustomerOrderHist.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Entry"
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
         Left            =   5430
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "AC_CustomerOrderHist.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Press F11 for Posting By Range"
         Top             =   60
         Width           =   765
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
         MouseIcon       =   "AC_CustomerOrderHist.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   60
         Width           =   765
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
         Left            =   3930
         MouseIcon       =   "AC_CustomerOrderHist.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":213D
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   60
         Width           =   765
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
         MouseIcon       =   "AC_CustomerOrderHist.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   60
         Width           =   765
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
         Left            =   2430
         MouseIcon       =   "AC_CustomerOrderHist.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   60
         Width           =   765
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
         Left            =   1680
         MouseIcon       =   "AC_CustomerOrderHist.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":2EF4
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   60
         Width           =   765
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
         Left            =   930
         MouseIcon       =   "AC_CustomerOrderHist.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   60
         Width           =   765
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
         Left            =   180
         MouseIcon       =   "AC_CustomerOrderHist.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   60
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   10050
      ScaleHeight     =   795
      ScaleWidth      =   1410
      TabIndex        =   73
      Top             =   5880
      Width           =   1470
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
         MouseIcon       =   "AC_CustomerOrderHist.frx":3B49
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":3C9B
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   0
         Width           =   705
      End
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
         Left            =   705
         MouseIcon       =   "AC_CustomerOrderHist.frx":3FEB
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":413D
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   2700
      ScaleHeight     =   2745
      ScaleWidth      =   8745
      TabIndex        =   26
      Top             =   75
      Width           =   8745
      Begin VB.ComboBox cboChargeTo 
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
         Left            =   5550
         TabIndex        =   9
         Text            =   "cboChargeTo"
         ToolTipText     =   "Select option from list."
         Top             =   60
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
         Height          =   885
         Left            =   4650
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         ToolTipText     =   "Type your message or remarks."
         Top             =   1725
         Width           =   3945
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
         Height          =   915
         Left            =   60
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         ToolTipText     =   "Type complete name of customer."
         Top             =   1350
         Width           =   4545
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
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   2
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
         TabIndex        =   10
         ToolTipText     =   "Type percentage to be added in the total amount. Do not include percent sign (e.g. 10, 15)"
         Top             =   945
         Width           =   525
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1740
         ScaleHeight     =   405
         ScaleWidth      =   765
         TabIndex        =   58
         Top             =   0
         Width           =   765
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
            Left            =   0
            MaxLength       =   3
            TabIndex        =   59
            Top             =   60
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1215
         Left            =   7110
         ScaleHeight     =   1215
         ScaleWidth      =   1515
         TabIndex        =   57
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
            TabIndex        =   64
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
            Left            =   30
            MaxLength       =   15
            TabIndex        =   63
            Top             =   440
            Width           =   1455
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
            TabIndex        =   62
            Top             =   60
            Width           =   1395
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
         TabIndex        =   11
         ToolTipText     =   "Input the type of the added amount."
         Top             =   950
         Width           =   1365
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
         Left            =   3660
         MaxLength       =   6
         TabIndex        =   5
         ToolTipText     =   "Input customer code (e.g. S01163)"
         Top             =   960
         Width           =   945
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
         Left            =   3660
         MaxLength       =   7
         TabIndex        =   3
         ToolTipText     =   "Type the transaction terms."
         Top             =   570
         Width           =   945
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
         Left            =   1140
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Type the transactin's RO number (e.g. A007541)"
         Top             =   960
         Width           =   1425
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
         Left            =   3300
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
         Top             =   60
         Width           =   975
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
         Left            =   1080
         TabIndex        =   8
         Text            =   "cboSMName"
         ToolTipText     =   "Select name of salesman from the list."
         Top             =   2310
         Width           =   3525
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
         TabIndex        =   7
         Text            =   "cboSalesMan"
         Top             =   2310
         Width           =   765
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         Top             =   2325
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
         Left            =   3810
         TabIndex        =   38
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
         Left            =   5370
         TabIndex        =   37
         Top             =   600
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
         Left            =   2580
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
         Left            =   3060
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
         Left            =   4560
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
         Left            =   2550
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
         TabIndex        =   27
         Top             =   90
         Width           =   1725
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   2700
      ScaleHeight     =   3045
      ScaleWidth      =   8745
      TabIndex        =   40
      Top             =   2820
      Width           =   8745
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   8010
         Top             =   1440
      End
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2985
         Left            =   30
         TabIndex        =   13
         Top             =   60
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   5265
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
   Begin VB.PictureBox fraAddTran 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   4710
      ScaleHeight     =   4425
      ScaleWidth      =   4575
      TabIndex        =   41
      Top             =   930
      Width           =   4605
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
         Left            =   2130
         MouseIcon       =   "AC_CustomerOrderHist.frx":447B
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":45CD
         Style           =   1  'Graphical
         TabIndex        =   91
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
         Left            =   1425
         MouseIcon       =   "AC_CustomerOrderHist.frx":490B
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":4A5D
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   2550
         Width           =   705
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
         Left            =   2850
         MouseIcon       =   "AC_CustomerOrderHist.frx":4DAD
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":4EFF
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   2550
         Width           =   705
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
         MaxLength       =   10
         TabIndex        =   16
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   420
         Width           =   585
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
         Left            =   1500
         TabIndex        =   54
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   1470
         Width           =   915
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part No."
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
         Left            =   570
         TabIndex        =   45
         Top             =   450
         Width           =   855
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   810
         Width           =   1275
      End
   End
   Begin Crystal.CrystalReport rptCustomerOrder 
      Left            =   2880
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   4605
      Left            =   4650
      TabIndex        =   60
      Top             =   870
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   8123
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
      MICON           =   "AC_CustomerOrderHist.frx":522A
   End
   Begin VB.PictureBox fraSignatories 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   4800
      ScaleHeight     =   2295
      ScaleWidth      =   4305
      TabIndex        =   49
      Top             =   1560
      Width           =   4335
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
         Left            =   1200
         MouseIcon       =   "AC_CustomerOrderHist.frx":5246
         MousePointer    =   99  'Custom
         Picture         =   "AC_CustomerOrderHist.frx":5398
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1530
         Width           =   855
      End
      Begin VB.CheckBox chkPreview 
         Height          =   255
         Left            =   4020
         TabIndex        =   24
         Top             =   1680
         Width           =   375
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   90
         Width           =   1065
      End
   End
   Begin wizButton.cmd cmdSignatories 
      Height          =   2475
      Left            =   4710
      TabIndex        =   61
      Top             =   1500
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4366
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
      MICON           =   "AC_CustomerOrderHist.frx":56FE
   End
   Begin VB.Frame fraDetails 
      Height          =   6795
      Left            =   60
      TabIndex        =   68
      Top             =   -60
      Width           =   2595
      Begin VB.OptionButton Option1 
         Caption         =   "Customer Name"
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
         TabIndex        =   92
         Top             =   930
         Width           =   2205
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
         TabIndex        =   71
         Top             =   390
         Value           =   -1  'True
         Width           =   2385
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
         TabIndex        =   70
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
         TabIndex        =   69
         Text            =   "Text1"
         Top             =   1230
         Width           =   2475
      End
      Begin MSComctlLib.ListView lstOrd_Hd 
         Height          =   5115
         Left            =   30
         TabIndex        =   72
         Top             =   1620
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   9022
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "AC_CustomerOrderHist.frx":571A
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
         TabIndex        =   67
         Top             =   150
         Width           =   1455
      End
   End
   Begin SHDocVwCtl.WebBrowser browRIV 
      Height          =   2625
      Left            =   2820
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
      Location        =   $"AC_CustomerOrderHist.frx":587C
   End
End
Attribute VB_Name = "frmPMISInquiry_CustomerOrderHist_AC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSORD_HIST, RSDAYTRAN, RSPARTMAS, RSTDAYTRAN       As ADODB.Recordset
Attribute RSDAYTRAN.VB_VarUserMemId = 1073938432
Attribute RSPARTMAS.VB_VarUserMemId = 1073938432
Attribute RSTDAYTRAN.VB_VarUserMemId = 1073938432
Dim RSSALESMAN, RSCUNTER, RSPROFILE                    As ADODB.Recordset
Attribute RSSALESMAN.VB_VarUserMemId = 1073938436
Attribute RSCUNTER.VB_VarUserMemId = 1073938436
Attribute RSPROFILE.VB_VarUserMemId = 1073938436
Dim rsSignatories                                      As ADODB.Recordset
Attribute rsSignatories.VB_VarUserMemId = 1073938439
Dim RSREPOR, rsCustomer                                As ADODB.Recordset
Attribute RSREPOR.VB_VarUserMemId = 1073938440
Attribute rsCustomer.VB_VarUserMemId = 1073938440
Dim KCNT                                               As Integer
Attribute KCNT.VB_VarUserMemId = 1073938442
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938443
Dim ORD_TOTUPRICE, ORD_TOTINVAMT, ORD_TOTVAT, ORD_TOTQTY As Double
Attribute ORD_TOTUPRICE.VB_VarUserMemId = 1073938444
Attribute ORD_TOTINVAMT.VB_VarUserMemId = 1073938444
Attribute ORD_TOTVAT.VB_VarUserMemId = 1073938444
Attribute ORD_TOTQTY.VB_VarUserMemId = 1073938444
Dim PREVORDTYPE, PREVORDNO                             As String
Attribute PREVORDTYPE.VB_VarUserMemId = 1073938448
Attribute PREVORDNO.VB_VarUserMemId = 1073938448

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
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC,srp from PMIS_STOCKMAS where type = 'A' and STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
        txtTranUPrice.Text = N2Str2Zero(RSPARTMAS!SRP)
    End If
End Function

Function SetSTOCKDESC2(pid As Variant)
    If pid <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select id,STOCKDESC,srp from PMIS_STOCKMAS where type = 'A' and id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
            txtTranUPrice.Text = Format(N2Str2Zero(RSPARTMAS!SRP), MAXIMUM_DIGIT)
        Else
            txtTranUPrice.Text = "0.00"
        End If
    End If
End Function

Function SetSTOCKNO(pid As Variant)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO,srp from PMIS_STOCKMAS where type = 'A' and id = " & pid, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
        txtTranUPrice.Text = Format(N2Str2Zero(RSPARTMAS!SRP), MAXIMUM_DIGIT)
    Else
        txtTranUPrice.Text = "0.00"
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_STOCKMAS where type = 'A' and STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_STOCKMAS where ltrim(rtrim(STOCKDESC)) = " & LTrim(RTrim(DDD)) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select srp,STOCKNO from PMIS_STOCKMAS where type = 'A' and STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetPartPrice = Format(N2Str2Zero(RSPARTMAS!SRP), MAXIMUM_DIGIT)
        End If
    End If
End Function

Function StorePartsEntry(ByVal ID As Variant)
    Set RSDAYTRAN = New ADODB.Recordset
    RSDAYTRAN.Open "select id,STOCK_ORD,STOCK_SUP,tranqty,itemno,tranuprice from PMIS_DayTran where type = 'A' and id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
        labDetID.Caption = RSDAYTRAN!ID
        labPartNo.Caption = Null2String(RSDAYTRAN!STOCK_ORD)
        labPrevOrdQty.Caption = N2Str2IntZero(RSDAYTRAN!TRANQTY)
        txtTranItemNo.Text = Format(Null2String(RSDAYTRAN!itemno), "0000")
        cboTranPartNo.Text = Null2String(RSDAYTRAN!STOCK_ORD)
        txtTranDescription.Text = SetSTOCKDESC(RSDAYTRAN!STOCK_ORD)
        txtTranQty.Text = N2Str2IntZero(RSDAYTRAN!TRANQTY)
        txtTranUPrice.Text = ToDoubleNumber(N2Str2Zero(RSDAYTRAN!TRANUPRICE))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2Zero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUPRICE))
    End If
End Function

Sub FindDupTranno(DDD As String)
    On Error Resume Next
    RSORD_HIST.Bookmark = rsFind(RSORD_HIST.Clone, "tranno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub CHGPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHG_HIst.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CHGDisc_Hist.RPT", "{ord_hd.TRANTYPE} = 'CHG' and {ord_hd.TYPE} = 'A' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub CSHPRINTING()
    If NumericVal(txtDS1.Text) = 0 Then
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSH_Hist.RPT", "{ord_hd.TYPE} = 'A' AND {ord_hd.TRANTYPE} = 'CSH' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    Else
        Screen.MousePointer = 11
        PrintSQLReport rptCustomerOrder, PMIS_REPORT_PATH & "CSHDisc_Hist.RPT", "{ord_hd.TYPE} = 'A' AND {ord_hd.TRANTYPE} = 'CSH' AND {ord_hd.tranno} = " & N2Str2Null(txtTranNo.Text), DMIS_REPORT_Connection, 1
        Screen.MousePointer = 0
    End If
End Sub

Sub SERVICEPISPRINTING()
    Screen.MousePointer = 11
    Dim cnt1, cnt2, cnt3                               As Integer
    Dim knt, cntCOPY                                   As Integer
    Dim TOTALQTY, TOTALPRICE                           As Double
    Set RSPROFILE = New ADODB.Recordset
    RSPROFILE.Open "select * from ALL_Profile where ModuleName = 'PMIS'", gconDMIS
    Open App.Path & "\PIS.HTML" For Output As #1
    Set RSTDAYTRAN = New ADODB.Recordset
    RSTDAYTRAN.Open "select tranno,trantype,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_AlldayTran where TYPE = 'A' AND tranno = " & N2Str2Null(RSORD_HIST!TRANNO) & " and trantype = 'RIV' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSTDAYTRAN.EOF And Not RSTDAYTRAN.BOF Then
        TOTALQTY = 0
        TOTALPRICE = 0
        cntCOPY = 1
        Print #1, "<html><body>"
        knt = 0
        For knt = 1 To cntCOPY
            If knt < 3 Then
                RSTDAYTRAN.MoveFirst
                TOTALQTY = 0: TOTALPRICE = 0
            Else
                If RSTDAYTRAN.EOF Then
                    RSTDAYTRAN.MoveLast
                Else
                    RSTDAYTRAN.MoveNext
                End If
            End If
            Print #1, "<table width=100% cellspacing=0 cellpadding=0>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNDATE: " & Format(LOGDATE, "MM/DD/YYYY") & "</font></td>"
            Print #1, "<td align=center width=60%><font size=3 FACE=TIMES NEW ROMAN>" & RSPROFILE!CompanyName & "</font></td>"
            Print #1, "<td align=right width=20%><font size=1 FACE=TIMES NEW ROMAN>COPY: " & knt & "</font></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td align=left width=20%><font size=1 FACE=TIMES NEW ROMAN>RUNTIME: " & Time & "</font></td>"
            Print #1, "<td align=center width=60%><font size=5 FACE=TIMES NEW ROMAN><strong>ACCESSORIES ISSUANCE SLIP</strong></font></td>"
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
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Number:</b></FONT><FONT SIZE=3 FACE=TIMES NEW ROMAN><b><i><u>" & "SERVICE PIS-" & Null2String(RSORD_HIST!TRANNO) & "</b></i></u></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN<b>Transaction Date: " & Null2String(RSORD_HIST!trandate) & "</b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>Customer: " & Null2String(RSORD_HIST!CUSTCODE) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "<tr>"
            Print #1, "<td width=60%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b>" & Null2String(RSORD_HIST!CUSTNAME) & "</b></FONT></td>"
            Print #1, "<td width=40%><FONT SIZE=2 FACE=TIMES NEW ROMAN><b></b></FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ITM #</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>ACCESSORIES CODE</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>DESCRIPTION</FONT></td>"
            Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>QTY</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>UNIT PRICE</FONT></td>"
            Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>TOTAL PRICE</FONT></td>"
            Print #1, "</tr>"
            Print #1, "</table>"
            Print #1, "-------------------------------------------------------------------------------------------------------------------------------------------"
            cnt1 = 0
            If RSTDAYTRAN.RecordCount > MAX_ISS_LINE Then
                cnt2 = 0
            Else
                cnt2 = MAX_ISS_LINE - RSTDAYTRAN.RecordCount
            End If
            If knt >= 3 Then cnt2 = MAX_ISS_LINE - (RSTDAYTRAN.RecordCount - MAX_ISS_LINE)
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            If RSTDAYTRAN.AbsolutePosition > MAX_ISS_LINE Then
                RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE + 1
            End If
            Do While Not RSTDAYTRAN.EOF
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!itemno) & "</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Null2String(RSTDAYTRAN!STOCK_ORD) & "</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & SetSTOCKDESC(Null2String(RSTDAYTRAN!STOCK_SUP)) & "</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & N2Str2IntZero(RSTDAYTRAN!TRANQTY) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=2 FACE=TIMES NEW ROMAN>" & Format(N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & "</FONT></td>"
                If knt <> 4 Then
                    TOTALQTY = TOTALQTY + N2Str2IntZero(RSTDAYTRAN!TRANQTY)
                    TOTALPRICE = TOTALPRICE + N2Str2Zero(RSTDAYTRAN!TRANQTY) * N2Str2Zero(RSTDAYTRAN!TRANUPRICE)
                End If
                Print #1, "</tr>"
                If RSTDAYTRAN.AbsolutePosition = MAX_ISS_LINE Then Exit Do
                RSTDAYTRAN.MoveNext
            Loop
            For cnt3 = 1 To cnt2
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
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
                Print #1, "<td width=15%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=8%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=5>&nbsp;</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            Else
                Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
                Print #1, "<tr>"
                Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td width=35%><FONT SIZE=2 FACE=TIMES NEW ROMAN>*** TOTAL AIS</FONT></td>"
                Print #1, "<td align=right width=8%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & TOTALQTY & "</FONT></td>"
                Print #1, "<td width=17%><FONT SIZE=2>&nbsp;</FONT></td>"
                Print #1, "<td align=right width=17%><FONT SIZE=3 FACE=TIMES NEW ROMAN>" & Format(TOTALPRICE, MAXIMUM_DIGIT) & "</FONT></td>"
                Print #1, "</tr>"
                Print #1, "</table>"
            End If
            Print #1, "<table border=0 cellpadding=0 cellspacing=0 width=100%>"
            Print #1, "<tr>"
            Print #1, "<td width=10%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=15%><FONT SIZE=2>&nbsp;</FONT></td>"
            Print #1, "<td width=35%><FONT SIZE=2>&nbsp;</FONT></td>"
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
        Open App.Path & "\PIS.HTML" For Input As #1
        If EOF(1) Then
            MsgSpeechBox "File Not Found!"
            MsgBoxXP "File Not Found!", "Error", XP_OKOnly, msg_Critical
        Else
            Close #1
            browRIV.Navigate "about:blank"
            browRIV.Refresh
            browRIV.Navigate App.Path & "\PIS.HTML"
            DoEvents
            browRIV.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
            Screen.MousePointer = 0
        End If
    End If
    Set RSPROFILE = Nothing
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    If LOGLEVEL = "RIV USER" Then
        If COUNTERTYPE = "RIV" Then
            Me.Caption = "Requisition Issuance Data Entry"
            Set RSORD_HIST = New ADODB.Recordset
            RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'RIV' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        End If
        InitCboChargeToWarehouse
    Else
        If LOGLEVEL = "SUPERVISOR" Or LOGLEVEL = "MANAGER" Or LOGLEVEL = "AUTHOR" Or LOGLEVEL = "ADM" Then
            If COUNTERTYPE = "CSH" Then
                Me.Caption = "Cash Counter Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If COUNTERTYPE = "CHG" Then
                Me.Caption = "Charge Counter Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If COUNTERTYPE = "RIV" Then
                Me.Caption = "Requisition Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'RIV' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If COUNTERTYPE = "DR" Then
                Me.Caption = "DR Out Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If COUNTERTYPE = "CSH" Then
                Me.Caption = "Cash Counter Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'CSH' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If COUNTERTYPE = "CHG" Then
                Me.Caption = "Charge Counter Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'CHG' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
            If COUNTERTYPE = "DR" Then
                Me.Caption = "DR Out Issuance Data Entry"
                Set RSORD_HIST = New ADODB.Recordset
                RSORD_HIST.Open "select * from PMIS_Ord_Hist where type = 'A' and trantype = 'DR' order by tranno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        InitCboChargeToCounter
    End If
End Sub

Sub InitCboChargeToWarehouse()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.Text = "MECHANICAL"
End Sub

Sub InitCboChargeToCounter()
    cboChargeTo.Clear
    cboChargeTo.AddItem "MECHANICAL"
    cboChargeTo.AddItem "WARRANTY"
    cboChargeTo.AddItem "TINSMITH"
    cboChargeTo.AddItem "VARIOUS"
    cboChargeTo.AddItem "FLEET"
    cboChargeTo.Text = "VARIOUS"
End Sub

Sub initMemvars()
    If COUNTERTYPE = "RIV" Then
        txtRONO.Enabled = True
        txtTerms.Enabled = False
        cboSalesMan.Enabled = False
        cboSMName.Enabled = False
    End If
    If COUNTERTYPE = "CSH" Then
        txtRONO.Enabled = False
        txtTerms.Enabled = False
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    If COUNTERTYPE = "CHG" Then
        txtRONO.Enabled = False
        txtTerms.Enabled = True
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    If COUNTERTYPE = "DR" Then
        txtRONO.Enabled = False
        txtTerms.Enabled = True
        cboSalesMan.Enabled = True
        cboSMName.Enabled = True
    End If
    If COUNTERTYPE = "ADB" Then
        txtRONO.Enabled = True
        txtTerms.Enabled = False
        cboSalesMan.Enabled = False
        cboSMName.Enabled = False
    End If
    txtTranDate.Text = LOGDATE
    txtCustCode.Text = "V00038"
    txtCustName.Text = ""
    txtChargeTo.Text = "VAR"
    txtRONO.Text = ""
    txtTerms.Text = ""
    txtTTLInvAmt.Text = "0.00"
    txtDS1.Text = "0"
    txtDS_Desc1.Text = "0.00"
    txtDS_Amt1.Text = "0.00"
    txtNetInvAmt.Text = "0.00"
    txtremarks.Text = "Pls Type Your Message Here!"
    labPosted.Caption = ""
    'InitCbo
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
    If Not RSORD_HIST.EOF And Not RSORD_HIST.BOF Then
        labid.Caption = RSORD_HIST!ID
        labSJ = "": labORNo = "": labDetails = "": labinvNo = ""
        labinvNo = Null2String(RSORD_HIST!TRANNO)
        labORNo = CheckORNum(Null2String(RSORD_HIST!TRANNO), "AI", COUNTERTYPE)
        labSJ = CheckSJNum(Null2String(RSORD_HIST!TRANNO), "AI")
        If labORNo = "" And labSJ = "" Then
            labDetails = ""
        ElseIf labORNo = "" And labSJ <> "" Then
            labDetails = "Imported Sales Journal"
        ElseIf labORNo <> "" And labSJ = "" Then
            labDetails = "OR Issued"
        Else
            labDetails = "OR Issued/Journal Posted"
        End If

        txtTranType.Text = Null2String(RSORD_HIST!TRANTYPE)
        If txtTranType.Text = "RIV" Then
            cboSalesMan.Enabled = False
            cboSMName.Enabled = False
        Else
            cboSalesMan.Enabled = True
            cboSMName.Enabled = True
        End If
        txtTranNo.Text = Null2String(RSORD_HIST!TRANNO)
        txtTranDate.Text = Null2String(RSORD_HIST!trandate)
        txtCustCode.Text = Null2String(RSORD_HIST!CUSTCODE)
        txtCustName.Text = Null2String(RSORD_HIST!CUSTNAME)
        If Null2String(RSORD_HIST!chargeto) = "MEC" Then
            cboChargeTo.Text = "MECHANICAL"
        ElseIf Null2String(RSORD_HIST!chargeto) = "COM" Then
            cboChargeTo.Text = "COMPANY"
        ElseIf Null2String(RSORD_HIST!chargeto) = "WAR" Then
            cboChargeTo.Text = "WARRANTY"
        ElseIf Null2String(RSORD_HIST!chargeto) = "TIN" Then
            cboChargeTo.Text = "TINSMITH"
        ElseIf Null2String(RSORD_HIST!chargeto) = "FLE" Then
            cboChargeTo.Text = "FLEET"
        ElseIf Null2String(RSORD_HIST!chargeto) = "VAR" Then
            cboChargeTo.Text = "VARIOUS"
        ElseIf Null2String(RSORD_HIST!chargeto) = "PCL" Then
            cboChargeTo.Text = "PARTS CLAIM"
        Else
            cboChargeTo.Text = ""
        End If
        txtRONO.Text = Null2String(RSORD_HIST!RONO)
        cboSMName.Text = FillSalesMan(Null2String(RSORD_HIST!salesman))
        txtTerms.Text = Null2String(RSORD_HIST!Terms)
        txtTTLInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HIST!ttlinvamt))
        txtDS1.Text = N2Str2IntZero(RSORD_HIST!ds1)
        txtDS_Desc1.Text = Null2String(RSORD_HIST!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSORD_HIST!DS_AMT1))
        txtNetInvAmt.Text = ToDoubleNumber(N2Str2Zero(RSORD_HIST!NETINVAMT))
        txtremarks.Text = Null2String(RSORD_HIST!REMARKS)
        If Null2String(RSORD_HIST!Status) = "C" Then
            labPosted.Caption = "CANCELLED"
            If LOGLEVEL = "41444D_]jUU" Then
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = False
            End If
        ElseIf Null2String(RSORD_HIST!Status) = "B" Then
            labPosted.Caption = "BILLED OUT"
            If LOGLEVEL = "41444D_]jUU" Then
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
            End If
        ElseIf Null2String(RSORD_HIST!Status) = "P" Then
            labPosted.Caption = "POSTED"
            If LOGLEVEL = "41444D_]jUU" Then
                cmdEdit.Enabled = False
                cmdCancelCO.Enabled = False
                cmdPost.Enabled = False
                cmdPrint.Enabled = True
            End If
        Else
            labPosted.Caption = ""
            If LOGLEVEL = "41444D_]jUU" Then
                cmdEdit.Enabled = True
                If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = True
                cmdPost.Enabled = True
                cmdPrint.Enabled = False
            End If
        End If
        If Null2String(RSORD_HIST!In_Process) = "N" Then
            labPosted.Caption = "RELEASED"
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdPost.Enabled = False
        End If
        cleargrid grdDetails
        FillDetails
    Else
        MsgBox "No record found on Issuance History Database... This Form will be unloaded...", vbInformation, "Info"
        Unload Me
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
        .ColWidth(6) = 1200
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
    Set RSDAYTRAN = New ADODB.Recordset
    RSDAYTRAN.Open "select trantype,tranno,id,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranuprice from PMIS_DayTran where type = 'A' and tranno = " & N2Str2Null(RSORD_HIST!TRANNO) & " and trantype = " & N2Str2Null(RSORD_HIST!TRANTYPE) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
        Screen.MousePointer = 11
        RSDAYTRAN.MoveFirst
        Do While Not RSDAYTRAN.EOF
            KCNT = KCNT + 1
            If txtTranType.Text = "ADB" Then
                STOCKDESCription = Null2String(RSDAYTRAN!STOCK_SUP)
            Else
                STOCKDESCription = SetSTOCKDESC(Null2String(RSDAYTRAN!STOCK_SUP))
            End If
            grdDetails.AddItem RSDAYTRAN!ID & Chr(9) & Format(Null2String(RSDAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(RSDAYTRAN!STOCK_ORD) & Chr(9) & _
                               STOCKDESCription & Chr(9) & _
                               N2Str2IntZero(RSDAYTRAN!TRANQTY) & Chr(9) & _
                               Format(N2Str2Zero(RSDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT) & Chr(9) & _
                               Format(N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUPRICE), MAXIMUM_DIGIT)
            ORD_TOTQTY = ORD_TOTQTY + N2Str2IntZero(RSDAYTRAN!TRANQTY)
            ORD_TOTUPRICE = ORD_TOTUPRICE + (N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUPRICE))
            ORD_TOTINVAMT = ORD_TOTINVAMT + (N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUPRICE))
            RSDAYTRAN.MoveNext
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
        cleargrid grdDetails
    End If
End Sub

Sub InitParts()
    txtTranItemNo.Text = Format(KCNT + 1, "0000")
    cboTranPartNo.Text = ""
    txtTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranUPrice.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
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

Sub FillGrid()
    Dim RSORD_HIST                                     As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HIST = New ADODB.Recordset
    Set RSORD_HIST = gconDMIS.Execute("select Tranno,ID from PMIS_Ord_Hist where type = 'A' and trantype = '" & COUNTERTYPE & "' order by Tranno asc")
    If Not (RSORD_HIST.EOF And RSORD_HIST.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HIST
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSORD_HIST                                     As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HIST = gconDMIS.Execute("select tranno, ID from PMIS_Ord_Hist where type = 'A' and trantype = '" & COUNTERTYPE & "' and tranno like '" & XXX & "%'")
    If Not (RSORD_HIST.EOF And RSORD_HIST.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HIST
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim RSORD_HIST                                     As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HIST = New ADODB.Recordset
    Set RSORD_HIST = gconDMIS.Execute("select rono,ID from PMIS_Ord_Hist where type = 'A' and trantype = '" & COUNTERTYPE & "' and rono is not null order by tranno asc")
    If Not (RSORD_HIST.EOF And RSORD_HIST.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HIST
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSORD_HIST                                     As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HIST = gconDMIS.Execute("select Rono, ID from PMIS_Ord_Hist where type = 'A' and trantype = '" & COUNTERTYPE & "' and rono like '" & XXX & "%' order by tranno asc")
    If Not (RSORD_HIST.EOF And RSORD_HIST.BOF) Then
        lstOrd_Hd.Enabled = True
        Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HIST
        lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillGrid3()
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Enabled = False
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    Set RSORD_HD = New ADODB.Recordset
    Set RSORD_HD = gconDMIS.Execute("select Custname,tranno from PMIS_Ord_Hist where [TYPE] = 'A' AND trantype = '" & COUNTERTYPE & "' order by CUSTNAME asc")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Sub FillSearchCusTomer(XXX As String)
    Dim RSORD_HD                                       As ADODB.Recordset
    lstOrd_Hd.Sorted = False: lstOrd_Hd.ListItems.Clear
    lstOrd_Hd.Enabled = False
    Set RSORD_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSORD_HD = gconDMIS.Execute("select custname, tranno from PMIS_Ord_Hist where [TYPE] = 'A' AND trantype = '" & COUNTERTYPE & "' and CUSTNAME  like '" & XXX & "%' order by CUSTNAME")
    If Not (RSORD_HD.EOF And RSORD_HD.BOF) Then
        lstOrd_Hd.Enabled = True: Listview_Loadval Me.lstOrd_Hd.ListItems, RSORD_HD: lstOrd_Hd.Refresh
    Else
        lstOrd_Hd.Enabled = False
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSORD_HIST.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSORD_HIST.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSORD_HIST.MoveNext
    If RSORD_HIST.EOF Then
        RSORD_HIST.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSORD_HIST.MovePrevious
    If RSORD_HIST.BOF Then
        RSORD_HIST.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If RSORD_HIST!TRANTYPE = "ADB" Or RSORD_HIST!TRANTYPE = "RIV" Then
        cmdSignatories.Visible = True
        cmdSignatories.ZOrder 0
        fraSignatories.Visible = True
        fraSignatories.ZOrder 0
        txtPreparedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", "")
        txtIssuedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", "")
        txtApprovedBy.Text = GetSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", "")
        On Error Resume Next
        txtRequestedBy.SetFocus
    End If
    If RSORD_HIST!TRANTYPE = "CSH" Then
        If MsgQuestionBox("Accessories Issuance Slip (CSH) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CSHPRINTING
        End If
    End If
    If RSORD_HIST!TRANTYPE = "CHG" Then
        If MsgQuestionBox("Accessories Issuance Slip (CHG) will be Printed. Are you Sure?", "Confirm Printing...") = True Then
            CHGPRINTING
        End If
    End If
    NEW_LogAudit "V", "ACCESSORIES ISSUANCE - HISTORY", "", "", "Accessories", txtTranNo, COUNTERTYPE, ""
End Sub

Private Sub cmdPrintRIV_Click()
    If RSORD_HIST!TRANTYPE = "RIV" Then
        SERVICEPISPRINTING
        'LogAudit "V", "ACCESSORIES HISTORY RIV PRINTING"
    End If

    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "PREPARED BY", txtPreparedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "ISSUED BY", txtIssuedBy)
    Call SaveSetting("DMIS", "SERVICE ISSUANCE", "APPROVED BY", txtApprovedBy)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemVars
            End If
        Case vbKeyF3
            If LOGLEVEL = "41444D_]jUU" Then
                If Picture1.Visible = True Then
                    If Null2String(RSORD_HIST!Status) = "C" Then
                        MsgSpeechBox "Transactions are Already Cancelled and cannot be Change..."
                    ElseIf Null2String(RSORD_HIST!Status) = "B" Then
                        MsgSpeechBox "Transactions are Already Billed-Out and cannot be Change..."
                    Else
                    End If
                End If
            Else
                MsgSpeechBox "History Transactions cannot be Changed..."
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    textSearch.Text = ""
    If COUNTERTYPE <> "RIV" And COUNTERTYPE <> "ADB" Then optRONo.Enabled = False
    If LOGLEVEL = "41444D_]jUU" Then
        cmdAdd.Enabled = False: cmdEdit.Enabled = False: cmdPost.Enabled = False
        If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = False: cmdPrint.Enabled = True
    Else
        cmdAdd.Enabled = False: cmdEdit.Enabled = False: cmdPost.Enabled = False
        If LOGLEVEL = "ADM" Then cmdCancelCO.Enabled = False: cmdPrint.Enabled = True
    End If
    rsRefresh
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISTrans_CustomerOrder = Nothing
    UnloadForm Me
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

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If
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

Private Sub lstOrd_Hd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then On Error Resume Next: textSearch.SetFocus
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
    textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If optTranno.Value = True Then
        If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    ElseIf optRONo.Value = True Then
        Dim RONOStr                                    As String
        RONOStr = textSearch.Text
        If Left(RONOStr, 2) = "R-" Then
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr) - 2)), "00000000")
        Else
            RONOStr = "R-" & Format(NumericVal(Right(RONOStr, Len(RONOStr))), "00000000")
        End If
        If Trim(textSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (RONOStr)
    Else
        FillSearchCusTomer (textSearch.Text)
    End If
End Sub

Private Sub Option1_Click()
    lstOrd_Hd.ColumnHeaders(1).Text = "CUSTOMER NAME"
    If textSearch = "" Then FillGrid3 Else FillSearchCusTomer (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub lstOrd_Hd_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    If optTranno.Value = True Then
        RSORD_HIST.Bookmark = rsFind(RSORD_HIST.Clone, "tranno", ITEM).Bookmark
    ElseIf optRONo.Value = True Then
        RSORD_HIST.Bookmark = rsFind(RSORD_HIST.Clone, "ID", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    Else
        RSORD_HIST.Bookmark = rsFind(RSORD_HIST.Clone, "tranno", lstOrd_Hd.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
End Sub

