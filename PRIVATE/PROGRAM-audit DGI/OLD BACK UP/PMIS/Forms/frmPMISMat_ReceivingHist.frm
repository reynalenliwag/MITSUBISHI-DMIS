VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMat_ReceivingHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receiving History"
   ClientHeight    =   6855
   ClientLeft      =   855
   ClientTop       =   855
   ClientWidth     =   11835
   ForeColor       =   &H00DEDFDE&
   Icon            =   "frmPMISMat_ReceivingHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   11835
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   3180
      ScaleHeight     =   870
      ScaleWidth      =   8655
      TabIndex        =   89
      Top             =   5580
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   92
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
         Left            =   7020
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelRR 
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
         Left            =   6240
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   5460
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   100
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":1B3D
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   94
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":1FEB
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":213D
         Style           =   1  'Graphical
         TabIndex        =   95
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":2450
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":25A2
         Style           =   1  'Graphical
         TabIndex        =   91
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":28F2
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   90
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":2DA2
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":2EF4
         Style           =   1  'Graphical
         TabIndex        =   96
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":31EE
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":3340
         Style           =   1  'Graphical
         TabIndex        =   97
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
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":3698
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":37EA
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   2250
      TabIndex        =   18
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtDS1 
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
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox txtINVNo 
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
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Type the Receiving Entry's Ref INV Number (e.g. 329874)"
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox txtDRNo 
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
         TabIndex        =   10
         ToolTipText     =   "Type the Receiving Entry DR Number,if there's any  (e.g. 555665)"
         Top             =   2670
         Width           =   1005
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
         Left            =   4620
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "frmPMISMat_ReceivingHist.frx":3B49
         ToolTipText     =   "Type your massage or remarks."
         Top             =   2010
         Width           =   4755
      End
      Begin VB.ComboBox cboClasscode 
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
         Left            =   5910
         TabIndex        =   2
         Text            =   "cboRecvd_Desc"
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtRecvd_Code 
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
         Left            =   1380
         TabIndex        =   6
         ToolTipText     =   "Type the supplier's code (e.g. 00001) "
         Top             =   1050
         Width           =   975
      End
      Begin VB.ComboBox cboRecvd_Desc 
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
         Text            =   "cboRecvd_Desc"
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1440
         Width           =   4395
      End
      Begin VB.TextBox txtRRNo 
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
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Type Receiving entry number (e.g 003294)"
         Top             =   180
         Width           =   1155
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
         Left            =   3210
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Type the terms of the transaction."
         Top             =   1050
         Width           =   1275
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
         Left            =   6210
         TabIndex        =   13
         ToolTipText     =   "Input the type of the additional amount (e.g. VAT)"
         Top             =   1200
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtPONo 
         Height          =   345
         Left            =   1380
         TabIndex        =   4
         ToolTipText     =   "Type purchase order number of the receiving entry (e.g. 02774)"
         Top             =   660
         Width           =   975
         _ExtentX        =   1720
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPODate 
         Height          =   345
         Left            =   3210
         TabIndex        =   5
         Top             =   660
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox txtRRDate 
         Height          =   345
         Left            =   3210
         TabIndex        =   1
         ToolTipText     =   "Type date of the receiving entry in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
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
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   825
         Left            =   90
         ScaleHeight     =   825
         ScaleWidth      =   4455
         TabIndex        =   31
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox txtDetails 
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
            Height          =   825
            Left            =   0
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   0
            Width           =   4365
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1245
         Left            =   7860
         ScaleHeight     =   1245
         ScaleWidth      =   1545
         TabIndex        =   16
         Top             =   750
         Width           =   1545
         Begin VB.TextBox txtTTLRRAmt 
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   78
            Top             =   60
            Width           =   1455
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
            Left            =   0
            MaxLength       =   15
            TabIndex        =   77
            Top             =   450
            Width           =   1515
         End
         Begin VB.TextBox txtNetRRAmt 
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   76
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8670
         Top             =   2520
      End
      Begin MSMask.MaskEdBox txtRIV_Tranno 
         Height          =   345
         Left            =   5280
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
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
      Begin VB.Label labRIV_TranNo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RIV #"
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
         Left            =   4650
         TabIndex        =   80
         Top             =   690
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label21 
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
         Height          =   255
         Left            =   5850
         TabIndex        =   79
         Top             =   1230
         Width           =   375
      End
      Begin VB.Label Label9 
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
         Left            =   6750
         TabIndex        =   25
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label Label10 
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
         Left            =   6780
         TabIndex        =   24
         Top             =   1650
         Width           =   1965
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref DR#"
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
         TabIndex        =   21
         Top             =   2700
         Width           =   795
      End
      Begin VB.Label Label8 
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
         TabIndex        =   52
         Top             =   1770
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PO NO"
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
         Left            =   90
         TabIndex        =   22
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
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
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   690
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RR Number"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RR Date"
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
         Index           =   0
         Left            =   2400
         TabIndex        =   29
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Classification"
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
         Left            =   4620
         TabIndex        =   28
         Top             =   240
         Width           =   1305
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
         Height          =   225
         Left            =   2400
         TabIndex        =   27
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Receive From"
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
         TabIndex        =   26
         Top             =   1080
         Width           =   1275
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
         Height          =   225
         Left            =   3660
         TabIndex        =   23
         Top             =   1470
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Ref INV#"
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
         Left            =   2490
         TabIndex        =   19
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label labRRsted 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   6900
         TabIndex        =   53
         Top             =   210
         Width           =   2475
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   90
      TabIndex        =   83
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
         TabIndex        =   86
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
         TabIndex        =   85
         Top             =   630
         Width           =   1875
      End
      Begin VB.OptionButton optRRNo 
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
         TabIndex        =   84
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstREC_Hist 
         Height          =   5115
         Left            =   60
         TabIndex        =   87
         Top             =   1320
         Width           =   1995
         _ExtentX        =   3519
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":3B63
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
         TabIndex        =   88
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2445
      Left            =   2250
      TabIndex        =   17
      Top             =   3030
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2175
         Left            =   60
         TabIndex        =   15
         Top             =   180
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   8
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
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
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
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   4620
      TabIndex        =   32
      Top             =   990
      Width           =   4575
      Begin VB.TextBox txtTranTotalAmt 
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
         Height          =   315
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   46
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox txtUnitCost 
         Alignment       =   1  'Right Justify
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
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   45
         Top             =   2340
         Width           =   1695
      End
      Begin VB.TextBox txtTranINVAmt 
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
         Height          =   315
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   44
         Top             =   1980
         Width           =   1695
      End
      Begin VB.TextBox txtTranQty 
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
         Height          =   315
         Left            =   1470
         MaxLength       =   15
         TabIndex        =   43
         Top             =   1620
         Width           =   885
      End
      Begin VB.CommandButton cmdTranCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2370
         MaskColor       =   &H0000FFFF&
         Picture         =   "frmPMISMat_ReceivingHist.frx":3CC5
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3120
         Width           =   1005
      End
      Begin VB.CommandButton cmdTranSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1230
         MaskColor       =   &H0000FFFF&
         Picture         =   "frmPMISMat_ReceivingHist.frx":3FD7
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtTranItemNo 
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
         Height          =   315
         Left            =   1470
         TabIndex        =   40
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdTranDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Courier New"
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
         Picture         =   "frmPMISMat_ReceivingHist.frx":4419
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3120
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   600
         Width           =   585
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Left            =   150
         TabIndex        =   33
         Top             =   2700
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Cost"
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
         Left            =   480
         TabIndex        =   51
         Top             =   2340
         Width           =   975
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
         Left            =   1620
         TabIndex        =   39
         Top             =   3330
         Width           =   285
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amt."
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
         Left            =   210
         TabIndex        =   38
         Top             =   1980
         Width           =   1245
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
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   1620
         Width           =   855
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
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   630
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
         TabIndex        =   35
         Top             =   270
         Width           =   885
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
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1125
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   4245
      Left            =   4530
      TabIndex        =   74
      Top             =   930
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   7488
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
      MICON           =   "frmPMISMat_ReceivingHist.frx":4723
   End
   Begin wizButton.cmd cmdUpdateMaster 
      Height          =   2235
      Left            =   4350
      TabIndex        =   75
      Top             =   1920
      Width           =   5205
      _ExtentX        =   9181
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
      MICON           =   "frmPMISMat_ReceivingHist.frx":473F
   End
   Begin VB.Frame fraUpdateMaster 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      Caption         =   "Confirm Update Master File"
      ForeColor       =   &H00000000&
      Height          =   2085
      Left            =   4440
      TabIndex        =   54
      Top             =   1980
      Width           =   5025
      Begin wizButton.cmd cmdOkUpdate 
         Height          =   345
         Left            =   3480
         TabIndex        =   55
         Top             =   1590
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         TX              =   "&Ok"
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
         MICON           =   "frmPMISMat_ReceivingHist.frx":475B
      End
      Begin VB.TextBox txtNewOH 
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
         Height          =   315
         Left            =   1980
         TabIndex        =   66
         Text            =   "Text"
         Top             =   1620
         Width           =   1260
      End
      Begin VB.TextBox txtNewSRP 
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
         Height          =   315
         Left            =   1980
         TabIndex        =   65
         Text            =   "Text"
         Top             =   1260
         Width           =   1260
      End
      Begin VB.TextBox txtNewDNP 
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
         Height          =   315
         Left            =   1980
         TabIndex        =   64
         Text            =   "Text"
         Top             =   900
         Width           =   1260
      End
      Begin VB.TextBox txtNewMAC 
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
         Height          =   315
         Left            =   1980
         TabIndex        =   63
         Text            =   "Text"
         Top             =   540
         Width           =   1260
      End
      Begin VB.TextBox txtOldOH 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   59
         Text            =   "Text"
         Top             =   1620
         Width           =   1260
      End
      Begin VB.TextBox txtOldSRP 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   58
         Text            =   "Text"
         Top             =   1260
         Width           =   1260
      End
      Begin VB.CheckBox chkUpdateDNP 
         BackColor       =   &H00DEDFDE&
         Caption         =   "Update DNP"
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
         Left            =   3360
         TabIndex        =   62
         Top             =   810
         Width           =   1485
      End
      Begin VB.CheckBox chkUpdateMAC 
         BackColor       =   &H00DEDFDE&
         Caption         =   "Update MAC"
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
         Left            =   3360
         TabIndex        =   61
         Top             =   540
         Width           =   1485
      End
      Begin VB.CheckBox chkUpdateSRP 
         BackColor       =   &H00DEDFDE&
         Caption         =   "Update SRP"
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
         Left            =   3360
         TabIndex        =   60
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txtOldDNP 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   57
         Text            =   "Text"
         Top             =   900
         Width           =   1260
      End
      Begin VB.TextBox txtOldMAC 
         Alignment       =   1  'Right Justify
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
         Left            =   630
         TabIndex        =   56
         Text            =   "Text"
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "SRP"
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
         Height          =   345
         Left            =   150
         TabIndex        =   73
         Top             =   1290
         Width           =   1125
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DNP"
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
         Height          =   345
         Left            =   150
         TabIndex        =   72
         Top             =   930
         Width           =   1125
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "MAC"
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
         Height          =   345
         Left            =   150
         TabIndex        =   71
         Top             =   540
         Width           =   1125
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OLD"
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
         Height          =   345
         Left            =   750
         TabIndex        =   70
         Top             =   210
         Width           =   585
      End
      Begin VB.Label Label16 
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
         Left            =   1620
         TabIndex        =   69
         Top             =   3000
         Width           =   285
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW"
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
         Height          =   345
         Left            =   2130
         TabIndex        =   68
         Top             =   210
         Width           =   885
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "OH"
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
         Height          =   345
         Left            =   150
         TabIndex        =   67
         Top             =   1650
         Width           =   1125
      End
   End
   Begin Crystal.CrystalReport rptReceiving 
      Left            =   2490
      Top             =   5910
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10305
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   101
      Top             =   5580
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
         Left            =   660
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":4777
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":48C9
         Style           =   1  'Graphical
         TabIndex        =   102
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
         Left            =   -120
         MouseIcon       =   "frmPMISMat_ReceivingHist.frx":4C07
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISMat_ReceivingHist.frx":4D59
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label Label3 
      Caption         =   "- required field"
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
      Left            =   10380
      TabIndex        =   82
      Top             =   6570
      Width           =   1395
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
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   10170
      TabIndex        =   81
      Top             =   6600
      Width           =   135
   End
End
Attribute VB_Name = "frmPMISMat_ReceivingHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsREC_HIST, rsPO_Hist, rsDaytran                   As ADODB.Recordset
Attribute rsPO_Hist.VB_VarUserMemId = 1073938432
Attribute rsDaytran.VB_VarUserMemId = 1073938432
Dim rsPartMas, rsSupplier                              As ADODB.Recordset
Attribute rsPartMas.VB_VarUserMemId = 1073938435
Attribute rsSupplier.VB_VarUserMemId = 1073938435
Dim rsCunter                                           As ADODB.Recordset
Attribute rsCunter.VB_VarUserMemId = 1073938437
Dim Pcnt                                               As Integer
Attribute Pcnt.VB_VarUserMemId = 1073938438
Dim AddorEdit                                          As String
Attribute AddorEdit.VB_VarUserMemId = 1073938439
Dim RR_TOTUCOST, RR_TOTINVAMT, RR_TOTVAT               As Double
Attribute RR_TOTUCOST.VB_VarUserMemId = 1073938440
Attribute RR_TOTINVAMT.VB_VarUserMemId = 1073938440
Attribute RR_TOTVAT.VB_VarUserMemId = 1073938440
Dim RR_QTY_REC                                         As Long
Attribute RR_QTY_REC.VB_VarUserMemId = 1073938443
Dim PrevRRNo                                           As String
Attribute PrevRRNo.VB_VarUserMemId = 1073938444
Dim PMIS_SUPPORT_Connection                            As String
Attribute PMIS_SUPPORT_Connection.VB_VarUserMemId = 1073938445
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasSRP              As Double
Attribute PrevPmasMAC.VB_VarUserMemId = 1073938446
Attribute PrevPmasDNP.VB_VarUserMemId = 1073938446
Attribute PrevPmasSRP.VB_VarUserMemId = 1073938446
Dim PrevPmasOnHand                                     As Integer
Attribute PrevPmasOnHand.VB_VarUserMemId = 1073938449
Dim NewPmasMAC, NewPmasDNP, NewPmasSRP                 As Double
Attribute NewPmasMAC.VB_VarUserMemId = 1073938450
Attribute NewPmasDNP.VB_VarUserMemId = 1073938450
Attribute NewPmasSRP.VB_VarUserMemId = 1073938450
Dim NewPmasOnHand, PrevTranQty                         As Integer
Attribute NewPmasOnHand.VB_VarUserMemId = 1073938453
Attribute PrevTranQty.VB_VarUserMemId = 1073938453
Dim ISNONVAT                                           As Boolean
Attribute ISNONVAT.VB_VarUserMemId = 1073938455

Private Sub cmdAddTran_Click()
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

Private Sub cmdCancelRR_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub

    On Error GoTo Errorcode:

    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
        Dim PCurOnOrder, PCurTRECQTY, PCurReceipts     As Integer
        Dim PCurLast_recq, PCurTpoQty                  As Integer
        Dim rsDAYTRANDup, rsPartmasDup                 As ADODB.Recordset
        Set rsDAYTRANDup = New ADODB.Recordset
        rsDAYTRANDup.Open "select trantype,tranno,tranqty,STOCK_ORD from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsREC_HIST!RRNO), gconDMIS
        If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
            rsDAYTRANDup.MoveFirst
            Do While Not rsDAYTRANDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select STOCKNO,onorder,trecqty,receipts,last_recq from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsDAYTRANDup!STOCK_ORD), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCurOnOrder = N2Str2IntZero(rsPartmasDup!onorder) + N2Str2IntZero(rsDAYTRANDup!tranqty)
                    PCurTRECQTY = N2Str2IntZero(rsPartmasDup!trecqty) - N2Str2IntZero(rsDAYTRANDup!tranqty)
                    PCurReceipts = N2Str2IntZero(rsPartmasDup!receipts) - N2Str2IntZero(rsDAYTRANDup!tranqty)
                    PCurLast_recq = N2Str2IntZero(rsPartmasDup!last_recq) - N2Str2IntZero(rsDAYTRANDup!tranqty)
                    gconDMIS.Execute "update PMIS_STOCKMAS set" & _
                                   " onorder = " & PCurOnOrder & "," & _
                                   " trecqty = " & PCurTRECQTY & "," & _
                                   " receipts = " & PCurReceipts & "," & _
                                   " last_recq = " & PCurLast_recq & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where STOCKNO = " & N2Str2Null(rsDAYTRANDup!STOCK_ORD)
                End If
                rsDAYTRANDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_Rec_Hist set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        gconDMIS.Execute "update PMIS_DayTran set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where tranno = " & N2Str2Null(rsREC_HIST!RRNO) & " and trantype = 'RR'"

        LogAudit "P", "Receiving History", txtRRNo
        rsRefresh
        On Error Resume Next
        rsREC_HIST.Find "id =" & labID.Caption
        StoreMemvars
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete") = False Then Exit Sub

End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub

    On Error GoTo Errorcode:

    Dim pmasOnOrder                                    As Integer
    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        Set rsDaytran = New ADODB.Recordset
        rsDaytran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsREC_HIST!RRNO) & " order by itemno asc", gconDMIS
        If Not rsDaytran.EOF And Not rsDaytran.BOF Then
            rsDaytran.MoveFirst
            Do While Not rsDaytran.EOF
                If N2Str2Zero(rsDaytran!TRANINVAMT) <= 0 Then
                    MsgSpeechBox "Transaction with Invoice Amount equal to Zero Encountered!"
                    Exit Sub
                End If
                rsDaytran.MoveNext
            Loop
            rsDaytran.MoveFirst
            Do While Not rsDaytran.EOF
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "Select STOCKNO,onhand,trecqty,onorder,receipts from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsDaytran!STOCK_ORD), gconDMIS
                If Not rsPartMas.EOF And Not rsPartMas.EOF Then
                    pmasOnOrder = N2Str2Zero(rsPartMas!onorder)
                    If pmasOnOrder <= 0 Then
                        pmasOnOrder = NumericVal(rsDaytran!tranqty)
                    End If
                    gconDMIS.Execute "update PMIS_STOCKMAS set onhand =" & N2Str2Zero(rsPartMas!ONHAND) + NumericVal(rsDaytran!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(rsPartMas!trecqty) + NumericVal(rsDaytran!tranqty) & ", " & _
                                   " onorder = " & pmasOnOrder - NumericVal(rsDaytran!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(rsPartMas!receipts) + NumericVal(rsDaytran!tranqty) & ", " & _
                                   " last_recq = " & N2Str2Zero(rsDaytran!tranqty) & ", " & _
                                   " last_recd = '" & LOGDATE & "', " & _
                                   " supcode = " & N2Str2Null(Mid(txtRecvd_Code.Text, 1, 5)) & _
                                   " where STOCKNO = " & N2Str2Null(rsPartMas!STOCKNO)
                    gconDMIS.Execute "update PMIS_DayTran set" & _
                                   " status = 'P'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsDaytran!ID
                End If
                rsDaytran.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_Rec_Hist set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        LogAudit "P", "Receiving History", txtRRNo
        rsRefresh
        On Error Resume Next
        rsREC_HIST.Find "id =" & labID.Caption
        StoreMemvars
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdPrint_Click()

    If Function_Access(LOGID, "Acess_Print", "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub

    On Error GoTo Errorcode:
    If MsgQuestionBox("Receiving Report Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        rptReceiving.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReceiving.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptReceiving, PMIS_REPORT_PATH & "rr_hist.rpt", "{RR_HD.type} = 'M' and {RR_HD.rrno} = '" & txtRRNo.Text & "'", DMIS_REPORT_Connection, 1
        LogAudit "V", "Receiving History", txtRRNo
        Screen.MousePointer = 0
    End If
    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdUnPost_Click()
    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then

        On Error GoTo Errorcode:

        Set rsDaytran = New ADODB.Recordset
        rsDaytran.Open "select id,itemno,trantype,tranno,STOCK_ORD,tranqty,traninvamt from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsREC_HIST!RRNO) & " order by itemno asc", gconDMIS
        If Not rsDaytran.EOF And Not rsDaytran.BOF Then
            rsDaytran.MoveFirst
            Do While Not rsDaytran.EOF
                Set rsPartMas = New ADODB.Recordset
                rsPartMas.Open "Select STOCKNO,onhand,trecqty,onorder,receipts from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(rsDaytran!STOCK_ORD), gconDMIS
                If Not rsPartMas.EOF And Not rsPartMas.EOF Then
                    gconDMIS.Execute "update PMIS_STOCKMAS set onhand =" & N2Str2Zero(rsPartMas!ONHAND) - NumericVal(rsDaytran!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(rsPartMas!trecqty) - NumericVal(rsDaytran!tranqty) & ", " & _
                                   " onorder = " & N2Str2Zero(rsPartMas!onorder) + NumericVal(rsDaytran!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(rsPartMas!receipts) - NumericVal(rsDaytran!tranqty) & ", " & _
                                   " last_recq = " & 0 & ", " & _
                                   " last_recd = NULL, " & _
                                   " supcode = NULL" & _
                                   " where STOCKNO = " & N2Str2Null(rsPartMas!STOCKNO)
                    gconDMIS.Execute "update PMIS_DayTran set" & _
                                   " status = 'N'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsDaytran!ID
                End If
                rsDaytran.MoveNext
            Loop
        End If
        gconDMIS.Execute "update PMIS_Rec_Hist set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labID.Caption
        rsRefresh
        On Error Resume Next
        rsREC_HIST.Find "id =" & labID.Caption
        StoreMemvars
    End If

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    On Error Resume Next
    txtRRNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "TRANSACTION HISTORY RECEIVING STORING") = False Then Exit Sub
    AddorEdit = "EDIT"
    grdDetails.Enabled = False
    PrevRRNo = Format(txtRRNo.Text, "000000")
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
End Sub

Sub FindDupRRno(DDD As String)
    rsREC_HIST.Bookmark = rsFind(rsREC_HIST.Clone, "rrno", Format(DDD, "000000")).Bookmark
    StoreMemvars
End Sub

Private Sub cmdFirst_Click()
    rsREC_HIST.MoveFirst
    StoreMemvars
End Sub

Private Sub cmdLast_Click()
    rsREC_HIST.MoveLast
    StoreMemvars
End Sub

Private Sub cmdNext_Click()
    rsREC_HIST.MoveNext
    If rsREC_HIST.EOF Then
        rsREC_HIST.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    rsREC_HIST.MovePrevious
    If rsREC_HIST.BOF Then
        rsREC_HIST.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                SendToBack
                StoreMemvars
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(rsREC_HIST!Status) = "P" Then
                    MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
                ElseIf Null2String(rsREC_HIST!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
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
    rsRefresh
    textSearch.Text = ""
    Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False
    chkUpdateMAC.Enabled = False: chkUpdateDNP.Enabled = False
    txtNewMAC.Enabled = False: txtNewDNP.Enabled = False
    txtPartID.Text = "": initMemvars: StoreMemvars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsREC_HIST = New ADODB.Recordset
    rsREC_HIST.Open "select * from PMIS_Rec_Hist where type = 'M' order by rrno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtRRNo.Text = ""
    txtPONo.Text = ""
    Set rsCunter = New ADODB.Recordset
    rsCunter.Open "select * from PMIS_Counter where modul = 'RR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCunter.EOF And Not rsCunter.BOF Then
        txtRRNo.Text = Format(Null2String(rsCunter!nextnumber), "000000")
    End If
    txtRRDate.Text = LOGDATE
    cboClasscode.Text = ""
    txtRIV_Tranno.Text = ""
    txtRecvd_Code.Text = ""
    FillCboRecvd
    txtDetails.Text = ""
    txtTerms.Text = ""
    txtPODate.Text = ""
    txtDRNo.Text = ""
    txtINVNo.Text = ""
    txtTTLRRAmt.Text = ""
    txtDS1.Text = ""
    txtDS_Desc1.Text = ""
    txtDS_Amt1.Text = ""
    txtNetRRAmt.Text = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    labRRsted.Caption = ""
    cleargrid grdDetails
    InitGrid
    InitCbo
    InitCboClasscode
    InitParts
End Sub

Sub StoreMemvars()
    If Not rsREC_HIST.EOF And Not rsREC_HIST.BOF Then
        labID.Caption = rsREC_HIST!ID
        txtRRNo.Text = Null2String(rsREC_HIST!RRNO)
        txtRRDate.Text = Null2String(rsREC_HIST!RRDATE)
        cboClasscode.Text = Null2String(rsREC_HIST!classcode)
        txtRIV_Tranno.Text = Null2String(rsREC_HIST!RIV_Tranno)
        txtRecvd_Code.Text = Null2String(rsREC_HIST!recvd_code)
        cboRecvd_Desc.Text = Null2String(rsREC_HIST!recvd_from)
        txtDetails.Text = Null2String(rsREC_HIST!Address)
        txtTerms.Text = Null2String(rsREC_HIST!Terms)
        txtPONo.Text = Null2String(rsREC_HIST!PONO)
        txtPODate.Text = Null2String(rsREC_HIST!PODATE)
        txtDRNo.Text = Null2String(rsREC_HIST!drno)
        txtINVNo.Text = Null2String(rsREC_HIST!invno)
        txtDS1.Text = N2Str2IntZero(rsREC_HIST!ds1)
        txtDS_Desc1.Text = Null2String(rsREC_HIST!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsREC_HIST!ds_amt1))
        txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsREC_HIST!ttlrramt))
        txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsREC_HIST!netrramt))
        txtRemarks.Text = Null2String(rsREC_HIST!remarks)
        If Null2String(rsREC_HIST!Status) = "P" Then
            labRRsted.Visible = True
            labRRsted.Caption = "POSTED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
        ElseIf Null2String(rsREC_HIST!Status) = "C" Then
            labRRsted.Visible = True
            labRRsted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
            cmdCancelRR.Enabled = False
        Else
            labRRsted.Visible = False
            labRRsted.Caption = ""
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
            If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = False
        End If
        cleargrid grdDetails
        FillDetails
    Else
        MsgBox "No record found on Receiving History Database... This form will be unloaded...", vbInformation, "Info"
        Unload Me
    End If
End Sub

Sub InitGrid()
    With grdDetails
        .ColWidth(0) = 1
        .ColWidth(1) = 800
        .ColWidth(2) = 1500
        .ColWidth(3) = 2500
        .ColWidth(4) = 500
        .ColWidth(5) = 1100
        .ColWidth(6) = 1100
        .ColWidth(7) = 1500
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
        .Text = "Inv. Amt."
        .Col = 6
        .Text = "Cost"
        .Col = 7
        .Text = "Total Amount"
    End With
End Sub

Sub FillDetails()
    On Error GoTo Errorcode
    Pcnt = 0
    RR_TOTUCOST = 0
    RR_TOTINVAMT = 0
    RR_TOTVAT = 0
    RR_QTY_REC = 0
    Set rsDaytran = New ADODB.Recordset
    rsDaytran.Open "select id,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt from PMIS_DayTran where type = 'M' and trantype = 'RR' and tranno = " & N2Str2Null(rsREC_HIST!RRNO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDaytran.EOF And Not rsDaytran.BOF Then
        Screen.MousePointer = 11
        rsDaytran.MoveFirst
        Do While Not rsDaytran.EOF
            Pcnt = Pcnt + 1
            grdDetails.AddItem rsDaytran!ID & Chr(9) & Null2String(rsDaytran!itemno) & Chr(9) & _
                               Null2String(rsDaytran!STOCK_ORD) & Chr(9) & _
                               SetSTOCKDESC(Null2String(rsDaytran!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(rsDaytran!tranqty) & Chr(9) & _
                               N2Str2Zero(rsDaytran!TRANINVAMT) & Chr(9) & _
                               N2Str2Zero(rsDaytran!TRANUCOST) & Chr(9) & _
                               Format(N2Str2IntZero(rsDaytran!tranqty) * N2Str2Zero(rsDaytran!TRANUCOST), MAXIMUM_DIGIT)
            RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsDaytran!tranqty)
            RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsDaytran!tranqty) * N2Str2Zero(rsDaytran!TRANUCOST))
            RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsDaytran!tranqty) * N2Str2Zero(rsDaytran!TRANINVAMT))
            rsDaytran.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        If Null2String(rsREC_HIST!classcode) = "PCS" Or Null2String(rsREC_HIST!classcode) = "PCG" Then
            RR_TOTVAT = ToDoubleNumber(RR_TOTINVAMT - RR_TOTUCOST)
        Else
            RR_TOTVAT = 0
        End If
        If NumericVal(RR_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtDS_Amt1.Text = RR_TOTVAT
            txtNetRRAmt.Text = NumericVal(txtTTLRRAmt.Text) + NumericVal(txtDS_Amt1.Text)
        Else
            txtDS1.Text = 0
            txtDS_Desc1.Text = ""
            txtDS_Amt1.Text = 0
            txtNetRRAmt.Text = NumericVal(txtTTLRRAmt.Text)
        End If
        txtDS_Amt1.Text = Format(txtDS_Amt1.Text, MAXIMUM_DIGIT)
        txtNetRRAmt.Text = Format(txtNetRRAmt.Text, MAXIMUM_DIGIT)
        Screen.MousePointer = 0
    Else
        cleargrid grdDetails
    End If
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Function SetSTOCKDESC(ppp As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select STOCKNO,STOCKDESC from PMIS_STOCKMAS where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKDESC = Null2String(rsPartMas!STOCKDESC)
    End If
End Function

Function SetSTOCKDESC2(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select id,STOCKDESC from PMIS_STOCKMAS where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetSTOCKDESC2 = Null2String(rsPartMas!STOCKDESC)
        End If
    End If
End Function

Function SetSTOCKNO(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_STOCKMAS where id = " & DDD, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetSTOCKNO = Null2String(rsPartMas!STOCKNO)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKNO from PMIS_STOCKMAS where STOCKNO = '" & DDD & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDSTOCKNO = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select id,STOCKDESC from PMIS_STOCKMAS where ltrim(rtrim(STOCKDESC))) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        SetPartIDDesc = Null2String(rsPartMas!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set rsPartMas = New ADODB.Recordset
        rsPartMas.Open "Select STOCKNO,mac from PMIS_STOCKMAS where STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsPartMas.EOF And Not rsPartMas.BOF Then
            SetPartPrice = Null2String(rsPartMas!Mac)
        End If
    End If
End Function

Sub FillCboRecvd()
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select supname from PMIS_vw_Supplier", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        rsSupplier.MoveFirst
        cboRecvd_Desc.Clear
        Do While Not rsSupplier.EOF
            cboRecvd_Desc.AddItem Null2String(rsSupplier!supname)
            rsSupplier.MoveNext
        Loop
    End If
End Sub

Function SetSupdesc(ppp As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supcode,supname,sup_addrs,vat_percnt,NONVAT from PMIS_vw_Supplier where supcode = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupdesc = Null2String(rsSupplier!supname)
        txtDetails.Text = Null2String(rsSupplier!sup_addrs)
        If Null2String(rsSupplier!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(rsSupplier!vat_percnt)
        End If
    Else
        txtDetails.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
    End If
End Function

Function SetSupCode(nnn As String)
    Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "Select supname,supcode,sup_addrs,vat_percnt,NONVAT from PMIS_vw_Supplier where supname = '" & nnn & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSupplier.EOF And Not rsSupplier.BOF Then
        SetSupCode = Null2String(rsSupplier!SupCode)
        txtDetails.Text = Null2String(rsSupplier!sup_addrs)
        If Null2String(rsSupplier!NONVAT) = "Y" Then
            ISNONVAT = True: txtDS1.Text = 0
        Else
            ISNONVAT = False: txtDS1.Text = Null2String(rsSupplier!vat_percnt)
        End If
    Else
        txtDetails.Text = ""
        txtDS1.Text = ""
        ISNONVAT = False
    End If
End Function

Sub InitParts()
    txtTranItemNo.Text = Format(Pcnt + 1, "0000")
    cboTranPartNo.Text = ""
    cboTranDescription.Text = ""
    txtTranQty.Text = 1
    txtTranINVAmt.Text = "0.00"
    txtTranTotalAmt.Text = "0.00"
End Sub

Function StorePartsEntry(ByVal ID As Variant)
    Set rsDaytran = New ADODB.Recordset
    rsDaytran.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_DayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsDaytran.EOF And Not rsDaytran.BOF Then
        labDetID.Caption = rsDaytran!ID
        txtTranItemNo.Text = Null2String(rsDaytran!itemno)
        cboTranPartNo.Text = Null2String(rsDaytran!STOCK_ORD)
        cboTranDescription.Text = SetSTOCKDESC(Null2String(rsDaytran!STOCK_SUP))
        txtTranQty.Text = N2Str2IntZero(rsDaytran!tranqty)
        txtTranINVAmt.Text = N2Str2Zero(rsDaytran!TRANINVAMT)
        txtUnitCost.Text = N2Str2Zero(rsDaytran!TRANUCOST)
        txtTranTotalAmt.Text = N2Str2IntZero(rsDaytran!tranqty) * N2Str2Zero(rsDaytran!TRANINVAMT)
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISReceiving2 = Nothing
    UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
    Dim fild                                           As String
    If Null2String(rsREC_HIST!Status) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
    ElseIf Null2String(rsREC_HIST!Status) = "C" Then
        MsgSpeechBox "Item(s) are Already Cancelled and cannot be edited"
    Else
        grdDetails.Row = grdDetails.Row
        grdDetails.Col = 0
        fild = grdDetails.Text
        If fild <> "" And fild <> "No Entry" Then
            AddorEdit = "EDIT"
            BringToFront
            cmdTranDelete.Enabled = True
            fraAddTran.Caption = "Edit Parts"
            StorePartsEntry (fild)
        Else
            MsgSpeechBox "No Entry on Parts"
            Exit Sub
        End If
    End If
End Sub

Sub SendToBack()
    cmdAddTran.ZOrder 1
    fraAddTran.ZOrder 1
    fraAddTran.Enabled = False
End Sub

Sub BringToFront()
    cmdAddTran.ZOrder 0
    fraAddTran.ZOrder 0
    fraAddTran.Enabled = True
End Sub

Sub InitCbo()
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "select STOCKNO,STOCKDESC from PMIS_STOCKMAS", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        rsPartMas.MoveFirst
        cboTranPartNo.Clear
        cboTranDescription.Clear
        Do While Not rsPartMas.EOF
            cboTranPartNo.AddItem Null2String(rsPartMas!STOCKNO)
            cboTranDescription.AddItem Null2String(rsPartMas!STOCKDESC)
            rsPartMas.MoveNext
        Loop
    End If
End Sub

Sub InitCboClasscode()
    cboClasscode.Clear
    cboClasscode.AddItem "IBT"
    cboClasscode.AddItem "PCG"
    cboClasscode.AddItem "PCS"
    cboClasscode.AddItem "RCG"
    cboClasscode.AddItem "RCS"
    cboClasscode.AddItem "REP"
    cboClasscode.AddItem "RRV"
    cboClasscode.Text = "PCG"
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboclasscode_LostFocus()
    If cboClasscode.Text <> "" Then
        If cboClasscode.Text = "RRV" Then
            labRIV_TranNo.Visible = True
            txtRIV_Tranno.Visible = True
        Else
            labRIV_TranNo.Visible = False
            txtRIV_Tranno.Visible = False
        End If
    Else
        MsgBoxXP "Invalid code. Please enter one of the following codes... " & vbCrLf & _
                 "IBT, PCG, PCS, RCG, RCS, REP, RRV", "Error Encountered", XP_OKOnly, msg_Information
    End If
End Sub

Private Sub Timer1_Timer()
    If labRRsted.Caption <> "" Then
        If labRRsted.Visible = True Then
            labRRsted.Visible = False
        Else
            labRRsted.Visible = True
        End If
    End If
End Sub

Private Sub txtDS1_LostFocus()
    txtDS1.Text = Format(txtDS1.Text, "##0")
End Sub

Private Sub txtPONo_GotFocus()
    If txtPONo.Text = "" Then
        Set rsCunter = New ADODB.Recordset
        rsCunter.Open "select * from PMIS_Counter where modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtPONo.Text = Format(N2Str2Zero(rsCunter!nextnumber) - 1, "000000")
        End If
    End If
End Sub

Private Sub txtPONo_LostFocus()
    If cboClasscode.Text = "PCG" Then
        If txtPONo.Text <> "" And AddorEdit = "ADD" And Len(txtPONo.Text) > 0 Then
            Dim rsREC_HISTDup                          As ADODB.Recordset
            Set rsREC_HISTDup = New ADODB.Recordset
            rsREC_HISTDup.Open "select pono from PMIS_Rec_Hist where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not rsREC_HISTDup.EOF And Not rsREC_HISTDup.BOF Then
                MsgBox "PO Number Already Received", vbInformation, "Invalid PO Number"
                Exit Sub
            End If
            Set rsPO_Hist = New ADODB.Recordset
            rsPO_Hist.Open "select pono,supcode,podate from PMIS_PO_Hist where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not rsPO_Hist.EOF And Not rsPO_Hist.BOF Then
                txtRecvd_Code.Text = Null2String(rsPO_Hist!SupCode)
                txtPODate.Text = Null2String(rsPO_Hist!PODATE)
                Pcnt = 0
                RR_TOTUCOST = 0
                RR_TOTINVAMT = 0
                RR_TOTVAT = 0
                RR_QTY_REC = 0
                Dim rsDAYTRANDup                       As ADODB.Recordset
                Set rsDAYTRANDup = New ADODB.Recordset
                rsDAYTRANDup.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_DayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_Hist!PONO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
                    Screen.MousePointer = 11
                    rsDAYTRANDup.MoveFirst
                    cleargrid grdDetails
                    Do While Not rsDAYTRANDup.EOF
                        Pcnt = Pcnt + 1
                        grdDetails.AddItem rsDAYTRANDup!ID & Chr(9) & Null2String(Format(rsDAYTRANDup!itemno, "0000")) & Chr(9) & _
                                           Null2String(rsDAYTRANDup!STOCK_ORD) & Chr(9) & _
                                           SetSTOCKDESC(Null2String(rsDAYTRANDup!STOCK_SUP)) & Chr(9) & _
                                           N2Str2IntZero(rsDAYTRANDup!tranqty) & Chr(9) & _
                                           N2Str2Zero(rsDAYTRANDup!TRANINVAMT) & Chr(9) & _
                                           N2Str2Zero(rsDAYTRANDup!TRANUCOST) & Chr(9) & _
                                           N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANUCOST)
                        RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANUCOST))
                        RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANINVAMT))
                        rsDAYTRANDup.MoveNext
                    Loop
                    If Pcnt <> 0 Then grdDetails.RemoveItem 1
                    Screen.MousePointer = 0
                Else
                    cleargrid grdDetails
                End If
            Else
                MsgSpeechBox "Invalid Purchase Order Number!"
                txtPONo.Text = ""
                txtPODate.Text = ""
                If AddorEdit = "ADD" Then
                    cleargrid grdDetails
                End If
                On Error Resume Next
                txtPONo.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtRecvd_Code_Change()
    cboRecvd_Desc.Text = SetSupdesc(txtRecvd_Code.Text)
End Sub

Private Sub txtRemarks_GotFocus()
    MsgSpeech "Pls Type Your Message Here!"
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRIV_Tranno_LostFocus()
    txtRIV_Tranno.Text = Format(txtRIV_Tranno, "000000")
End Sub

Private Sub txttranQty_Change()
    If txtTranQty.Text <> "" Then
        If Not rsREC_HIST.EOF And Not rsREC_HIST.BOF Then
            If Null2String(rsREC_HIST!classcode) = "PCS" Or Null2String(rsREC_HIST!classcode) = "PCG" Then
                If ISNONVAT = True Then
                    txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
                Else
                    txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
                End If
                txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
                txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
            End If
        End If
    End If
End Sub

Private Sub txtTranQty_GotFocus()
    If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_LostFocus()
    If txtTranQty.Text <> "" Then
        If Null2String(rsREC_HIST!classcode) = "PCS" Or Null2String(rsREC_HIST!classcode) = "PCG" Then
            If ISNONVAT = True Then
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        Else
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    Else
        txtTranQty.Text = 1
        If Null2String(rsREC_HIST!classcode) = "PCS" Or Null2String(rsREC_HIST!classcode) = "PCG" Then
            If ISNONVAT = True Then
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        Else
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    End If
    txtTranQty.Text = Format(txtTranQty.Text, DIGIT_FORMAT)
End Sub

Private Sub txtTranINVAmt_Change()
    On Error Resume Next
    If Null2String(rsREC_HIST!classcode) = "PCS" Or Null2String(rsREC_HIST!classcode) = "PCG" Then
        If NumericVal(txtTranINVAmt.Text) <> 0 Then
            If ISNONVAT = True Then
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    Else
        If NumericVal(txtTranINVAmt.Text) <> 0 Then
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    End If
End Sub

Private Sub txtTranINVAmt_GotFocus()
    If NumericVal(txtTranINVAmt.Text) = 0 Then txtTranINVAmt.Text = ""
End Sub

Private Sub txtTranINVAmt_LostFocus()
    If txtTranINVAmt.Text = "" Then txtTranINVAmt.Text = 0
    txtTranINVAmt.Text = Format(txtTranINVAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtUnitPrice_LostFocus()
    If Null2String(rsREC_HIST!RRNO) = "PCS" Or Null2String(rsREC_HIST!RRNO) = "PCS" Then
        If txtTranINVAmt.Text <> "" Then
            If ISNONVAT = True Then
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            Else
                txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
            End If
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    Else
        If txtTranINVAmt.Text <> "" Then
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
            txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
        End If
    End If
End Sub

Private Sub txtTranTotalAmt_Change()
    txtTranTotalAmt.Text = Format(txtTranTotalAmt.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtUnitCost_LostFocus()
    txtUnitCost.Text = Format(txtUnitCost.Text, MAXIMUM_DIGIT)
End Sub

Private Sub lstREC_HIST_GotFocus()
    rsREC_HIST.Bookmark = rsFind(rsREC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
    StoreMemvars
End Sub

Private Sub lstREC_HIST_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If optRRNo.Value = True Then
        rsREC_HIST.Bookmark = rsFind(rsREC_HIST.Clone, "rrno", Item).Bookmark
    Else
        rsREC_HIST.Bookmark = rsFind(rsREC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemvars
End Sub

Private Sub lstREC_HIST_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstREC_Hist
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

Private Sub lstREC_HIST_DblClick()
    cmdEdit.Value = False
End Sub

Private Sub lstREC_HIST_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_Change()
    If optRRNo.Value = True Then
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
    Dim rsREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    lstREC_Hist.Enabled = False
    Set rsREC_HIST = New ADODB.Recordset
    Set rsREC_HIST = gconDMIS.Execute("select rrno,ID from PMIS_Rec_Hist where type = 'M' order by rrno asc")
    If Not (rsREC_HIST.EOF And rsREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, rsREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Enabled = False
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set rsREC_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsREC_HIST = gconDMIS.Execute("select rrno, ID from PMIS_Rec_Hist where type = 'M' and rrno like'" & XXX & "%'")
    If Not (rsREC_HIST.EOF And rsREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, rsREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Enabled = False
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set rsREC_HIST = New ADODB.Recordset
    Set rsREC_HIST = gconDMIS.Execute("select recvd_from, ID from PMIS_Rec_Hist where type = 'M' order by rrno asc")
    If Not (rsREC_HIST.EOF And rsREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, rsREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim rsREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set rsREC_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set rsREC_HIST = gconDMIS.Execute("select recvd_from, ID from PMIS_Rec_Hist where type = 'M' and recvd_from like '" & XXX & "%' order by rrno asc")
    If Not (rsREC_HIST.EOF And rsREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, rsREC_HIST
        lstREC_Hist.Refresh
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstREC_Hist.ListItems.Count > 0 And lstREC_Hist.Enabled = True Then: lstREC_Hist.SetFocus
    End If
End Sub

Private Sub optRONo_Click()
    lstREC_Hist.ColumnHeaders(1).Text = "Sup. Name"
    lstREC_Hist.ColumnHeaders(1).Width = 4000
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optRRNo_Click()
    lstREC_Hist.ColumnHeaders(1).Text = "Tran. No."
    lstREC_Hist.ColumnHeaders(1).Width = 2150
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub
