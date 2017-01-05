VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMIOSRECON_ReceivingHist 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECON Receiving History"
   ClientHeight    =   6570
   ClientLeft      =   855
   ClientTop       =   855
   ClientWidth     =   11835
   ForeColor       =   &H00DEDFDE&
   Icon            =   "RECON_ReceivingHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11835
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   2250
      TabIndex        =   34
      Top             =   0
      Width           =   9495
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
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   12
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox txtINVNo 
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
         Left            =   3480
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Type the Receiving Entry's Ref INV Number (e.g. 329874)"
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox txtDRNo 
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
         Left            =   1410
         MaxLength       =   15
         TabIndex        =   10
         ToolTipText     =   "Type the Receiving Entry DR Number,if there's any  (e.g. 555665)"
         Top             =   2670
         Width           =   1005
      End
      Begin VB.TextBox txtRemarks 
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
         Height          =   1005
         Left            =   4620
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "RECON_ReceivingHist.frx":08CA
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
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Type Receiving entry number (e.g 003294)"
         Top             =   180
         Width           =   1155
      End
      Begin VB.TextBox txtTerms 
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
         Left            =   3210
         MaxLength       =   4
         TabIndex        =   7
         ToolTipText     =   "Type the terms of the transaction."
         Top             =   1050
         Width           =   1275
      End
      Begin VB.TextBox txtDS_Desc1 
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
         Appearance      =   0
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
         Appearance      =   0
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
         Appearance      =   0
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
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   90
         ScaleHeight     =   825
         ScaleWidth      =   4455
         TabIndex        =   47
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox txtDetails 
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
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   7860
         ScaleHeight     =   1245
         ScaleWidth      =   1545
         TabIndex        =   30
         Top             =   750
         Width           =   1545
         Begin VB.TextBox txtTTLRRAmt 
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   94
            Top             =   60
            Width           =   1455
         End
         Begin VB.TextBox txtDS_Amt1 
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
            Left            =   0
            MaxLength       =   15
            TabIndex        =   93
            Top             =   450
            Width           =   1515
         End
         Begin VB.TextBox txtNetRRAmt 
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   92
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
         Appearance      =   0
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
         TabIndex        =   96
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
         TabIndex        =   95
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   37
         Top             =   2700
         Width           =   795
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4560
         X2              =   9450
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4560
         X2              =   30
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   4560
         X2              =   4560
         Y1              =   120
         Y2              =   3180
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
         TabIndex        =   68
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
         TabIndex        =   38
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
         TabIndex        =   36
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   39
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
         TabIndex        =   35
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
         TabIndex        =   69
         Top             =   210
         Width           =   2475
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2250
      ScaleHeight     =   945
      ScaleWidth      =   9465
      TabIndex        =   31
      Top             =   5520
      Width           =   9495
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
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
         Left            =   8610
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":08E4
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":0BEE
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Close window"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P&rint"
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
         Left            =   7830
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":0EF8
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":1202
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Print transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelRR 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel Transaction"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7050
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":1ACC
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":1DD6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel selected transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdUnPost 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Unpost Transaction"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6270
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":2218
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":2522
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Unpost selected transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPost 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Post Transaction"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5490
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":2964
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":2C6E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Post selected transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
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
         Left            =   4710
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":3538
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":3842
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Edit selected transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
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
         Left            =   3930
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":410C
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":4416
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Add transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3150
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":4CE0
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":4FEA
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "View last record"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "F&irst"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2370
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":542C
         MousePointer    =   99  'Custom
         OLEDropMode     =   1  'Manual
         Picture         =   "RECON_ReceivingHist.frx":5736
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "View first record"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Find"
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
         Left            =   1590
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":5B78
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":5E82
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Search for a transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Next"
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
         Left            =   810
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":674C
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":6A56
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "View next transaction"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prev"
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
         Left            =   30
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":6E98
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":71A2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "View previous transaction"
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2250
      ScaleHeight     =   945
      ScaleWidth      =   9465
      TabIndex        =   32
      Top             =   5520
      Width           =   9495
      Begin VB.CommandButton cmdCancel 
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
         Left            =   8610
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":75E4
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":78EE
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Discard changes"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdSave 
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
         Left            =   7830
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "RECON_ReceivingHist.frx":8930
         MousePointer    =   99  'Custom
         Picture         =   "RECON_ReceivingHist.frx":8C3A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save changes"
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   6405
      Left            =   90
      ScaleHeight     =   6345
      ScaleWidth      =   2055
      TabIndex        =   99
      Top             =   90
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   11640
         Left            =   -210
         Picture         =   "RECON_ReceivingHist.frx":907C
         Top             =   -30
         Width           =   2550
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   90
      TabIndex        =   100
      Top             =   0
      Width           =   2115
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
         TabIndex        =   103
         Text            =   "TEXT"
         Top             =   960
         Width           =   1995
      End
      Begin VB.OptionButton optRONo 
         BackColor       =   &H00DEDFDE&
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
         TabIndex        =   102
         Top             =   630
         Width           =   1875
      End
      Begin VB.OptionButton optRRNo 
         BackColor       =   &H00DEDFDE&
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
         TabIndex        =   101
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstREC_Hist 
         Height          =   5115
         Left            =   60
         TabIndex        =   104
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
         MouseIcon       =   "RECON_ReceivingHist.frx":1D9F4
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
         BackColor       =   &H00DEDFDE&
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
         TabIndex        =   105
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   2250
      TabIndex        =   33
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
         BackColorFixed  =   14606302
         ForeColorFixed  =   0
         BackColorSel    =   14606302
         ForeColorSel    =   0
         BackColorBkg    =   14606302
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
      TabIndex        =   48
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         Picture         =   "RECON_ReceivingHist.frx":1DB56
         Style           =   1  'Graphical
         TabIndex        =   64
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
         Picture         =   "RECON_ReceivingHist.frx":1DE68
         Style           =   1  'Graphical
         TabIndex        =   63
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
         TabIndex        =   56
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
         Picture         =   "RECON_ReceivingHist.frx":1E2AA
         Style           =   1  'Graphical
         TabIndex        =   65
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   66
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
         TabIndex        =   49
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
         TabIndex        =   67
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   960
         Width           =   1125
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   4245
      Left            =   4530
      TabIndex        =   90
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
      MICON           =   "RECON_ReceivingHist.frx":1E5B4
   End
   Begin wizButton.cmd cmdUpdateMaster 
      Height          =   2235
      Left            =   4350
      TabIndex        =   91
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
      MICON           =   "RECON_ReceivingHist.frx":1E5D0
   End
   Begin VB.Frame fraUpdateMaster 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      Caption         =   "Confirm Update Master File"
      ForeColor       =   &H00000000&
      Height          =   2085
      Left            =   4440
      TabIndex        =   70
      Top             =   1980
      Width           =   5025
      Begin wizButton.cmd cmdOkUpdate 
         Height          =   345
         Left            =   3480
         TabIndex        =   71
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
         MICON           =   "RECON_ReceivingHist.frx":1E5EC
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   75
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
         TabIndex        =   74
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   73
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
         TabIndex        =   72
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
         TabIndex        =   89
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
         TabIndex        =   88
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
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
         TabIndex        =   84
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
         TabIndex        =   83
         Top             =   1650
         Width           =   1125
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
      TabIndex        =   98
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
      TabIndex        =   97
      Top             =   6600
      Width           =   135
   End
End
Attribute VB_Name = "frmPMIOSRECON_ReceivingHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRECON_REC_HIST, rsRECON_PO_HIST, rsRECON_DAYTRAN As ADODB.Recordset
Dim rsRECON_PARTMAS, rsSupplier As ADODB.Recordset
Dim rsCunter As ADODB.Recordset
Dim Pcnt As Integer
Dim AddorEdit As String
Dim RR_TOTUCOST, RR_TOTINVAMT, RR_TOTVAT As Double
Dim RR_QTY_REC As Long
Dim PrevRRNo As String
Dim PMIOS_SUPPORT_Connection As String
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasSRP As Double
Dim PrevPmasOnHand As Integer
Dim NewPmasMAC, NewPmasDNP, NewPmasSRP As Double
Dim NewPmasOnHand, PrevTranQty As Integer
Dim ISNONVAT As Boolean

Private Sub cboClasscode_Change()
If cboClasscode.Text = "RRV" Then
   labRIV_TranNo.Visible = True
   txtRIV_Tranno.Visible = True
Else
   labRIV_TranNo.Visible = False
   txtRIV_Tranno.Visible = False
End If
End Sub

Private Sub cboClasscode_Click()
If cboClasscode.Text = "RRV" Then
   labRIV_TranNo.Visible = True
   txtRIV_Tranno.Visible = True
Else
   labRIV_TranNo.Visible = False
   txtRIV_Tranno.Visible = False
End If
End Sub

Private Sub cboRecvd_Desc_Click()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DragDrop(Source As Control, X As Single, Y As Single)
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DropDown()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_LostFocus()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboTranDescription_Click()
If cboTranDescription.Text <> "" Then
   txtPartID.Text = SetPartIDDesc(cboTranDescription.Text)
   cboTranPartNo.Text = SetPartNo(txtPartID.Text)
   cboTranDescription.Text = SetPartDesc2(txtPartID.Text)
End If
End Sub

Private Sub cboTranDescription_LostFocus()
cboTranDescription.Text = UCase(cboTranDescription.Text)
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

Private Sub cmdAddTran_Click()
If Picture1.Visible = True Then
   SendToBack
   cmdAddTran.ZOrder 0
   fraAddTran.ZOrder 0
   cmdTranDelete.Visible = False
   fraAddTran.Enabled = True
   AddorEdit = "ADD"
   InitParts
   cboTranPartNo.SetFocus
End If
End Sub

Private Sub cmdCancelRR_Click()
If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
   Dim PCurOnOrder, PCurTRECQTY, PCurReceipts As Integer
   Dim PCurLast_recq, PCurTpoQty As Integer
   Dim rsRECON_DAYTRANDup, rsRECON_PARTMASDup As ADODB.Recordset
   Set rsRECON_DAYTRANDup = New ADODB.Recordset
       rsRECON_DAYTRANDup.Open "select trantype,tranno,tranqty,part_ord from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno), gconPMIOS
   If Not rsRECON_DAYTRANDup.EOF And Not rsRECON_DAYTRANDup.BOF Then
      rsRECON_DAYTRANDup.MoveFirst
      Do While Not rsRECON_DAYTRANDup.EOF
         Set rsRECON_PARTMASDup = New ADODB.Recordset
             rsRECON_PARTMASDup.Open "select partno,onorder,trecqty,receipts,last_recq from RECON_PARTMAS where partno = " & N2Str2Null(rsRECON_DAYTRANDup!part_ord), gconPMIOS
         If Not rsRECON_PARTMASDup.EOF And Not rsRECON_PARTMASDup.BOF Then
            PCurOnOrder = N2Str2IntZero(rsRECON_PARTMASDup!onorder) + N2Str2IntZero(rsRECON_DAYTRANDup!tranqty)
            PCurTRECQTY = N2Str2IntZero(rsRECON_PARTMASDup!trecqty) - N2Str2IntZero(rsRECON_DAYTRANDup!tranqty)
            PCurReceipts = N2Str2IntZero(rsRECON_PARTMASDup!RECEIPTS) - N2Str2IntZero(rsRECON_DAYTRANDup!tranqty)
            PCurLast_recq = N2Str2IntZero(rsRECON_PARTMASDup!last_recq) - N2Str2IntZero(rsRECON_DAYTRANDup!tranqty)
            gconPMIOS.Execute "update RECON_PARTMAS set" & _
                             " onorder = " & PCurOnOrder & "," & _
                             " trecqty = " & PCurTRECQTY & "," & _
                             " receipts = " & PCurReceipts & "," & _
                             " last_recq = " & PCurLast_recq & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where partno = " & N2Str2Null(rsRECON_DAYTRANDup!part_ord)
         End If
         rsRECON_DAYTRANDup.MoveNext
      Loop
   End If
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                    " status = 'C'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   gconPMIOS.Execute "update RECON_DAYTRAN set" & _
                    " status = 'C'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno) & " and trantype = 'RR'"
   rsRefresh
   On Error Resume Next
   rsRECON_REC_HIST.Find "id =" & labID.Caption
   StoreMemvars
End If
End Sub

Private Sub cmdOkUpdate_Click()
If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 1 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                    " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                    " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 0 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                    " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 0 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 1 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 1 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                    " mac = " & NumericVal(txtNewMAC.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 0 And chkUpdateSRP.Value = 1 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 0 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
If chkUpdateMAC.Value = 0 And chkUpdateDNP.Value = 1 And chkUpdateSRP.Value = 1 Then
   gconPMIOS.Execute "update RECON_PARTMAS set" & _
                    " dnp = " & NumericVal(txtNewDNP.Text) & ", " & _
                    " srp = " & NumericVal(txtNewSRP.Text) & ", " & _
                    " partdesc = " & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & _
                    " where partno = " & UCase(N2Str2Null(cboTranPartNo.Text))
End If
cleargrid grdDetails
FillDetails
If NumericVal(txtDS1.Text) > 0 Then
   RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = '" & "VAT" & "'," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
Else
   RR_TOTVAT = 0
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = NULL," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsRECON_REC_HIST.Find "id = " & labID.Caption
cmdTranCancel.Value = True
If AddorEdit = "ADD" Then cmdAddTran_Click
Screen.MousePointer = 0
Send2BackConfirm
End Sub

Private Sub cmdPost_Click()
Dim pmasOnOrder As Integer
If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
   Set rsRECON_DAYTRAN = New ADODB.Recordset
       rsRECON_DAYTRAN.Open "select id,itemno,trantype,tranno,part_ord,tranqty,traninvamt from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno) & " order by itemno asc", gconPMIOS
   If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
      rsRECON_DAYTRAN.MoveFirst
      Do While Not rsRECON_DAYTRAN.EOF
         If N2Str2Zero(rsRECON_DAYTRAN!traninvamt) <= 0 Then
            MsgSpeechBox "Transaction with Invoice Amount equal to Zero Encountered!"
            Exit Sub
         End If
         rsRECON_DAYTRAN.MoveNext
      Loop
      rsRECON_DAYTRAN.MoveFirst
      Do While Not rsRECON_DAYTRAN.EOF
         Set rsRECON_PARTMAS = New ADODB.Recordset
             rsRECON_PARTMAS.Open "Select partno,onhand,trecqty,onorder,receipts from RECON_PARTMAS where partno = " & N2Str2Null(rsRECON_DAYTRAN!part_ord), gconPMIOS
         If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.EOF Then
            pmasOnOrder = N2Str2Zero(rsRECON_PARTMAS!onorder)
            If pmasOnOrder <= 0 Then
               pmasOnOrder = NumericVal(rsRECON_DAYTRAN!tranqty)
            End If
            gconPMIOS.Execute "update RECON_PARTMAS set onhand =" & N2Str2Zero(rsRECON_PARTMAS!Onhand) + NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " trecqty = " & N2Str2Zero(rsRECON_PARTMAS!trecqty) + NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " onorder = " & pmasOnOrder - NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " receipts = " & N2Str2Zero(rsRECON_PARTMAS!RECEIPTS) + NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " last_recq = " & N2Str2Zero(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " last_recd = '" & LOGDATE & "', " & _
                             " supcode = " & N2Str2Null(Mid(txtRecvd_Code.Text, 1, 5)) & _
                             " where partno = " & N2Str2Null(rsRECON_PARTMAS!PartNo)
            gconPMIOS.Execute "update RECON_DAYTRAN set" & _
                             " status = 'P'" & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where id = " & rsRECON_DAYTRAN!ID
         End If
         rsRECON_DAYTRAN.MoveNext
      Loop
   End If
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                    " status = 'P'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   rsRefresh
   On Error Resume Next
   rsRECON_REC_HIST.Find "id =" & labID.Caption
   StoreMemvars
End If
End Sub

Private Sub cmdPrint_Click()
Screen.MousePointer = 11
PrintSQLReport rptReceiving, PMIOS_REPORT_PATH & "rr.rpt", "{RECON_REC_HIST.rrno} = '" & txtRRNo.Text & "'", PMIOS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdTranCancel_Click()
SendToBack
StoreMemvars
End Sub

Private Sub cmdTranDelete_Click()
If MsgQuestionBox("Delete This Parts, Are you Sure?", "Delete Parts Entry") = True Then
   gconPMIOS.Execute "delete from RECON_DAYTRAN where id = " & labDetID.Caption
End If
Dim Cnt As Integer
Dim rsRECON_DAYTRANDup As ADODB.Recordset
Set rsRECON_DAYTRANDup = New ADODB.Recordset
    rsRECON_DAYTRANDup.Open "select id,itemno from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno) & " order by itemno asc", gconPMIOS
If Not rsRECON_DAYTRANDup.EOF And Not rsRECON_DAYTRANDup.BOF Then
   rsRECON_DAYTRANDup.MoveFirst
   Cnt = 0
   Do While Not rsRECON_DAYTRANDup.EOF
      Cnt = Cnt + 1
      gconPMIOS.Execute "update RECON_DAYTRAN set itemno = " & Format(Cnt, "0000") & " where id = " & rsRECON_DAYTRANDup!ID
      rsRECON_DAYTRANDup.MoveNext
   Loop
End If
FillDetails
If NumericVal(txtDS1.Text) > 0 Then
   RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = '" & "VAT" & "'," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
Else
   RR_TOTVAT = 0
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = NULL," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsRECON_REC_HIST.Find "id = " & labID.Caption
cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
On Error GoTo ErrorCode

If cboTranPartNo.Text = "" Then
   MsgSpeechBox "Part Number must have a value"
   On Error Resume Next
   cboTranPartNo.SetFocus
   Exit Sub
End If

If AddorEdit = "ADD" Then
   Dim rsRECON_DAYTRANClone As ADODB.Recordset
   Set rsRECON_DAYTRANClone = New ADODB.Recordset
       rsRECON_DAYTRANClone.Open "select trantype,tranno,itemno,part_ord from RECON_DAYTRAN where part_ord = '" & UCase(cboTranPartNo.Text) & "' and trantype = 'RR' and tranno =" & N2Str2Null(rsRECON_REC_HIST!rrno) & " order by itemno asc", gconPMIOS
   If Not rsRECON_DAYTRANClone.EOF And Not rsRECON_DAYTRANClone.BOF Then
      MsgSpeechBox "Part Number already used in this transaction"
      On Error Resume Next
      cboTranPartNo.SetFocus
      Exit Sub
   End If
End If

Dim RRTRANDATE, RRTRANNO, RRTRANTYPE As String
Dim RRITEMNO, RRPART_ORD, RRPART_SUP As String
Dim RRTRANQTY As Integer
Dim RRTRANUCOST, RRTRANINVAMT As Double
Dim RRSTATUS, RRIN_OUT As String

RRTRANDATE = N2Date2Null(txtRRDate.Text)
RRTRANTYPE = "'RR'"
RRTRANNO = N2Str2Null(txtRRNo.Text)
RRITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
RRPART_ORD = N2Str2Null(UCase(cboTranPartNo.Text))
RRPART_SUP = N2Str2Null(UCase(cboTranPartNo.Text))
RRTRANQTY = NumericVal(txtTranQty.Text)
RRTRANINVAMT = NumericVal(txtTranINVAmt.Text)
RRTRANUCOST = NumericVal(txtUnitCost.Text)
RRIN_OUT = "'I'"
RRSTATUS = "'N'"

Screen.MousePointer = 11
If RRTRANINVAMT <= 0 Then
   MsgSpeechBox "Warning: Invoice Amount must not be zero"
   Screen.MousePointer = 0
   Exit Sub
End If

If AddorEdit = "ADD" Then
   gconPMIOS.Execute "insert into RECON_DAYTRAN " & _
                    "(trandate,trantype,tranno,itemno,part_ord,part_sup,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                    " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                    " " & RRITEMNO & "," & RRPART_ORD & "," & _
                    " " & RRPART_SUP & ", " & RRTRANQTY & "," & _
                    " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
Else
   gconPMIOS.Execute "update RECON_DAYTRAN set" & _
                    " trandate = " & RRTRANDATE & "," & _
                    " trantype = " & RRTRANTYPE & "," & _
                    " tranno = " & RRTRANNO & "," & _
                    " itemno = " & RRITEMNO & "," & _
                    " part_ord = " & RRPART_ORD & "," & _
                    " part_sup = " & RRPART_SUP & "," & _
                    " tranqty = " & RRTRANQTY & "," & _
                    " tranucost = " & RRTRANUCOST & "," & _
                    " traninvamt = " & RRTRANINVAMT & "," & _
                    " lastupdate = '" & LOGDATE & "'," & _
                    " status = " & RRSTATUS & "," & _
                    " in_out = " & RRIN_OUT & "," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "" & _
                    " where id = " & labDetID.Caption
End If

Dim varPmasTrecqty, varPmasOnOrder, varPmasTpoqty As Long
If cboClasscode.Text = "PCS" Or cboClasscode.Text = "PCG" Then
   Dim rsRECON_PARTMASClone As ADODB.Recordset
   Set rsRECON_PARTMASClone = New ADODB.Recordset
       rsRECON_PARTMASClone.Open "select partno,tpoqty,onorder,mac,dnp,srp,onhand from RECON_PARTMAS where partno = " & RRPART_ORD, gconPMIOS
   If Not rsRECON_PARTMASClone.EOF And Not rsRECON_PARTMASClone.BOF Then
      PrevPmasMAC = Format(N2Str2Zero(rsRECON_PARTMASClone!MAC), MAXIMUM_DIGIT)
      PrevPmasDNP = Format(N2Str2Zero(rsRECON_PARTMASClone!dnp), MAXIMUM_DIGIT)
      PrevPmasSRP = Format(N2Str2Zero(rsRECON_PARTMASClone!srp), MAXIMUM_DIGIT)
      PrevPmasOnHand = N2Str2Zero(rsRECON_PARTMASClone!Onhand)
      NewPmasOnHand = RRTRANQTY
      If NumericVal(txtDS1.Text) <= 0 Then
         NewPmasDNP = RRTRANUCOST
      Else
         NewPmasDNP = RRTRANUCOST * 1.1
      End If
      If txtRecvd_Code.Text = vPAMCOR Then
         If PrevPmasOnHand <= 0 Then
            NewPmasMAC = (RRTRANUCOST * RRTRANQTY) / NewPmasOnHand
         Else
            NewPmasMAC = ((PrevPmasMAC * PrevPmasOnHand) + (RRTRANUCOST * RRTRANQTY)) / (NewPmasOnHand + PrevPmasOnHand)
         End If
         NewPmasSRP = Format(NewPmasDNP * 1.43 * 1.05, DIGIT_FORMAT)
      Else
         If PrevPmasOnHand <= 0 Then
            NewPmasMAC = (RRTRANUCOST * RRTRANQTY) / NewPmasOnHand
         Else
            NewPmasMAC = ((PrevPmasMAC * PrevPmasOnHand) + (RRTRANUCOST * RRTRANQTY)) / (NewPmasOnHand + PrevPmasOnHand)
         End If
         NewPmasSRP = Format(NewPmasDNP * 1.43 * 1.05, DIGIT_FORMAT)
      End If
      Send2FrontConfirm
      txtOldMAC.Text = Format(PrevPmasMAC, MAXIMUM_DIGIT)
      txtOldDNP.Text = Format(PrevPmasDNP, DIGIT_FORMAT)
      txtOldSRP.Text = Format(PrevPmasSRP, DIGIT_FORMAT)
      txtOldOH.Text = Format(PrevPmasOnHand, DIGIT_FORMAT)
      txtNewMAC.Text = Format(NewPmasMAC, MAXIMUM_DIGIT)
      txtNewDNP.Text = Format(NewPmasDNP, DIGIT_FORMAT)
      txtNewSRP.Text = Format(NewPmasSRP, DIGIT_FORMAT)
      txtNewOH.Text = Format(NewPmasOnHand, DIGIT_FORMAT)
      Screen.MousePointer = 0
   Else
      PrevPmasMAC = "0.00": PrevPmasDNP = "0.00": PrevPmasSRP = "0.00": PrevPmasOnHand = "0"
      NewPmasOnHand = RRTRANQTY
      If NumericVal(txtDS1.Text) <= 0 Then
         NewPmasDNP = RRTRANUCOST
      Else
         NewPmasDNP = RRTRANUCOST * 1.1
      End If
      If txtRecvd_Code.Text = vPAMCOR Then
         NewPmasMAC = (RRTRANUCOST * RRTRANQTY) / NewPmasOnHand
         NewPmasSRP = Format(NewPmasDNP * 1.43 * 1.05, DIGIT_FORMAT)
      Else
         NewPmasMAC = (RRTRANUCOST * RRTRANQTY) / NewPmasOnHand
         NewPmasSRP = Format(NewPmasDNP * 1.43 * 1.05, DIGIT_FORMAT)
      End If
      Send2FrontConfirm
      txtOldMAC.Text = Format(PrevPmasMAC, MAXIMUM_DIGIT)
      txtOldDNP.Text = Format(PrevPmasDNP, DIGIT_FORMAT)
      txtOldSRP.Text = Format(PrevPmasSRP, DIGIT_FORMAT)
      txtOldOH.Text = Format(PrevPmasOnHand, DIGIT_FORMAT)
      txtNewMAC.Text = Format(NewPmasMAC, MAXIMUM_DIGIT)
      txtNewDNP.Text = Format(NewPmasDNP, DIGIT_FORMAT)
      txtNewSRP.Text = Format(NewPmasSRP, DIGIT_FORMAT)
      txtNewOH.Text = Format(NewPmasOnHand, DIGIT_FORMAT)
      Screen.MousePointer = 0
      If txtRecvd_Code.Text = vPAMCOR Then
         gconPMIOS.Execute "insert into RECON_PARTMAS " & _
                           "(partno,partdesc,date_entered)" & _
                           " values (" & RRPART_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
      Else
         gconPMIOS.Execute "insert into RECON_PARTMAS " & _
                           "(partno,partdesc,date_entered)" & _
                           " values (" & RRPART_ORD & "," & UCase(N2Str2Null(Mid(cboTranDescription.Text, 1, 16))) & ",'" & LOGDATE & "')"
      End If
   End If
Else
   cleargrid grdDetails
   FillDetails
   If NumericVal(txtDS1.Text) > 0 Then
      RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
      gconPMIOS.Execute "update RECON_REC_HIST set" & _
                        " ttlrramt = " & RR_TOTUCOST & "," & _
                        " netrramt = " & RR_TOTINVAMT & "," & _
                        " ds_desc1 = '" & "VAT" & "'," & _
                        " ds_amt1 = " & RR_TOTVAT & "," & _
                        " ds1 = " & NumericVal(txtDS1.Text) & _
                        " where id = " & labID.Caption
   Else
      RR_TOTVAT = 0
      gconPMIOS.Execute "update RECON_REC_HIST set" & _
                        " ttlrramt = " & RR_TOTUCOST & "," & _
                        " netrramt = " & RR_TOTINVAMT & "," & _
                        " ds_desc1 = NULL," & _
                        " ds_amt1 = " & RR_TOTVAT & "," & _
                        " ds1 = " & NumericVal(txtDS1.Text) & _
                        " where id = " & labID.Caption
   End If
   rsRefresh
   On Error Resume Next
   rsRECON_REC_HIST.Find "id = " & labID.Caption
   cmdTranCancel.Value = True
   If AddorEdit = "ADD" Then
      cmdAddTran_Click
   End If
End If
Screen.MousePointer = 0
Exit Sub

ErrorCode:
ShowVBError
Screen.MousePointer = 0
Exit Sub
End Sub

Sub Send2FrontConfirm()
Frame1.Enabled = False
Picture1.Enabled = False
fraDetails.Enabled = False
cmdAddTran.Enabled = False
fraAddTran.Enabled = False
cmdUpdateMaster.ZOrder 0
fraUpdateMaster.ZOrder 0
txtOldMAC.Text = 0
txtOldDNP.Text = 0
txtOldSRP.Text = 0
txtOldOH.Text = 0
txtNewMAC.Text = 0
txtNewDNP.Text = 0
txtNewSRP.Text = 0
txtNewOH.Text = 0
chkUpdateMAC.Value = 1
chkUpdateDNP.Value = 1
chkUpdateSRP.Value = 1
cmdOkUpdate.SetFocus
End Sub

Sub Send2BackConfirm()
Frame1.Enabled = True
Picture1.Enabled = True
fraDetails.Enabled = True
cmdAddTran.Enabled = True
fraAddTran.Enabled = True
cmdUpdateMaster.ZOrder 1
fraUpdateMaster.ZOrder 1
txtOldMAC.Text = 0
txtOldDNP.Text = 0
txtOldSRP.Text = 0
txtOldOH.Text = 0
txtNewMAC.Text = 0
txtNewDNP.Text = 0
txtNewSRP.Text = 0
txtNewOH.Text = 0
chkUpdateMAC.Value = 1
chkUpdateDNP.Value = 1
chkUpdateSRP.Value = 1
End Sub

Private Sub cmdUnPost_Click()
If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
   Set rsRECON_DAYTRAN = New ADODB.Recordset
       rsRECON_DAYTRAN.Open "select id,itemno,trantype,tranno,part_ord,tranqty,traninvamt from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno) & " order by itemno asc", gconPMIOS
   If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
      rsRECON_DAYTRAN.MoveFirst
      Do While Not rsRECON_DAYTRAN.EOF
         Set rsRECON_PARTMAS = New ADODB.Recordset
             rsRECON_PARTMAS.Open "Select partno,onhand,trecqty,onorder,receipts from RECON_PARTMAS where partno = " & N2Str2Null(rsRECON_DAYTRAN!part_ord), gconPMIOS
         If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.EOF Then
            gconPMIOS.Execute "update RECON_PARTMAS set onhand =" & N2Str2Zero(rsRECON_PARTMAS!Onhand) - NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " trecqty = " & N2Str2Zero(rsRECON_PARTMAS!trecqty) - NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " onorder = " & N2Str2Zero(rsRECON_PARTMAS!onorder) + NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " receipts = " & N2Str2Zero(rsRECON_PARTMAS!RECEIPTS) - NumericVal(rsRECON_DAYTRAN!tranqty) & ", " & _
                             " last_recq = " & 0 & ", " & _
                             " last_recd = NULL, " & _
                             " supcode = NULL" & _
                             " where partno = " & N2Str2Null(rsRECON_PARTMAS!PartNo)
            gconPMIOS.Execute "update RECON_DAYTRAN set" & _
                             " status = 'N'" & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where id = " & rsRECON_DAYTRAN!ID
         End If
         rsRECON_DAYTRAN.MoveNext
      Loop
   End If
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                    " status = 'N'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   rsRefresh
   On Error Resume Next
   rsRECON_REC_HIST.Find "id =" & labID.Caption
   StoreMemvars
End If
End Sub

Private Sub cmdAdd_Click()
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
textSearch.SetFocus
Picture5.Visible = False
'Dim findStr As String
'findStr = InputBoxXP("Please Input Tran. No.", txtRRNo.Text)
'findStr = Format(findStr, "000000")
'if findStr <> "" Then
'   On Error GoTo ErrorCode
'   rsRECON_REC_HIST.Bookmark = rsFind(rsRECON_REC_HIST.Clone, "rrno", findStr).Bookmark
'End If
'StoreMemvars
'Exit Sub

'ErrorCode:
'If Err.Number = 3021 Then
'   ShowCantFind findStr
'   Resume Next
'End If
End Sub

Sub FindDupRRno(DDD As String)
rsRECON_REC_HIST.Bookmark = rsFind(rsRECON_REC_HIST.Clone, "rrno", Format(DDD, "000000")).Bookmark
StoreMemvars
End Sub

Private Sub cmdFirst_Click()
rsRECON_REC_HIST.MoveFirst
StoreMemvars
End Sub

Private Sub cmdLast_Click()
rsRECON_REC_HIST.MoveLast
StoreMemvars
End Sub

Private Sub cmdNext_Click()
rsRECON_REC_HIST.MoveNext
If rsRECON_REC_HIST.EOF Then
   rsRECON_REC_HIST.MoveLast
   ShowLastRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
rsRECON_REC_HIST.MovePrevious
If rsRECON_REC_HIST.BOF Then
   rsRECON_REC_HIST.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemvars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim rsRECON_REC_HISTDup As ADODB.Recordset
      
If txtINVNo.Text = "" Then
   MsgSpeechBox "Reference Invoice Number must be inputed!"
   On Error Resume Next
   txtINVNo.SetFocus
   Exit Sub
End If
If cboClasscode.Text = "PCG" Then
   If cboRecvd_Desc.Text <> "PAMCOR" Then
      If txtTerms.Text = "" Then
         MsgSpeechBox "Warning: Terms must be Inputed"
         On Error Resume Next
         txtTerms.SetFocus
         Exit Sub
      End If
   Else
      If txtTerms.Text = "" Then txtTerms.Text = "90"
   End If
End If
If Trim(txtRRNo.Text) = "" Then
   MsgSpeechBox "MRR Number must not be empty"
   On Error Resume Next
   txtRRNo.SetFocus
   Exit Sub
Else
   If AddorEdit = "ADD" Then
      Dim rsfindDup As ADODB.Recordset
      Set rsfindDup = New ADODB.Recordset
          rsfindDup.Open "select rrno from RECON_REC_HIST where rrno = '" & txtRRNo.Text & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
      If Not rsfindDup.EOF And Not rsfindDup.BOF Then
         MsgSpeechBox "MRR Number already exist!"
         On Error Resume Next
         txtRRNo.SetFocus
         Exit Sub
      End If
      Set rsRECON_REC_HISTDup = New ADODB.Recordset
          rsRECON_REC_HISTDup.Open "select pono from RECON_REC_HIST where pono = '" & txtPONo.Text & "'", gconPMIOS
      If Not rsRECON_REC_HISTDup.EOF And Not rsRECON_REC_HISTDup.BOF Then
         MsgSpeechBox "Purchase Order Number Already Received"
         On Error Resume Next
         txtPONo.SetFocus
         Exit Sub
      End If
   End If
End If
If txtRRDate.Text = "" Or IsDate(txtRRDate.Text) = False Then
   MsgSpeechBox "Invalid MRR Date!"
   On Error Resume Next
   txtRRDate.SetFocus
   Exit Sub
End If

Dim NewRRCunTer As String
NewRRCunTer = NumericVal(txtRRNo.Text) + 1

Dim VTXTRRNo, VTXTRRDate, Vcboclasscode As String
Dim VTXTRecvd_Code, VTXTRecvd_From, VTXTAddress As String
Dim VTXTTerms, VTXTPONo, VTXTPODate As String
Dim VTXTDRNo, VTXTINVNo As String
Dim VTXTTTLRRAmt, VTXTDS1 As Double
Dim VTXTDS_Desc1 As String
Dim VTXTDS_Amt1, VTXTNetRRAmt As Double
Dim VTXTRemarks As String
Dim VTXTRIV_Tranno As String
Dim RRTRANDATE, RRTRANNO, RRTRANTYPE As String
Dim RRITEMNO, RRPART_ORD, RRPART_SUP As String
Dim RRTRANQTY As Integer
Dim RRTRANUCOST, RRTRANINVAMT As Double
Dim RRIN_OUT, RRSTATUS As String

VTXTRRNo = N2Str2Null(txtRRNo.Text)
VTXTRRDate = N2Date2Null(txtRRDate.Text)
Vcboclasscode = N2Str2Null(UCase(cboClasscode.Text))
VTXTRIV_Tranno = N2Str2Null(txtRIV_Tranno.Text)
VTXTRecvd_Code = N2Str2Null(txtRecvd_Code.Text)
VTXTRecvd_From = N2Str2Null(cboRecvd_Desc.Text)
VTXTAddress = N2Str2Null(txtDetails.Text)
VTXTTerms = N2Str2Null(txtTerms.Text)
VTXTPONo = N2Str2Null(txtPONo.Text)
VTXTPODate = N2Date2Null(txtPODate.Text)
VTXTDRNo = N2Str2Null(txtDRNo.Text)
VTXTINVNo = N2Str2Null(txtINVNo.Text)
VTXTTTLRRAmt = NumericVal(txtTTLRRAmt.Text)
VTXTDS1 = NumericVal(txtDS1.Text)
VTXTDS_Desc1 = N2Str2Null(txtDS_Desc1.Text)
VTXTDS_Amt1 = NumericVal(txtDS_Amt1.Text)
VTXTNetRRAmt = NumericVal(txtNetRRAmt.Text)
If txtRemarks.Text = "Pls Type Your Message Here!" Then
   VTXTRemarks = "NULL"
Else
   VTXTRemarks = N2Str2Null(Trim(txtRemarks.Text))
End If

If AddorEdit = "ADD" Then
   Set rsRECON_REC_HISTDup = New ADODB.Recordset
       rsRECON_REC_HISTDup.Open "select id from RECON_REC_HIST order by id desc", gconPMIOS
   If Not rsRECON_REC_HISTDup.EOF And Not rsRECON_REC_HISTDup.BOF Then
      rsRECON_REC_HISTDup.MoveFirst
      labID.Caption = NumericVal(rsRECON_REC_HISTDup!ID) + 1
   End If
   gconPMIOS.Execute "Insert into RECON_REC_HIST" & _
                    " (rrno,rrdate,classcode,RIV_Tranno,recvd_code,recvd_from,address,terms,pono,podate,drno,invno,ttlrramt,ds1,ds_desc1,ds_amt1,netrramt,usercode,lastupdate,remarks)" & _
                    " values (" & VTXTRRNo & ", " & VTXTRRDate & ", " & Vcboclasscode & ", " & VTXTRIV_Tranno & _
                    ", " & VTXTRecvd_Code & ", " & VTXTRecvd_From & ", " & VTXTAddress & ", " & VTXTTerms & _
                    ", " & VTXTPONo & ", " & VTXTPODate & ", " & VTXTDRNo & ", " & VTXTINVNo & _
                    ", " & VTXTTTLRRAmt & _
                    ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                    ", " & VTXTNetRRAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
Else
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                    " rrno = " & VTXTRRNo & "," & _
                    " rrdate = " & VTXTRRDate & "," & _
                    " classcode = " & Vcboclasscode & "," & _
                    " RIV_Tranno = " & VTXTRIV_Tranno & "," & _
                    " recvd_code = " & VTXTRecvd_Code & "," & _
                    " recvd_from = " & VTXTRecvd_From & "," & _
                    " address = " & VTXTAddress & "," & _
                    " terms = " & VTXTTerms & "," & _
                    " pono = " & VTXTPONo & "," & _
                    " podate = " & VTXTPODate & "," & _
                    " drno = " & VTXTDRNo & "," & _
                    " invno = " & VTXTINVNo & "," & _
                    " ttlrramt = " & VTXTTTLRRAmt & "," & _
                    " ds1 = " & VTXTDS1 & "," & _
                    " ds_desc1 = " & VTXTDS_Desc1 & "," & _
                    " ds_amt1 = " & VTXTDS_Amt1 & "," & _
                    " netrramt = " & VTXTNetRRAmt & "," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'," & _
                    " remarks = " & VTXTRemarks & _
                    " where id = " & labID.Caption
   gconPMIOS.Execute "update RECON_DAYTRAN set" & _
                    " trandate = " & VTXTRRDate & "," & _
                    " tranno = " & VTXTRRNo & _
                    " where trantype = 'RR' and tranno = '" & PrevRRNo & "'"
End If
If AddorEdit = "ADD" Then
   gconPMIOS.Execute "update cunter set nextnumber = '" & NewRRCunTer & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where modul = 'RR'"
End If
rsRefresh
On Error Resume Next
rsRECON_REC_HIST.Find "rrno = " & VTXTRRNo
cmdCancel.Value = True
On Error GoTo ErrorCode
If AddorEdit = "ADD" Then
   Dim rsRECON_DAYTRANDup, rsRECON_DAYTRANDUp2 As ADODB.Recordset
   Dim varPmasTrecqty, varPmasOnOrder, varPmasOnhand As Long
   Dim rsRECON_PARTMASClone As ADODB.Recordset
   Set rsRECON_DAYTRANDup = New ADODB.Recordset
       rsRECON_DAYTRANDup.Open "select trantype,tranno from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno), gconPMIOS
   If rsRECON_DAYTRANDup.EOF And rsRECON_DAYTRANDup.BOF Then
      rsRECON_DAYTRANDup.Close
      Set rsRECON_DAYTRANDUp2 = New ADODB.Recordset
          rsRECON_DAYTRANDUp2.Open "select trantype,tranno,part_ord,part_sup,itemno,tranqty,traninvamt,tranucost from RECON_DAYTRAN where trantype = 'PO' and tranno = " & N2Str2Null(rsRECON_REC_HIST!pono), gconPMIOS
      If Not rsRECON_DAYTRANDUp2.EOF And Not rsRECON_DAYTRANDUp2.BOF Then
         rsRECON_DAYTRANDUp2.MoveFirst
         Do While Not rsRECON_DAYTRANDUp2.EOF
            RRTRANDATE = N2Date2Null(txtPODate.Text)
            RRTRANTYPE = "'RR'"
            RRTRANNO = N2Str2Null(rsRECON_REC_HIST!rrno)
            RRITEMNO = N2Str2Null(Null2String(rsRECON_DAYTRANDUp2!itemno))
            RRPART_ORD = UCase(N2Str2Null(rsRECON_DAYTRANDUp2!part_ord))
            RRPART_SUP = UCase(N2Str2Null(rsRECON_DAYTRANDUp2!part_sup))
            RRTRANQTY = N2Str2IntZero(rsRECON_DAYTRANDUp2!tranqty)
            RRTRANINVAMT = N2Str2Zero(rsRECON_DAYTRANDUp2!traninvamt)
            RRTRANUCOST = N2Str2Zero(rsRECON_DAYTRANDUp2!tranucost)
            RRIN_OUT = "'I'"
            RRSTATUS = "'N'"
            
            gconPMIOS.Execute "insert into RECON_DAYTRAN " & _
                             "(trandate,trantype,tranno,itemno,part_ord,part_sup,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                             " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                             " " & RRITEMNO & "," & RRPART_ORD & "," & _
                             " " & RRPART_SUP & ", " & RRTRANQTY & "," & _
                             " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
            rsRECON_DAYTRANDUp2.MoveNext
         Loop
      End If
   Else
      cleargrid grdDetails
      FillDetails
      cmdAddTran_Click
   End If
End If

cleargrid grdDetails
FillDetails
If NumericVal(txtDS1.Text) > 0 Then
   RR_TOTVAT = RR_TOTINVAMT - RR_TOTUCOST
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = '" & "VAT" & "'," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
Else
   RR_TOTVAT = 0
   gconPMIOS.Execute "update RECON_REC_HIST set" & _
                     " ttlrramt = " & RR_TOTUCOST & "," & _
                     " netrramt = " & RR_TOTINVAMT & "," & _
                     " ds_desc1 = NULL," & _
                     " ds_amt1 = " & RR_TOTVAT & "," & _
                     " ds1 = " & NumericVal(txtDS1.Text) & _
                     " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsRECON_REC_HIST.Find "rrno = " & VTXTRRNo
StoreMemvars
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
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
               If Null2String(rsRECON_REC_HIST!Status) = "P" Then
                  MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
               ElseIf Null2String(rsRECON_REC_HIST!Status) = "C" Then
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
textSearch.Text = "": Picture5.ZOrder 0
Frame1.Enabled = False: Picture1.Visible = True: Picture2.Visible = False
txtPartID.Text = "": initMemvars: StoreMemvars
chkUpdateMAC.Enabled = False: chkUpdateDNP.Enabled = False
txtNewMAC.Enabled = False: txtNewDNP.Enabled = False
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsRECON_REC_HIST = New ADODB.Recordset
    rsRECON_REC_HIST.Open "select * from RECON_REC_HIST order by rrno desc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
txtRRNo.Text = ""
txtPONo.Text = ""
Set rsCunter = New ADODB.Recordset
    rsCunter.Open "select * from cunter where modul = 'RR'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
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
initGrid
InitCbo
InitCboClasscode
InitParts
End Sub

Sub StoreMemvars()
If Not rsRECON_REC_HIST.EOF And Not rsRECON_REC_HIST.BOF Then
   labID.Caption = rsRECON_REC_HIST!ID
   txtRRNo.Text = Null2String(rsRECON_REC_HIST!rrno)
   txtRRDate.Text = Null2String(rsRECON_REC_HIST!rrdate)
   cboClasscode.Text = Null2String(rsRECON_REC_HIST!classcode)
   txtRIV_Tranno.Text = Null2String(rsRECON_REC_HIST!RIV_Tranno)
   txtRecvd_Code.Text = Null2String(rsRECON_REC_HIST!recvd_code)
   cboRecvd_Desc.Text = Null2String(rsRECON_REC_HIST!recvd_from)
   txtDetails.Text = Null2String(rsRECON_REC_HIST!Address)
   txtTerms.Text = Null2String(rsRECON_REC_HIST!terms)
   txtPONo.Text = Null2String(rsRECON_REC_HIST!pono)
   txtPODate.Text = Null2String(rsRECON_REC_HIST!podate)
   txtDRNo.Text = Null2String(rsRECON_REC_HIST!drno)
   txtINVNo.Text = Null2String(rsRECON_REC_HIST!invno)
   txtDS1.Text = N2Str2IntZero(rsRECON_REC_HIST!ds1)
   txtDS_Desc1.Text = Null2String(rsRECON_REC_HIST!ds_desc1)
   txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsRECON_REC_HIST!ds_amt1))
   txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsRECON_REC_HIST!ttlrramt))
   txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsRECON_REC_HIST!netrramt))
   txtRemarks.Text = Null2String(rsRECON_REC_HIST!remarks)
   If Null2String(rsRECON_REC_HIST!Status) = "P" Then
      labRRsted.Visible = True
      labRRsted.Caption = "POSTED"
      cmdEdit.Enabled = False
      cmdPost.Enabled = False
      cmdPrint.Enabled = True
      If LOGLEVEL = "ADMIN" Then cmdCancelRR.Enabled = True
      If LOGLEVEL = "ADMIN" Then cmdUnPost.Enabled = True
   ElseIf Null2String(rsRECON_REC_HIST!Status) = "C" Then
      labRRsted.Visible = True
      labRRsted.Caption = "CANCELLED"
      cmdEdit.Enabled = False
      cmdPost.Enabled = False
      cmdUnPost.Enabled = False
      cmdPrint.Enabled = False
      cmdCancelRR.Enabled = False
      cmdUnPost.Enabled = False
   Else
      labRRsted.Visible = False
      labRRsted.Caption = ""
      cmdEdit.Enabled = True
      cmdPost.Enabled = True
      cmdPrint.Enabled = True
      If LOGLEVEL = "ADMIN" Then cmdCancelRR.Enabled = True
      cmdUnPost.Enabled = False
   End If
   cleargrid grdDetails
   FillDetails
Else
   ShowNoRecord
   cmdAdd.Value = True
End If
End Sub

Sub initGrid()
With grdDetails
   .ColWidth(0) = 1
   .ColWidth(1) = 800
   .ColWidth(2) = 1500
   .ColWidth(3) = 2500
   .ColWidth(4) = 500
   .ColWidth(5) = 1100
   .ColWidth(6) = 1100
   .ColWidth(7) = 1500
   
   '.ColWidth(0) = 1
   '.ColWidth(1) = 800
   '.ColWidth(2) = 1500
   '.ColWidth(3) = 2200
   '.ColWidth(4) = 500
   '.ColWidth(5) = 1
   '.ColWidth(6) = 900
   '.ColWidth(7) = 1200
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
On Error GoTo ErrorCode
Pcnt = 0
RR_TOTUCOST = 0
RR_TOTINVAMT = 0
RR_TOTVAT = 0
RR_QTY_REC = 0
Set rsRECON_DAYTRAN = New ADODB.Recordset
    rsRECON_DAYTRAN.Open "select id,trantype,tranno,itemno,part_ord,part_sup,tranqty,tranucost,traninvamt from RECON_DAYTRAN where trantype = 'RR' and tranno = " & N2Str2Null(rsRECON_REC_HIST!rrno) & " order by itemno asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
   Screen.MousePointer = 11
   rsRECON_DAYTRAN.MoveFirst
   Do While Not rsRECON_DAYTRAN.EOF
      Pcnt = Pcnt + 1
      grdDetails.AddItem rsRECON_DAYTRAN!ID & Chr(9) & Null2String(rsRECON_DAYTRAN!itemno) & Chr(9) & _
                      Null2String(rsRECON_DAYTRAN!part_ord) & Chr(9) & _
                      SetPartDesc(Null2String(rsRECON_DAYTRAN!part_sup)) & Chr(9) & _
                      N2Str2IntZero(rsRECON_DAYTRAN!tranqty) & Chr(9) & _
                      N2Str2Zero(rsRECON_DAYTRAN!traninvamt) & Chr(9) & _
                      N2Str2Zero(rsRECON_DAYTRAN!tranucost) & Chr(9) & _
                      Format(N2Str2IntZero(rsRECON_DAYTRAN!tranqty) * N2Str2Zero(rsRECON_DAYTRAN!tranucost), MAXIMUM_DIGIT)
      RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsRECON_DAYTRAN!tranqty)
      RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsRECON_DAYTRAN!tranqty) * N2Str2Zero(rsRECON_DAYTRAN!tranucost))
      RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsRECON_DAYTRAN!tranqty) * N2Str2Zero(rsRECON_DAYTRAN!traninvamt))
      rsRECON_DAYTRAN.MoveNext
   Loop
   If Pcnt <> 0 Then grdDetails.RemoveItem 1
   If Null2String(rsRECON_REC_HIST!classcode) = "PCS" Or Null2String(rsRECON_REC_HIST!classcode) = "PCG" Then
      RR_TOTVAT = ToDoubleNumber(RR_TOTINVAMT - RR_TOTUCOST)
      'If ISNONVAT = True Then
      '   RR_TOTVAT = 0
      'Else
      '   RR_TOTVAT = (RR_TOTUCOST * 1.1) - RR_TOTUCOST
      'End If
   Else
      RR_TOTVAT = 0
   End If
   If NumericVal(RR_TOTVAT) <> 0 Then
      txtDS1.Text = 10
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

ErrorCode:
ShowVBError
Exit Sub
End Sub

Function SetPartDesc(ppp As String)
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "Select partno,partdesc from RECON_PARTMAS where partno= '" & ppp & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   SetPartDesc = Null2String(rsRECON_PARTMAS!PartDesc)
End If
End Function

Function SetPartDesc2(ppp As String)
If ppp <> "" Then
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "Select id,partdesc from RECON_PARTMAS where id = " & ppp, gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   SetPartDesc2 = Null2String(rsRECON_PARTMAS!PartDesc)
End If
End If
End Function

Function SetPartNo(DDD As String)
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "Select id,partno from RECON_PARTMAS where id = " & DDD, gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   SetPartNo = Null2String(rsRECON_PARTMAS!PartNo)
End If
End Function

Function SetPartIDPartNo(DDD As String)
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "Select id,partno from RECON_PARTMAS where partno = '" & DDD & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   SetPartIDPartNo = Null2String(rsRECON_PARTMAS!ID)
End If
End Function

Function SetPartIDDesc(DDD As String)
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "Select id,partdesc from RECON_PARTMAS where ucase(ltrim(rtrim(partdesc))) = '" & UCase(LTrim(RTrim(DDD))) & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   SetPartIDDesc = Null2String(rsRECON_PARTMAS!ID)
End If
End Function

Function SetPartPrice(ppp As String)
If ppp <> "" Then
  Set rsRECON_PARTMAS = New ADODB.Recordset
      rsRECON_PARTMAS.Open "Select partno,mac from RECON_PARTMAS where partno = '" & ppp & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
  If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
     SetPartPrice = Null2String(rsRECON_PARTMAS!MAC)
  End If
End If
End Function

Sub FillCboRecvd()
Set rsSupplier = New ADODB.Recordset
    rsSupplier.Open "select supname from supplier", gconPMIOS, adOpenForwardOnly, adLockReadOnly
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
    rsSupplier.Open "Select supcode,supname,sup_addrs,vat_percnt,NONVAT from supplier where supcode = '" & ppp & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
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
    rsSupplier.Open "Select supname,supcode,sup_addrs,vat_percnt,NONVAT from supplier where supname = '" & nnn & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
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
Set rsRECON_DAYTRAN = New ADODB.Recordset
    rsRECON_DAYTRAN.Open "select id,itemno,part_ord,part_sup,tranqty,traninvamt,tranucost from RECON_DAYTRAN where id = " & ID, gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_DAYTRAN.EOF And Not rsRECON_DAYTRAN.BOF Then
   labDetID.Caption = rsRECON_DAYTRAN!ID
   txtTranItemNo.Text = Null2String(rsRECON_DAYTRAN!itemno)
   cboTranPartNo.Text = Null2String(rsRECON_DAYTRAN!part_ord)
   cboTranDescription.Text = SetPartDesc(Null2String(rsRECON_DAYTRAN!part_sup))
   txtTranQty.Text = N2Str2IntZero(rsRECON_DAYTRAN!tranqty)
   txtTranINVAmt.Text = N2Str2Zero(rsRECON_DAYTRAN!traninvamt)
   txtUnitCost.Text = N2Str2Zero(rsRECON_DAYTRAN!tranucost)
   txtTranTotalAmt.Text = N2Str2IntZero(rsRECON_DAYTRAN!tranqty) * N2Str2Zero(rsRECON_DAYTRAN!traninvamt)
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Set frmPMIOSReceiving2 = Nothing
UnloadForm Me
End Sub

Private Sub grdDetails_DblClick()
Dim Fild As String
If Null2String(rsRECON_REC_HIST!Status) = "P" Then
   MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
ElseIf Null2String(rsRECON_REC_HIST!Status) = "C" Then
   MsgSpeechBox "Item(s) are Already Cancelled and cannot be edited"
Else
   grdDetails.Row = grdDetails.Row
   grdDetails.Col = 0
   Fild = grdDetails.Text
   If Fild <> "" And Fild <> "No Entry" Then
      AddorEdit = "EDIT"
      BringToFront
      cmdTranDelete.Visible = True
      fraAddTran.Caption = "Edit Parts"
      StorePartsEntry (Fild)
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
Send2BackConfirm
End Sub

Sub BringToFront()
cmdAddTran.ZOrder 0
fraAddTran.ZOrder 0
fraAddTran.Enabled = True
End Sub

Sub InitCbo()
Set rsRECON_PARTMAS = New ADODB.Recordset
    rsRECON_PARTMAS.Open "select partno,partdesc from RECON_PARTMAS", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsRECON_PARTMAS.EOF And Not rsRECON_PARTMAS.BOF Then
   rsRECON_PARTMAS.MoveFirst
   cboTranPartNo.Clear
   cboTranDescription.Clear
   Do While Not rsRECON_PARTMAS.EOF
      cboTranPartNo.AddItem Null2String(rsRECON_PARTMAS!PartNo)
      cboTranDescription.AddItem Null2String(rsRECON_PARTMAS!PartDesc)
      rsRECON_PARTMAS.MoveNext
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
   cboClasscode.Text = UCase(cboClasscode.Text)
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
       rsCunter.Open "select * from cunter where modul = 'PO'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsCunter.EOF And Not rsCunter.BOF Then
      txtPONo.Text = Format(N2Str2Zero(rsCunter!nextnumber) - 1, "000000")
   End If
End If
End Sub

Private Sub txtPONo_LostFocus()
If cboClasscode.Text = "PCG" Then
   If txtPONo.Text <> "" And AddorEdit = "ADD" And Len(txtPONo.Text) > 0 Then
      Dim rsRECON_REC_HISTDup As ADODB.Recordset
      Set rsRECON_REC_HISTDup = New ADODB.Recordset
          rsRECON_REC_HISTDup.Open "select pono from RECON_REC_HIST where pono = '" & txtPONo.Text & "'", gconPMIOS
      If Not rsRECON_REC_HISTDup.EOF And Not rsRECON_REC_HISTDup.BOF Then
         MsgBox "PO Number Already Received", vbInformation, "Invalid PO Number"
         Exit Sub
      End If
      Set rsRECON_PO_HIST = New ADODB.Recordset
          rsRECON_PO_HIST.Open "select pono,supcode,podate from RECON_PO_HIST where pono = '" & txtPONo.Text & "'", gconPMIOS
      If Not rsRECON_PO_HIST.EOF And Not rsRECON_PO_HIST.BOF Then
         txtRecvd_Code.Text = Null2String(rsRECON_PO_HIST!SupCode)
         txtPODate.Text = Null2String(rsRECON_PO_HIST!podate)
         Pcnt = 0
         RR_TOTUCOST = 0
         RR_TOTINVAMT = 0
         RR_TOTVAT = 0
         RR_QTY_REC = 0
         Dim rsRECON_DAYTRANDup As ADODB.Recordset
         Set rsRECON_DAYTRANDup = New ADODB.Recordset
             rsRECON_DAYTRANDup.Open "select id,itemno,part_ord,part_sup,tranqty,traninvamt,tranucost from RECON_DAYTRAN where trantype = 'PO' and tranno = " & N2Str2Null(rsRECON_PO_HIST!pono) & " order by itemno asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
         If Not rsRECON_DAYTRANDup.EOF And Not rsRECON_DAYTRANDup.BOF Then
            Screen.MousePointer = 11
            rsRECON_DAYTRANDup.MoveFirst
            cleargrid grdDetails
            Do While Not rsRECON_DAYTRANDup.EOF
               Pcnt = Pcnt + 1
               grdDetails.AddItem rsRECON_DAYTRANDup!ID & Chr(9) & Null2String(Format(rsRECON_DAYTRANDup!itemno, "0000")) & Chr(9) & _
                                  Null2String(rsRECON_DAYTRANDup!part_ord) & Chr(9) & _
                                  SetPartDesc(Null2String(rsRECON_DAYTRANDup!part_sup)) & Chr(9) & _
                                  N2Str2IntZero(rsRECON_DAYTRANDup!tranqty) & Chr(9) & _
                                  N2Str2Zero(rsRECON_DAYTRANDup!traninvamt) & Chr(9) & _
                                  N2Str2Zero(rsRECON_DAYTRANDup!tranucost) & Chr(9) & _
                                  N2Str2IntZero(rsRECON_DAYTRANDup!tranqty) * N2Str2Zero(rsRECON_DAYTRANDup!tranucost)
               RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsRECON_DAYTRANDup!tranqty) * N2Str2Zero(rsRECON_DAYTRANDup!tranucost))
               RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsRECON_DAYTRANDup!tranqty) * N2Str2Zero(rsRECON_DAYTRANDup!traninvamt))
               rsRECON_DAYTRANDup.MoveNext
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
   If Not rsRECON_REC_HIST.EOF And Not rsRECON_REC_HIST.BOF Then
      If Null2String(rsRECON_REC_HIST!classcode) = "PCS" Or Null2String(rsRECON_REC_HIST!classcode) = "PCG" Then
         If ISNONVAT = True Then
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
         Else
            txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / 1.1)
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
   If Null2String(rsRECON_REC_HIST!classcode) = "PCS" Or Null2String(rsRECON_REC_HIST!classcode) = "PCG" Then
      If ISNONVAT = True Then
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
      Else
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / 1.1)
      End If
      txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
   Else
      txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
      txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
   End If
Else
   txtTranQty.Text = 1
   If Null2String(rsRECON_REC_HIST!classcode) = "PCS" Or Null2String(rsRECON_REC_HIST!classcode) = "PCG" Then
      If ISNONVAT = True Then
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
      Else
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / 1.1)
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
If Null2String(rsRECON_REC_HIST!classcode) = "PCS" Or Null2String(rsRECON_REC_HIST!classcode) = "PCG" Then
   If NumericVal(txtTranINVAmt.Text) <> 0 Then
      If ISNONVAT = True Then
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
      Else
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / 1.1)
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
If Null2String(rsRECON_REC_HIST!rrno) = "PCS" Or Null2String(rsRECON_REC_HIST!rrno) = "PCS" Then
   If txtTranINVAmt.Text <> "" Then
      If ISNONVAT = True Then
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
      Else
         txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / 1.1)
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

'SEARCH MODULE
Private Sub lstREC_HIST_GotFocus()
rsRECON_REC_HIST.Bookmark = rsFind(rsRECON_REC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
StoreMemvars
End Sub

Private Sub lstREC_HIST_ItemClick(ByVal Item As MSComctlLib.ListItem)
If optRRNo.Value = True Then
   rsRECON_REC_HIST.Bookmark = rsFind(rsRECON_REC_HIST.Clone, "rrno", Item).Bookmark
Else
   rsRECON_REC_HIST.Bookmark = rsFind(rsRECON_REC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
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
cmdEdit.Value = True
End Sub

Private Sub lstREC_HIST_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then textSearch.SetFocus
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
Dim rsRECON_REC_HIST As ADODB.Recordset
lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
Set rsRECON_REC_HIST = New ADODB.Recordset
Set rsRECON_REC_HIST = gconPMIOS.Execute("select rrno,ID from RECON_REC_HIST order by rrno asc")
If Not (rsRECON_REC_HIST.EOF And rsRECON_REC_HIST.BOF) Then
   lstREC_Hist.Enabled = True
   Listview_Loadval Me.lstREC_Hist.ListItems, rsRECON_REC_HIST
   lstREC_Hist.Refresh
Else
   lstREC_Hist.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsRECON_REC_HIST As ADODB.Recordset
lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
Set rsRECON_REC_HIST = New ADODB.Recordset
Set rsRECON_REC_HIST = gconPMIOS.Execute("select rrno, ID from RECON_REC_HIST where rrno like'" & XXX & "%'")
If Not (rsRECON_REC_HIST.EOF And rsRECON_REC_HIST.BOF) Then
   lstREC_Hist.Enabled = True
   Listview_Loadval Me.lstREC_Hist.ListItems, rsRECON_REC_HIST
   lstREC_Hist.Refresh
Else
   lstREC_Hist.Enabled = False
End If
End Sub

Sub FillGrid2()
Dim rsRECON_REC_HIST As ADODB.Recordset
lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
Set rsRECON_REC_HIST = New ADODB.Recordset
Set rsRECON_REC_HIST = gconPMIOS.Execute("select recvd_from, ID from RECON_REC_HIST order by rrno asc")
If Not (rsRECON_REC_HIST.EOF And rsRECON_REC_HIST.BOF) Then
   lstREC_Hist.Enabled = True
   Listview_Loadval Me.lstREC_Hist.ListItems, rsRECON_REC_HIST
   lstREC_Hist.Refresh
Else
   lstREC_Hist.Enabled = True
End If
End Sub

Sub FillSearchGrid2(XXX As String)
Dim rsRECON_REC_HIST As ADODB.Recordset
lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
Set rsRECON_REC_HIST = New ADODB.Recordset
Set rsRECON_REC_HIST = gconPMIOS.Execute("select recvd_from, ID from RECON_REC_HIST where recvd_from like '" & XXX & "%' order by rrno asc")
If Not (rsRECON_REC_HIST.EOF And rsRECON_REC_HIST.BOF) Then
   lstREC_Hist.Enabled = True
   Listview_Loadval Me.lstREC_Hist.ListItems, rsRECON_REC_HIST
   lstREC_Hist.Refresh
Else
   lstREC_Hist.Enabled = False
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstREC_Hist.SetFocus
End Sub

Private Sub optRONo_Click()
lstREC_Hist.ColumnHeaders(1).Text = "Sup. Name"
lstREC_Hist.ColumnHeaders(1).Width = 4000
If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
textSearch.SetFocus
End Sub

Private Sub optRRNo_Click()
lstREC_Hist.ColumnHeaders(1).Text = "Tran. No."
lstREC_Hist.ColumnHeaders(1).Width = 2150
If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
textSearch.SetFocus
End Sub
