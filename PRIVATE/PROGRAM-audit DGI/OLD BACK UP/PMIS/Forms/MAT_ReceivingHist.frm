VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMISMat_ReceivingHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Receiving History"
   ClientHeight    =   7965
   ClientLeft      =   855
   ClientTop       =   855
   ClientWidth     =   12315
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_ReceivingHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12315
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   2190
      TabIndex        =   23
      Top             =   0
      Width           =   10065
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8670
         Top             =   2520
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1245
         Left            =   8250
         ScaleHeight     =   1245
         ScaleWidth      =   1725
         TabIndex        =   39
         Top             =   750
         Width           =   1725
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
            TabIndex        =   42
            Top             =   840
            Width           =   1635
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
            Left            =   60
            MaxLength       =   15
            TabIndex        =   41
            Top             =   450
            Width           =   1635
         End
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
            TabIndex        =   40
            Top             =   60
            Width           =   1635
         End
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   825
         Left            =   90
         ScaleHeight     =   825
         ScaleWidth      =   4455
         TabIndex        =   37
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
            TabIndex        =   38
            Top             =   0
            Width           =   4365
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
         Left            =   6660
         TabIndex        =   33
         ToolTipText     =   "Input the type of the additional amount (e.g. VAT)"
         Top             =   1200
         Width           =   1605
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
         TabIndex        =   32
         ToolTipText     =   "Type the terms of the transaction."
         Top             =   1050
         Width           =   1275
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
         TabIndex        =   31
         ToolTipText     =   "Type Receiving entry number (e.g 003294)"
         Top             =   180
         Width           =   1155
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
         TabIndex        =   30
         Text            =   "cboRecvd_Desc"
         ToolTipText     =   "Select the name of supplier from the list."
         Top             =   1440
         Width           =   4395
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
         Left            =   1200
         TabIndex        =   29
         ToolTipText     =   "Type the supplier's code (e.g. 00001) "
         Top             =   1020
         Width           =   1155
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
         Left            =   5280
         TabIndex        =   28
         Text            =   "cboRecvd_Desc"
         Top             =   390
         Width           =   2745
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
         TabIndex        =   27
         Text            =   "MAT_ReceivingHist.frx":08CA
         ToolTipText     =   "Type your massage or remarks."
         Top             =   2010
         Width           =   5355
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
         TabIndex        =   26
         ToolTipText     =   "Type the Receiving Entry DR Number,if there's any  (e.g. 555665)"
         Top             =   2670
         Width           =   1005
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
         TabIndex        =   25
         ToolTipText     =   "Type the Receiving Entry's Ref INV Number (e.g. 329874)"
         Top             =   2670
         Width           =   1005
      End
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
         TabIndex        =   24
         ToolTipText     =   "Type the percentage of the amount to be added. Do not include % sign (e.g. 10, 15)"
         Top             =   1230
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtPONo 
         Height          =   345
         Left            =   1200
         TabIndex        =   34
         ToolTipText     =   "Type purchase order number of the receiving entry (e.g. 02774)"
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
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
         TabIndex        =   35
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
         TabIndex        =   36
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
      Begin MSMask.MaskEdBox txtRIV_Tranno 
         Height          =   345
         Left            =   5280
         TabIndex        =   43
         Top             =   810
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
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
         TabIndex        =   59
         Top             =   150
         Width           =   2595
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
         TabIndex        =   58
         Top             =   2700
         Width           =   855
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
         TabIndex        =   57
         Top             =   1470
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Receive From"
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
         Height          =   405
         Left            =   90
         TabIndex        =   56
         Top             =   990
         Width           =   1005
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
         TabIndex        =   55
         Top             =   1110
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
         Left            =   5280
         TabIndex        =   54
         Top             =   150
         Width           =   1305
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
         TabIndex        =   53
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "RR #"
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
         Top             =   240
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
         TabIndex        =   51
         Top             =   690
         Width           =   795
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
         TabIndex        =   50
         Top             =   690
         Width           =   1275
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
         TabIndex        =   49
         Top             =   1770
         Width           =   885
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
         TabIndex        =   48
         Top             =   2700
         Width           =   795
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
         Left            =   7080
         TabIndex        =   47
         Top             =   1650
         Width           =   1965
      End
      Begin VB.Label Label9 
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
         Height          =   285
         Left            =   6990
         TabIndex        =   46
         Top             =   870
         Width           =   1965
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
         Left            =   6330
         TabIndex        =   45
         Top             =   1260
         Width           =   375
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
         TabIndex        =   44
         Top             =   870
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture6 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   12315
      TabIndex        =   19
      Top             =   7620
      Width           =   12315
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         BackColor       =   &H00C4F4CD&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " APJ #:"
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
         Index           =   2
         Left            =   60
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
      Begin VB.Label labAPJ 
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
         Left            =   930
         TabIndex        =   21
         Top             =   0
         Width           =   2145
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
         Left            =   3090
         TabIndex        =   20
         Top             =   0
         Width           =   9195
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7545
      Left            =   30
      TabIndex        =   2
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstREC_Hist 
         Height          =   6165
         Left            =   60
         TabIndex        =   6
         Top             =   1320
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   10874
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
         MouseIcon       =   "MAT_ReceivingHist.frx":08E4
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
         TabIndex        =   7
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3705
      Left            =   2190
      TabIndex        =   1
      Top             =   3030
      Width           =   10095
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   3435
         Left            =   60
         TabIndex        =   0
         Top             =   180
         Width           =   9945
         _ExtentX        =   17542
         _ExtentY        =   6059
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
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   3810
      ScaleHeight     =   870
      ScaleWidth      =   8535
      TabIndex        =   8
      Top             =   6750
      Width           =   8535
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
         MouseIcon       =   "MAT_ReceivingHist.frx":0A46
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":0B98
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
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
         MouseIcon       =   "MAT_ReceivingHist.frx":0EFE
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":1050
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   30
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
         Left            =   6180
         MouseIcon       =   "MAT_ReceivingHist.frx":13B6
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":1508
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
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
         Left            =   5430
         MouseIcon       =   "MAT_ReceivingHist.frx":1858
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":19AA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
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
         Left            =   4680
         MouseIcon       =   "MAT_ReceivingHist.frx":1D08
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":1E5A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   30
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
         Left            =   3930
         MouseIcon       =   "MAT_ReceivingHist.frx":2154
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":22A6
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   30
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
         Left            =   3180
         MouseIcon       =   "MAT_ReceivingHist.frx":25FE
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":2750
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10875
      ScaleHeight     =   855
      ScaleWidth      =   1470
      TabIndex        =   16
      Top             =   6750
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
         MouseIcon       =   "MAT_ReceivingHist.frx":2AAF
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":2C01
         Style           =   1  'Graphical
         TabIndex        =   18
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
         MouseIcon       =   "MAT_ReceivingHist.frx":2F51
         MousePointer    =   99  'Custom
         Picture         =   "MAT_ReceivingHist.frx":30A3
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmPMISMat_ReceivingHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSREC_HIST                                         As ADODB.Recordset
Dim RSPO_HIST                                          As ADODB.Recordset
Dim RSDAYTRAN                                          As ADODB.Recordset
Dim RSPARTMAS                                          As ADODB.Recordset
Dim rsSupplier                                         As ADODB.Recordset
Dim RSCUNTER                                           As ADODB.Recordset
Dim Pcnt                                               As Integer
Dim AddorEdit                                          As String
Dim RR_TOTUCOST                                        As Double
Dim RR_TOTINVAMT                                       As Double
Dim RR_TOTVAT                                          As Double
Dim RR_QTY_REC                                         As Long
Dim ISNONVAT                                           As Boolean

Function GetRecClassCode(XXX)
    Select Case XXX
        Case "IBT": GetRecClassCode = "INTER BRANCH TRANSFER"
        Case "PCG": GetRecClassCode = "PURCHASED CHARGE"
        Case "PCS": GetRecClassCode = "PURCHASED CASH"
        Case "RCG": GetRecClassCode = "RETURN FROM CHARGE"
        Case "RCS": GetRecClassCode = "RETURN FROM CASH"
        Case "REP": GetRecClassCode = "REPLACEMENT"
        Case "RRV": GetRecClassCode = "RETURNED FROM SERVICE"
    End Select

End Function

Function SetSTOCKDESC(ppp As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select STOCKNO,STOCKDESC from PMIS_STOCKMAS where STOCKNO= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKDESC = Null2String(RSPARTMAS!STOCKDESC)
    End If
End Function

Function SetSTOCKDESC2(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select id,STOCKDESC from PMIS_STOCKMAS where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetSTOCKDESC2 = Null2String(RSPARTMAS!STOCKDESC)
        End If
    End If
End Function

Function SetSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_STOCKMAS where id = " & DDD, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetSTOCKNO = Null2String(RSPARTMAS!STOCKNO)
    End If
End Function

Function SetPartIDSTOCKNO(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKNO from PMIS_STOCKMAS where STOCKNO = " & N2Str2Null(DDD) & "", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDSTOCKNO = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartIDDesc(DDD As String)
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select id,STOCKDESC from PMIS_STOCKMAS where ltrim(rtrim(STOCKDESC))) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        SetPartIDDesc = Null2String(RSPARTMAS!ID)
    End If
End Function

Function SetPartPrice(ppp As String)
    If ppp <> "" Then
        Set RSPARTMAS = New ADODB.Recordset
        RSPARTMAS.Open "Select STOCKNO,mac from PMIS_STOCKMAS where STOCKNO = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
            SetPartPrice = Null2String(RSPARTMAS!MAC)
        End If
    End If
End Function

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

Sub FindDupRRno(DDD As String)
    RSREC_HIST.Bookmark = rsFind(RSREC_HIST.Clone, "rrno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    Set RSREC_HIST = New ADODB.Recordset
    RSREC_HIST.Open "select * from PMIS_Rec_Hist where type = 'M' order by ID desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtRRNo.Text = ""
    txtPONo.Text = ""
    Set RSCUNTER = New ADODB.Recordset
    RSCUNTER.Open "select * from PMIS_Counter where modul = 'RR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
        txtRRNo.Text = Format(Null2String(RSCUNTER!nextnumber), "000000")
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
    txtremarks.Text = "Pls Type Your Message Here!"
    labRRsted.Caption = ""
    cleargrid grdDetails
    InitGrid

    InitCboClasscode

End Sub

Sub StoreMemVars()

    If Not RSREC_HIST.EOF And Not RSREC_HIST.BOF Then
        labAPJ = "": labDetails = ""
        labid.Caption = RSREC_HIST!ID
        txtRRNo.Text = Null2String(RSREC_HIST!RRNO)
        txtRRDate.Text = Null2String(RSREC_HIST!RRDATE)
        cboClasscode.Text = GetRecClassCode(Null2String(RSREC_HIST!classcode))
        txtRIV_Tranno.Text = Null2String(RSREC_HIST!RIV_Tranno)
        txtRecvd_Code.Text = Null2String(RSREC_HIST!recvd_code)
        cboRecvd_Desc.Text = Null2String(RSREC_HIST!recvd_from)
        txtDetails.Text = Null2String(RSREC_HIST!Address)
        txtTerms.Text = Null2String(RSREC_HIST!Terms)
        txtPONo.Text = Null2String(RSREC_HIST!PONO)

        txtPODate.Text = Null2String(RSREC_HIST!PODATE)
        txtDRNo.Text = Null2String(RSREC_HIST!drno)
        txtINVNo.Text = Null2String(RSREC_HIST!invno)
        txtDS1.Text = N2Str2IntZero(RSREC_HIST!ds1)
        txtDS_Desc1.Text = Null2String(RSREC_HIST!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(RSREC_HIST!DS_AMT1))
        txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(RSREC_HIST!ttlrramt))
        txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(RSREC_HIST!netrramt))
        txtremarks.Text = Null2String(RSREC_HIST!REMARKS)
        labAPJ = CheckAPJNum(Null2String(RSREC_HIST!RRNO), "MI")
        If Null2String(RSREC_HIST!Status) = "P" Then
            If labAPJ <> "" Then
                labDetails = "TRANSACTION IMPORTED TO ACCOUNTING"
            End If
            labRRsted.Visible = True
            labRRsted.Caption = "POSTED"


            cmdPrint.Enabled = True
        ElseIf Null2String(RSREC_HIST!Status) = "C" Then
            labRRsted.Visible = True
            labRRsted.Caption = "CANCELLED"


            cmdPrint.Enabled = False

        Else
            labRRsted.Visible = False
            labRRsted.Caption = ""


            cmdPrint.Enabled = False

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
        .ColAlignment(2) = 2
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
    Set RSDAYTRAN = New ADODB.Recordset
    RSDAYTRAN.Open "select id,trantype,tranno,itemno,STOCK_ORD,STOCK_SUP,tranqty,tranucost,traninvamt from PMIS_DayTran where type = 'M' and trantype = 'RR' and tranno = " & N2Str2Null(RSREC_HIST!RRNO) & " and trandate = '" & Null2String(RSREC_HIST!RRDATE) & "' order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not RSDAYTRAN.EOF And Not RSDAYTRAN.BOF Then
        Screen.MousePointer = 11
        RSDAYTRAN.MoveFirst
        Do While Not RSDAYTRAN.EOF
            Pcnt = Pcnt + 1
            grdDetails.AddItem RSDAYTRAN!ID & Chr(9) & Format(Null2String(RSDAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(RSDAYTRAN!STOCK_ORD) & Chr(9) & _
                               SetSTOCKDESC(Null2String(RSDAYTRAN!STOCK_SUP)) & Chr(9) & _
                               N2Str2IntZero(RSDAYTRAN!TRANQTY) & Chr(9) & _
                               N2Str2Zero(RSDAYTRAN!TRANINVAMT) & Chr(9) & _
                               N2Str2Zero(RSDAYTRAN!TRANUCOST) & Chr(9) & _
                               Format(N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUCOST), MAXIMUM_DIGIT)
            RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(RSDAYTRAN!TRANQTY)
            RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANUCOST))
            RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(RSDAYTRAN!TRANQTY) * N2Str2Zero(RSDAYTRAN!TRANINVAMT))
            RSDAYTRAN.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        If Null2String(RSREC_HIST!classcode) = "PCS" Or Null2String(RSREC_HIST!classcode) = "PCG" Then
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





Sub InitCboClasscode()
    cboClasscode.Clear
    cboClasscode.AddItem "INTER BRANCH TRANSFER"
    cboClasscode.AddItem "PURCHASED CHARGE"
    cboClasscode.AddItem "PURCHASED CASH"
    cboClasscode.AddItem "RETURN FROM CHARGE"
    cboClasscode.AddItem "RETURN FROM CASH"
    cboClasscode.AddItem "REPLACEMENT"
    cboClasscode.AddItem "RETURNED FROM SERVICE"
    cboClasscode.Text = "PURCHASED CHARGE"
End Sub

Sub FillGrid()
    Dim RSREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    lstREC_Hist.Enabled = False
    Set RSREC_HIST = New ADODB.Recordset
    Set RSREC_HIST = gconDMIS.Execute("select rrno,ID from PMIS_Rec_Hist where type = 'M' order by ID desc")
    If Not (RSREC_HIST.EOF And RSREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, RSREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Enabled = False
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set RSREC_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSREC_HIST = gconDMIS.Execute("select rrno, ID from PMIS_Rec_Hist where type = 'M' and rrno like'" & XXX & "%' order by ID desc")
    If Not (RSREC_HIST.EOF And RSREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, RSREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim RSREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Enabled = False
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set RSREC_HIST = New ADODB.Recordset
    Set RSREC_HIST = gconDMIS.Execute("select recvd_from, ID from PMIS_Rec_Hist where type = 'M' order by ID desc")
    If Not (RSREC_HIST.EOF And RSREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, RSREC_HIST
        lstREC_Hist.Refresh
        lstREC_Hist.Enabled = True
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(XXX As String)
    Dim RSREC_HIST                                     As ADODB.Recordset
    lstREC_Hist.Sorted = False: lstREC_Hist.ListItems.Clear
    Set RSREC_HIST = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set RSREC_HIST = gconDMIS.Execute("select recvd_from, ID from PMIS_Rec_Hist where type = 'M' and recvd_from like '" & XXX & "%' order by ID desc")
    If Not (RSREC_HIST.EOF And RSREC_HIST.BOF) Then
        lstREC_Hist.Enabled = True
        Listview_Loadval Me.lstREC_Hist.ListItems, RSREC_HIST
        lstREC_Hist.Refresh
    Else
        lstREC_Hist.Enabled = False
    End If
End Sub



Private Sub cmdPrint_Click()
    On Error GoTo Errorcode:
    If MsgQuestionBox("Receiving Report Transaction will be Printed. Are you Sure?", "Confirm Printing...") = True Then
        Screen.MousePointer = 11
        rptReceiving.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptReceiving.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptReceiving, PMIS_REPORT_PATH & "rr_hist.rpt", "{RR_HD.type} = 'M' and {RR_HD.rrno} = '" & txtRRNo.Text & "'", DMIS_REPORT_Connection, 1
        NEW_LogAudit "V", "TRANSACTION HISTORY RECEIVING STORING", "", "", "Materials", txtRRNo, "Receiving", ""
        Screen.MousePointer = 0
    End If
    Exit Sub
Errorcode:
    ShowVBError

End Sub


Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    StoreMemVars
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    RSREC_HIST.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    RSREC_HIST.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    RSREC_HIST.MoveNext
    If RSREC_HIST.EOF Then
        RSREC_HIST.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    RSREC_HIST.MovePrevious
    If RSREC_HIST.BOF Then
        RSREC_HIST.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Picture1.Visible = False Then Exit Sub
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MATERIALS RECEIVING)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "MATERIALS RECEIVING")
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If Picture1.Visible = True Then
                StoreMemVars
            End If
        Case vbKeyF3
            If Picture1.Visible = True Then
                If Null2String(RSREC_HIST!Status) = "P" Then
                    MsgSpeechBox "Item(s) are Already Posted and cannot be edited"
                ElseIf Null2String(RSREC_HIST!Status) = "C" Then
                    MsgSpeechBox "Transactions are Already Cancelled and cannot be Change"
                Else

                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    textSearch.Text = ""
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False





    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPMISTrans_Receiving2 = Nothing
    UnloadForm Me
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
        Set RSCUNTER = New ADODB.Recordset
        RSCUNTER.Open "select * from PMIS_Counter where modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not RSCUNTER.EOF And Not RSCUNTER.BOF Then
            txtPONo.Text = Format(N2Str2Zero(RSCUNTER!nextnumber) - 1, "000000")
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
            Set RSPO_HIST = New ADODB.Recordset
            RSPO_HIST.Open "select pono,supcode,podate from PMIS_PO_Hist where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not RSPO_HIST.EOF And Not RSPO_HIST.BOF Then
                txtRecvd_Code.Text = Null2String(RSPO_HIST!SupCode)
                txtPODate.Text = Null2String(RSPO_HIST!PODATE)
                Pcnt = 0
                RR_TOTUCOST = 0
                RR_TOTINVAMT = 0
                RR_TOTVAT = 0
                RR_QTY_REC = 0
                Dim rsDAYTRANDup                       As ADODB.Recordset
                Set rsDAYTRANDup = New ADODB.Recordset
                rsDAYTRANDup.Open "select id,itemno,STOCK_ORD,STOCK_SUP,tranqty,traninvamt,tranucost from PMIS_DayTran where trantype = 'PO' and tranno = " & N2Str2Null(RSPO_HIST!PONO) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
                    Screen.MousePointer = 11
                    rsDAYTRANDup.MoveFirst
                    cleargrid grdDetails
                    Do While Not rsDAYTRANDup.EOF
                        Pcnt = Pcnt + 1
                        grdDetails.AddItem rsDAYTRANDup!ID & Chr(9) & Format(Null2String(rsDAYTRANDup!itemno), "0000") & Chr(9) & _
                                           Null2String(rsDAYTRANDup!STOCK_ORD) & Chr(9) & _
                                           SetSTOCKDESC(Null2String(rsDAYTRANDup!STOCK_SUP)) & Chr(9) & _
                                           N2Str2IntZero(rsDAYTRANDup!TRANQTY) & Chr(9) & _
                                           N2Str2Zero(rsDAYTRANDup!TRANINVAMT) & Chr(9) & _
                                           N2Str2Zero(rsDAYTRANDup!TRANUCOST) & Chr(9) & _
                                           N2Str2IntZero(rsDAYTRANDup!TRANQTY) * N2Str2Zero(rsDAYTRANDup!TRANUCOST)
                        RR_TOTUCOST = RR_TOTUCOST + (N2Str2IntZero(rsDAYTRANDup!TRANQTY) * N2Str2Zero(rsDAYTRANDup!TRANUCOST))
                        RR_TOTINVAMT = RR_TOTINVAMT + (N2Str2IntZero(rsDAYTRANDup!TRANQTY) * N2Str2Zero(rsDAYTRANDup!TRANINVAMT))
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
    If txtremarks.Text = "Pls Type Your Message Here!" Then txtremarks.Text = ""
End Sub

Private Sub txtRIV_Tranno_LostFocus()
    txtRIV_Tranno.Text = Format(txtRIV_Tranno, "000000")
End Sub










Private Sub lstREC_HIST_GotFocus()
    RSREC_HIST.Bookmark = rsFind(RSREC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstREC_HIST_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    If optRRNo.Value = True Then
        RSREC_HIST.Bookmark = rsFind(RSREC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
    Else
        RSREC_HIST.Bookmark = rsFind(RSREC_HIST.Clone, "ID", lstREC_Hist.SelectedItem.SubItems(1)).Bookmark
    End If
    StoreMemVars
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

