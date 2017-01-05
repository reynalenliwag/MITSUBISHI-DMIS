VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEstimate 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estimate Data Entry"
   ClientHeight    =   6720
   ClientLeft      =   705
   ClientTop       =   1020
   ClientWidth     =   10350
   ForeColor       =   &H8000000D&
   Icon            =   "Estimate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   10350
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   60
      TabIndex        =   77
      Top             =   -30
      Width           =   10245
      Begin Crystal.CrystalReport rptEstimate 
         Left            =   90
         Top             =   2010
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Print Out of Repair Order Estimate"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.ComboBox cboModel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   7920
         TabIndex        =   17
         Text            =   "cboModel"
         Top             =   2340
         Width           =   2175
      End
      Begin VB.ComboBox cboRecd_by 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   4830
         TabIndex        =   10
         Text            =   "cboRecd_by"
         Top             =   1620
         Width           =   2295
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1170
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   540
         Width           =   5955
      End
      Begin VB.TextBox txtNiym 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   4800
         MaxLength       =   60
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   5295
      End
      Begin VB.TextBox txtEstimateno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox txtPlate_No 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox txtAcct_No 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtROType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1980
         Width           =   1815
      End
      Begin VB.TextBox txtSvc_No 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2340
         Width           =   1815
      End
      Begin VB.TextBox txtTerm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1380
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2700
         Width           =   1815
      End
      Begin VB.TextBox txtKm_rdg 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   4830
         MaxLength       =   9
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   900
         Width           =   2295
      End
      Begin VB.TextBox txtSektion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   4830
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1260
         Width           =   2295
      End
      Begin VB.TextBox txtParticipat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   8730
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   900
         Width           =   1365
      End
      Begin VB.TextBox txtCertific8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   7260
         MaxLength       =   9
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1980
         Width           =   2835
      End
      Begin VB.TextBox txtMake 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   7920
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2700
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txtDte_Rel 
         Height          =   315
         Left            =   4830
         TabIndex        =   13
         Top             =   2700
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDte_recd 
         Height          =   315
         Left            =   4830
         TabIndex        =   11
         Top             =   1980
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDte_comp 
         Height          =   315
         Left            =   4830
         TabIndex        =   12
         Top             =   2340
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd-mmm-yy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPart_amt 
         Height          =   315
         Left            =   8730
         TabIndex        =   15
         Top             =   1260
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label labPrinted 
         BackColor       =   &H8000000D&
         Caption         =   "PRINTED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8790
         TabIndex        =   135
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3330
         TabIndex        =   97
         Top             =   210
         Width           =   1635
      End
      Begin VB.Label labID 
         BackColor       =   &H00000000&
         Caption         =   "Label18"
         Height          =   165
         Left            =   3390
         TabIndex        =   96
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   420
         TabIndex        =   95
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Estimate #"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   94
         Top             =   210
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Plate No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   510
         TabIndex        =   93
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   450
         TabIndex        =   92
         Top             =   1650
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "ROType"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   91
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   630
         TabIndex        =   90
         Top             =   2370
         Width           =   1035
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "Pay Term"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   450
         TabIndex        =   89
         Top             =   2730
         Width           =   1035
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "KM Reading"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   88
         Top             =   930
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         Caption         =   "Section No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3750
         TabIndex        =   87
         Top             =   1290
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "Sales Advisor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3540
         TabIndex        =   86
         Top             =   1650
         Width           =   1515
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         Caption         =   "Date Recorded"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3450
         TabIndex        =   85
         Top             =   2010
         Width           =   1605
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         Caption         =   "Date Completed"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3330
         TabIndex        =   84
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   "Date Released"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3480
         TabIndex        =   83
         Top             =   2730
         Width           =   1395
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         Caption         =   "Participation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7620
         TabIndex        =   82
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7980
         TabIndex        =   81
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Caption         =   "Warranty Certificate Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7410
         TabIndex        =   80
         Top             =   1710
         Width           =   2775
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7320
         TabIndex        =   79
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000D&
         Caption         =   "Make"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7320
         TabIndex        =   78
         Top             =   2730
         Width           =   1035
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   810
      ScaleHeight     =   945
      ScaleWidth      =   7155
      TabIndex        =   19
      Top             =   5730
      Width           =   7185
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
         Left            =   6300
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   5520
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   4740
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   3960
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Last Record"
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
         Left            =   3180
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Caption         =   "F&irst Record"
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
         Left            =   2400
         MaskColor       =   &H0000FFFF&
         OLEDropMode     =   1  'Manual
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   1620
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Left            =   840
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Left            =   60
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   6300
      ScaleHeight     =   945
      ScaleWidth      =   1665
      TabIndex        =   21
      Top             =   5730
      Width           =   1695
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
         Left            =   810
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   30
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":1C9A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   30
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Estimate Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2625
      Left            =   60
      TabIndex        =   67
      Top             =   3060
      Width           =   10245
      Begin TabDlg.SSTab SSTab1 
         Height          =   2385
         Left            =   60
         TabIndex        =   68
         Top             =   210
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   4207
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Details"
         TabPicture(0)   =   "Estimate.frx":20DC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraDetails"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Jobs"
         TabPicture(1)   =   "Estimate.frx":20F8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fraJobs"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Parts"
         TabPicture(2)   =   "Estimate.frx":2114
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraParts"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Materials"
         TabPicture(3)   =   "Estimate.frx":2130
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fraMaterials"
         Tab(3).ControlCount=   1
         Begin VB.Frame fraDetails 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1905
            Left            =   90
            TabIndex        =   75
            Top             =   390
            Width           =   9945
            Begin MSFlexGridLib.MSFlexGrid grdDetails 
               Height          =   1815
               Left            =   30
               TabIndex        =   76
               Top             =   60
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   3201
               _Version        =   393216
               Rows            =   5
               Cols            =   8
               ForeColor       =   0
               BackColorFixed  =   -2147483635
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fraJobs 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1905
            Left            =   -74910
            TabIndex        =   73
            Top             =   390
            Width           =   9945
            Begin MSFlexGridLib.MSFlexGrid grdJobs 
               Height          =   1815
               Left            =   60
               TabIndex        =   74
               Top             =   60
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   3201
               _Version        =   393216
               Cols            =   7
               ForeColor       =   0
               BackColorFixed  =   -2147483635
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fraParts 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1905
            Left            =   -74910
            TabIndex        =   71
            Top             =   390
            Width           =   9945
            Begin MSFlexGridLib.MSFlexGrid grdParts 
               Height          =   1815
               Left            =   60
               TabIndex        =   72
               Top             =   60
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   3201
               _Version        =   393216
               Cols            =   10
               ForeColor       =   0
               BackColorFixed  =   -2147483635
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Frame fraMaterials 
            BackColor       =   &H8000000D&
            BorderStyle     =   0  'None
            Height          =   1905
            Left            =   -74910
            TabIndex        =   69
            Top             =   390
            Width           =   9945
            Begin MSFlexGridLib.MSFlexGrid grdMaterials 
               Height          =   1815
               Left            =   60
               TabIndex        =   70
               Top             =   60
               Width           =   9855
               _ExtentX        =   17383
               _ExtentY        =   3201
               _Version        =   393216
               Cols            =   10
               ForeColor       =   0
               BackColorFixed  =   -2147483635
               ForeColorFixed  =   16777215
               BackColorBkg    =   -2147483635
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
      End
   End
   Begin VB.Frame fraAddJobs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Add/Edit Jobs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4545
      Left            =   2070
      TabIndex        =   113
      Top             =   780
      Width           =   5535
      Begin VB.ComboBox cboJobChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1320
         Width           =   585
      End
      Begin VB.TextBox txtJobLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1110
         TabIndex        =   66
         Text            =   "Text1"
         Top             =   240
         Width           =   585
      End
      Begin VB.TextBox txtJobPostCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1110
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   960
         Width           =   585
      End
      Begin VB.CommandButton cmdJobSave 
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
         Left            =   1620
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":214C
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdJobCancel 
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
         Left            =   2730
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":258E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3600
         Width           =   1005
      End
      Begin VB.ComboBox cboJobCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1110
         TabIndex        =   33
         Text            =   "Combo1"
         Top             =   600
         Width           =   4305
      End
      Begin VB.CommandButton cmdJobDelete 
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
         Left            =   4410
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":28A0
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3600
         Width           =   1005
      End
      Begin VB.ComboBox cboJcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   3990
         TabIndex        =   32
         Text            =   "Combo1"
         Top             =   240
         Width           =   1425
      End
      Begin MSMask.MaskEdBox txtJobRate 
         Height          =   315
         Left            =   1110
         TabIndex        =   36
         Top             =   1680
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin RichTextLib.RichTextBox txtJobDetail 
         Height          =   795
         Left            =   90
         TabIndex        =   38
         Top             =   2700
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   1402
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Estimate.frx":2BAA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtJobDiscount 
         Height          =   315
         Left            =   1110
         TabIndex        =   37
         Top             =   2040
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin VB.Label labJOBDetVol 
         Alignment       =   2  'Center
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
         Left            =   2070
         TabIndex        =   134
         Top             =   3120
         Width           =   1725
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   121
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000D&
         Caption         =   "Job Desc."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   120
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000D&
         Caption         =   "Post Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   119
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000D&
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   150
         TabIndex        =   118
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000D&
         Caption         =   "Job Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   300
         TabIndex        =   117
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000D&
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   300
         TabIndex        =   116
         Top             =   2070
         Width           =   1305
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Enter Job Description"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   840
         TabIndex        =   115
         Top             =   2460
         Width           =   3705
      End
      Begin VB.Label Label49 
         BackColor       =   &H8000000D&
         Caption         =   "Job Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   3120
         TabIndex        =   114
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdAddJobs 
      BackColor       =   &H8000000D&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1950
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   750
      Width           =   5745
   End
   Begin VB.Frame fraAddParts 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Add/Edit Parts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4845
      Left            =   2550
      TabIndex        =   102
      Top             =   630
      Width           =   4575
      Begin VB.ComboBox cboChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   3030
         Width           =   585
      End
      Begin VB.CommandButton cmdPartsCancel 
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
         Picture         =   "Estimate.frx":2C31
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3870
         Width           =   1005
      End
      Begin VB.CommandButton cmdPartsSave 
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
         Picture         =   "Estimate.frx":2F43
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3870
         Width           =   975
      End
      Begin VB.TextBox txtPartCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1590
         Width           =   2295
      End
      Begin VB.TextBox txtPartsLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1230
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdPartsDelete 
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
         Picture         =   "Estimate.frx":3385
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3870
         Width           =   1005
      End
      Begin VB.ComboBox cboDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Text            =   "Combo1"
         Top             =   1230
         Width           =   4335
      End
      Begin VB.ComboBox cboPartNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   600
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtQty 
         Height          =   315
         Left            =   1230
         TabIndex        =   46
         Top             =   1950
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtUnitPrice 
         Height          =   315
         Left            =   1230
         TabIndex        =   47
         Top             =   2310
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPartAmount 
         Height          =   315
         Left            =   1230
         TabIndex        =   48
         Top             =   2670
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPartDiscount 
         Height          =   315
         Left            =   1230
         TabIndex        =   50
         Top             =   3390
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin VB.Label labPID 
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
         Left            =   3780
         TabIndex        =   136
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000D&
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   330
         TabIndex        =   111
         Top             =   2340
         Width           =   1305
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   420
         TabIndex        =   110
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000D&
         Caption         =   "Part Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   300
         TabIndex        =   109
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   420
         TabIndex        =   108
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label Label35 
         BackColor       =   &H8000000D&
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   450
         TabIndex        =   107
         Top             =   270
         Width           =   1635
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   180
         TabIndex        =   106
         Top             =   990
         Width           =   1305
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   390
         TabIndex        =   105
         Top             =   3420
         Width           =   1305
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   240
         TabIndex        =   104
         Top             =   3030
         Width           =   1305
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   450
         TabIndex        =   103
         Top             =   2670
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdAddParts 
      BackColor       =   &H8000000D&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   2490
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   570
      Width           =   4725
   End
   Begin VB.Frame fraAddMaterials 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Add/Edit Materials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4845
      Left            =   3030
      TabIndex        =   124
      Top             =   660
      Width           =   3765
      Begin VB.ComboBox cboMatChargeTo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   2670
         Width           =   585
      End
      Begin VB.TextBox txtMatLineNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   1230
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton cmdMatSave 
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
         Left            =   450
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":368F
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3870
         Width           =   975
      End
      Begin VB.CommandButton cmdMatCancel 
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
         Left            =   1560
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":3AD1
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3870
         Width           =   1005
      End
      Begin VB.TextBox txtMatPOCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         TabIndex        =   62
         Text            =   "Text1"
         Top             =   3450
         Width           =   2295
      End
      Begin VB.ComboBox cboMatCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1230
         TabIndex        =   55
         Text            =   "Combo1"
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cboMaterial 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Text            =   "Combo1"
         Top             =   1200
         Width           =   3435
      End
      Begin VB.CommandButton cmdMatDelete 
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
         Left            =   2670
         MaskColor       =   &H0000FFFF&
         Picture         =   "Estimate.frx":3DE3
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3870
         Width           =   1005
      End
      Begin MSMask.MaskEdBox txtMatQty 
         Height          =   315
         Left            =   1230
         TabIndex        =   57
         Top             =   1590
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMatUnitPrice 
         Height          =   315
         Left            =   1230
         TabIndex        =   58
         Top             =   1950
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMatAmount 
         Height          =   315
         Left            =   1230
         TabIndex        =   59
         Top             =   2310
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtMatDiscount 
         Height          =   315
         Left            =   1230
         TabIndex        =   61
         Top             =   3090
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin VB.Label Label39 
         BackColor       =   &H8000000D&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   133
         Top             =   2310
         Width           =   1305
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000D&
         Caption         =   "Charge To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   270
         TabIndex        =   132
         Top             =   2700
         Width           =   1305
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000D&
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   420
         TabIndex        =   131
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label42 
         BackColor       =   &H8000000D&
         Caption         =   "Material"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   150
         TabIndex        =   130
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label43 
         BackColor       =   &H8000000D&
         Caption         =   "Line No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   450
         TabIndex        =   129
         Top             =   270
         Width           =   1665
      End
      Begin VB.Label Label46 
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   450
         TabIndex        =   128
         Top             =   1590
         Width           =   1305
      End
      Begin VB.Label Label47 
         BackColor       =   &H8000000D&
         Caption         =   "Unit Price"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   127
         Top             =   1950
         Width           =   1305
      End
      Begin VB.Label Label44 
         BackColor       =   &H8000000D&
         Caption         =   "Mat. Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   126
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000D&
         Caption         =   "PO code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   420
         TabIndex        =   125
         Top             =   3480
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdAddMaterials 
      BackColor       =   &H8000000D&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5025
      Left            =   2940
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   600
      Width           =   3945
   End
   Begin VB.CommandButton cmdDiscount 
      BackColor       =   &H8000000D&
      Enabled         =   0   'False
      Height          =   1395
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   2160
      Width           =   3705
   End
   Begin VB.Frame fraDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "Enter Discount Percentage"
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
      Height          =   1245
      Left            =   3060
      TabIndex        =   137
      Top             =   2190
      Width           =   3495
      Begin VB.CommandButton cmdOkDisc 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Ok"
         Height          =   345
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   810
         Width           =   1005
      End
      Begin VB.CommandButton cmdCancelDisk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ca&ncel"
         Height          =   345
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   810
         Width           =   1005
      End
      Begin MSMask.MaskEdBox txtDiscAmt 
         Height          =   435
         Left            =   840
         TabIndex        =   140
         Top             =   300
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   767
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   14.25
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
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "F6 - Discount"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8130
      TabIndex        =   142
      Top             =   6450
      Width           =   2115
   End
   Begin VB.Label labDetId 
      BackColor       =   &H8000000D&
      Caption         =   "Label48"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   9780
      TabIndex        =   101
      Top             =   5790
      Width           =   375
   End
   Begin VB.Label Label23 
      BackColor       =   &H8000000D&
      Caption         =   "F5 - Materials"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8130
      TabIndex        =   100
      Top             =   6180
      Width           =   2115
   End
   Begin VB.Label Label22 
      BackColor       =   &H8000000D&
      Caption         =   "F4 - Parts"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8130
      TabIndex        =   99
      Top             =   5940
      Width           =   1635
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000D&
      Caption         =   "F3 - Jobs"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8130
      TabIndex        =   98
      Top             =   5700
      Width           =   1635
   End
End
Attribute VB_Name = "frmEstimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsESTI_HD, rsEsti_Det, rsCusmas As Recordset
Dim rsEmpNo, rsS_Model, rsROJOBS As Recordset
Dim rsEsti_Rem As Recordset

Dim rsORD_HD, rsTdaytran, rsPartMas, rsMATMas As Recordset
Dim JobTotal, JobComTotal, JobSalesTotal As Double
Dim JobWarTotal, JobDiscTotal, JobVatTotal As Double

Dim PartsTotal, PartsComTotal, PartsSalesTotal As Double
Dim PartsWarTotal, PartsDiscTotal, PartsVatTotal As Double

Dim MatTotal, MatComTotal, MatSalesTotal As Double
Dim MatWarTotal, MatDiscTotal, MatVatTotal As Double
Dim ROTotal, roTax, roDisc As Double

Dim AddorEdit As String
Dim kcnt, Pcnt, Mcnt As Integer

Dim DiscTotal As Double

Private Sub cboDescription_Click()
labPID.Caption = SetPartID(cboDescription.Text)
cboPartNo.Text = SetPartIDNo(labPID.Caption)
txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
SetPartStatus (cboPartNo.Text)
End Sub

Private Sub cboJobCode_LostFocus()
cboJcode.Text = setJobCode(cboJobCode.Text)
txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
txtJobRate.Text = setJobRate(cboJobCode.Text)
If AddorEdit = "ADD" Then
   txtJobDetail.Text = setJobDetail(cboJobCode.Text)
End If
End Sub

Private Sub cboMaterial_Click()
cboMatCode.Text = SetMatCode(cboMaterial.Text)
txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
txtMatAmount.Text = txtMatUnitPrice.Text
End Sub

Private Sub cboModel_Change()
txtMake.Text = SetMake(cboModel.Text)
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdAdd.SetFocus
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdCancel.SetFocus
End Sub

Private Sub cmdCancelDisk_Click()
SendToBackDisc
End Sub

Private Sub cmdEdit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdEdit.SetFocus
End Sub

Private Sub cmdExit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdExit.SetFocus
End Sub

Private Sub cmdFind_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdFind.SetFocus
End Sub

Private Sub cmdNext_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdNext.SetFocus
End Sub

Private Sub cmdOkDisc_Click()
Screen.MousePointer = 11
Dim varESTIMATENO As String
Dim varDISVAL As Double
varESTIMATENO = N2Str2Null(rsESTI_HD!estimateno)
varDISVAL = Val(txtDiscAmt.Text) / 100

If SSTab1.Tab = 1 Then
   Dim JOBID As Long
   Dim JOBESTIMATENO, JOBLEVEL, JOBLINE_NO As String
   Dim JOBDETCDE, JOBDETDSC, JOBDETUNT As String
   Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
   Dim JOBCODE, JOBWCODE, JOBTAXRATE As Double
   Dim JOBDISCRATE, JOBTAXVAL, JOBDISVAL As Double
   Dim JOBPOCODE, JOBrep_or2, JOBDETAIL As String
   Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2 As Double
   Dim JOBREMARKS As String

   Set rsEsti_Det = New Recordset
       rsEsti_Det.Open "Select * from esti_det where EstimateNo = " & varESTIMATENO & " and livil = '1' order by LINE_NO asc", gconCSMIOS
   If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
      rsEsti_Det.MoveFirst
      Do While Not rsEsti_Det.EOF
         JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
         JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

         JOBID = rsEsti_Det!ID
         JOBESTIMATENO = varESTIMATENO
         JOBLEVEL = "'1'"
         JOBLINE_NO = Format(N2Str2Null(rsEsti_Det!line_no), "00")
         JOBDETCDE = N2Str2Null(rsEsti_Det!detcde)
         JOBDETDSC = N2Str2Null(rsEsti_Det!detdsc)
         JOBDETUNT = N2Str2Null(rsEsti_Det!detunt)
         JOBDETVOL = N2Str2IntZero(rsEsti_Det!detvol)
         JOBDETPRC = N2Str2IntZero(rsEsti_Det!detprc)
         JOBCODE = N2Str2Null(rsEsti_Det!code)
         JOBWCODE = N2Str2Null(rsEsti_Det!wCode)
         JOBTAXRATE = 0.1
         JOBDISCRATE = varDISVAL
         JOBDETAMT = JOBDETPRC / 1.1
         JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
         JOBPOCODE = N2Str2Null(rsEsti_Det!pocode)
         JOBrep_or2 = "NULL"
         JOBDETAIL = N2Str2Null(rsEsti_Det!detail)
         JOBDET_AMT = JOBDETPRC
         JOBDIS_VAL = JOBDISVAL * JOBTAXRATE
         JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
         JOBTAXVAL = (JOBDETAMT - JOBDISCOUNT_2) * JOBTAXRATE
    
         gconCSMIOS.Execute "update esti_det set" & _
                            " EstimateNo = " & JOBESTIMATENO & "," & _
                            " livil = " & JOBLEVEL & "," & _
                            " LINE_NO = " & JOBLINE_NO & "," & _
                            " detcde = " & JOBDETCDE & "," & _
                            " detdsc = " & JOBDETDSC & "," & _
                            " detunt = " & JOBDETUNT & "," & _
                            " detvol = " & JOBDETVOL & "," & _
                            " detprc = " & JOBDETPRC & "," & _
                            " detamt = " & JOBDETAMT & "," & _
                            " code = " & JOBCODE & "," & _
                            " wcode = " & JOBWCODE & "," & _
                            " taxrate = " & (JOBTAXRATE * 100) & "," & _
                            " discrate = " & (JOBDISCRATE * 100) & "," & _
                            " taxval = " & JOBTAXVAL & "," & _
                            " disval = " & JOBDISVAL & "," & _
                            " pocode = " & JOBPOCODE & "," & _
                            " rep_or2 = " & JOBrep_or2 & "," & _
                            " detail = " & JOBDETAIL & "," & _
                            " det_amt = " & JOBDET_AMT & "," & _
                            " dis_val = " & JOBDIS_VAL & "," & _
                            " discount_2 = " & JOBDISCOUNT_2 & _
                            " where id = " & JOBID
         rsEsti_Det.MoveNext
      Loop
      FillJobs
      ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
      gconCSMIOS.Execute "update esti_hd set" & _
                         " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                         " l_amtvalue = " & TOTJOBAMT & "," & _
                         " l_disc = " & TOTJOBDISCVAL & "," & _
                         " l_disc2 = " & TOTJOBDISC * 0.1 & "," & _
                         " l_taxval = " & TOTJOBTAX & "," & _
                         " l_discount = " & TOTJOBDISC & "," & _
                         " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                         " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                         " wl_amt = " & 0 & "," & _
                         " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                         " where id = " & labID.Caption
   End If
End If
If SSTab1.Tab = 2 Then
   Dim PARTSID As Long
   Dim PARTSESTIMATENO, PARTSLEVEL, PARTSLINE_NO As String
   Dim PARTSDETCDE, PARTSDETDSC, PARTSDETUNT As String
   Dim PARTSDETVOL, PARTSDETPRC, PARTSDETAMT As Double
   Dim PARTSCODE, PARTSWCODE As String
   Dim PARTSTAXRATE, PARTSDISCRATE, PARTSTAXVAL As Double
   Dim PARTSDISVAL As Double
   Dim PARTSPOCODE, PARTSrep_or2, PARTSDETAIL As String
   Dim PARTSDET_AMT, PARTSDIS_VAL, PARTSDISCOUNT_2 As Double
   Dim PARTSREMARKS As String

   Set rsEsti_Det = New Recordset
       rsEsti_Det.Open "select * from esti_det where EstimateNo = " & varESTIMATENO & " and livil = '2' order by LINE_NO asc", gconCSMIOS
   If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
      rsEsti_Det.MoveFirst
      Do While Not rsEsti_Det.EOF
         PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
         PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0
  
         PARTSID = rsEsti_Det!ID
         PARTSESTIMATENO = varESTIMATENO
         PARTSLEVEL = "'2'"
         PARTSLINE_NO = Format(N2Str2Null(rsEsti_Det!line_no), "00")
         PARTSDETCDE = N2Str2Null(rsEsti_Det!detcde)
         PARTSDETDSC = N2Str2Null(rsEsti_Det!detdsc)
         PARTSDETUNT = N2Str2Null(rsEsti_Det!detunt)
         PARTSDETVOL = N2Str2Zero(rsEsti_Det!detvol)
         PARTSDETPRC = N2Str2Zero(rsEsti_Det!detprc)
         PARTSDETAMT = N2Str2Zero(rsEsti_Det!detamt)
         PARTSCODE = N2Str2Null(rsEsti_Det!code)
         PARTSWCODE = N2Str2Null(rsEsti_Det!wCode)
         PARTSTAXRATE = 0.1
         PARTSDISCRATE = varDISVAL
         PARTSDISVAL = (PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE)
         PARTSPOCODE = N2Str2Null(rsEsti_Det!pocode)
         PARTSrep_or2 = "NULL"
         PARTSDETAIL = "NULL"
         PARTSDET_AMT = N2Str2Zero(rsEsti_Det!det_amt)
         PARTSDIS_VAL = PARTSDISVAL * PARTSTAXRATE
         PARTSDISCOUNT_2 = PARTSDET_AMT * PARTSDISCRATE
         PARTSTAXVAL = (PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE
  
         gconCSMIOS.Execute "update esti_det set" & _
                            " EstimateNo = " & PARTSESTIMATENO & "," & _
                            " livil = " & PARTSLEVEL & "," & _
                            " LINE_NO = " & PARTSLINE_NO & "," & _
                            " detcde = " & PARTSDETCDE & "," & _
                            " detdsc = " & PARTSDETDSC & "," & _
                            " detunt = " & PARTSDETUNT & "," & _
                            " detvol = " & PARTSDETVOL & "," & _
                            " detprc = " & PARTSDETPRC & "," & _
                            " detamt = " & PARTSDETAMT & "," & _
                            " code = " & PARTSCODE & "," & _
                            " wcode = " & PARTSWCODE & "," & _
                            " taxrate = " & PARTSTAXRATE * 100 & "," & _
                            " discrate = " & PARTSDISCRATE * 100 & "," & _
                            " taxval = " & PARTSTAXVAL & "," & _
                            " disval = " & PARTSDISVAL & "," & _
                            " pocode = " & PARTSPOCODE & "," & _
                            " rep_or2 = " & PARTSrep_or2 & "," & _
                            " detail = " & PARTSDETAIL & "," & _
                            " det_amt = " & PARTSDET_AMT & "," & _
                            " dis_val = " & PARTSDIS_VAL & "," & _
                            " discount_2 = " & PARTSDISCOUNT_2 & _
                            " where id = " & PARTSID
         rsEsti_Det.MoveNext
      Loop
      FillParts
      ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
      gconCSMIOS.Execute "update esti_hd set" & _
                         " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                         " p_amtvalue = " & TOTPARTSAMT & "," & _
                         " p_disc = " & TOTPARTSDISCVAL & "," & _
                         " p_disc2 = " & TOTPARTSDISC * 0.1 & "," & _
                         " p_taxval = " & TOTPARTSTAX & "," & _
                         " p_discount = " & TOTPARTSDISC & "," & _
                         " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                         " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                         " wp_amt = " & 0 & "," & _
                         " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                         " where id = " & labID.Caption

   End If
End If

If SSTab1.Tab = 3 Then
   Dim MATID As Long
   Dim MATESTIMATENO, MATLEVEL, MATLINE_NO As String
   Dim MATDETCDE, MATDETDSC, MATDETUNT As String
   Dim MATDETVOL, MATDETPRC, MATDETAMT As Double
   Dim MATCODE, MATWCODE As String
   Dim MATTAXRATE, MATDISCRATE, MATTAXVAL As Double
   Dim MATDISVAL As Double
   Dim MATPOCODE, MATrep_or2, MATDETAIL As String
   Dim MATDET_AMT, MATDIS_VAL, MATDISCOUNT_2 As Double

   Set rsEsti_Det = New Recordset
       rsEsti_Det.Open "select * from esti_det where EstimateNo = " & varESTIMATENO & " and livil = '3' order by LINE_NO asc", gconCSMIOS
   If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
      rsEsti_Det.MoveFirst
      Do While Not rsEsti_Det.EOF
         MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
         MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0
        
         MATID = rsEsti_Det!ID
         MATESTIMATENO = varESTIMATENO
         MATLEVEL = "'3'"
         MATLINE_NO = Format(N2Str2Null(rsEsti_Det!line_no), "00")
         MATDETCDE = N2Str2Null(rsEsti_Det!detcde)
         MATDETDSC = N2Str2Null(rsEsti_Det!detdsc)
         MATDETUNT = N2Str2Null(rsEsti_Det!detunt)
         MATDETVOL = N2Str2Zero(rsEsti_Det!detvol)
         MATDETPRC = N2Str2Zero(rsEsti_Det!detprc)
         MATDETAMT = N2Str2Zero(rsEsti_Det!detamt)
         MATCODE = N2Str2Null(rsEsti_Det!code)
         MATWCODE = N2Str2Null(rsEsti_Det!wCode)
         MATTAXRATE = 0.1
         MATDISCRATE = varDISVAL
         MATDISVAL = (MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE)
         MATPOCODE = N2Str2Null(rsEsti_Det!pocode)
         MATrep_or2 = "NULL"
         MATDETAIL = "NULL"
         MATDET_AMT = N2Str2Zero(rsEsti_Det!det_amt)
         MATDIS_VAL = MATDISVAL * MATTAXRATE
         MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
         MATTAXVAL = (MATDETAMT - MATDISCOUNT_2) * MATTAXRATE
  
         gconCSMIOS.Execute "update esti_det set" & _
                            " EstimateNo = " & MATESTIMATENO & "," & _
                            " livil = " & MATLEVEL & "," & _
                            " LINE_NO = " & MATLINE_NO & "," & _
                            " detcde = " & MATDETCDE & "," & _
                            " detdsc = " & MATDETDSC & "," & _
                            " detunt = " & MATDETUNT & "," & _
                            " detvol = " & MATDETVOL & "," & _
                            " detprc = " & MATDETPRC & "," & _
                            " detamt = " & MATDETAMT & "," & _
                            " code = " & MATCODE & "," & _
                            " wcode = " & MATWCODE & "," & _
                            " taxrate = " & MATTAXRATE * 100 & "," & _
                            " discrate = " & MATDISCRATE * 100 & "," & _
                            " taxval = " & MATTAXVAL & "," & _
                            " disval = " & MATDISVAL & "," & _
                            " pocode = " & MATPOCODE & "," & _
                            " rep_or2 = " & MATrep_or2 & "," & _
                            " detail = " & MATDETAIL & "," & _
                            " det_amt = " & MATDET_AMT & "," & _
                            " dis_val = " & MATDIS_VAL & "," & _
                            " discount_2 = " & MATDISCOUNT_2 & _
                            " where id = " & MATID
         rsEsti_Det.MoveNext
      Loop
      FillMaterials
      ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
      gconCSMIOS.Execute "update esti_hd set" & _
                         " material = " & TOTMATAMT - TOTMATTAX & "," & _
                         " m_amtvalue = " & TOTMATAMT & "," & _
                         " m_disc = " & TOTMATDISCVAL & "," & _
                         " m_disc2 = " & TOTMATDISC * 0.1 & "," & _
                         " m_taxval = " & TOTMATTAX & "," & _
                         " m_discount = " & TOTMATDISC & "," & _
                         " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                         " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                         " wm_amt = " & 0 & "," & _
                         " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                         " where id = " & labID.Caption
   End If
End If
Screen.MousePointer = 0
SendToBackDisc
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
StoreMemvars
End Sub

Private Sub cmdPrevious_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPrevious.SetFocus
End Sub

Private Sub cmdfirst_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdFirst.SetFocus
End Sub

Private Sub cmdlast_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdLast.SetFocus
End Sub

Private Sub cmdPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdPrint.SetFocus
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
cmdSave.SetFocus
End Sub

Private Sub cmdAdd_Click()
AddorEdit = "ADD"
initMemvars
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
RO_OR_ESTI_OR_PART = "ESTI"
Me.Enabled = False
Me.ZOrder 1
DoEvents
frmCustomer.Show
frmCustomer.ZOrder 0
End Sub

Private Sub cmdCancel_Click()
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
StoreMemvars
End Sub

Private Sub cmdEdit_Click()
AddorEdit = "EDIT"
Frame1.Enabled = True
Picture1.Visible = False
Picture2.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim findStr As String
findStr = InputBox("Please Input Estimate No. or Name ...", "Find", txtEstimateno.Text)
If findStr <> "" Then
   On Error Resume Next
   rsESTI_HD.Bookmark = rsFind(rsESTI_HD.Clone, "estimateno", findStr).Bookmark
   If Err.Number = 3021 Then
      On Error GoTo ErrorCode
      rsESTI_HD.Bookmark = rsFind(rsESTI_HD.Clone, "niym", findStr).Bookmark
   End If
End If
StoreMemvars
Exit Sub

ErrorCode:
If Err.Number = 3021 Then
   MsgBox "Can't find " & findStr, vbOKOnly + vbExclamation, "Not Found"
   Resume Next
End If
End Sub

Private Sub cmdFirst_Click()
rsESTI_HD.MoveFirst
StoreMemvars
End Sub

Private Sub cmdJobCancel_Click()
SendToBack
cmdCancel.Value = True
cleargrid grdJobs
FillJobs
FillDetails
End Sub

Private Sub cmdJobDelete_Click()
If MsgBox("Delete This Job, Are you Sure?", vbQuestion + vbYesNo, "Delete Job Entry") = vbYes Then
   gconCSMIOS.Execute "delete from esti_det where id = " & labDetId.Caption
End If
Dim cnt As Integer
Dim rsesti_detDup As Recordset
Set rsesti_detDup = New Recordset
    rsesti_detDup.Open "select id,LINE_NO from esti_det where estimateno = " & N2Str2Null(rsESTI_HD!estimateno) & " and livil = '1' order by LINE_NO asc", gconCSMIOS
If Not rsesti_detDup.EOF And Not rsesti_detDup.BOF Then
   cnt = 0
   rsesti_detDup.MoveFirst
   Do While Not rsesti_detDup.EOF
      cnt = cnt + 1
      gconCSMIOS.Execute "update esti_det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsesti_detDup!ID
      rsesti_detDup.MoveNext
   Loop
End If
FillJobs
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update esti_hd set" & _
                 " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                 " l_amtvalue = " & TOTJOBAMT & "," & _
                 " l_disc = " & TOTJOBDISCVAL & "," & _
                 " l_disc2 = " & TOTJOBDISC * 0.1 & "," & _
                 " l_taxval = " & TOTJOBTAX & "," & _
                 " l_discount = " & TOTJOBDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdJobCancel.Value = True
End Sub

Private Sub cmdJobSave_Click()
On Error GoTo ErrorCode
Screen.MousePointer = 11

If cboJcode.Text = "" Then
   MsgBox "Cannot find Job Description" & vbCrLf & _
          "Please repeat choosing Job Description", vbCritical, "Error"
   Exit Sub
End If

Dim JOBESTIMATENO, JOBLEVEL, JOBLINE_NO As String
Dim JOBDETCDE, JOBDETDSC, JOBDETUNT As String
Dim JOBDETVOL, JOBDETPRC, JOBDETAMT As Double
Dim JOBCODE, JOBWCODE As String
Dim JOBTAXRATE, JOBDISCRATE, JOBTAXVAL As Double
Dim JOBDISVAL As Double
Dim JOBPOCODE, JOBrep_or2, JOBDETAIL As String
Dim JOBDET_AMT, JOBDIS_VAL, JOBDISCOUNT_2 As Double
Dim JOBREMARKS As String

JOBDISVAL = 0: JOBTAXVAL = 0: JOBDETAMT = 0
JOBDIS_VAL = 0: JOBDISCOUNT_2 = 0: JOBDISCRATE = 0

JOBESTIMATENO = N2Str2Null(txtEstimateno.Text)
JOBLEVEL = "'1'"
JOBLINE_NO = N2Str2Null(Format(txtJobLineNo.Text, "00"))
JOBDETCDE = N2Str2Null(cboJcode.Text)
JOBDETDSC = N2Str2Null(Mid(cboJobCode.Text, 1, 15))
JOBDETUNT = "NULL"
JOBDETVOL = Val(labDetId.Caption)
JOBDETPRC = N2Str2Zero(txtJobRate.Text)
JOBCODE = "NULL"
JOBWCODE = N2Str2Null(cboJobChargeTo.Text)
JOBTAXRATE = 0.1
JOBDISCRATE = Val(txtJobDiscount.Text) / 100
JOBDETAMT = JOBDETPRC / 1.1
JOBDISVAL = (JOBDETPRC * JOBDISCRATE) - ((JOBDETPRC * JOBDISCRATE) * JOBTAXRATE)
JOBPOCODE = N2Str2Null(txtJobPostCode.Text)
JOBrep_or2 = "NULL"
JOBDETAIL = N2Str2Null(CheckChar(txtJobDetail.Text))
JOBDET_AMT = JOBDETPRC
JOBDIS_VAL = JOBDISVAL * JOBTAXRATE
JOBDISCOUNT_2 = JOBDET_AMT * JOBDISCRATE
JOBREMARKS = N2Str2Null(CheckChar(txtJobDetail.Text))
JOBTAXVAL = (JOBDETAMT - JOBDISCOUNT_2) * JOBTAXRATE

If AddorEdit = "ADD" Then
   gconCSMIOS.Execute "insert into Esti_Det " & _
                    "(EstimateNo,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                    " values (" & JOBESTIMATENO & ", " & JOBLEVEL & ", " & JOBLINE_NO & "," & _
                    " " & JOBDETCDE & "," & JOBDETDSC & "," & _
                    " " & JOBDETUNT & ", " & JOBDETVOL & "," & _
                    " " & JOBDETPRC & ", " & JOBDETAMT & ", " & JOBCODE & _
                    ", " & JOBWCODE & ", " & (JOBTAXRATE * 100) & ", " & (JOBDISCRATE * 100) & _
                    ", " & JOBTAXVAL & ", " & JOBDISVAL & ", " & JOBPOCODE & _
                    ", " & JOBrep_or2 & ", " & JOBDETAIL & ", " & JOBDET_AMT & _
                    ", " & JOBDIS_VAL & ", " & JOBDISCOUNT_2 & ")"
Else
   gconCSMIOS.Execute "update Esti_Det set" & _
                    " EstimateNo = " & JOBESTIMATENO & "," & _
                    " livil = " & JOBLEVEL & "," & _
                    " LINE_NO = " & JOBLINE_NO & "," & _
                    " detcde = " & JOBDETCDE & "," & _
                    " detdsc = " & JOBDETDSC & "," & _
                    " detunt = " & JOBDETUNT & "," & _
                    " detvol = " & JOBDETVOL & "," & _
                    " detprc = " & JOBDETPRC & "," & _
                    " detamt = " & JOBDETAMT & "," & _
                    " code = " & JOBCODE & "," & _
                    " wcode = " & JOBWCODE & "," & _
                    " taxrate = " & (JOBTAXRATE * 100) & "," & _
                    " discrate = " & (JOBDISCRATE * 100) & "," & _
                    " taxval = " & JOBTAXVAL & "," & _
                    " disval = " & JOBDISVAL & "," & _
                    " pocode = " & JOBPOCODE & "," & _
                    " rep_or2 = " & JOBrep_or2 & "," & _
                    " detail = " & JOBDETAIL & "," & _
                    " det_amt = " & JOBDET_AMT & "," & _
                    " dis_val = " & JOBDIS_VAL & "," & _
                    " discount_2 = " & JOBDISCOUNT_2 & _
                    " where id = " & labDetId.Caption
End If

FillJobs
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update Esti_Hd set" & _
                 " labor = " & TOTJOBAMT - TOTJOBTAX & "," & _
                 " l_amtvalue = " & TOTJOBAMT & "," & _
                 " l_disc = " & TOTJOBDISCVAL & "," & _
                 " l_disc2 = " & TOTJOBDISC * 0.1 & "," & _
                 " l_taxval = " & TOTJOBTAX & "," & _
                 " l_discount = " & TOTJOBDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " wl_amt = " & 0 & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdJobCancel.Value = True
Screen.MousePointer = 0
If AddorEdit = "ADD" Then
   AddJobs
End If
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub cmdLast_Click()
rsESTI_HD.MoveLast
StoreMemvars
End Sub

Private Sub cmdMatCancel_Click()
SendToBack
cmdCancel.Value = True
End Sub

Private Sub cmdMatDelete_Click()
If MsgBox("Delete This Materials, Are you Sure?", vbQuestion + vbYesNo, "Delete Materials Entry") = vbYes Then
   gconCSMIOS.Execute "delete from esti_det where id = " & labDetId.Caption
End If
Dim cnt As Integer
Dim rsesti_detDup As Recordset
Set rsesti_detDup = New Recordset
    rsesti_detDup.Open "select id,LINE_NO from esti_det where estimateno = " & N2Str2Null(rsESTI_HD!estimateno) & " and livil = '3' order by LINE_NO asc", gconCSMIOS
If Not rsesti_detDup.EOF And Not rsesti_detDup.BOF Then
   cnt = 0
   rsesti_detDup.MoveFirst
   Do While Not rsesti_detDup.EOF
      cnt = cnt + 1
      gconCSMIOS.Execute "update esti_det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsesti_detDup!ID
      rsesti_detDup.MoveNext
   Loop
End If
FillMaterials
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update esti_hd set" & _
                 " material = " & TOTMATAMT - TOTMATTAX & "," & _
                 " m_amtvalue = " & TOTMATAMT & "," & _
                 " m_disc = " & TOTMATDISCVAL & "," & _
                 " m_disc2 = " & TOTMATDISC * 0.1 & "," & _
                 " m_taxval = " & TOTMATTAX & "," & _
                 " m_discount = " & TOTMATDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdMatCancel.Value = True
End Sub

Private Sub cmdMatSave_Click()
Screen.MousePointer = 11
On Error GoTo ErrorCode

Dim MATESTIMATENO, MATLEVEL, MATLINE_NO As String
Dim MATDETCDE, MATDETDSC, MATDETUNT As String
Dim MATDETVOL, MATDETPRC, MATDETAMT As Double
Dim MATCODE, MATWCODE As String
Dim MATTAXRATE, MATDISCRATE, MATTAXVAL As Double
Dim MATDISVAL As Double
Dim MATPOCODE, MATrep_or2, MATDETAIL As String
Dim MATDET_AMT, MATDIS_VAL, MATDISCOUNT_2 As Double

MATDISVAL = 0: MATTAXVAL = 0: MATDETAMT = 0
MATDIS_VAL = 0: MATDISCOUNT_2 = 0: MATDISCRATE = 0

MATESTIMATENO = N2Str2Null(txtEstimateno.Text)
MATLEVEL = "'3'"
MATLINE_NO = N2Str2Null(Format(txtMatLineNo.Text, "00"))
MATDETCDE = N2Str2Null(cboMatCode.Text)
MATDETDSC = N2Str2Null(Mid(cboMaterial.Text, 1, 15))
MATDETUNT = "NULL"
MATDETVOL = N2Str2Zero(txtMatQty.Text)
MATDETPRC = Val(txtMatUnitPrice.Text)
MATDETAMT = Val(txtMatAmount.Text) / 1.1
MATCODE = "NULL"
MATWCODE = N2Str2Null(cboMatChargeTo.Text)
MATTAXRATE = 0.1
MATDISCRATE = Val(txtMatDiscount.Text) / 100
MATDISVAL = (MATDETPRC * MATDISCRATE) - ((MATDETPRC * MATDISCRATE) * MATTAXRATE)
MATPOCODE = N2Str2Null(txtMatPOCode.Text)
MATrep_or2 = "NULL"
MATDETAIL = "NULL"
MATDET_AMT = Val(txtMatAmount.Text)
MATDIS_VAL = MATDISVAL * MATTAXRATE
MATDISCOUNT_2 = MATDET_AMT * MATDISCRATE
MATTAXVAL = (MATDETAMT - MATDISCOUNT_2) * MATTAXRATE

If AddorEdit = "ADD" Then
   gconCSMIOS.Execute "insert into Esti_Det " & _
                    "(EstimateNo,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                    " values (" & MATESTIMATENO & ", " & MATLEVEL & ", " & MATLINE_NO & "," & _
                    " " & MATDETCDE & "," & MATDETDSC & "," & _
                    " " & MATDETUNT & ", " & MATDETVOL & "," & _
                    " " & MATDETPRC & ", " & MATDETAMT & ", " & MATCODE & _
                    ", " & MATWCODE & ", " & MATTAXRATE * 100 & ", " & MATDISCRATE * 100 & _
                    ", " & MATTAXVAL & ", " & MATDISVAL & ", " & MATPOCODE & _
                    ", " & MATrep_or2 & ", " & MATDETAIL & ", " & MATDET_AMT & _
                    ", " & MATDIS_VAL & ", " & MATDISCOUNT_2 & ")"
Else
   gconCSMIOS.Execute "update Esti_Det set" & _
                    " EstimateNo = " & MATESTIMATENO & "," & _
                    " livil = " & MATLEVEL & "," & _
                    " LINE_NO = " & MATLINE_NO & "," & _
                    " detcde = " & MATDETCDE & "," & _
                    " detdsc = " & MATDETDSC & "," & _
                    " detunt = " & MATDETUNT & "," & _
                    " detvol = " & MATDETVOL & "," & _
                    " detprc = " & MATDETPRC & "," & _
                    " detamt = " & MATDETAMT & "," & _
                    " code = " & MATCODE & "," & _
                    " wcode = " & MATWCODE & "," & _
                    " taxrate = " & MATTAXRATE * 100 & "," & _
                    " discrate = " & MATDISCRATE * 100 & "," & _
                    " taxval = " & MATTAXVAL & "," & _
                    " disval = " & MATDISVAL & "," & _
                    " pocode = " & MATPOCODE & "," & _
                    " rep_or2 = " & MATrep_or2 & "," & _
                    " detail = " & MATDETAIL & "," & _
                    " det_amt = " & MATDET_AMT & "," & _
                    " dis_val = " & MATDIS_VAL & "," & _
                    " discount_2 = " & MATDISCOUNT_2 & _
                    " where id = " & labDetId.Caption
End If
FillMaterials
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update Esti_Hd set" & _
                 " material = " & TOTMATAMT - TOTMATTAX & "," & _
                 " m_amtvalue = " & TOTMATAMT & "," & _
                 " m_disc = " & TOTMATDISCVAL & "," & _
                 " m_disc2 = " & TOTMATDISC * 0.1 & "," & _
                 " m_taxval = " & TOTMATTAX & "," & _
                 " m_discount = " & TOTMATDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " wm_amt = " & 0 & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdMatCancel.Value = True
Screen.MousePointer = 0
If AddorEdit = "ADD" Then
   AddMaterials
End If
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub cmdNext_Click()
rsESTI_HD.MoveNext
If rsESTI_HD.EOF Then
   rsESTI_HD.MoveLast
   MsgBox "Last Record"
End If
StoreMemvars
End Sub

Private Sub cmdPartsCancel_Click()
SendToBack
cmdCancel.Value = True
cleargrid grdParts
FillParts
FillDetails
End Sub

Private Sub cmdPartsDelete_Click()
If MsgBox("Delete This Parts, Are you Sure?", vbQuestion + vbYesNo, "Delete Parts Entry") = vbYes Then
   gconCSMIOS.Execute "delete from esti_det where id = " & labDetId.Caption
End If
Dim cnt As Integer
Dim rsesti_detDup As Recordset
Set rsesti_detDup = New Recordset
    rsesti_detDup.Open "select id,LINE_NO from esti_det where estimateno = " & N2Str2Null(rsESTI_HD!estimateno) & " and livil = '2' order by LINE_NO asc", gconCSMIOS
If Not rsesti_detDup.EOF And Not rsesti_detDup.BOF Then
   cnt = 0
   rsesti_detDup.MoveFirst
   Do While Not rsesti_detDup.EOF
      cnt = cnt + 1
      gconCSMIOS.Execute "update esti_det set LINE_NO = " & N2Str2Null(Format(cnt, "00")) & " where id = " & rsesti_detDup!ID
      rsesti_detDup.MoveNext
   Loop
End If
FillParts
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update esti_hd set" & _
                 " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                 " p_amtvalue = " & TOTPARTSAMT & "," & _
                 " p_disc = " & TOTPARTSDISCVAL & "," & _
                 " p_disc2 = " & TOTPARTSDISC * 0.1 & "," & _
                 " p_taxval = " & TOTPARTSTAX & "," & _
                 " p_discount = " & TOTPARTSDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdPartsCancel.Value = True
End Sub

Private Sub cmdPartsSave_Click()
Screen.MousePointer = 11
On Error GoTo ErrorCode

Dim PARTSESTIMATENO, PARTSLEVEL, PARTSLINE_NO As String
Dim PARTSDETCDE, PARTSDETDSC, PARTSDETUNT As String
Dim PARTSDETVOL, PARTSDETPRC, PARTSDETAMT As Double
Dim PARTSCODE, PARTSWCODE As String
Dim PARTSTAXRATE, PARTSDISCRATE, PARTSTAXVAL As Double
Dim PARTSDISVAL As Double
Dim PARTSPOCODE, PARTSrep_or2, PARTSDETAIL As String
Dim PARTSDET_AMT, PARTSDIS_VAL, PARTSDISCOUNT_2 As Double
Dim PARTSREMARKS As String

PARTSDISVAL = 0: PARTSTAXVAL = 0: PARTSDETAMT = 0
PARTSDIS_VAL = 0: PARTSDISCOUNT_2 = 0: PARTSDISCRATE = 0

PARTSESTIMATENO = N2Str2Null(txtEstimateno.Text)
PARTSLEVEL = "'2'"
PARTSLINE_NO = N2Str2Null(Format(txtPartsLineNo.Text, "00"))
PARTSDETCDE = N2Str2Null(cboPartNo.Text)
PARTSDETDSC = N2Str2Null(Mid(cboDescription.Text, 1, 15))
PARTSDETUNT = "NULL"
PARTSDETVOL = N2Str2Zero(txtQty.Text)
PARTSDETPRC = Val(txtUnitPrice.Text)
PARTSDETAMT = Val(txtPartAmount.Text) / 1.1
PARTSCODE = "NULL"
PARTSWCODE = N2Str2Null(cboChargeTo.Text)
PARTSTAXRATE = 0.1
PARTSDISCRATE = Val(txtPartDiscount.Text) / 100
PARTSDISVAL = (PARTSDETPRC * PARTSDISCRATE) - ((PARTSDETPRC * PARTSDISCRATE) * PARTSTAXRATE)
PARTSPOCODE = N2Str2Null(txtPartCode.Text)
PARTSrep_or2 = "NULL"
PARTSDETAIL = "NULL"
PARTSDET_AMT = Val(txtPartAmount.Text)
PARTSDIS_VAL = PARTSDISVAL * PARTSTAXRATE
PARTSDISCOUNT_2 = PARTSDET_AMT * PARTSDISCRATE
PARTSTAXVAL = (PARTSDETAMT - PARTSDISCOUNT_2) * PARTSTAXRATE

If AddorEdit = "ADD" Then
   gconCSMIOS.Execute "insert into Esti_Det " & _
                    "(EstimateNo,livil,LINE_NO,detcde,detdsc,detunt,detvol,detprc,detamt,code,wcode,taxrate,discrate,taxval,disval,pocode,rep_or2,detail,det_amt,dis_val,discount_2)" & _
                    " values (" & PARTSESTIMATENO & ", " & PARTSLEVEL & ", " & PARTSLINE_NO & "," & _
                    " " & PARTSDETCDE & "," & PARTSDETDSC & "," & _
                    " " & PARTSDETUNT & ", " & PARTSDETVOL & "," & _
                    " " & PARTSDETPRC & ", " & PARTSDETAMT & ", " & PARTSCODE & _
                    ", " & PARTSWCODE & ", " & PARTSTAXRATE * 100 & ", " & PARTSDISCRATE * 100 & _
                    ", " & PARTSTAXVAL & ", " & PARTSDISVAL & ", " & PARTSPOCODE & _
                    ", " & PARTSrep_or2 & ", " & PARTSDETAIL & ", " & PARTSDET_AMT & _
                    ", " & PARTSDIS_VAL & ", " & PARTSDISCOUNT_2 & ")"
Else
   gconCSMIOS.Execute "update Esti_Det set" & _
                    " EstimateNo = " & PARTSESTIMATENO & "," & _
                    " livil = " & PARTSLEVEL & "," & _
                    " LINE_NO = " & PARTSLINE_NO & "," & _
                    " detcde = " & PARTSDETCDE & "," & _
                    " detdsc = " & PARTSDETDSC & "," & _
                    " detunt = " & PARTSDETUNT & "," & _
                    " detvol = " & PARTSDETVOL & "," & _
                    " detprc = " & PARTSDETPRC & "," & _
                    " detamt = " & PARTSDETAMT & "," & _
                    " code = " & PARTSCODE & "," & _
                    " wcode = " & PARTSWCODE & "," & _
                    " taxrate = " & PARTSTAXRATE * 100 & "," & _
                    " discrate = " & PARTSDISCRATE * 100 & "," & _
                    " taxval = " & PARTSTAXVAL & "," & _
                    " disval = " & PARTSDISVAL & "," & _
                    " pocode = " & PARTSPOCODE & "," & _
                    " rep_or2 = " & PARTSrep_or2 & "," & _
                    " detail = " & PARTSDETAIL & "," & _
                    " det_amt = " & PARTSDET_AMT & "," & _
                    " dis_val = " & PARTSDIS_VAL & "," & _
                    " discount_2 = " & PARTSDISCOUNT_2 & _
                    " where id = " & labDetId.Caption
End If
FillParts
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
gconCSMIOS.Execute "update Esti_Hd set" & _
                 " parts = " & TOTPARTSAMT - TOTPARTSTAX & "," & _
                 " p_amtvalue = " & TOTPARTSAMT & "," & _
                 " p_disc = " & TOTPARTSDISCVAL & "," & _
                 " p_disc2 = " & TOTPARTSDISC * 0.1 & "," & _
                 " p_taxval = " & TOTPARTSTAX & "," & _
                 " p_discount = " & TOTPARTSDISC & "," & _
                 " amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & "," & _
                 " rovat = " & TOTJOBTAX + TOTPARTSTAX + TOTMATTAX & "," & _
                 " wp_amt = " & 0 & "," & _
                 " ro_amount = " & ROTotal - TOTJOBDISC - TOTPARTSDISC - TOTMATDISC & _
                 " where id = " & labID.Caption
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdPartsCancel.Value = True
Screen.MousePointer = 0
If AddorEdit = "ADD" Then
   AddParts
End If
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub cmdPrevious_Click()
rsESTI_HD.MovePrevious
If rsESTI_HD.BOF Then
   rsESTI_HD.MoveFirst
   MsgBox "First Record"
End If
StoreMemvars
End Sub

Private Sub cmdPrint_Click()
Dim Filter As String
If Null2String(rsESTI_HD!prin_dte) = "" Then
   If MsgBox("Print this Estimate?", vbQuestion + vbYesNo, "Print Estimate") = vbYes Then
      gconCSMIOS.Execute "update esti_hd set prin_dte = '" & Date & "' where id = " & labID.Caption
      StoreMemvars
      If DiscTotal > 0 Then
         PRINTESTIDISC
      Else
         PRINTESTI
      End If
   End If
Else
   If MsgBox("Estimate Already Printed! Print Another?", vbQuestion + vbYesNo, "Print Estimate") = vbYes Then
      If DiscTotal > 0 Then
         PRINTESTIDISC
      Else
         PRINTESTI
      End If
   End If
End If
End Sub

Sub PRINTESTIDISC()
Screen.MousePointer = 11
PrintSQLReport rptEstimate, CSMIOS_REPORT_PATH & "estimatedisc.rpt", "{esti_hd.estimateno} = '" & txtEstimateno.Text & "'", CSMIOS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Sub PRINTESTI()
Screen.MousePointer = 11
PrintSQLReport rptEstimate, CSMIOS_REPORT_PATH & "estimate.rpt", "{esti_hd.estimateno} = '" & txtEstimateno.Text & "'", CSMIOS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode

If txtNiym.Text = "" Then
   MsgBox "Customer must have a name", vbInformation, "Invalid Name"
   Exit Sub
End If

If cboRecd_by.Text = "" Then
   MsgBox "Record By must not be Empty!"
   Exit Sub
End If

If AddorEdit = "ADD" Then
   Dim rsESTI_HDDup As Recordset
   Set rsESTI_HDDup = New Recordset
       rsESTI_HDDup.Open "select id from ESTI_HD order by id asc", gconCSMIOS
   If Not rsESTI_HDDup.EOF And Not rsESTI_HDDup.BOF Then
      rsESTI_HDDup.MoveLast
      labID.Caption = Val(rsESTI_HDDup!ID) + 1
   End If
   Dim rsDupESTI_HD As Recordset
   Set rsDupESTI_HD = New Recordset
       rsDupESTI_HD.Open "select estimateno from ESTI_HD where estimateno = " & N2Str2Null(txtEstimateno.Text), gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsDupESTI_HD.EOF And Not rsDupESTI_HD.BOF Then
      MsgBox "Repair Order No. Already Exist!", "Invalid RO Number"
      Exit Sub
   End If
End If
         
Dim VTXTestimateno, VTXTROType, VTXTSvc_No As String
Dim VTXTAcct_No, VTXTNiym, VTXTPlate_No As String
Dim VcboModel, VTXTMake, VTXTTerm As String
Dim VTXTSektion, VTXTKm_rdg, VTXTDte_recd As String
Dim VTXTCertific8, VTXTDte_comp, VTXTDte_Rel As String
Dim VTXTPart_amt As Double
Dim VTXTParticipat, VcboRecd_by As String

VTXTestimateno = N2Str2Null(txtEstimateno.Text)
VTXTROType = N2Str2Null(txtROType.Text)
VTXTSvc_No = N2Str2Null(txtSvc_No.Text)
VTXTAcct_No = N2Str2Null(txtAcct_No.Text)
VTXTNiym = N2Str2Null(txtNiym.Text)
VTXTPlate_No = N2Str2Null(txtPlate_No.Text)
VcboModel = N2Str2Null(cboModel.Text)
VTXTMake = N2Str2Null(txtMake.Text)
VTXTTerm = N2Str2Null(txtTerm.Text)
VTXTSektion = N2Str2Null(txtSektion.Text)
VTXTKm_rdg = N2Str2Null(txtKm_rdg.Text)
VTXTDte_recd = N2Date2Null(txtDte_recd.Text)
VTXTCertific8 = N2Str2Null(txtCertific8.Text)
VTXTDte_comp = N2Date2Null(txtDte_comp.Text)
VTXTDte_Rel = N2Date2Null(txtDte_Rel.Text)
VTXTPart_amt = N2Str2Zero(txtPart_amt.Text)
VTXTParticipat = N2Str2Null(txtParticipat.Text)
VcboRecd_by = N2Str2Null(SetCodeSA(cboRecd_by.Text))

If AddorEdit = "ADD" Then
   gconCSMIOS.Execute "insert into ESTI_HD " & _
                    "(estimateno,rotype,svc_no,acct_no,niym,plate_no,model,term,sektion,Recd_by,km_rdg,dte_recd,certific8,dte_comp,dte_rel,part_amt,participat)" & _
                    " values (" & VTXTestimateno & ", " & VTXTROType & ", " & VTXTSvc_No & _
                    ", " & VTXTAcct_No & ", " & VTXTNiym & ", " & VTXTPlate_No & ", " & VcboModel & ", " & VTXTTerm & ", " & VTXTSektion & _
                    ", " & VcboRecd_by & ", " & VTXTKm_rdg & ", " & VTXTDte_recd & ", " & VTXTCertific8 & _
                    ", " & VTXTDte_comp & ", " & VTXTDte_Rel & ", " & VTXTPart_amt & ", " & VTXTParticipat & ")"
Else
   gconCSMIOS.Execute "update ESTI_HD set" & _
                    " estimateno = " & VTXTestimateno & "," & _
                    " rotype = " & VTXTROType & "," & _
                    " svc_no = " & VTXTSvc_No & "," & _
                    " acct_no = " & VTXTAcct_No & "," & _
                    " niym = " & VTXTNiym & "," & _
                    " plate_no = " & VTXTPlate_No & "," & _
                    " model = " & VcboModel & "," & _
                    " term = " & VTXTTerm & "," & _
                    " sektion = " & VTXTSektion & "," & _
                    " recd_by = " & VcboRecd_by & "," & _
                    " km_rdg = " & VTXTKm_rdg & "," & _
                    " dte_recd = " & VTXTDte_recd & "," & _
                    " certific8 = " & VTXTCertific8 & "," & _
                    " dte_comp = " & VTXTDte_comp & "," & _
                    " dte_rel = " & VTXTDte_Rel & "," & _
                    " part_amt = " & VTXTPart_amt & "," & _
                    " participat = " & VTXTParticipat & _
                    " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsESTI_HD.Find "id = " & labID.Caption
cmdCancel.Value = True
Exit Sub

ErrorCode:
MsgBox "Error:" & Err & " " & Error, vbOKOnly, "Error"
cmdCancel.Value = True
Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "" Then
   KeyAscii = 0
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
       Case vbKeyReturn
            If Me.ActiveControl.Name = "txtCertific8" Then
               Me.Enabled = False
               frmCusveh.Show
               frmCusveh.ZOrder 0
            Else
               If Mid(Me.ActiveControl.Name, 1, 3) = "txt" Or Mid(Me.ActiveControl.Name, 1, 3) = "opt" Or Mid(Me.ActiveControl.Name, 1, 3) = "cbo" Then
                  SendKeys "{TAB}"
               End If
            End If
       Case vbKeyEscape
            SSTab1.Tab = 0
            SendToBack
       Case vbKeyF3
            If Picture1.Visible = True Then
               AddJobs
            End If
       Case vbKeyF4
            If Picture1.Visible = True Then
               AddParts
            End If
       Case vbKeyF5
            If Picture1.Visible = True Then
               AddMaterials
            End If
       Case vbKeyF6
            If Picture1.Visible = True Then
               SendToFrontDisc
            End If
End Select
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
CenterMe frmMain, Me, 1
Set rsESTI_HD = New Recordset
    rsESTI_HD.Open "select * from ESTI_HD order by estimateno desc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
fraAddJobs.Enabled = False
fraAddParts.Enabled = False
fraAddMaterials.Enabled = False
RO_OR_ESTI_OR_PART = "ESTI"
ESTISHOW = True
initMemvars
StoreMemvars
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsESTI_HD = New Recordset
    rsESTI_HD.Open "select * from ESTI_HD order by estimateno desc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
txtEstimateno.Text = ""
txtAddress.Text = ""
txtROType.Text = "0"
txtSvc_No.Text = ""
txtAcct_No.Text = ""
txtNiym.Text = ""
txtPlate_No.Text = ""
txtMake.Text = ""
txtTerm.Text = "CSH"
txtSektion.Text = ""
txtKm_rdg.Text = ""
txtDte_recd.Text = Date
txtCertific8.Text = ""
txtDte_comp.Text = ""
txtDte_comp.Text = ""
txtPart_amt.Text = ""
txtParticipat.Text = ""
InitCbo
cleargrid grdDetails
cleargrid grdJobs
cleargrid grdParts
cleargrid grdMaterials
initGrid
InitJobs
InitParts
InitMaterials
End Sub

Sub InitMaterials()
txtMatLineNo.Text = Format(Mcnt + 1, "00")
cboMatCode.Text = ""
cboMaterial.Text = ""
txtMatQty.Text = 1
txtMatUnitPrice.Text = 0#
txtMatAmount.Text = 0#
cboMatChargeTo.Clear
cboMatChargeTo.AddItem ""
cboMatChargeTo.AddItem "W"
cboMatChargeTo.AddItem "S"
cboMatChargeTo.AddItem "C"
txtMatDiscount.Text = 0#
txtMatPOCode.Text = "01"
End Sub

Sub InitParts()
txtPartsLineNo.Text = Format(Pcnt + 1, "00")
cboPartNo.Text = ""
cboDescription.Text = ""
txtPartCode.Text = "01"
txtQty.Text = 1
txtUnitPrice.Text = 0#
txtPartAmount.Text = 0#
cboChargeTo.Clear
cboChargeTo.AddItem ""
cboChargeTo.AddItem "W"
cboChargeTo.AddItem "S"
cboChargeTo.AddItem "C"
txtPartDiscount.Text = 0#
End Sub

Sub InitJobs()
cboJcode.Text = ""
txtJobLineNo.Text = Format(kcnt + 1, "00")
cboJobCode.Text = ""
txtJobPostCode.Text = ""
cboJobChargeTo.Clear
cboJobChargeTo.AddItem ""
cboJobChargeTo.AddItem "W"
cboJobChargeTo.AddItem "S"
cboJobChargeTo.AddItem "C"
txtJobRate.Text = "0"
txtJobDiscount.Text = "0"
txtJobDetail.Text = ""
End Sub

Sub InitCbo()
Set rsEmpNo = New Recordset
    rsEmpNo.Open "select naym from empno", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
   rsEmpNo.MoveFirst
   cboRecd_by.Clear
   cboRecd_by.Text = Null2String(rsEmpNo!naym)
   Do While Not rsEmpNo.EOF
      cboRecd_by.AddItem Null2String(rsEmpNo!naym)
      rsEmpNo.MoveNext
   Loop
End If
Set rsS_Model = New Recordset
    rsS_Model.Open "Select model from s_model", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsS_Model.EOF And Not rsS_Model.BOF Then
   rsS_Model.MoveFirst
   cboModel.Clear
   Do While Not rsS_Model.EOF
      cboModel.AddItem Null2String(rsS_Model!model)
      rsS_Model.MoveNext
   Loop
End If
Set rsROJOBS = New Recordset
    rsROJOBS.Open "Select jcode from ROJOBS order by jcode asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
   rsROJOBS.MoveFirst
   cboJcode.Clear
   Do While Not rsROJOBS.EOF
      cboJcode.AddItem Null2String(rsROJOBS!jcode)
      rsROJOBS.MoveNext
   Loop
End If
Set rsROJOBS = New Recordset
    rsROJOBS.Open "Select desc1 from ROJOBS order by desc1 asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
   rsROJOBS.MoveFirst
   cboJobCode.Clear
   Do While Not rsROJOBS.EOF
      cboJobCode.AddItem Null2String(rsROJOBS!desc1)
      rsROJOBS.MoveNext
   Loop
End If
Set rsPartMas = New Recordset
    rsPartMas.Open "select partno from partmas order by partno asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsPartMas.EOF And Not rsPartMas.BOF Then
   rsPartMas.MoveFirst
   cboPartNo.Clear
   Do While Not rsPartMas.EOF
      cboPartNo.AddItem Null2String(rsPartMas!PartNo)
      rsPartMas.MoveNext
   Loop
End If
Set rsPartMas = New Recordset
    rsPartMas.Open "select partdesc from partmas order by partdesc asc", gconPMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsPartMas.EOF And Not rsPartMas.BOF Then
   rsPartMas.MoveFirst
   cboDescription.Clear
   Do While Not rsPartMas.EOF
      cboDescription.AddItem Null2String(rsPartMas!partdesc)
      rsPartMas.MoveNext
   Loop
End If
Set rsMATMas = New Recordset
    rsMATMas.Open "select matcde from matmas order by matcde asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsMATMas.EOF And Not rsMATMas.BOF Then
   rsMATMas.MoveFirst
   cboMatCode.Clear
   Do While Not rsMATMas.EOF
      cboMatCode.AddItem Null2String(rsMATMas!matcde)
      rsMATMas.MoveNext
   Loop
End If
Set rsMATMas = New Recordset
    rsMATMas.Open "select matdsc from matmas order by matdsc asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsMATMas.EOF And Not rsMATMas.BOF Then
   rsMATMas.MoveFirst
   cboMaterial.Clear
   Do While Not rsMATMas.EOF
      cboMaterial.AddItem Null2String(rsMATMas!matdsc)
      rsMATMas.MoveNext
   Loop
End If
End Sub

Function setJobCode(jjj As String)
If jjj <> "" Then
   Set rsROJOBS = New Recordset
       rsROJOBS.Open "Select jcode,desc1 from ROJOBS where desc1 = '" & jjj & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
      setJobCode = Null2String(rsROJOBS!jcode)
   Else
      setJobCode = ""
   End If
End If
End Function

Function setJobDesc(jjj As String)
If jjj <> "" Then
   Set rsROJOBS = New Recordset
       rsROJOBS.Open "Select jcode,desc1 from ROJOBS where jcode = '" & jjj & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
      setJobDesc = Null2String(rsROJOBS!desc1)
   Else
      setJobDesc = ""
   End If
End If
End Function

Function setJobPOcode(ppp As String)
If ppp <> "" Then
  Set rsROJOBS = New Recordset
      rsROJOBS.Open "Select desc1,pocode from ROJOBS where desc1 = '" & ppp & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
  If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
     setJobPOcode = Null2String(rsROJOBS!pocode)
  Else
     setJobPOcode = ""
  End If
End If
End Function

Function setJobDetail(ppp As String)
If ppp <> "" Then
   Set rsROJOBS = New Recordset
       rsROJOBS.Open "Select desc1,detail from ROJOBS where desc1 = '" & ppp & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
      setJobDetail = Null2String(rsROJOBS!detail)
   Else
      setJobDetail = ""
   End If
End If
End Function

Function setJobRate(ppp As String)
If ppp <> "" Then
   Set rsROJOBS = New Recordset
       rsROJOBS.Open "Select desc1,flatrate from ROJOBS where desc1 = '" & ppp & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsROJOBS.EOF And Not rsROJOBS.BOF Then
      setJobRate = Null2String(rsROJOBS!flatrate)
   Else
      setJobRate = 0#
   End If
End If
End Function

Sub StoreMemvars()
If Not rsESTI_HD.EOF And Not rsESTI_HD.BOF Then
   labID.Caption = rsESTI_HD!ID
   txtEstimateno.Text = Null2String(rsESTI_HD!estimateno)
   txtROType.Text = Null2String(rsESTI_HD!rotype)
   txtSvc_No.Text = Null2String(rsESTI_HD!svc_no)
   txtAcct_No.Text = Null2String(rsESTI_HD!acct_no)
   If IsNull(rsESTI_HD!acct_no) = False Then
      SetAdres (rsESTI_HD!acct_no)
   End If
   txtNiym.Text = Null2String(rsESTI_HD!Niym)
   txtPlate_No.Text = Null2String(rsESTI_HD!plate_no)
   cboModel.Text = Null2String(rsESTI_HD!model)
   txtMake.Text = SetMake(Null2String(rsESTI_HD!model))
   txtTerm.Text = Null2String(rsESTI_HD!term)
   txtSektion.Text = Null2String(rsESTI_HD!sektion)
   cboRecd_by.Text = SetSA(Null2String(rsESTI_HD!recd_by))
   txtKm_rdg.Text = Null2String(rsESTI_HD!km_rdg)
   txtCertific8.Text = Null2String(rsESTI_HD!certific8)
   txtDte_recd.Text = Null2String(rsESTI_HD!dte_recd)
   txtDte_comp.Text = Null2String(rsESTI_HD!dte_comp)
   txtDte_Rel.Text = Null2String(rsESTI_HD!dte_rel)
   If Null2String(rsESTI_HD!prin_dte) <> "" Then
      labPrinted.Visible = True
   Else
      labPrinted.Visible = False
   End If
   txtPart_amt.Text = Null2String(rsESTI_HD!part_amt)
   txtParticipat.Text = Null2String(rsESTI_HD!participat)
   cleargrid grdJobs
   cleargrid grdParts
   cleargrid grdMaterials
   FillJobs
   FillParts
   FillMaterials
   FillDetails
Else
   cmdFirst.Enabled = False
   cmdLast.Enabled = False
   cmdPrevious.Enabled = False
   cmdNext.Enabled = False
   cmdEdit.Enabled = False
   cmdPrint.Enabled = False
End If
End Sub

Sub SetAdres(CCC As String)
Set rsCusmas = New Recordset
    rsCusmas.Open "Select cuscde,cusadd from cusmas where cuscde = '" & CCC & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsCusmas.EOF And Not rsCusmas.BOF Then
   txtAddress.Text = Null2String(rsCusmas!Cusadd)
Else
   txtAddress.Text = ""
End If
End Sub

Function SetPhone(CCC As String)
Set rsCusmas = New Recordset
    rsCusmas.Open "Select cuscde,cusphon1 from cusmas where cuscde = '" & CCC & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsCusmas.EOF And Not rsCusmas.BOF Then
   SetPhone = Null2String(rsCusmas!cusphon1)
Else
   SetPhone = ""
End If
End Function

Function SetMake(mmm As String)
Set rsS_Model = New Recordset
    rsS_Model.Open "Select make,model from s_model where model = '" & mmm & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsS_Model.EOF And Not rsS_Model.BOF Then
   SetMake = Null2String(rsS_Model!make)
Else
   SetMake = ""
End If
End Function

Function SetSA(emp As String)
Set rsEmpNo = New Recordset
    rsEmpNo.Open "Select code,naym from empno where code = '" & emp & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
   SetSA = Null2String(rsEmpNo!naym)
End If
End Function

Function SetCodeSA(nam As String)
Set rsEmpNo = New Recordset
    rsEmpNo.Open "Select code,naym from empno where naym = '" & nam & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEmpNo.EOF And Not rsEmpNo.BOF Then
   SetCodeSA = Null2String(rsEmpNo!code)
End If
End Function

Sub initGrid()
With grdDetails
   .Rows = 7
   .ColWidth(0) = 1350
   .ColWidth(1) = 1350
   .ColWidth(2) = 1350
   .ColWidth(3) = 1350
   .ColWidth(4) = 1350
   .ColWidth(5) = 1350
   .ColWidth(6) = 1350
   .ColWidth(7) = 1
   .Row = 0
   .Col = 1
   .Text = "Amount"
   .Col = 2
   .Text = "Company"
   .Col = 3
   .Text = "Sales"
   .Col = 4
   .Text = "Warranty"
   .Col = 5
   .Text = "Discount"
   .Col = 6
   .Text = "Vat"
   .Col = 7
   .Text = "ID"
   .Col = 0
   .Row = 2
   .Text = "Labor"
   .Row = 3
   .Text = "Parts"
   .Row = 4
   .Text = "Materials"
   .Row = 5
   .Text = "TOTAL"
   .Row = 6
   .Text = "RO Amount"
End With
grdDetails.RemoveItem 1
With grdJobs
   .Rows = 2
   .ColWidth(0) = 1
   .ColWidth(1) = 1000
   .ColWidth(2) = 1000
   .ColWidth(3) = 3000
   .ColWidth(4) = 1500
   .ColWidth(5) = 1500
   .ColWidth(6) = 1520
   .Row = 0
   .Col = 0
   .Text = "ID"
   .Col = 1
   .Text = "Line No."
   .Col = 2
   .Text = "Job Code"
   .Col = 3
   .Text = "Job Description"
   .Col = 4
   .Text = "Details Amount"
   .Col = 5
   .Text = "WC"
   .Col = 6
   .Text = "Discount Value"
End With
With grdParts
   .Rows = 2
   .ColWidth(0) = 1
   .ColWidth(1) = 700
   .ColWidth(2) = 1000
   .ColWidth(3) = 2400
   .ColWidth(4) = 1100
   .ColWidth(5) = 1000
   .ColWidth(6) = 1000
   .ColWidth(7) = 500
   .ColWidth(8) = 1000
   .ColWidth(9) = 800
   .Row = 0
   .Col = 0
   .Text = "ID"
   .Col = 1
   .Text = "Line No."
   .Col = 2
   .Text = "Part Number"
   .Col = 3
   .Text = "Description"
   .Col = 4
   .Text = "Detail Volume"
   .Col = 5
   .Text = "Det. Price"
   .Col = 6
   .Text = "Det. Amount"
   .Col = 7
   .Text = "WC"
   .Col = 8
   .Text = "Disc. Value"
   .Col = 9
   .Text = "PO Code"
End With
With grdMaterials
   .Rows = 2
   .ColWidth(0) = 1
   .ColWidth(1) = 700
   .ColWidth(2) = 1000
   .ColWidth(3) = 2400
   .ColWidth(4) = 1000
   .ColWidth(5) = 1000
   .ColWidth(6) = 1000
   .ColWidth(7) = 1100
   .ColWidth(8) = 500
   .ColWidth(9) = 800
   .Row = 0
   .Col = 0
   .Text = "ID"
   .Col = 1
   .Text = "Line No."
   .Col = 2
   .Text = "Mat. Code"
   .Col = 3
   .Text = "Material"
   .Col = 4
   .Text = "Quantity"
   .Col = 5
   .Text = "Detail Price"
   .Col = 6
   .Text = "Det. Amount"
   .Col = 7
   .Text = "Disc. Value"
   .Col = 8
   .Text = "WC"
   .Col = 9
   .Text = "PO Code"
End With
End Sub

Sub AddJobs()
SSTab1.Tab = 1
SendToBack
cmdAddJobs.ZOrder 0
fraAddJobs.ZOrder 0
fraAddJobs.Enabled = True
AddorEdit = "ADD"
InitJobs
cboJcode.SetFocus
End Sub

Sub AddParts()
SSTab1.Tab = 2
SendToBack
cmdAddParts.ZOrder 0
fraAddParts.ZOrder 0
fraAddParts.Enabled = True
AddorEdit = "ADD"
InitParts
cboPartNo.SetFocus
End Sub

Sub AddMaterials()
SSTab1.Tab = 3
SendToBack
cmdAddMaterials.ZOrder 0
fraAddMaterials.ZOrder 0
fraAddMaterials.Enabled = True
AddorEdit = "ADD"
InitMaterials
cboMatCode.SetFocus
End Sub

Sub SendToBackDisc()
cmdDiscount.ZOrder 1
fraDiscount.ZOrder 1
txtDiscAmt.Text = 0
End Sub

Sub SendToFrontDisc()
cmdDiscount.ZOrder 0
fraDiscount.ZOrder 0
txtDiscAmt.Text = 5
End Sub

Sub FillDetails()
Dim COMTotal, SALESTotal, WARTotal, VATTotal As Double
Screen.MousePointer = 11
ROTotal = TOTJOBAMT + TOTPARTSAMT + TOTMATAMT
COMTotal = JobComTotal + PartsComTotal + MatComTotal
SALESTotal = JobSalesTotal + PartsSalesTotal + MatSalesTotal
WARTotal = JobWarTotal + PartsWarTotal + MatWarTotal

DiscTotal = N2Str2IntZero(rsESTI_HD!l_discount) + N2Str2IntZero(rsESTI_HD!p_discount) + N2Str2IntZero(rsESTI_HD!m_discount)
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select Estimateno from Esti_Det where Estimateno = " & N2Str2Null(rsESTI_HD!estimateno) & " order by LINE_NO asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   With grdDetails
        .Rows = 6
        .Col = 1
        .Row = 1
        .Text = Format(TOTJOBAMT, "###,###,##0.00")
        .Row = 2
        .Text = Format(TOTPARTSAMT, "###,###,##0.00")
        .Row = 3
        .Text = Format(TOTMATAMT, "###,###,##0.00")
        .Row = 4
        .Text = Format(ROTotal, "###,###,##0.00")
        .Row = 5
        .Text = Format(N2Str2Zero(rsESTI_HD!ro_amount), "###,###,##0.00")
        .Col = 2
        .Row = 1
        .Text = Format(JobComTotal, "###,###,##0.00")
        .Row = 2
        .Text = Format(PartsComTotal, "###,###,##0.00")
        .Row = 3
        .Text = Format(PartsComTotal, "###,###,##0.00")
        .Row = 4
        .Text = Format(COMTotal, "###,###,##0.00")
        .Col = 3
        .Row = 1
        .Text = Format(JobSalesTotal, "###,###,##0.00")
        .Row = 2
        .Text = Format(PartsSalesTotal, "###,###,##0.00")
        .Row = 3
        .Text = Format(PartsSalesTotal, "###,###,##0.00")
        .Row = 4
        .Text = Format(SALESTotal, "###,###,##0.00")
        .Col = 4
        .Row = 1
        .Text = Format(JobWarTotal, "###,###,##0.00")
        .Row = 2
        .Text = Format(PartsWarTotal, "###,###,##0.00")
        .Row = 3
        .Text = Format(PartsWarTotal, "###,###,##0.00")
        .Row = 4
        .Text = Format(WARTotal, "###,###,##0.00")
        .Col = 5
        .Row = 1
        .Text = Format(N2Str2Zero(rsESTI_HD!l_discount), "###,###,##0.00")
        .Row = 2
        .Text = Format(N2Str2Zero(rsESTI_HD!p_discount), "###,###,##0.00")
        .Row = 3
        .Text = Format(N2Str2Zero(rsESTI_HD!m_discount), "###,###,##0.00")
        .Row = 4
        .Text = Format(DiscTotal, "###,###,##0.00")
        .Col = 6
        .Row = 1
        .Text = Format(N2Str2Zero(rsESTI_HD!l_taxval), "###,###,##0.00")
        .Row = 2
        .Text = Format(N2Str2Zero(rsESTI_HD!p_taxval), "###,###,##0.00")
        .Row = 3
        .Text = Format(N2Str2Zero(rsESTI_HD!m_taxval), "###,###,##0.00")
        .Row = 4
        .Text = Format(VATTotal, "###,###,##0.00")
   End With
Else
   cleargrid grdDetails
   initGrid
End If
Screen.MousePointer = 0
End Sub

Sub FillJobs()
kcnt = 0
TOTJOBAMT = 0
TOTJOBDISC = 0
TOTJOBDISCVAL = 0
TOTJOBTAX = 0
JobComTotal = 0
JobSalesTotal = 0
JobWarTotal = 0
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,Estimateno,LINE_NO,detcde,detdsc,det_amt,wcode,discount_2,pocode,taxval,disval,livil from Esti_Det where Estimateno = " & N2Str2Null(rsESTI_HD!estimateno) & " order by LINE_NO asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   Screen.MousePointer = 11
   rsEsti_Det.MoveFirst
   Do While Not rsEsti_Det.EOF
      If rsEsti_Det!livil = "1" Then
         kcnt = kcnt + 1
         grdJobs.AddItem rsEsti_Det!ID & Chr(9) & Null2String(rsEsti_Det!line_no) & Chr(9) & _
                         Null2String(rsEsti_Det!detcde) & Chr(9) & _
                         Null2String(rsEsti_Det!detdsc) & Chr(9) & _
                         N2Str2Zero(rsEsti_Det!det_amt) & Chr(9) & _
                         Null2String(rsEsti_Det!wCode) & Chr(9) & _
                         N2Str2Zero(rsEsti_Det!discount_2) & Chr(9) & _
                         Null2String(rsEsti_Det!pocode)
         If Null2String(rsEsti_Det!wCode) = "C" Then
            JobComTotal = JobComTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "S" Then
            JobSalesTotal = JobSalesTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "W" Then
            JobWarTotal = JobWarTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) <> "C" And Null2String(rsEsti_Det!wCode) <> "S" And Null2String(rsEsti_Det!wCode) <> "W" Then
            TOTJOBAMT = TOTJOBAMT + N2Str2Zero(rsEsti_Det!det_amt)
            TOTJOBDISC = TOTJOBDISC + N2Str2Zero(rsEsti_Det!discount_2)
            TOTJOBDISCVAL = TOTJOBDISCVAL + N2Str2Zero(rsEsti_Det!disval)
            TOTJOBTAX = TOTJOBTAX + N2Str2Zero(rsEsti_Det!taxval)
         End If
      End If
      rsEsti_Det.MoveNext
   Loop
   If kcnt <> 0 Then grdJobs.RemoveItem 1
   Screen.MousePointer = 0
End If
End Sub

Sub FillParts()
Pcnt = 0
TOTPARTSAMT = 0
TOTPARTSDISC = 0
TOTPARTSDISCVAL = 0
TOTPARTSTAX = 0
PartsComTotal = 0
PartsSalesTotal = 0
PartsWarTotal = 0
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,Estimateno,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discount_2,pocode,taxval,disval,livil from Esti_Det where Estimateno = '" & rsESTI_HD!estimateno & "' order by LINE_NO asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   Screen.MousePointer = 11
   rsEsti_Det.MoveFirst
   Do While Not rsEsti_Det.EOF
      If rsEsti_Det!livil = "2" Then
         Pcnt = Pcnt + 1
         grdParts.AddItem rsEsti_Det!ID & Chr(9) & Null2String(rsEsti_Det!line_no) & Chr(9) & _
                          Null2String(rsEsti_Det!detcde) & Chr(9) & _
                          Null2String(rsEsti_Det!detdsc) & Chr(9) & _
                          N2Str2IntZero(rsEsti_Det!detvol) & Chr(9) & _
                          N2Str2Zero(rsEsti_Det!detprc) & Chr(9) & _
                          N2Str2Zero(rsEsti_Det!det_amt) & Chr(9) & _
                          Null2String(rsEsti_Det!wCode) & Chr(9) & _
                          N2Str2Zero(rsEsti_Det!discount_2) & Chr(9) & _
                          Null2String(rsEsti_Det!pocode)
         If Null2String(rsEsti_Det!wCode) = "C" Then
            PartsComTotal = PartsComTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "S" Then
            PartsSalesTotal = PartsSalesTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "W" Then
            PartsWarTotal = PartsWarTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) <> "C" And Null2String(rsEsti_Det!wCode) <> "S" And Null2String(rsEsti_Det!wCode) <> "W" Then
            TOTPARTSAMT = TOTPARTSAMT + N2Str2Zero(rsEsti_Det!det_amt)
            TOTPARTSDISC = TOTPARTSDISC + N2Str2Zero(rsEsti_Det!discount_2)
            TOTPARTSDISCVAL = TOTPARTSDISCVAL + N2Str2Zero(rsEsti_Det!disval)
            TOTPARTSTAX = TOTPARTSTAX + N2Str2Zero(rsEsti_Det!taxval)
         End If
      End If
      rsEsti_Det.MoveNext
   Loop
   If Pcnt <> 0 Then grdParts.RemoveItem 1
   Screen.MousePointer = 0
End If
End Sub

Sub FillMaterials()
Mcnt = 0
TOTMATAMT = 0
TOTMATDISC = 0
TOTMATDISCVAL = 0
TOTMATTAX = 0
MatComTotal = 0
MatSalesTotal = 0
MatWarTotal = 0
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,Estimateno,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,discount_2,taxval,disval,wcode,pocode,livil from Esti_Det where Estimateno = '" & rsESTI_HD!estimateno & "' order by LINE_NO asc", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   Screen.MousePointer = 11
   rsEsti_Det.MoveFirst
   Do While Not rsEsti_Det.EOF
      If rsEsti_Det!livil = "3" Then
         grdMaterials.AddItem rsEsti_Det!ID & Chr(9) & Null2String(rsEsti_Det!line_no) & Chr(9) & _
                              Null2String(rsEsti_Det!detcde) & Chr(9) & _
                              Null2String(rsEsti_Det!detdsc) & Chr(9) & _
                              N2Str2IntZero(rsEsti_Det!detvol) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!detprc) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!det_amt) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!discount_2) & Chr(9) & _
                              Null2String(rsEsti_Det!wCode) & Chr(9) & _
                              Null2String(rsEsti_Det!pocode)
         Mcnt = Mcnt + 1
         If Null2String(rsEsti_Det!wCode) = "C" Then
            MatComTotal = MatComTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "S" Then
            MatSalesTotal = MatSalesTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "W" Then
            MatWarTotal = MatWarTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) <> "C" And Null2String(rsEsti_Det!wCode) <> "S" And Null2String(rsEsti_Det!wCode) <> "W" Then
            TOTMATAMT = TOTMATAMT + N2Str2Zero(rsEsti_Det!det_amt)
            TOTMATDISC = TOTMATDISC + N2Str2Zero(rsEsti_Det!discount_2)
            TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsEsti_Det!disval)
            TOTMATTAX = TOTMATTAX + N2Str2Zero(rsEsti_Det!taxval)
         End If
      ElseIf rsEsti_Det!livil = "2" And (rsEsti_Det!detcde = "RUGS" Or rsEsti_Det!detdsc = "RUGS") Then
         grdMaterials.AddItem rsEsti_Det!ID & Chr(9) & Null2String(rsEsti_Det!line_no) & Chr(9) & _
                              Null2String(rsEsti_Det!detcde) & Chr(9) & _
                              Null2String(rsEsti_Det!detdsc) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!detvol) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!detprc) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!det_amt) & Chr(9) & _
                              N2Str2Zero(rsEsti_Det!discount_2) & Chr(9) & _
                              Null2String(rsEsti_Det!wCode) & Chr(9) & _
                              Null2String(rsEsti_Det!pocode)
         Mcnt = Mcnt + 1
         If Null2String(rsEsti_Det!wCode) = "C" Then
            MatComTotal = MatComTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "S" Then
            MatSalesTotal = MatSalesTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) = "W" Then
            MatWarTotal = MatWarTotal + N2Str2Zero(rsEsti_Det!det_amt)
         End If
         If Null2String(rsEsti_Det!wCode) <> "C" And Null2String(rsEsti_Det!wCode) <> "S" And Null2String(rsEsti_Det!wCode) <> "W" Then
            TOTMATAMT = TOTMATAMT + N2Str2Zero(rsEsti_Det!det_amt)
            TOTMATDISC = TOTMATDISC + N2Str2Zero(rsEsti_Det!discount_2)
            TOTMATDISCVAL = TOTMATDISCVAL + N2Str2Zero(rsEsti_Det!disval)
            TOTMATTAX = TOTMATTAX + N2Str2Zero(rsEsti_Det!taxval)
         End If
      End If
      rsEsti_Det.MoveNext
   Loop
   If Mcnt <> 0 Then grdMaterials.RemoveItem 1
   Screen.MousePointer = 0
End If
End Sub

Sub SendToBack()
cmdAddJobs.ZOrder 1
fraAddJobs.ZOrder 1
fraAddJobs.Enabled = False
cmdAddParts.ZOrder 1
fraAddParts.ZOrder 1
fraAddParts.Enabled = False
cmdAddMaterials.ZOrder 1
fraAddMaterials.ZOrder 1
fraAddMaterials.Enabled = False
End Sub

Function StoreJobsEntry(ByVal ID As Variant)
Dim retVal As Boolean
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,pocode,wcode,det_amt,discrate,EstimateNo,livil,detail from Esti_Det where id = " & ID, gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   labDetId.Caption = rsEsti_Det!ID
   txtJobLineNo.Text = Null2String(rsEsti_Det!line_no)
   cboJcode.Text = Null2String(rsEsti_Det!detcde)
   cboJobCode.Text = setJobDesc(Null2String(rsEsti_Det!detcde))
   txtJobPostCode.Text = Null2String(rsEsti_Det!pocode)
   If Null2String(rsEsti_Det!wCode) <> "" Then
      If UCase(rsEsti_Det!wCode) <> "0" And UCase(rsEsti_Det!wCode) <> "1" And UCase(rsEsti_Det!wCode) <> "D" Then
         cboJobChargeTo.Text = rsEsti_Det!wCode
      End If
   End If
   txtJobRate.Text = N2Str2Zero(rsEsti_Det!det_amt)
   txtJobDiscount.Text = N2Str2Zero(rsEsti_Det!discrate)
   txtJobDetail.Text = Null2String(rsEsti_Det!detail)
   retVal = True
Else
   retVal = False
End If
StoreJobsEntry = retVal
End Function

Private Sub Form_Unload(Cancel As Integer)
ESTISHOW = False
End Sub

Private Sub grdDetails_Click()
Dim Fild As String
If grdDetails.Row = 1 Then
   SSTab1.Tab = 1
End If
If grdDetails.Row = 2 Then
   SSTab1.Tab = 2
End If
If grdDetails.Row = 3 Then
   SSTab1.Tab = 3
End If
End Sub

Private Sub grdJobs_DblClick()
Dim Fild As String
grdJobs.Row = grdJobs.Row
grdJobs.Col = 0
Fild = grdJobs.Text
If Fild <> "" And Fild <> "No Entry" Then
   AddorEdit = "EDIT"
   SSTab1.Tab = 1
   SendToBack
   cmdAddJobs.ZOrder 0
   fraAddJobs.ZOrder 0
   fraAddJobs.Enabled = True
   fraAddJobs.Caption = "Edit Jobs"
   StoreJobsEntry (Fild)
Else
   MsgBox "No Entry on Jobs"
End If
End Sub

Function StorePartsEntry(ByVal ID As Variant)
Dim retVal As Boolean
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,detdsc,pocode,detvol,detprc,det_amt,wcode,discrate from Esti_Det where id = " & ID, gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   labDetId.Caption = rsEsti_Det!ID
   txtPartsLineNo.Text = Null2String(rsEsti_Det!line_no)
   cboPartNo.Text = Null2String(rsEsti_Det!detcde)
   cboDescription.Text = Null2String(rsEsti_Det!detdsc)
   txtPartCode.Text = Null2String(rsEsti_Det!pocode)
   txtQty.Text = N2Str2Zero(rsEsti_Det!detvol)
   txtUnitPrice.Text = Null2String(rsEsti_Det!detprc)
   txtPartAmount.Text = N2Str2Zero(rsEsti_Det!det_amt)
   If Null2String(rsEsti_Det!wCode) <> "" Then
      If UCase(rsEsti_Det!wCode) <> "0" And UCase(rsEsti_Det!wCode) <> "1" And UCase(rsEsti_Det!wCode) <> "D" Then
         cboChargeTo.Text = rsEsti_Det!wCode
      End If
   End If
   txtPartDiscount.Text = N2Str2IntZero(rsEsti_Det!discrate)
   retVal = True
Else
   retVal = False
End If
StorePartsEntry = retVal
End Function

Private Sub grdParts_DblClick()
Dim Fild As String
grdParts.Row = grdParts.Row
grdParts.Col = 0
Fild = grdParts.Text
If Fild <> "" And Fild <> "No Entry" Then
   AddorEdit = "EDIT"
   SSTab1.Tab = 2
   SendToBack
   cmdAddParts.ZOrder 0
   fraAddParts.ZOrder 0
   fraAddParts.Enabled = True
   fraAddParts.Caption = "Edit Parts"
   StorePartsEntry (Fild)
Else
   MsgBox "No Entry on Parts"
   Exit Sub
End If
End Sub

Function StoreMatEntry(ByVal ID As Variant)
Dim retVal As Boolean
Set rsEsti_Det = New Recordset
    rsEsti_Det.Open "select id,LINE_NO,detcde,detdsc,detvol,detprc,det_amt,wcode,discrate,pocode from Esti_Det where id = " & ID, gconCSMIOS, adOpenForwardOnly, adLockReadOnly
If Not rsEsti_Det.EOF And Not rsEsti_Det.BOF Then
   labDetId.Caption = rsEsti_Det!ID
   txtMatLineNo.Text = Null2String(rsEsti_Det!line_no)
   cboMatCode.Text = Null2String(rsEsti_Det!detcde)
   cboMaterial.Text = Null2String(rsEsti_Det!detdsc)
   txtMatQty.Text = N2Str2Zero(rsEsti_Det!detvol)
   txtMatUnitPrice.Text = Null2String(rsEsti_Det!detprc)
   txtMatAmount.Text = N2Str2Zero(rsEsti_Det!det_amt)
   If Null2String(rsEsti_Det!wCode) <> "" Then
      If UCase(rsEsti_Det!wCode) <> "0" And UCase(rsEsti_Det!wCode) <> "1" And UCase(rsEsti_Det!wCode) <> "D" Then
         cboMatChargeTo.Text = rsEsti_Det!wCode
      End If
   End If
   txtMatDiscount.Text = N2Str2Zero(rsEsti_Det!discrate)
   txtMatPOCode.Text = Null2String(rsEsti_Det!pocode)
   retVal = True
Else
   retVal = False
End If
StoreMatEntry = retVal
End Function

Private Sub grdMaterials_DblClick()
Dim Fild As String
grdMaterials.Row = grdMaterials.Row
grdMaterials.Col = 0
Fild = grdMaterials.Text
If Fild <> "" And Fild <> "No Entry" Then
   AddorEdit = "EDIT"
   SSTab1.Tab = 3
   SendToBack
   cmdAddMaterials.ZOrder 0
   fraAddMaterials.ZOrder 0
   fraAddMaterials.Enabled = True
   fraAddMaterials.Caption = "Edit Parts"
   StoreMatEntry (Fild)
Else
   MsgBox "No Entry on Materials"
   Exit Sub
End If
End Sub

Private Sub cbojcode_LostFocus()
cboJcode.Text = UCase(cboJcode.Text)
If cboJcode.Text <> "" Then
   cboJobCode.Text = setJobDesc(cboJcode.Text)
   txtJobPostCode.Text = setJobPOcode(cboJobCode.Text)
   txtJobRate.Text = setJobRate(cboJobCode.Text)
   If AddorEdit = "ADD" Then
      txtJobDetail.Text = setJobDetail(cboJobCode.Text)
   End If
End If
End Sub

Private Sub cbomatcode_LostFocus()
cboMaterial.Text = SetMatDisc(cboMatCode.Text)
txtMatUnitPrice.Text = SetMatPrice(cboMatCode.Text)
txtMatPOCode.Text = SetMatPOCode(cboMatCode.Text)
txtMatAmount.Text = txtMatUnitPrice.Text
End Sub

Function SetMatCode(mmm As String)
If mmm <> "" Then
   Set rsMATMas = New Recordset
       rsMATMas.Open "select matcde,matdsc from matmas where matdsc = '" & mmm & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsMATMas.EOF And Not rsMATMas.BOF Then
      SetMatCode = Null2String(rsMATMas!matcde)
   Else
      SetMatCode = ""
      MsgBox "Material Not Found!"
   End If
End If
End Function

Function SetMatDisc(mmm As String)
If mmm <> "" Then
   Set rsMATMas = New Recordset
       rsMATMas.Open "select matcde,matdsc from matmas where matcde = '" & mmm & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsMATMas.EOF And Not rsMATMas.BOF Then
      SetMatDisc = Null2String(rsMATMas!matdsc)
   Else
      SetMatDisc = ""
      MsgBox "Material Not Found!"
   End If
End If
End Function

Function SetMatPrice(mmm As String)
If mmm <> "" Then
   Set rsMATMas = New Recordset
       rsMATMas.Open "select matcde,s_price from matmas where matcde = '" & mmm & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsMATMas.EOF And Not rsMATMas.BOF Then
      SetMatPrice = Null2String(rsMATMas!s_price)
   Else
      SetMatPrice = ""
   End If
End If
End Function

Function SetMatPOCode(mmm As String)
If mmm <> "" Then
   Set rsMATMas = New Recordset
       rsMATMas.Open "select matcde,pocode from matmas where matcde = '" & mmm & "'", gconCSMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsMATMas.EOF And Not rsMATMas.BOF Then
      SetMatPOCode = Null2String(rsMATMas!pocode)
   Else
      SetMatPOCode = ""
   End If
End If
End Function

Private Sub txtCertific8_Click()
Me.Enabled = False
frmCusveh.Show
frmCusveh.ZOrder 0
End Sub

Private Sub txtMatAmount_GotFocus()
txtMatAmount.Text = Val(txtMatQty.Text) * Val(txtMatUnitPrice.Text)
End Sub

Private Sub txtMatQty_Change()
If txtMatQty.Text <> "" Then
   txtMatAmount.Text = Val(txtMatQty.Text) * Val(txtMatUnitPrice.Text)
End If
End Sub

Private Sub txtMatQty_LostFocus()
If txtMatQty.Text <> "" Then
   txtMatAmount.Text = Val(txtMatQty.Text) * Val(txtMatUnitPrice.Text)
End If
End Sub

Private Sub txtMatUnitPrice_Change()
If txtMatUnitPrice.Text <> "" Then
   txtMatAmount.Text = Val(txtMatQty.Text) * Val(txtMatUnitPrice.Text)
End If
End Sub

Private Sub txtMatUnitPrice_LostFocus()
If txtMatUnitPrice.Text <> "" Then
   txtMatAmount.Text = Val(txtMatQty.Text) * Val(txtMatUnitPrice.Text)
End If
End Sub

Private Sub txtPartAmount_GotFocus()
txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
End Sub

Private Sub cbopartno_LostFocus()
cboPartNo.Text = UCase(cboPartNo.Text)
cboDescription.Text = SetPartDisc(cboPartNo.Text)
txtUnitPrice.Text = SetPartPrice(cboPartNo.Text)
txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
SetPartStatus (cboPartNo.Text)
End Sub

Function SetPartID(xx As String)
If xx <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select id,partdesc from partmas where partdesc = '" & xx & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      SetPartID = Null2String(rsPartMas!ID)
   Else
      SetPartID = ""
      MsgBox "Part No. Not Found!"
   End If
End If
End Function

Function SetPartIDNo(xx As String)
If xx <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select id,partno from partmas where id = " & xx, gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      SetPartIDNo = Null2String(rsPartMas!PartNo)
   Else
      SetPartIDNo = ""
      MsgBox "Part No. Not Found!"
   End If
End If
End Function

Function SetPartNo(xx As String)
If xx <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select partno,partdesc from partmas where partdesc = '" & xx & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      SetPartNo = Null2String(rsPartMas!PartNo)
   Else
      SetPartNo = ""
      MsgBox "Part No. Not Found!"
   End If
End If
End Function

Function SetPartDisc(xx As String)
If xx <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select partno,partdesc from partmas where partno = '" & xx & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      SetPartDisc = Null2String(rsPartMas!partdesc)
   Else
      SetPartDisc = ""
      MsgBox "Part No. Not Found!"
   End If
End If
End Function

Function SetPartPrice(ppp As String)
If ppp <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select partno,srp from PARTMAS where partno = '" & ppp & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      SetPartPrice = Null2String(rsPartMas!srp)
   Else
      SetPartPrice = ""
   End If
End If
End Function

Sub SetPartStatus(ppp As String)
If ppp <> "" Then
   Set rsPartMas = New Recordset
       rsPartMas.Open "Select partno,onhand,partdesc from PARTMAS where partno = '" & ppp & "'", gconPMIOS, adOpenForwardOnly, adLockReadOnly
   If Not rsPartMas.EOF And Not rsPartMas.BOF Then
      If N2Str2IntZero(rsPartMas!onhand) <= 0 Then
         MsgBox "Warning: " & Null2String(rsPartMas!partdesc) & " is not availabled on Stock", vbCritical, "Out of Stock!"
      End If
   End If
End If
End Sub

Private Sub txtQty_Change()
If txtQty.Text <> "" Then
   txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
End If
End Sub

Private Sub txtQty_LostFocus()
If txtQty.Text <> "" Then
   txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
End If
End Sub

Private Sub txtUnitPrice_Change()
If txtUnitPrice.Text <> "" Then
   txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
End If
End Sub

Private Sub txtUnitPrice_LostFocus()
If txtUnitPrice.Text <> "" Then
   txtPartAmount.Text = Val(txtQty.Text) * Val(txtUnitPrice.Text)
End If
End Sub
