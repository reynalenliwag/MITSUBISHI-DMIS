VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSMATReceivingHist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Receiving Entry"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MATreceivingHist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   11760
   Begin VB.Frame Frame1 
      Height          =   3105
      Left            =   2220
      TabIndex        =   20
      Top             =   0
      Width           =   9495
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
         Left            =   5970
         TabIndex        =   2
         Text            =   "cboRecvd_Desc"
         Top             =   180
         Width           =   915
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   885
         Left            =   4620
         TabIndex        =   11
         Top             =   2130
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   1561
         _Version        =   393217
         BackColor       =   16777215
         TextRTF         =   $"MATreceivingHist.frx":08CA
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
         Left            =   1410
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1050
         Width           =   915
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
         TabIndex        =   7
         Text            =   "cboRecvd_Desc"
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
         Left            =   1410
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   180
         Width           =   915
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
         Left            =   3240
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1050
         Width           =   1245
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
         Left            =   6240
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   1200
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtNetRRAmt 
         Height          =   345
         Left            =   8010
         TabIndex        =   18
         Top             =   1620
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTTLRRAmt 
         Height          =   345
         Left            =   7950
         TabIndex        =   15
         Top             =   780
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDS1 
         Height          =   345
         Left            =   5400
         TabIndex        =   13
         Top             =   1200
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox txtPONo 
         Height          =   345
         Left            =   1410
         TabIndex        =   3
         Top             =   660
         Width           =   915
         _ExtentX        =   1614
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
      Begin MSMask.MaskEdBox txtDRNo 
         Height          =   345
         Left            =   1410
         TabIndex        =   9
         Top             =   2670
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   3240
         TabIndex        =   4
         Top             =   660
         Width           =   1245
         _ExtentX        =   2196
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
      Begin MSMask.MaskEdBox txtINVNo 
         Height          =   345
         Left            =   3390
         TabIndex        =   10
         Top             =   2670
         Width           =   1095
         _ExtentX        =   1931
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
      Begin MSMask.MaskEdBox txtRRDate 
         Height          =   345
         Left            =   3240
         TabIndex        =   1
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
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
         TabIndex        =   33
         Top             =   1800
         Width           =   4455
         Begin RichTextLib.RichTextBox txtDetails 
            Height          =   735
            Left            =   0
            TabIndex        =   8
            Top             =   30
            Width           =   4395
            _ExtentX        =   7752
            _ExtentY        =   1296
            _Version        =   393217
            BackColor       =   16777215
            TextRTF         =   $"MATreceivingHist.frx":095D
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
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   7860
         ScaleHeight     =   405
         ScaleWidth      =   1575
         TabIndex        =   17
         Top             =   1170
         Width           =   1575
         Begin MSMask.MaskEdBox txtDS_Amt1 
            Height          =   345
            Left            =   30
            TabIndex        =   16
            Top             =   30
            Width           =   1515
            _ExtentX        =   2672
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
      Begin VB.Label Label5 
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
         TabIndex        =   23
         Top             =   2670
         Width           =   795
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
         Height          =   315
         Left            =   7620
         TabIndex        =   55
         Top             =   210
         Width           =   1725
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
         Height          =   255
         Left            =   4620
         TabIndex        =   54
         Top             =   1830
         Width           =   885
      End
      Begin VB.Label Label3 
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
         Left            =   90
         TabIndex        =   24
         Top             =   690
         Width           =   1275
      End
      Begin VB.Label Label11 
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
         Left            =   2370
         TabIndex        =   22
         Top             =   690
         Width           =   795
      End
      Begin VB.Label Label1 
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
         TabIndex        =   32
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label2 
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
         Left            =   2370
         TabIndex        =   31
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label4 
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
         TabIndex        =   30
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label6 
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
         Left            =   2370
         TabIndex        =   29
         Top             =   1110
         Width           =   795
      End
      Begin VB.Label Label7 
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
         TabIndex        =   28
         Top             =   1080
         Width           =   1275
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
         Left            =   6720
         TabIndex        =   27
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
         Left            =   6750
         TabIndex        =   26
         Top             =   1650
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
         Height          =   225
         Left            =   3660
         TabIndex        =   25
         Top             =   1470
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label12 
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
         TabIndex        =   21
         Top             =   2670
         Width           =   855
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2445
      Left            =   2220
      TabIndex        =   19
      Top             =   3000
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2175
         Left            =   60
         TabIndex        =   12
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
   Begin Crystal.CrystalReport rptReceiving 
      Left            =   2370
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Materials Receiving Printout..."
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
   Begin wizButton.cmd cmdAddTran 
      Height          =   4245
      Left            =   4500
      TabIndex        =   56
      Top             =   900
      Width           =   4755
      _ExtentX        =   8387
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
      MICON           =   "MATreceivingHist.frx":09E5
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
      Left            =   4590
      TabIndex        =   34
      Top             =   960
      Width           =   4575
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
         Picture         =   "MATreceivingHist.frx":0A01
         Style           =   1  'Graphical
         TabIndex        =   50
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
         Picture         =   "MATreceivingHist.frx":0D13
         Style           =   1  'Graphical
         TabIndex        =   49
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
         TabIndex        =   42
         Text            =   "Text1"
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
         Picture         =   "MATreceivingHist.frx":1155
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3120
         Width           =   1005
      End
      Begin VB.ComboBox cboTranMatDsc 
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
         TabIndex        =   44
         Text            =   "Combo1"
         Top             =   1230
         Width           =   4335
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
         Left            =   1470
         Sorted          =   -1  'True
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   600
         Width           =   2295
      End
      Begin MSMask.MaskEdBox txtTranQty 
         Height          =   315
         Left            =   1470
         TabIndex        =   45
         Top             =   1620
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTranINVAmt 
         Height          =   315
         Left            =   1470
         TabIndex        =   46
         Top             =   1980
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
         Left            =   1470
         TabIndex        =   48
         Top             =   2700
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
         Left            =   1590
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   600
         Width           =   585
      End
      Begin MSMask.MaskEdBox txtUnitCost 
         Height          =   315
         Left            =   1470
         TabIndex        =   47
         Top             =   2340
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
      Begin VB.Label Label38 
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
         TabIndex        =   35
         Top             =   2700
         Width           =   1305
      End
      Begin VB.Label Label14 
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
         TabIndex        =   53
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
         TabIndex        =   41
         Top             =   3330
         Width           =   285
      End
      Begin VB.Label Label30 
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
         TabIndex        =   40
         Top             =   1980
         Width           =   1245
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
         Height          =   255
         Left            =   600
         TabIndex        =   39
         Top             =   1620
         Width           =   855
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
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   630
         Width           =   1335
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
         TabIndex        =   37
         Top             =   270
         Width           =   885
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   60
      TabIndex        =   57
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
         TabIndex        =   58
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstMATREC 
         Height          =   5085
         Left            =   60
         TabIndex        =   61
         Top             =   1350
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   8969
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
         MouseIcon       =   "MATreceivingHist.frx":145F
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
   Begin VB.PictureBox Picture1 
      Height          =   870
      Left            =   3180
      ScaleHeight     =   810
      ScaleWidth      =   8475
      TabIndex        =   63
      Top             =   5550
      Width           =   8535
      Begin VB.CommandButton cmdPost 
         Caption         =   "Post Entry"
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
         Left            =   4935
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATreceivingHist.frx":15C1
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":1713
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Press F11 for Posting By Range"
         Top             =   15
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelRR 
         Caption         =   "Cancel Transaction"
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
         Left            =   5655
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATreceivingHist.frx":1A38
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":1B8A
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   15
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
         Left            =   0
         MouseIcon       =   "MATreceivingHist.frx":1EC4
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":2016
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   15
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
         Left            =   705
         MouseIcon       =   "MATreceivingHist.frx":2375
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":24C7
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   15
         Width           =   705
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
         Left            =   1410
         MouseIcon       =   "MATreceivingHist.frx":281F
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":2971
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   15
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
         Left            =   3525
         MouseIcon       =   "MATreceivingHist.frx":2C6B
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":2DBD
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   15
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
         Left            =   4230
         MouseIcon       =   "MATreceivingHist.frx":30D0
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":3222
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   15
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
         Left            =   6360
         MouseIcon       =   "MATreceivingHist.frx":357E
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":36D0
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   15
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
         Left            =   7065
         MouseIcon       =   "MATreceivingHist.frx":39FB
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":3B4D
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   15
         Width           =   705
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
         Left            =   7770
         MouseIcon       =   "MATreceivingHist.frx":3EB3
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":4005
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   15
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
         Left            =   2820
         MouseIcon       =   "MATreceivingHist.frx":436B
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":44BD
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   15
         Width           =   705
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
         Left            =   2115
         MouseIcon       =   "MATreceivingHist.frx":480D
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":495F
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   15
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   855
      Left            =   10215
      ScaleHeight     =   795
      ScaleWidth      =   1410
      TabIndex        =   76
      Top             =   5550
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
         MouseIcon       =   "MATreceivingHist.frx":4CBD
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":4E0F
         Style           =   1  'Graphical
         TabIndex        =   78
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
         MouseIcon       =   "MATreceivingHist.frx":515F
         MousePointer    =   99  'Custom
         Picture         =   "MATreceivingHist.frx":52B1
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   0
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMSMATReceivingHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMATREC_HIST, rsPO_HD, rsDAYTRAN As ADODB.Recordset
Dim rsMatMas, rsSupplier As ADODB.Recordset
Dim rsCunter As ADODB.Recordset
Dim Pcnt As Integer
Dim AddorEdit As String
Dim MATREC_HIST_TOTUCOST, MATREC_HIST_TOTINVAMT, MATREC_HIST_TOTVAT As Double
Dim RR_QTY_REC As Long
Dim PrevRRNo As String
Dim ISNONVAT As Boolean

Private Sub cboRecvd_Desc_Click()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DragDrop(Source As Control, x As Single, y As Single)
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_DropDown()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboRecvd_Desc_LostFocus()
txtRecvd_Code.Text = SetSupCode(cboRecvd_Desc.Text)
End Sub

Private Sub cboTranMatDsc_Click()
If cboTranMatDsc.Text <> "" Then
   txtMaterialID.Text = SetMatIDDesc(cboTranMatDsc.Text)
   cboTranMatCde.Text = Setmatcde(txtMaterialID.Text)
   cboTranMatDsc.Text = Setmatdsc2(txtMaterialID.Text)
End If
End Sub

Private Sub cboTranMatDsc_LostFocus()
cboTranMatDsc.Text = cboTranMatDsc.Text
End Sub

Private Sub cboTranmatcde_Change()
If cboTranMatCde.Text <> "" Then
   txtMaterialID.Text = SetMatIDmatcde(cboTranMatCde.Text)
   cboTranMatDsc.Text = Setmatdsc2(txtMaterialID.Text)
End If
End Sub

Private Sub cboTranmatcde_Click()
If cboTranMatCde.Text <> "" Then
   txtMaterialID.Text = SetMatIDmatcde(cboTranMatCde.Text)
   cboTranMatDsc.Text = Setmatdsc2(txtMaterialID.Text)
End If
End Sub

Private Sub cboTranmatcde_LostFocus()
cboTranMatCde.Text = cboTranMatCde.Text
End Sub

Private Sub cmdAddTran_Click()
If Picture1.Visible = True Then
   SendToBack
   cmdAddTran.ZOrder 0
   fraAddTran.ZOrder 0
   cmdTranDelete.Visible = False
   fraAddTran.Enabled = True
   AddorEdit = "ADD"
   InitMaterials
   cboTranMatCde.SetFocus
End If
End Sub

Private Sub cmdCancelRR_Click()
If Function_Access(LOGID, "Acess_CancelEntry") = False Then Exit Sub

If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
   Dim PCurOnOrder, PCurTRECQTY, PCurReceipts As Integer
   Dim PCurLast_recq, PCurTpoQty As Integer
   Dim rsDAYTRANDup, rsPartmasDup As ADODB.Recordset
   Set rsDAYTRANDup = New ADODB.Recordset
       rsDAYTRANDup.Open "select trantype,tranno,tranqty,matcde from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno), gconDMIS
   If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
      rsDAYTRANDup.MoveFirst
      Do While Not rsDAYTRANDup.EOF
         Set rsPartmasDup = New ADODB.Recordset
             rsPartmasDup.Open "select matcde,onorder,trecqty,receipts,last_recq from MatMas where matcde = " & N2Str2Null(rsDAYTRANDup!MATCDE), gconDMIS
         If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
            PCurOnOrder = N2Str2IntZero(rsPartmasDup!onorder) + N2Str2IntZero(rsDAYTRANDup!tranqty)
            PCurTRECQTY = N2Str2IntZero(rsPartmasDup!trecqty) - N2Str2IntZero(rsDAYTRANDup!tranqty)
            PCurReceipts = N2Str2IntZero(rsPartmasDup!receipts) - N2Str2IntZero(rsDAYTRANDup!tranqty)
            PCurLast_recq = N2Str2IntZero(rsPartmasDup!last_recq) - N2Str2IntZero(rsDAYTRANDup!tranqty)
            gconDMIS.Execute "update MatMas set" & _
                             " onorder = " & PCurOnOrder & "," & _
                             " trecqty = " & PCurTRECQTY & "," & _
                             " receipts = " & PCurReceipts & "," & _
                             " last_recq = " & PCurLast_recq & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where matcde = " & N2Str2Null(rsDAYTRANDup!MATCDE)
         End If
         rsDAYTRANDup.MoveNext
      Loop
   End If
   gconDMIS.Execute "update MATREC_HIST set" & _
                    " status = 'C'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   gconDMIS.Execute "update PMIS_DayTran set" & _
                    " status = 'C'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where tranno = " & N2Str2Null(rsMATREC_HIST!rrno) & " and trantype = 'RR'"
   rsRefresh
   On Error Resume Next
   rsMATREC_HIST.Find "id =" & labID.Caption
   StoreMemVars
End If
End Sub

Private Sub cmdDelete_Click()
If Function_Access(LOGID, "Acess_Delete") = False Then Exit Sub

End Sub

Private Sub cmdPost_Click()
If Function_Access(LOGID, "Acess_Post") = False Then Exit Sub

Dim mmasOnOrder As Integer
If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
   Set rsDAYTRAN = New ADODB.Recordset
       rsDAYTRAN.Open "select id,itemno,trantype,tranno,matcde,tranqty,traninvamt from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno) & " order by itemno asc", gconDMIS
   If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
      rsDAYTRAN.MoveFirst
      Do While Not rsDAYTRAN.EOF
         If N2Str2Zero(rsDAYTRAN!TRANINVAMT) <= 0 Then
            MsgSpeechBox "Transaction with Invoice Amount equal to Zero Encountered!"
            Exit Sub
         End If
         rsDAYTRAN.MoveNext
      Loop
      rsDAYTRAN.MoveFirst
      Do While Not rsDAYTRAN.EOF
         Set rsMatMas = New ADODB.Recordset
             rsMatMas.Open "Select matcde,onhand,trecqty,onorder,receipts from MatMas where matcde = " & N2Str2Null(rsDAYTRAN!MATCDE), gconDMIS
         If Not rsMatMas.EOF And Not rsMatMas.EOF Then
            mmasOnOrder = N2Str2Zero(rsMatMas!onorder)
            If mmasOnOrder <= 0 Then
               mmasOnOrder = NumericVal(rsDAYTRAN!tranqty)
            End If
            gconDMIS.Execute "update MatMas set onhand =" & N2Str2Zero(rsMatMas!ONHAND) + NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " trecqty = " & N2Str2Zero(rsMatMas!trecqty) + NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " onorder = " & mmasOnOrder - NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " receipts = " & N2Str2Zero(rsMatMas!receipts) + NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " last_recq = " & N2Str2Zero(rsDAYTRAN!tranqty) & ", " & _
                             " last_recd = '" & LOGDATE & "', " & _
                             " supcode = " & N2Str2Null(Mid(txtRecvd_Code.Text, 1, 5)) & _
                             " where matcde = " & N2Str2Null(rsMatMas!MATCDE)
            gconDMIS.Execute "update PMIS_DayTran set" & _
                             " status = 'P'" & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where id = " & rsDAYTRAN!ID
         End If
         rsDAYTRAN.MoveNext
      Loop
   End If
   gconDMIS.Execute "update MATREC_HIST set" & _
                    " status = 'P'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   rsRefresh
   On Error Resume Next
   rsMATREC_HIST.Find "id =" & labID.Caption
   StoreMemVars
End If
End Sub

Private Sub cmdPrint_Click()
If Function_Access(LOGID, "Acess_Print") = False Then Exit Sub

Screen.MousePointer = 11
      rptReceiving.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
      rptReceiving.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
PrintSQLReport rptReceiving, CSMS_REPORT_PATH & "rr.RPT", "{MATREC_HIST.rrno} = '" & txtRRNo.Text & "'", DMIS_REPORT_Connection, 1
Screen.MousePointer = 0
End Sub

Private Sub cmdTranCancel_Click()
SendToBack
StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Material Entry") = True Then
   gconDMIS.Execute "delete from PMIS_DayTran where id = " & labDetID.Caption
   ShowDeletedMsg
End If
Dim cnt As Integer
Dim rsDAYTRANDup As ADODB.Recordset
Set rsDAYTRANDup = New ADODB.Recordset
    rsDAYTRANDup.Open "select id,itemno from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno) & " order by itemno asc", gconDMIS
If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
   rsDAYTRANDup.MoveFirst
   cnt = 0
   Do While Not rsDAYTRANDup.EOF
      cnt = cnt + 1
      gconDMIS.Execute "update PMIS_DayTran set itemno = " & Format(cnt, "0000") & " where id = " & rsDAYTRANDup!ID
      rsDAYTRANDup.MoveNext
   Loop
End If
FillDetails
If NumericVal(txtDS1.Text) > 0 Then
   MATREC_HIST_TOTVAT = MATREC_HIST_TOTINVAMT - MATREC_HIST_TOTUCOST
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = '" & "VAT" & "'," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
Else
   MATREC_HIST_TOTVAT = 0
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = NULL," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsMATREC_HIST.Find "id = " & labID.Caption
cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
Screen.MousePointer = 11
On Error GoTo ErrorCode

If cboTranMatCde.Text = "" Then
   MsgSpeechBox "Material Code must have a value"
   Exit Sub
End If

If AddorEdit = "ADD" Then
   Dim rsDAYTRANClone As ADODB.Recordset
   Set rsDAYTRANClone = New ADODB.Recordset
       rsDAYTRANClone.Open "select trantype,tranno,itemno,matcde from PMIS_DayTran where matcde = '" & cboTranMatCde.Text & "' and trantype = 'RR' and tranno =" & N2Str2Null(rsMATREC_HIST!rrno) & " order by itemno asc", gconDMIS
   If Not rsDAYTRANClone.EOF And Not rsDAYTRANClone.BOF Then
      MsgSpeechBox "Material Code already used in this transaction!"
      Exit Sub
   End If
End If

Dim RRTRANDATE, RRTRANNO, RRTRANTYPE As String
Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP As String
Dim RRTRANQTY As Integer
Dim RRTRANUCOST, RRTRANINVAMT As Double
Dim RRSTATUS, RRIN_OUT As String

RRTRANDATE = N2Date2Null(txtRRDate.Text)
RRTRANTYPE = "'RR'"
RRTRANNO = N2Str2Null(txtRRNo.Text)
RRITEMNO = N2Str2Null(Format(txtTranItemNo.Text, "0000"))
RRSTOCK_ORD = N2Str2Null(cboTranMatCde.Text)
RRSTOCK_SUP = N2Str2Null(cboTranMatDsc.Text)
RRTRANQTY = NumericVal(txtTranQty.Text)
RRTRANINVAMT = NumericVal(txtTranINVAmt.Text)
RRTRANUCOST = NumericVal(txtUnitCost.Text)
RRIN_OUT = "'I'"
RRSTATUS = "'N'"

If RRTRANINVAMT <= 0 Then
   Screen.MousePointer = 0
   MsgSpeechBox "Invoice Amount must not be zero."
   Exit Sub
End If

If AddorEdit = "ADD" Then
   gconDMIS.Execute "insert into PMIS_DayTran " & _
                    "(trandate,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                    " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                    " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                    " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                    " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
Else
   gconDMIS.Execute "update PMIS_DayTran set" & _
                    " trandate = " & RRTRANDATE & "," & _
                    " trantype = " & RRTRANTYPE & "," & _
                    " tranno = " & RRTRANNO & "," & _
                    " itemno = " & RRITEMNO & "," & _
                    " matcde = " & RRSTOCK_ORD & "," & _
                    " matdsc = " & RRSTOCK_SUP & "," & _
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
Dim rsMatMasClone As ADODB.Recordset
Set rsMatMasClone = New ADODB.Recordset
    rsMatMasClone.Open "select matcde,trecqty,receipts,onorder,tpoqty from MatMas where matcde = " & RRSTOCK_ORD, gconDMIS
If rsMatMasClone.EOF And rsMatMasClone.BOF Then
   gconDMIS.Execute "insert into MatMas" & _
                    "(matcde,matdsc,date_entered)" & _
                    " values (" & N2Str2Null(cboTranMatCde.Text) & ", " & N2Str2Null(Mid(cboTranMatDsc.Text, 1, 50)) & _
                    ",'" & LOGDATE & "')"
End If
cleargrid grdDetails
FillDetails
If NumericVal(txtDS1.Text) > 0 Then
   MATREC_HIST_TOTVAT = MATREC_HIST_TOTINVAMT - MATREC_HIST_TOTUCOST
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = '" & "VAT" & "'," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
Else
   MATREC_HIST_TOTVAT = 0
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = NULL," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsMATREC_HIST.Find "id = " & labID.Caption
cmdTranCancel.Value = True
If AddorEdit = "ADD" Then cmdAddTran_Click
Screen.MousePointer = 0
Exit Sub

ErrorCode:
ShowVBError
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub cmdUnPost_Click()
If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
   Set rsDAYTRAN = New ADODB.Recordset
       rsDAYTRAN.Open "select id,itemno,trantype,tranno,matcde,tranqty,traninvamt from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno) & " order by itemno asc", gconDMIS
   If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
      rsDAYTRAN.MoveFirst
      Do While Not rsDAYTRAN.EOF
         Set rsMatMas = New ADODB.Recordset
             rsMatMas.Open "Select matcde,onhand,trecqty,onorder,receipts from MatMas where matcde = " & N2Str2Null(rsDAYTRAN!MATCDE), gconDMIS
         If Not rsMatMas.EOF And Not rsMatMas.EOF Then
            gconDMIS.Execute "update MatMas set onhand =" & N2Str2Zero(rsMatMas!ONHAND) - NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " trecqty = " & N2Str2Zero(rsMatMas!trecqty) - NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " onorder = " & N2Str2Zero(rsMatMas!onorder) + NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " receipts = " & N2Str2Zero(rsMatMas!receipts) - NumericVal(rsDAYTRAN!tranqty) & ", " & _
                             " last_recq = " & 0 & ", " & _
                             " last_recd = NULL, " & _
                             " supcode = NULL" & _
                             " where matcde = " & N2Str2Null(rsMatMas!MATCDE)
            gconDMIS.Execute "update PMIS_DayTran set" & _
                             " status = 'N'" & "," & _
                             " usercode = " & N2Str2Null(LOGCODE) & "," & _
                             " lastupdate = '" & LOGDATE & "'" & _
                             " where id = " & rsDAYTRAN!ID
         End If
         rsDAYTRAN.MoveNext
      Loop
   End If
   gconDMIS.Execute "update MATREC_HIST set" & _
                    " status = 'N'," & _
                    " usercode = " & N2Str2Null(LOGCODE) & "," & _
                    " lastupdate = '" & LOGDATE & "'" & _
                    " where id = " & labID.Caption
   rsRefresh
   On Error Resume Next
   rsMATREC_HIST.Find "id =" & labID.Caption
   StoreMemVars
End If
End Sub

Private Sub cmdAdd_Click()
If Function_Access(LOGID, "Acess_Add") = False Then Exit Sub

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
StoreMemVars
End Sub

Private Sub cmdEdit_Click()
If Function_Access(LOGID, "Acess_Edit") = False Then Exit Sub

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
'Picture5.Visible = False
'Dim findStr As String
'findStr = InputSpeechBox("Please Input Transaction Number", txtRRNo.Text)
'If findStr <> "" Then
'   On Error GoTo ErrorCode
'   rsMATREC_HIST.Bookmark = rsFind(rsMATREC_HIST.Clone, "rrno", findStr).Bookmark
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
rsMATREC_HIST.Bookmark = rsFind(rsMATREC_HIST.Clone, "rrno", Format(DDD, "000000")).Bookmark
StoreMemVars
End Sub

Private Sub cmdFirst_Click()
rsMATREC_HIST.MoveFirst
StoreMemVars
End Sub

Private Sub cmdLast_Click()
rsMATREC_HIST.MoveLast
StoreMemVars
End Sub

Private Sub cmdNext_Click()
rsMATREC_HIST.MoveNext
If rsMATREC_HIST.EOF Then
   rsMATREC_HIST.MoveLast
   ShowLastRecordMsg
End If
StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
rsMATREC_HIST.MovePrevious
If rsMATREC_HIST.BOF Then
   rsMATREC_HIST.MoveFirst
   ShowFirstRecordMsg
End If
StoreMemVars
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrorCode
Dim rsMATREC_HISTDup As ADODB.Recordset
      
If Trim(txtRRNo.Text) = "" Then
   MsgSpeechBox "RR Number must not be empty"
   On Error Resume Next
   txtRRNo.SetFocus
   Exit Sub
Else
   If AddorEdit = "ADD" Then
      Dim rsfindDup As ADODB.Recordset
      Set rsfindDup = New ADODB.Recordset
          rsfindDup.Open "select rrno from MATREC_HIST where rrno = '" & txtRRNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
      If Not rsfindDup.EOF And Not rsfindDup.BOF Then
         MsgSpeechBox "RR Number already exist!"
         On Error Resume Next
         txtRRNo.SetFocus
         Exit Sub
      End If
      Set rsMATREC_HISTDup = New ADODB.Recordset
          rsMATREC_HISTDup.Open "select pono from MATREC_HIST where pono = '" & txtPONo.Text & "'", gconDMIS
      If Not rsMATREC_HISTDup.EOF And Not rsMATREC_HISTDup.BOF Then
         MsgSpeechBox "PO Number Already Received"
         Exit Sub
      End If
   End If
   
End If
If txtRRDate.Text = "" Or IsDate(txtRRDate.Text) = False Then
   MsgSpeechBox "Invalid RR Date!"
   On Error Resume Next
   txtRRDate.SetFocus
   Exit Sub
End If

Dim NewRRCunTer As String
NewRRCunTer = NumericVal(txtRRNo.Text) + 1

Dim VTXTRRNo, VTXTRRDate, Vcboclasscode As String
Dim VTXTRecvd_Code, VTXTRecvd_From, VtxtAddress As String
Dim VTXTTerms, VTXTPONo, VTXTPODate As String
Dim VTXTDRNo, VTXTINVNo As String
Dim VTXTTTLRRAmt, VTXTDS1 As Double
Dim VTXTDS_Desc1 As String
Dim VTXTDS_Amt1, VTXTNetRRAmt As Double
Dim VTXTRemarks As String

Dim RRTRANDATE, RRTRANNO, RRTRANTYPE As String
Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP As String
Dim RRTRANQTY As Integer
Dim RRTRANUCOST, RRTRANINVAMT As Double
Dim RRIN_OUT, RRSTATUS As String

VTXTRRNo = N2Str2Null(txtRRNo.Text)
VTXTRRDate = N2Date2Null(txtRRDate.Text)
Vcboclasscode = N2Str2Null(cboClasscode.Text)
VTXTRecvd_Code = N2Str2Null(txtRecvd_Code.Text)
VTXTRecvd_From = N2Str2Null(cboRecvd_Desc.Text)
VtxtAddress = N2Str2Null(txtDetails.Text)
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
   gconDMIS.Execute "Insert into MATREC_HIST" & _
                    " (rrno,rrdate,classcode,recvd_code,recvd_from,address,terms,pono,podate,drno,invno,ttlrramt,ds1,ds_desc1,ds_amt1,netrramt,usercode,lastupdate,remarks)" & _
                    " values (" & VTXTRRNo & ", " & VTXTRRDate & ", " & Vcboclasscode & ", " & _
                    " " & VTXTRecvd_Code & ", " & VTXTRecvd_From & ", " & VtxtAddress & ", " & VTXTTerms & _
                    ", " & VTXTPONo & ", " & VTXTPODate & ", " & VTXTDRNo & ", " & VTXTINVNo & _
                    ", " & VTXTTTLRRAmt & _
                    ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                    ", " & VTXTNetRRAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
Else
   gconDMIS.Execute "update MATREC_HIST set" & _
                    " rrno = " & VTXTRRNo & "," & _
                    " rrdate = " & VTXTRRDate & "," & _
                    " classcode = " & Vcboclasscode & "," & _
                    " recvd_code = " & VTXTRecvd_Code & "," & _
                    " recvd_from = " & VTXTRecvd_From & "," & _
                    " address = " & VtxtAddress & "," & _
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
   gconDMIS.Execute "update PMIS_DayTran set" & _
                    " trandate = " & VTXTRRDate & "," & _
                    " tranno = " & VTXTRRNo & _
                    " where trantype = 'RR' and tranno = '" & PrevRRNo & "'"
End If
If AddorEdit = "ADD" Then
   gconDMIS.Execute "update PMIS_Counter set nextnumber = '" & NewRRCunTer & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where modul = 'RR'"
End If
rsRefresh
On Error Resume Next
rsMATREC_HIST.Find "rrno = " & VTXTRRNo
cmdCancel.Value = True
On Error GoTo ErrorCode
If AddorEdit = "ADD" Then
   Dim rsDAYTRANDup, rsDAYTRANDUp2 As ADODB.Recordset
   Dim varPmasTrecqty, varPmasOnOrder, varPmasOnhand As Long
   Dim rsMatMasClone As ADODB.Recordset
   Set rsDAYTRANDup = New ADODB.Recordset
       rsDAYTRANDup.Open "select trantype,tranno from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno), gconDMIS
   If rsDAYTRANDup.EOF And rsDAYTRANDup.BOF Then
      rsDAYTRANDup.Close
      Set rsDAYTRANDUp2 = New ADODB.Recordset
          rsDAYTRANDUp2.Open "select trantype,tranno,matcde,matdsc,itemno,tranqty,traninvamt,tranucost from PMIS_DayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsMATREC_HIST!pono), gconDMIS
      If Not rsDAYTRANDUp2.EOF And Not rsDAYTRANDUp2.BOF Then
         rsDAYTRANDUp2.MoveFirst
         Do While Not rsDAYTRANDUp2.EOF
            RRTRANDATE = N2Date2Null(txtPODate.Text)
            RRTRANTYPE = "'RR'"
            RRTRANNO = N2Str2Null(rsMATREC_HIST!rrno)
            RRITEMNO = N2Str2Null(Null2String(rsDAYTRANDUp2!itemno))
            RRSTOCK_ORD = N2Str2Null(rsDAYTRANDUp2!MATCDE)
            RRSTOCK_SUP = N2Str2Null(rsDAYTRANDUp2!MatDsc)
            RRTRANQTY = N2Str2IntZero(rsDAYTRANDUp2!tranqty)
            RRTRANINVAMT = N2Str2Zero(rsDAYTRANDUp2!TRANINVAMT)
            RRTRANUCOST = N2Str2Zero(rsDAYTRANDUp2!TRANUCOST)
            RRIN_OUT = "'I'"
            RRSTATUS = "'N'"
            
            gconDMIS.Execute "insert into PMIS_DayTran " & _
                             "(trandate,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                             " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                             " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                             " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                             " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
            rsDAYTRANDUp2.MoveNext
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
   MATREC_HIST_TOTVAT = MATREC_HIST_TOTINVAMT - MATREC_HIST_TOTUCOST
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = '" & "VAT" & "'," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
Else
   MATREC_HIST_TOTVAT = 0
   gconDMIS.Execute "update MATREC_HIST set" & _
                      " ttlrramt = " & MATREC_HIST_TOTUCOST & "," & _
                      " netrramt = " & MATREC_HIST_TOTINVAMT & "," & _
                      " ds_desc1 = NULL," & _
                      " ds_amt1 = " & MATREC_HIST_TOTVAT & "," & _
                      " ds1 = " & NumericVal(txtDS1.Text) & _
                      " where id = " & labID.Caption
End If
rsRefresh
On Error Resume Next
rsMATREC_HIST.Find "rrno = " & VTXTRRNo
StoreMemVars
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
               StoreMemVars
            End If
       Case vbKeyF3
            If Picture1.Visible = True Then
               If Null2String(rsMATREC_HIST!Status) = "P" Then
                  MsgSpeechBox "Item(s) are Already Posted, and cannot be change..."
               ElseIf Null2String(rsMATREC_HIST!Status) = "C" Then
                  MsgSpeechBox "Transactions are Already Cancelled, and cannot be Change..."
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
textSearch.Text = "": 'Picture5.ZOrder 0
Frame1.Enabled = False
Picture1.Visible = True
Picture2.Visible = False
txtMaterialID.Text = ""
initMemvars
StoreMemVars
Screen.MousePointer = 0
End Sub

Sub rsRefresh()
Set rsMATREC_HIST = New ADODB.Recordset
    rsMATREC_HIST.Open "select * from CSMS_MATREC_HIST order by id desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
cleargrid grdDetails
InitGrid
InitCbo
InitCboClasscode
InitMaterials
End Sub

Sub StoreMemVars()
If Not rsMATREC_HIST.EOF And Not rsMATREC_HIST.BOF Then
   labID.Caption = rsMATREC_HIST!ID
   txtRRNo.Text = Null2String(rsMATREC_HIST!rrno)
   txtRRDate.Text = Null2String(rsMATREC_HIST!rrdate)
   cboClasscode.Text = Null2String(rsMATREC_HIST!classcode)
   txtRecvd_Code.Text = Null2String(rsMATREC_HIST!recvd_code)
   cboRecvd_Desc.Text = SetSupdesc(Null2String(rsMATREC_HIST!recvd_code))
   txtDetails.Text = Null2String(rsMATREC_HIST!Address)
   txtTerms.Text = Null2String(rsMATREC_HIST!terms)
   txtPONo.Text = Null2String(rsMATREC_HIST!pono)
   txtPODate.Text = Null2String(rsMATREC_HIST!podate)
   txtDRNo.Text = Null2String(rsMATREC_HIST!drno)
   txtINVNo.Text = Null2String(rsMATREC_HIST!invno)
   txtDS1.Text = N2Str2IntZero(rsMATREC_HIST!ds1)
   txtDS_Desc1.Text = Null2String(rsMATREC_HIST!ds_desc1)
   txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsMATREC_HIST!ds_amt1))
   txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATREC_HIST!ttlrramt))
   txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATREC_HIST!netrramt))
   txtRemarks.Text = ToDoubleNumber(Null2String(rsMATREC_HIST!remarks))
   If Null2String(rsMATREC_HIST!Status) = "P" Then
      labRRsted.Visible = True
      labRRsted.Caption = "POSTED"
      cmdEdit.Enabled = False
      cmdPost.Enabled = False
      cmdPrint.Enabled = True
      If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = True
   ElseIf Null2String(rsMATREC_HIST!Status) = "C" Then
      labRRsted.Visible = True
      labRRsted.Caption = "CANCELLED"
      cmdEdit.Enabled = False
      cmdPost.Enabled = False
      'cmdUnPost.Enabled = False
      cmdPrint.Enabled = False
      cmdCancelRR.Enabled = False
   Else
      labRRsted.Visible = False
      cmdEdit.Enabled = True
      cmdPost.Enabled = True
      cmdPrint.Enabled = True
      If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = True
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
   .ColWidth(0) = 1
   .ColWidth(1) = 800
   .ColWidth(2) = 1500
   .ColWidth(3) = 4100
   .ColWidth(4) = 500
   .ColWidth(5) = 1
   .ColWidth(6) = 900
   .ColWidth(7) = 1200
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
MATREC_HIST_TOTUCOST = 0
MATREC_HIST_TOTINVAMT = 0
MATREC_HIST_TOTVAT = 0
RR_QTY_REC = 0
Set rsDAYTRAN = New ADODB.Recordset
    rsDAYTRAN.Open "select id,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt from PMIS_DayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC_HIST!rrno) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
   Screen.MousePointer = 11
   rsDAYTRAN.MoveFirst
   Do While Not rsDAYTRAN.EOF
      Pcnt = Pcnt + 1
      grdDetails.AddItem rsDAYTRAN!ID & Chr(9) & Format(Null2String(rsDAYTRAN!itemno), "0000") & Chr(9) & _
                      Null2String(rsDAYTRAN!MATCDE) & Chr(9) & _
                      Null2String(rsDAYTRAN!MatDsc) & Chr(9) & _
                      N2Str2IntZero(rsDAYTRAN!tranqty) & Chr(9) & _
                      N2Str2Zero(rsDAYTRAN!TRANINVAMT) & Chr(9) & _
                      N2Str2Zero(rsDAYTRAN!TRANUCOST) & Chr(9) & _
                      Format(N2Str2IntZero(rsDAYTRAN!tranqty) * N2Str2Zero(rsDAYTRAN!TRANUCOST), MAXIMUM_DIGIT)
      RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsDAYTRAN!tranqty)
      MATREC_HIST_TOTUCOST = MATREC_HIST_TOTUCOST + (N2Str2IntZero(rsDAYTRAN!tranqty) * N2Str2Zero(rsDAYTRAN!TRANUCOST))
      MATREC_HIST_TOTINVAMT = MATREC_HIST_TOTINVAMT + (N2Str2IntZero(rsDAYTRAN!tranqty) * N2Str2Zero(rsDAYTRAN!TRANINVAMT))
      rsDAYTRAN.MoveNext
   Loop
   If Pcnt <> 0 Then grdDetails.RemoveItem 1
   If Null2String(rsMATREC_HIST!classcode) = "PCS" Or Null2String(rsMATREC_HIST!classcode) = "PCG" Then
      If ISNONVAT = True Then
         MATREC_HIST_TOTVAT = 0
      Else
         MATREC_HIST_TOTVAT = (MATREC_HIST_TOTUCOST * ConvertToBIRDecimalFormat(VAT_RATE)) - MATREC_HIST_TOTUCOST
      End If
   Else
      MATREC_HIST_TOTVAT = 0
   End If
   If NumericVal(MATREC_HIST_TOTVAT) <> 0 Then
      txtDS1.Text = VAT_RATE
      txtDS_Desc1.Text = "VAT"
      txtDS_Amt1.Text = ToDoubleNumber(MATREC_HIST_TOTVAT)
      txtNetRRAmt.Text = ToDoubleNumber(NumericVal(txtTTLRRAmt.Text) + NumericVal(txtDS_Amt1.Text))
   Else
      txtDS1.Text = 0
      txtDS_Desc1.Text = ""
      txtDS_Amt1.Text = 0
      txtNetRRAmt.Text = ToDoubleNumber(NumericVal(txtTTLRRAmt.Text))
   End If
   Screen.MousePointer = 0
Else
   cleargrid grdDetails
End If
Exit Sub

ErrorCode:
ShowVBError
Exit Sub
End Sub

Function Setmatdsc(ppp As String)
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select matcde,matdsc from MatMas where matcde= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   Setmatdsc = Null2String(rsMatMas!MatDsc)
End If
End Function

Function Setmatdsc2(ppp As String)
If ppp <> "" Then
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matdsc from MatMas where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   Setmatdsc2 = Null2String(rsMatMas!MatDsc)
End If
End If
End Function

Function Setmatcde(DDD As String)
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matcde from MatMas where id = " & DDD, gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   Setmatcde = Null2String(rsMatMas!MATCDE)
End If
End Function

Function SetMatIDmatcde(DDD As String)
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matcde from MatMas where matcde = '" & DDD & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   SetMatIDmatcde = Null2String(rsMatMas!ID)
End If
End Function

Function SetMatIDDesc(DDD As String)
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matdsc from MatMas where ltrim(rtrim(matdsc)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   SetMatIDDesc = Null2String(rsMatMas!ID)
End If
End Function

Function SetPartPrice(ppp As String)
If ppp <> "" Then
   Set rsMatMas = New ADODB.Recordset
       rsMatMas.Open "Select matcde,mac from PMIS_STOCKMAS where matcde = '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
   If Not rsMatMas.EOF And Not rsMatMas.BOF Then
      SetPartPrice = Null2String(rsMatMas!Mac)
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
   ISNONVAT = False
   txtDS1.Text = ""
   txtDetails.Text = ""
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
   ISNONVAT = False
   txtDS1.Text = ""
   txtDetails.Text = ""
End If
End Function

Sub InitMaterials()
txtTranItemNo.Text = Format(Pcnt + 1, "0000")
cboTranMatCde.Text = ""
cboTranMatDsc.Text = ""
txtTranQty.Text = 1
txtTranINVAmt.Text = ZERO
txtTranTotalAmt.Text = ZERO
End Sub

Function StorePartsEntry(ByVal ID As Variant)
Set rsDAYTRAN = New ADODB.Recordset
    rsDAYTRAN.Open "select id,itemno,matcde,matdsc,tranqty,traninvamt,tranucost from PMIS_DayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsDAYTRAN.EOF And Not rsDAYTRAN.BOF Then
   labDetID.Caption = rsDAYTRAN!ID
   txtTranItemNo.Text = Null2String(rsDAYTRAN!itemno)
   cboTranMatCde.Text = Null2String(rsDAYTRAN!MATCDE)
   cboTranMatDsc.Text = Null2String(rsDAYTRAN!MatDsc)
   txtTranQty.Text = N2Str2IntZero(rsDAYTRAN!tranqty)
   txtTranINVAmt.Text = ToDoubleNumber(N2Str2Zero(rsDAYTRAN!TRANINVAMT))
   txtUnitCost.Text = ToDoubleNumber(N2Str2Zero(rsDAYTRAN!TRANUCOST))
   txtTranTotalAmt.Text = ToDoubleNumber(N2Str2IntZero(rsDAYTRAN!tranqty) * N2Str2Zero(rsDAYTRAN!TRANINVAMT))
End If
End Function

Private Sub grdDetails_DblClick()
Dim Fild As String
If Null2String(rsMATREC_HIST!Status) = "P" Then
   MsgSpeechBox "Item(s) are Already Posted, and cannot be change..."
ElseIf Null2String(rsMATREC_HIST!Status) = "C" Then
   MsgSpeechBox "Item(s) are Already Cancelled, and cannot be edited"
Else
   grdDetails.Row = grdDetails.Row
   grdDetails.Col = 0
   Fild = grdDetails.Text
   If Fild <> "" And Fild <> "No Entry" Then
      AddorEdit = "EDIT"
      cmdTranDelete.Visible = True
      BringToFront
      fraAddTran.Caption = "Edit Materials"
      StorePartsEntry (Fild)
   Else
      MsgSpeechBox "No Entry on Materials"
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
Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "select matcde,matdsc from CSMS_MatMas", gconDMIS, adOpenForwardOnly, adLockReadOnly
If Not rsMatMas.EOF And Not rsMatMas.BOF Then
   rsMatMas.MoveFirst
   cboTranMatCde.Clear
   cboTranMatDsc.Clear
   Do While Not rsMatMas.EOF
      cboTranMatCde.AddItem Null2String(rsMatMas!MATCDE)
      cboTranMatDsc.AddItem Null2String(rsMatMas!MatDsc)
      rsMatMas.MoveNext
   Loop
End If
End Sub

Sub InitCboClasscode()
cboClasscode.Clear
cboClasscode.AddItem "IBT"
cboClasscode.AddItem "PCG"
cboClasscode.AddItem "PCS"
cboClasscode.AddItem "RCG"
cboClasscode.AddItem "REP"
cboClasscode.AddItem "RRV"
cboClasscode.Text = "PCG"
End Sub

Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cboclasscode_LostFocus()
If cboClasscode.Text <> "" Then
   cboClasscode.Text = cboClasscode.Text
Else
   MsgBoxXP "Invalid code. Please enter one of the following codes... " & vbCrLf & _
            "IBT, PCG, PCS, RCG, RCS, REP, RRV", "Error Encountered", XP_OKOnly, msg_Critical
End If
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
      Dim rsMATREC_HISTDup As ADODB.Recordset
      Set rsMATREC_HISTDup = New ADODB.Recordset
          rsMATREC_HISTDup.Open "select pono from MATIss where pono = '" & txtPONo.Text & "'", gconDMIS
      If Not rsMATREC_HISTDup.EOF And Not rsMATREC_HISTDup.BOF Then
         MsgSpeechBox "PO Number Already Received"
         Exit Sub
      End If
      Set rsPO_HD = New ADODB.Recordset
          rsPO_HD.Open "select pono,supcode,podate from PMIS_PO_Hd where pono = '" & txtPONo.Text & "'", gconDMIS
      If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
         txtRecvd_Code.Text = Null2String(rsPO_HD!SupCode)
         txtPODate.Text = Null2String(rsPO_HD!podate)
         Pcnt = 0
         MATREC_HIST_TOTUCOST = 0
         MATREC_HIST_TOTINVAMT = 0
         MATREC_HIST_TOTVAT = 0
         RR_QTY_REC = 0
         Dim rsDAYTRANDup As ADODB.Recordset
         Set rsDAYTRANDup = New ADODB.Recordset
             rsDAYTRANDup.Open "select id,itemno,matcde,matdsc,tranqty,traninvamt,tranucost from PMIS_DayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_HD!pono) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
         If Not rsDAYTRANDup.EOF And Not rsDAYTRANDup.BOF Then
            Screen.MousePointer = 11
            rsDAYTRANDup.MoveFirst
            cleargrid grdDetails
            Do While Not rsDAYTRANDup.EOF
               Pcnt = Pcnt + 1
               grdDetails.AddItem rsDAYTRANDup!ID & Chr(9) & Null2String(Format(rsDAYTRANDup!itemno, "0000")) & Chr(9) & _
                                  Null2String(rsDAYTRANDup!MATCDE) & Chr(9) & _
                                  Setmatdsc(Null2String(rsDAYTRANDup!MatDsc)) & Chr(9) & _
                                  N2Str2IntZero(rsDAYTRANDup!tranqty) & Chr(9) & _
                                  N2Str2Zero(rsDAYTRANDup!TRANINVAMT) & Chr(9) & _
                                  N2Str2Zero(rsDAYTRANDup!TRANUCOST) & Chr(9) & _
                                  N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANUCOST)
               MATREC_HIST_TOTUCOST = MATREC_HIST_TOTUCOST + (N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANUCOST))
               MATREC_HIST_TOTINVAMT = MATREC_HIST_TOTINVAMT + (N2Str2IntZero(rsDAYTRANDup!tranqty) * N2Str2Zero(rsDAYTRANDup!TRANINVAMT))
               rsDAYTRANDup.MoveNext
            Loop
            If Pcnt <> 0 Then grdDetails.RemoveItem 1
            Screen.MousePointer = 0
         Else
            cleargrid grdDetails
         End If
      Else
         MsgSpeechBox "Invalid PO Number!"
         txtPONo.Text = "": txtPODate.Text = ""
         If AddorEdit = "ADD" Then cleargrid grdDetails
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
If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txttranQty_Change()
If txtTranQty.Text <> "" Then
   If ISNONVAT = True Then
      txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
   Else
      txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
   End If
   txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
End If
End Sub

Private Sub txtTranQty_GotFocus()
If NumericVal(txtTranQty.Text) = 1 Then txtTranQty.Text = ""
End Sub

Private Sub txtTranQty_LostFocus()
If txtTranQty.Text <> "" Then
   If ISNONVAT = True Then
      txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text))
   Else
      txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
   End If
   txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
Else
   txtTranQty.Text = 1
   txtUnitCost.Text = ToDoubleNumber(NumericVal(txtTranINVAmt.Text) / ConvertToBIRDecimalFormat(VAT_RATE))
   txtTranTotalAmt.Text = ToDoubleNumber(NumericVal(txtTranQty.Text) * NumericVal(txtTranINVAmt.Text))
End If
End Sub

Private Sub txtTranINVAmt_Change()
On Error Resume Next
If Null2String(rsMATREC_HIST!classcode) = "PCS" Or Null2String(rsMATREC_HIST!classcode) = "PCG" Then
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
End Sub

Private Sub txtUnitPrice_LostFocus()
If Null2String(rsMATREC_HIST!rrno) = "PCS" Or Null2String(rsMATREC_HIST!rrno) = "PCG" Then
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

'SEARCH MODULE
Private Sub lstMATREC_GotFocus()
rsMATREC_HIST.Bookmark = rsFind(rsMATREC_HIST.Clone, "tranno", lstMATREC.SelectedItem).Bookmark
StoreMemVars
End Sub

Private Sub lstMATREC_ItemClick(ByVal Item As MSComctlLib.ListItem)
If optRRNo.Value = True Then
   rsMATREC_HIST.Bookmark = rsFind(rsMATREC_HIST.Clone, "rrno", Item).Bookmark
Else
   rsMATREC_HIST.Bookmark = rsFind(rsMATREC_HIST.Clone, "ID", lstMATREC.SelectedItem.SubItems(1)).Bookmark
End If
StoreMemVars
End Sub

Private Sub lstMATREC_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With lstMATREC
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

Private Sub lstMATREC_DblClick()
cmdEdit.Value = True
End Sub

Private Sub lstMATREC_KeyDown(KeyCode As Integer, Shift As Integer)
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
Dim rsMATREC_HIST As ADODB.Recordset
lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
Set rsMATREC_HIST = New ADODB.Recordset
Set rsMATREC_HIST = gconDMIS.Execute("select rrno,ID from CSMS_MATREC_HIST order by rrno asc")
If Not (rsMATREC_HIST.EOF And rsMATREC_HIST.BOF) Then
   lstMATREC.Enabled = True
   Listview_Loadval Me.lstMATREC.ListItems, rsMATREC_HIST
   lstMATREC.Refresh
Else
   lstMATREC.Enabled = False
End If
End Sub

Sub FillSearchGrid(XXX As String)
Dim rsMATREC_HIST As ADODB.Recordset
lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
Set rsMATREC_HIST = New ADODB.Recordset
Set rsMATREC_HIST = gconDMIS.Execute("select rrno, ID from MATREC_HIST where rrno like'" & XXX & "%'")
If Not (rsMATREC_HIST.EOF And rsMATREC_HIST.BOF) Then
   lstMATREC.Enabled = True
   Listview_Loadval Me.lstMATREC.ListItems, rsMATREC_HIST
   lstMATREC.Refresh
Else
   lstMATREC.Enabled = False
End If
End Sub

Sub FillGrid2()
Dim rsMATREC_HIST As ADODB.Recordset
lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
Set rsMATREC_HIST = New ADODB.Recordset
Set rsMATREC_HIST = gconDMIS.Execute("select recvd_from, ID from MATREC_HIST order by rrno asc")
If Not (rsMATREC_HIST.EOF And rsMATREC_HIST.BOF) Then
   lstMATREC.Enabled = True
   Listview_Loadval Me.lstMATREC.ListItems, rsMATREC_HIST
   lstMATREC.Refresh
Else
   lstMATREC.Enabled = False
End If
End Sub

Sub FillSearchGrid2(XXX As String)
Dim rsMATREC_HIST As ADODB.Recordset
lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
Set rsMATREC_HIST = New ADODB.Recordset
Set rsMATREC_HIST = gconDMIS.Execute("select recvd_from, ID from MATREC_HIST where recvd_from like '" & XXX & "%' order by rrno asc")
If Not (rsMATREC_HIST.EOF And rsMATREC_HIST.BOF) Then
   lstMATREC.Enabled = True
   Listview_Loadval Me.lstMATREC.ListItems, rsMATREC_HIST
   lstMATREC.Refresh
Else
   lstMATREC.Enabled = False
End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then lstMATREC.SetFocus
End Sub

Private Sub optRONo_Click()
lstMATREC.ColumnHeaders(1).Text = "Sup. Name"
lstMATREC.ColumnHeaders(1).Width = 4000
If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
textSearch.SetFocus
End Sub

Private Sub optRRNo_Click()
lstMATREC.ColumnHeaders(1).Text = "Tran. No."
lstMATREC.ColumnHeaders(1).Width = 2150
If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
textSearch.SetFocus
End Sub
