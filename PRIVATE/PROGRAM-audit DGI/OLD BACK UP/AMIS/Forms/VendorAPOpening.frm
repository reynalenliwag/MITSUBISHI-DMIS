VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "wizMACBut.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmAMISVendorAPOpening 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   6405
   ClientLeft      =   10305
   ClientTop       =   3990
   ClientWidth     =   9615
   ForeColor       =   &H00FFFFFF&
   Icon            =   "VendorAPOpening.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   9615
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   60
      ScaleHeight     =   3135
      ScaleWidth      =   9690
      TabIndex        =   1
      Top             =   195
      Width           =   9690
      Begin VB.PictureBox picNameofVendor 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2490
         ScaleHeight     =   405
         ScaleWidth      =   4095
         TabIndex        =   208
         Top             =   450
         Width           =   4095
         Begin VB.TextBox txtNameofVendor 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   210
            Text            =   "000226"
            Top             =   0
            Width           =   3705
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   209
            Top             =   30
            Width           =   345
         End
      End
      Begin VB.ComboBox cboCOBAcctName 
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
         Left            =   1530
         TabIndex        =   205
         Top             =   2670
         Width           =   6045
      End
      Begin VB.TextBox txtCode 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   7
         Text            =   "000226"
         Top             =   460
         Width           =   1005
      End
      Begin VB.TextBox txtJDate 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   60
         Width           =   1545
      End
      Begin VB.TextBox txtVoucherNo 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "000226"
         Top             =   60
         Width           =   1005
      End
      Begin VB.ComboBox cboNameofVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   360
         Left            =   2490
         TabIndex        =   9
         Text            =   "cboRecvd_Desc"
         Top             =   450
         Width           =   4080
      End
      Begin VB.TextBox txtDueDate 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox txtJNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2EEE9&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7950
         MaxLength       =   6
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cboATCTAXRATE 
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
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   1440
         TabIndex        =   13
         Text            =   "cboATCTAXRATE"
         Top             =   810
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.PictureBox picReceivable 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   120
         ScaleHeight     =   2085
         ScaleWidth      =   9585
         TabIndex        =   41
         Top             =   4350
         Width           =   9585
         Begin VB.TextBox txtInvoiceNo 
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   53
            Text            =   "000000"
            Top             =   930
            Width           =   1485
         End
         Begin VB.CheckBox chkNonVat 
            BackColor       =   &H00DEDFDE&
            Caption         =   "Non-Vat"
            Height          =   285
            Left            =   1140
            TabIndex        =   52
            Top             =   930
            Width           =   915
         End
         Begin VB.ComboBox cboBankName2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   4560
            TabIndex        =   59
            Text            =   "Invoice Type"
            Top             =   930
            Width           =   4920
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
            Left            =   4560
            MaxLength       =   5
            TabIndex        =   57
            Top             =   930
            Width           =   2085
         End
         Begin VB.TextBox txtDealer 
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
            Left            =   7710
            MaxLength       =   50
            TabIndex        =   58
            Top             =   930
            Width           =   1755
         End
         Begin RichTextLib.RichTextBox txtRemarks2 
            Height          =   705
            Left            =   4560
            TabIndex        =   63
            Top             =   1350
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   1244
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"VendorAPOpening.frx":08CA
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtRefDate 
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
            Left            =   7710
            MaxLength       =   10
            TabIndex        =   50
            Top             =   540
            Width           =   1755
         End
         Begin VB.TextBox txtRefNo 
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
            Left            =   4560
            MaxLength       =   10
            TabIndex        =   49
            Top             =   540
            Width           =   2085
         End
         Begin VB.ComboBox cboInvoiceType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   1530
            TabIndex        =   46
            Text            =   "Invoice Type"
            Top             =   510
            Width           =   1500
         End
         Begin VB.ComboBox cboCustName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   2520
            TabIndex        =   42
            Text            =   "Customer Name"
            Top             =   30
            Width           =   4080
         End
         Begin VB.TextBox txtInvoiceAmt 
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
            Left            =   1530
            MaxLength       =   15
            TabIndex        =   65
            Text            =   "0.00"
            Top             =   1710
            Width           =   1485
         End
         Begin VB.TextBox txtInvoiceDate2 
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
            Left            =   1530
            MaxLength       =   10
            TabIndex        =   61
            Text            =   "88/88/8888"
            Top             =   1320
            Width           =   1485
         End
         Begin VB.TextBox txtCustCode 
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
            Height          =   345
            Left            =   1440
            MaxLength       =   6
            TabIndex        =   43
            Text            =   "000226"
            Top             =   45
            Width           =   1005
         End
         Begin VB.Line Line4 
            BorderWidth     =   2
            X1              =   0
            X2              =   9570
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label labTerms 
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
            Left            =   3180
            TabIndex        =   55
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label labDealer 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer"
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
            TabIndex        =   60
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label RefCRJ 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ref CRJ# 000000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   6720
            TabIndex        =   45
            Top             =   60
            Width           =   2775
         End
         Begin VB.Label labBankName 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
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
            Left            =   3180
            TabIndex        =   56
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label labRefDate 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ref. Date"
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
            TabIndex        =   51
            Top             =   570
            Width           =   1335
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   3090
            X2              =   3090
            Y1              =   480
            Y2              =   2070
         End
         Begin VB.Label labRefNo 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Reference No."
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
            Left            =   3180
            TabIndex        =   48
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label labType 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type"
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
            Left            =   150
            TabIndex        =   47
            Top             =   570
            Width           =   1425
         End
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   6660
            X2              =   6660
            Y1              =   480
            Y2              =   0
         End
         Begin VB.Label Label32 
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
            Left            =   180
            TabIndex        =   44
            Top             =   60
            Width           =   1935
         End
         Begin VB.Label labParticulars 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   3180
            TabIndex        =   64
            Top             =   1350
            Width           =   1695
         End
         Begin VB.Label labAmt 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Amount"
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
            Left            =   150
            TabIndex        =   66
            Top             =   1740
            Width           =   1425
         End
         Begin VB.Label labDate 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. Date"
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
            Left            =   150
            TabIndex        =   62
            Top             =   1350
            Width           =   1425
         End
         Begin VB.Label LabNo 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. No."
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
            Left            =   150
            TabIndex        =   54
            Top             =   960
            Width           =   1425
         End
      End
      Begin VB.PictureBox picDisbursement 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   120
         ScaleHeight     =   1245
         ScaleWidth      =   9525
         TabIndex        =   30
         Top             =   3660
         Width           =   9525
         Begin VB.TextBox txtCheckDate 
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
            MaxLength       =   10
            TabIndex        =   40
            Text            =   "000226"
            Top             =   810
            Width           =   1815
         End
         Begin VB.TextBox txtCheckNo 
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
            MaxLength       =   6
            TabIndex        =   35
            Text            =   "000226"
            Top             =   420
            Width           =   1815
         End
         Begin VB.ComboBox cboBankName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   360
            Left            =   4380
            TabIndex        =   33
            Text            =   "cboRecvd_Desc"
            Top             =   30
            Width           =   5070
         End
         Begin VB.TextBox txtBankCode 
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
            MaxLength       =   8
            TabIndex        =   31
            Text            =   "000226"
            Top             =   30
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtParticulars 
            Height          =   735
            Left            =   4380
            TabIndex        =   38
            Top             =   420
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1296
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"VendorAPOpening.frx":0961
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            Caption         =   "Bank Name"
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
            Left            =   3270
            TabIndex        =   34
            Top             =   90
            Width           =   1935
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   3270
            TabIndex        =   39
            Top             =   450
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
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
            TabIndex        =   37
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label12 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check No."
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
            TabIndex        =   36
            Top             =   450
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Code"
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
            TabIndex        =   32
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.PictureBox picPayables 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   30
         ScaleHeight     =   1275
         ScaleWidth      =   9465
         TabIndex        =   18
         Top             =   1230
         Width           =   9465
         Begin VB.TextBox txtTaxBaseAmount 
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
            ForeColor       =   &H00701E2A&
            Height          =   315
            Left            =   6210
            MaxLength       =   15
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   60
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox txtPayCode 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1410
            MaxLength       =   4
            TabIndex        =   19
            Text            =   "000226"
            Top             =   60
            Width           =   645
         End
         Begin VB.TextBox txtAmountToPay 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1410
            MaxLength       =   15
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtInvoiceDate 
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
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   1410
            MaxLength       =   10
            TabIndex        =   25
            Text            =   "88/88/8888"
            Top             =   450
            Width           =   1695
         End
         Begin VB.ComboBox cboPayType 
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
            Height          =   330
            Left            =   2100
            TabIndex        =   21
            Text            =   "cboPayType"
            Top             =   60
            Width           =   2205
         End
         Begin RichTextLib.RichTextBox txtRemarks 
            Height          =   765
            Left            =   4320
            TabIndex        =   27
            Top             =   420
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   1349
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            TextRTF         =   $"VendorAPOpening.frx":09F5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   4350
            TabIndex        =   22
            Top             =   90
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount to Pay"
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
            Left            =   0
            TabIndex        =   26
            Top             =   810
            Width           =   1845
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Date"
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
            Left            =   180
            TabIndex        =   24
            Top             =   450
            Width           =   1185
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Type"
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
            Left            =   30
            TabIndex        =   20
            Top             =   90
            Width           =   1725
         End
      End
      Begin RichTextLib.RichTextBox txtCOBAcctNo 
         Height          =   315
         Left            =   7620
         TabIndex        =   206
         Top             =   2670
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393217
         BackColor       =   16777215
         MultiLine       =   0   'False
         TextRTF         =   $"VendorAPOpening.frx":0A8C
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
      Begin VB.Label Label44 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Name."
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
         TabIndex        =   207
         Top             =   2700
         Width           =   1545
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   9570
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label txtAddress 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Supplier Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   90
         TabIndex        =   14
         Top             =   840
         Width           =   6465
      End
      Begin VB.Label labPosted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "*** POSTED ***"
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
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   4065
      End
      Begin VB.Label labSupplierPayTo 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Code"
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
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label labDueDate 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Due Date"
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
         TabIndex        =   12
         Top             =   510
         Width           =   1185
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "TAX"
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
         Left            =   9150
         TabIndex        =   29
         Top             =   2280
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
         Height          =   285
         Left            =   4110
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal Date"
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
         Left            =   6690
         TabIndex        =   6
         Top             =   90
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher No."
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
         TabIndex        =   3
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Journal No."
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
         Left            =   6810
         TabIndex        =   17
         Top             =   870
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ATC TAX RATE"
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
         Height          =   285
         Left            =   60
         TabIndex        =   15
         Top             =   870
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin TabDlg.SSTab JournalTAB 
      Height          =   2775
      Left            =   60
      TabIndex        =   129
      Top             =   2730
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   4895
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "[<F3> Add &Journals]   [<Ctrl> + <J> View &Journals]   "
      TabPicture(0)   =   "VendorAPOpening.frx":0B1F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDetails"
      Tab(0).Control(1)=   "fraAddJournal"
      Tab(0).Control(2)=   "cmdAddJournal"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[<F4> Add &Details]   [<Ctrl> + <D> View &Details]   "
      TabPicture(1)   =   "VendorAPOpening.frx":0B3B
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "picPV_Entry"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdPV_Entry"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "picPV_Detail"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.PictureBox picPV_Detail 
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   -60
         ScaleHeight     =   2415
         ScaleWidth      =   9615
         TabIndex        =   166
         Top             =   420
         Width           =   9615
         Begin MSComctlLib.ListView lstPV_Detail 
            Height          =   1755
            Left            =   60
            TabIndex        =   167
            Top             =   150
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   3096
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
            MouseIcon       =   "VendorAPOpening.frx":0B57
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEM #"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PO NUMBER"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "MRR NUMBER"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "INVOICE NO."
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "PRODUCT NO."
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "AMOUNT"
               Object.Width           =   2823
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
         Begin MSMask.MaskEdBox txtTotalPV_Amount 
            Height          =   345
            Left            =   8100
            TabIndex        =   168
            Top             =   1920
            Width           =   1350
            _ExtentX        =   2381
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
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   7500
            TabIndex        =   169
            Top             =   1980
            Width           =   1275
         End
      End
      Begin VB.PictureBox fraDetails 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   -74955
         ScaleHeight     =   2295
         ScaleWidth      =   9405
         TabIndex        =   130
         Top             =   90
         Width           =   9405
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   30
            Top             =   2280
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Enabled         =   0   'False
            Height          =   405
            Left            =   0
            TabIndex        =   132
            Top             =   1830
            Width           =   9135
            Begin VB.PictureBox picChat 
               BorderStyle     =   0  'None
               Height          =   345
               Left            =   60
               ScaleHeight     =   345
               ScaleWidth      =   5895
               TabIndex        =   133
               Top             =   30
               Visible         =   0   'False
               Width           =   5895
               Begin VB.Label Label40 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Warning: Sales Details Amount is not Balance with Journal Details Amount"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   60
                  TabIndex        =   134
                  Top             =   30
                  Width           =   5685
               End
            End
            Begin VB.TextBox txtOutBalance 
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
               Left            =   1320
               MaxLength       =   14
               TabIndex        =   136
               Text            =   "Text1"
               Top             =   30
               Width           =   1515
            End
            Begin VB.TextBox txtTotDebit 
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
               Left            =   6000
               MaxLength       =   14
               TabIndex        =   138
               Text            =   "Text1"
               Top             =   30
               Width           =   1485
            End
            Begin VB.TextBox txtTotCredit 
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
               Left            =   7470
               MaxLength       =   14
               TabIndex        =   137
               Text            =   "Text1"
               Top             =   30
               Width           =   1485
            End
            Begin VB.Label labOutBalance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Out Balance"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   0
               TabIndex        =   135
               Top             =   45
               Width           =   1275
            End
         End
         Begin MSComctlLib.ListView lstDetails 
            Height          =   1785
            Left            =   30
            TabIndex        =   131
            Top             =   30
            Width           =   9285
            _ExtentX        =   16378
            _ExtentY        =   3149
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
            MouseIcon       =   "VendorAPOpening.frx":0CB9
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ITEM #"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ACCOUNT CODE"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ACCOUNT DESCRIPTION"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "DEBIT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "CREDIT"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "ID"
               Object.Width           =   2
            EndProperty
         End
      End
      Begin VB.PictureBox fraAddJournal 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   -74865
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   140
         Top             =   600
         Width           =   9135
         Begin VB.CommandButton cmdJournalDelete 
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
            Left            =   90
            MouseIcon       =   "VendorAPOpening.frx":0E1B
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":0F6D
            Style           =   1  'Graphical
            TabIndex        =   156
            Top             =   765
            Width           =   705
         End
         Begin VB.CommandButton cmdJournalCancel 
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
            Left            =   8295
            MouseIcon       =   "VendorAPOpening.frx":1298
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":13EA
            Style           =   1  'Graphical
            TabIndex        =   165
            Top             =   765
            Width           =   705
         End
         Begin VB.CommandButton cmdJournalSave 
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
            Left            =   7560
            MouseIcon       =   "VendorAPOpening.frx":1728
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":187A
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   765
            Width           =   705
         End
         Begin VB.TextBox txtCredit 
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
            Left            =   7950
            MaxLength       =   10
            TabIndex        =   149
            Top             =   330
            Width           =   1100
         End
         Begin VB.TextBox txtDebit 
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
            Left            =   6780
            MaxLength       =   10
            TabIndex        =   147
            Top             =   330
            Width           =   1100
         End
         Begin VB.Frame fraComp 
            Height          =   915
            Left            =   2340
            TabIndex        =   157
            Top             =   660
            Width           =   4365
            Begin VB.TextBox txtNetAmt 
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
               Left            =   2910
               MaxLength       =   10
               TabIndex        =   163
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtTax 
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
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   162
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtGrossAmt 
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
               Left            =   150
               MaxLength       =   10
               TabIndex        =   161
               Top             =   510
               Width           =   1300
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Net Amount"
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
               Left            =   2910
               TabIndex        =   160
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label labTax 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Output Tax"
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
               Left            =   1560
               TabIndex        =   159
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Gross Amt."
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
               TabIndex        =   158
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   2310
            TabIndex        =   151
            Top             =   -30
            Width           =   4425
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   315
               Left            =   30
               TabIndex        =   153
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393217
               BackColor       =   16777215
               MultiLine       =   0   'False
               TextRTF         =   $"VendorAPOpening.frx":1BCA
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
            Begin VB.Label Label33 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Account Name"
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
               Left            =   60
               TabIndex        =   152
               Top             =   90
               Width           =   2205
            End
         End
         Begin VB.ComboBox cboAcct_Code 
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
            Left            =   60
            TabIndex        =   146
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin VB.TextBox txtAcctID 
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
            Left            =   840
            TabIndex        =   144
            Text            =   "Text1"
            Top             =   330
            Width           =   585
         End
         Begin VB.TextBox txtJItemNo 
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
            Height          =   255
            Left            =   690
            MaxLength       =   4
            TabIndex        =   145
            Text            =   "Text1"
            Top             =   330
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
            Left            =   390
            TabIndex        =   150
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label34 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
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
            Left            =   90
            TabIndex        =   141
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label30 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
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
            Left            =   7050
            TabIndex        =   142
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label38 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
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
            Left            =   8130
            TabIndex        =   143
            Top             =   60
            Width           =   795
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
            Left            =   6840
            TabIndex        =   148
            Top             =   360
            Width           =   855
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
            Left            =   2880
            TabIndex        =   155
            Top             =   420
            Width           =   915
         End
         Begin VB.Label labPartNo 
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
            Height          =   315
            Left            =   2340
            TabIndex        =   154
            Top             =   420
            Width           =   2685
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1785
         Left            =   -74925
         TabIndex        =   139
         Top             =   540
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3149
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
         MICON           =   "VendorAPOpening.frx":1C5D
      End
      Begin wizButton.cmd cmdPV_Entry 
         Height          =   1785
         Left            =   120
         TabIndex        =   170
         Top             =   540
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3149
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
         MICON           =   "VendorAPOpening.frx":1C79
      End
      Begin VB.PictureBox picPV_Entry 
         Height          =   1665
         Left            =   180
         ScaleHeight     =   1605
         ScaleWidth      =   9075
         TabIndex        =   171
         Top             =   600
         Width           =   9135
         Begin VB.CommandButton cmdPVDelete 
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
            Left            =   90
            MouseIcon       =   "VendorAPOpening.frx":1C95
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":1DE7
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   720
            Width           =   705
         End
         Begin VB.CommandButton cmdPVCancel 
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
            Left            =   7755
            MouseIcon       =   "VendorAPOpening.frx":2112
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":2264
            Style           =   1  'Graphical
            TabIndex        =   187
            Top             =   720
            Width           =   705
         End
         Begin VB.CommandButton cmdPVSave 
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
            Left            =   7020
            MouseIcon       =   "VendorAPOpening.frx":25A2
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":26F4
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   720
            Width           =   705
         End
         Begin MSMask.MaskEdBox txtMRR_No 
            Height          =   315
            Left            =   1950
            TabIndex        =   179
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtPVAmount 
            Height          =   315
            Left            =   7620
            TabIndex        =   183
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
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
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   315
            Left            =   7140
            TabIndex        =   186
            Top             =   780
            Width           =   435
            _ExtentX        =   767
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
         Begin MSMask.MaskEdBox txtINV_No 
            Height          =   315
            Left            =   3840
            TabIndex        =   181
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtPO_No 
            Height          =   315
            Left            =   60
            TabIndex        =   177
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtProd_No 
            Height          =   315
            Left            =   5730
            TabIndex        =   182
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   0
            MaxLength       =   20
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
         Begin MSMask.MaskEdBox txtPVItemNo 
            Height          =   315
            Left            =   1020
            TabIndex        =   178
            Top             =   330
            Width           =   735
            _ExtentX        =   1296
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
         Begin VB.Label Label18 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   7830
            TabIndex        =   176
            Top             =   60
            Width           =   795
         End
         Begin VB.Label labPV1 
            BackColor       =   &H8000000D&
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
            Height          =   225
            Left            =   90
            TabIndex        =   172
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label labPV2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MRR Number"
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
            Left            =   1980
            TabIndex        =   173
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label labPV3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Number"
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
            Left            =   3870
            TabIndex        =   174
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label labPV4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Product Number"
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
            Left            =   5730
            TabIndex        =   175
            Top             =   60
            Width           =   1875
         End
         Begin VB.Label labPVID 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "MRR Number"
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
            Left            =   2040
            TabIndex        =   180
            Top             =   390
            Width           =   1305
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   360
      ScaleHeight     =   900
      ScaleWidth      =   12585
      TabIndex        =   188
      Top             =   5490
      Width           =   12585
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
         Left            =   7815
         MouseIcon       =   "VendorAPOpening.frx":2A44
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":2B96
         Style           =   1  'Graphical
         TabIndex        =   201
         ToolTipText     =   "Exit Window"
         Top             =   45
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
         Left            =   7125
         MouseIcon       =   "VendorAPOpening.frx":2EFC
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":304E
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "Print this Record"
         Top             =   45
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
         Left            =   7125
         MouseIcon       =   "VendorAPOpening.frx":33B4
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":3506
         Style           =   1  'Graphical
         TabIndex        =   199
         ToolTipText     =   "Delete Selected Record"
         Top             =   45
         Width           =   705
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
         Left            =   6435
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VendorAPOpening.frx":3831
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":3983
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Cancel this Transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdUnPost 
         Caption         =   "Unpost Transaction"
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
         Left            =   5745
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VendorAPOpening.frx":3CBD
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":3E0F
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Unpost this Transaction"
         Top             =   45
         Width           =   705
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
         Left            =   5055
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "VendorAPOpening.frx":4154
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":42A6
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Post this Transaction"
         Top             =   45
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
         Left            =   4365
         MouseIcon       =   "VendorAPOpening.frx":45CB
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":471D
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
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
         Left            =   3675
         MouseIcon       =   "VendorAPOpening.frx":4A79
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":4BCB
         Style           =   1  'Graphical
         TabIndex        =   194
         ToolTipText     =   "Add Record"
         Top             =   45
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
         Left            =   2985
         MouseIcon       =   "VendorAPOpening.frx":4EDE
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":5030
         Style           =   1  'Graphical
         TabIndex        =   193
         ToolTipText     =   "Move to Last Record"
         Top             =   45
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
         Left            =   2295
         MouseIcon       =   "VendorAPOpening.frx":5380
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":54D2
         Style           =   1  'Graphical
         TabIndex        =   192
         ToolTipText     =   "Move to First Record"
         Top             =   45
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
         Left            =   1605
         MouseIcon       =   "VendorAPOpening.frx":5830
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":5982
         Style           =   1  'Graphical
         TabIndex        =   191
         ToolTipText     =   "Find a Record"
         Top             =   45
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
         Left            =   915
         MouseIcon       =   "VendorAPOpening.frx":5C7C
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":5DCE
         Style           =   1  'Graphical
         TabIndex        =   190
         ToolTipText     =   "Move to Next Record"
         Top             =   45
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
         Left            =   225
         MouseIcon       =   "VendorAPOpening.frx":6126
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":6278
         Style           =   1  'Graphical
         TabIndex        =   189
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   8175
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   202
      Top             =   5505
      Width           =   1980
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
         MouseIcon       =   "VendorAPOpening.frx":65D7
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":6729
         Style           =   1  'Graphical
         TabIndex        =   204
         ToolTipText     =   "Cancel"
         Top             =   30
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
         MouseIcon       =   "VendorAPOpening.frx":6A67
         MousePointer    =   99  'Custom
         Picture         =   "VendorAPOpening.frx":6BB9
         Style           =   1  'Graphical
         TabIndex        =   203
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picRefCDJ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6790
      ScaleHeight     =   345
      ScaleWidth      =   2775
      TabIndex        =   105
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label RefCDJ 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ref CDJ# 000000"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   2775
      End
   End
   Begin Crystal.CrystalReport rptAP 
      Left            =   9090
      Top             =   5460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Accounts Payable Printout"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin wizButton.cmd cmdShowPostRange 
      Height          =   2175
      Left            =   3540
      TabIndex        =   111
      Top             =   1710
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3836
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
      MICON           =   "VendorAPOpening.frx":6F09
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H00FFFFFF&
      Height          =   2265
      Left            =   3570
      ScaleHeight     =   2205
      ScaleWidth      =   2535
      TabIndex        =   121
      Top             =   1920
      Width           =   2595
      Begin VB.PictureBox picPrintCheck 
         BackColor       =   &H00DEDFDE&
         Enabled         =   0   'False
         Height          =   885
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   2355
         TabIndex        =   123
         Top             =   450
         Width           =   2415
         Begin VB.OptionButton optSECBANK 
            BackColor       =   &H00DEDFDE&
            Caption         =   "SECURITY BANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   124
            Top             =   -30
            Value           =   -1  'True
            Width           =   2355
         End
         Begin VB.OptionButton optPRUDBANK 
            BackColor       =   &H00DEDFDE&
            Caption         =   "PRUDENTIAL BANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   125
            Top             =   240
            Width           =   2355
         End
         Begin VB.OptionButton optCHINBANK 
            BackColor       =   &H00DEDFDE&
            Caption         =   "CHINABANK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   60
            TabIndex        =   126
            Top             =   510
            Width           =   2355
         End
      End
      Begin VB.OptionButton optPrintVoucher 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRINT VOUCHER"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   1380
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton optPrintCheck 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PRINT CHECK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   60
         Width           =   2415
      End
      Begin wizMacBut.MacBut cmdOkPrint 
         Height          =   345
         Left            =   360
         TabIndex        =   128
         Top             =   1800
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "OK"
         Caption_Xpos    =   700
      End
   End
   Begin VB.PictureBox picShowPostRange 
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   3600
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   112
      Top             =   1770
      Width           =   2535
      Begin wizProgBar.Prg prgPostRange 
         Height          =   285
         Left            =   90
         TabIndex        =   119
         Top             =   1650
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         Picture         =   "VendorAPOpening.frx":6F25
         BarPicture      =   "VendorAPOpening.frx":6F41
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin wizMacBut.MacBut cmdPostRange 
         Height          =   345
         Left            =   390
         TabIndex        =   118
         Top             =   1230
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "    POST"
      End
      Begin VB.TextBox txtToVNo 
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
         Left            =   870
         MaxLength       =   10
         TabIndex        =   117
         Top             =   780
         Width           =   1485
      End
      Begin VB.TextBox txtFromVNo 
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
         Left            =   870
         MaxLength       =   10
         TabIndex        =   115
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Post By Range"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   30
         TabIndex        =   113
         Top             =   30
         Width           =   2415
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To     :"
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
         TabIndex        =   116
         Top             =   810
         Width           =   735
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
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
         TabIndex        =   114
         Top             =   420
         Width           =   735
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   4665
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   8229
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
      MICON           =   "VendorAPOpening.frx":6F5D
   End
   Begin VB.Frame fraFindAccount 
      Caption         =   "Chart of Accounts"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   90
      TabIndex        =   67
      Top             =   270
      Width           =   9375
      Begin VB.TextBox txtSearch 
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
         MaxLength       =   50
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   270
         Width           =   9195
      End
      Begin VB.CommandButton cmdAddAccount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add Account"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   5850
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   3960
         Visible         =   0   'False
         Width           =   45
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   3975
         Left            =   60
         TabIndex        =   70
         Top             =   630
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   7011
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
         MouseIcon       =   "VendorAPOpening.frx":6F79
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   11819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "TYPE"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Press <F9>  to Add Account Entries From Template"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   60
         TabIndex        =   73
         Top             =   4950
         Width           =   9225
      End
      Begin VB.Label labAccountCode 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   285
         Left            =   90
         TabIndex        =   69
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "[Press <Enter> to Accept]      [Press <Ctrl> + <A> to Add Account]       [<F8> Change Search]"
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
         Height          =   285
         Left            =   75
         TabIndex        =   72
         Top             =   4650
         Width           =   9225
      End
   End
   Begin VB.CommandButton cmdPrinting 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Command1"
      Height          =   2445
      Left            =   3480
      TabIndex        =   120
      Top             =   1680
      Width           =   2775
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1200
      TabIndex        =   104
      Top             =   930
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7488
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
      MICON           =   "VendorAPOpening.frx":70DB
   End
   Begin VB.PictureBox picGJ 
      Height          =   4875
      Left            =   90
      ScaleHeight     =   4815
      ScaleWidth      =   9495
      TabIndex        =   74
      Top             =   315
      Width           =   9555
      Begin MSComctlLib.ListView lstGJ 
         Height          =   3315
         Left            =   60
         TabIndex        =   77
         Top             =   1080
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   5847
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "VendorAPOpening.frx":70F7
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ITEM #"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ACCOUNT CODE"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ACCOUNT DESCRIPTION"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "DEBIT"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "CREDIT"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         Height          =   405
         Left            =   90
         TabIndex        =   99
         Top             =   4410
         Width           =   9135
         Begin VB.TextBox txtGJTotCredit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   7440
            MaxLength       =   14
            TabIndex        =   102
            Text            =   "Text1"
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtGJTotDebit 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   5940
            MaxLength       =   14
            TabIndex        =   103
            Text            =   "Text1"
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtGJOutBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1320
            MaxLength       =   14
            TabIndex        =   101
            Text            =   "Text1"
            Top             =   30
            Width           =   1515
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Out Balance"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   0
            TabIndex        =   100
            Top             =   60
            Width           =   1275
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   90
         Top             =   3870
      End
      Begin VB.PictureBox picGJEntry 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   120
         ScaleHeight     =   1635
         ScaleWidth      =   9225
         TabIndex        =   79
         Top             =   2640
         Width           =   9255
         Begin VB.CommandButton cmdGJCancel 
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
            Left            =   8160
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VendorAPOpening.frx":7259
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":73AB
            Style           =   1  'Graphical
            TabIndex        =   97
            Top             =   690
            Width           =   1005
         End
         Begin VB.CommandButton cmdGJSave 
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
            Left            =   7140
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VendorAPOpening.frx":76BD
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":780F
            Style           =   1  'Graphical
            TabIndex        =   94
            Top             =   690
            Width           =   975
         End
         Begin VB.CommandButton cmdGJDelete 
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
            Left            =   90
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VendorAPOpening.frx":7C51
            MousePointer    =   99  'Custom
            Picture         =   "VendorAPOpening.frx":7DA3
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   720
            Width           =   1005
         End
         Begin VB.ComboBox cboGJAccountNo 
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
            Left            =   60
            TabIndex        =   84
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin RichTextLib.RichTextBox txtGJAccountName 
            Height          =   315
            Left            =   2340
            TabIndex        =   85
            Top             =   330
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   16777215
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"VendorAPOpening.frx":80AD
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
         Begin MSMask.MaskEdBox txtGJDebit 
            Height          =   315
            Left            =   6690
            TabIndex        =   87
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox txtGJCredit 
            Height          =   315
            Left            =   7950
            TabIndex        =   89
            Top             =   330
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   315
            Left            =   7140
            TabIndex        =   96
            Top             =   780
            Width           =   435
            _ExtentX        =   767
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
         Begin VB.TextBox txtGJItemNo 
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
            Left            =   2580
            MaxLength       =   4
            TabIndex        =   86
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin RichTextLib.RichTextBox txtGJAccountParticulars 
            Height          =   885
            Left            =   2340
            TabIndex        =   93
            Top             =   690
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1561
            _Version        =   393217
            BackColor       =   16777215
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"VendorAPOpening.frx":8140
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label labGJID 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Left            =   1890
            TabIndex        =   98
            Top             =   1620
            Width           =   2205
         End
         Begin VB.Label Label29 
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
            Height          =   315
            Left            =   2340
            TabIndex        =   91
            Top             =   420
            Width           =   2685
         End
         Begin VB.Label Label28 
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
            Left            =   2880
            TabIndex        =   92
            Top             =   420
            Width           =   915
         End
         Begin VB.Label Label27 
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
            Left            =   6840
            TabIndex        =   88
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label26 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Credit"
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
            Left            =   7950
            TabIndex        =   83
            Top             =   60
            Width           =   795
         End
         Begin VB.Label Label25 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Debit"
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
            Left            =   6720
            TabIndex        =   82
            Top             =   60
            Width           =   885
         End
         Begin VB.Label Label24 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account No."
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
            Left            =   90
            TabIndex        =   80
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label Label23 
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
            Left            =   390
            TabIndex        =   90
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label22 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            TabIndex        =   81
            Top             =   60
            Width           =   2205
         End
      End
      Begin wizButton.cmd cmdGJEntry 
         Height          =   1785
         Left            =   60
         TabIndex        =   78
         Top             =   2580
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3149
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
         MICON           =   "VendorAPOpening.frx":81D7
      End
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   60
         TabIndex        =   76
         Top             =   330
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"VendorAPOpening.frx":81F3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
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
         TabIndex        =   75
         Top             =   60
         Width           =   1695
      End
   End
   Begin VB.PictureBox picTemplates 
      BackColor       =   &H00FFFFFF&
      Height          =   4125
      Left            =   1260
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   107
      Top             =   990
      Width           =   7185
      Begin VB.TextBox txtSearchTemplates 
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
         MaxLength       =   50
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   109
         Top             =   450
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   5583
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
         MouseIcon       =   "VendorAPOpening.frx":828A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DESCRIPTION"
            Object.Width           =   11819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Press <Enter> to Insert Account Entries From Template"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   30
         TabIndex        =   110
         Top             =   3750
         Width           =   7035
      End
   End
End
Attribute VB_Name = "frmAMISVendorAPOpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                  As ADODB.Recordset
Dim rsJournal_Det                                 As ADODB.Recordset
Dim rsPV_Detail                                   As ADODB.Recordset
Dim rsCV_Detail                                   As ADODB.Recordset
Dim rsCRJ_Detail                                  As ADODB.Recordset
Dim rsJV_detail                                   As ADODB.Recordset
Dim rsChartAccount                                As ADODB.Recordset
Dim rsJournal_HD2                                 As ADODB.Recordset
Dim rsProfile                                     As ADODB.Recordset
Dim rsCheckJournal_HD                             As ADODB.Recordset
Attribute rsCheckJournal_HD.VB_VarUserMemId = 1073938441
Dim rsVENDOR                                      As ADODB.Recordset
Attribute rsVENDOR.VB_VarUserMemId = 1073938442
Dim rsPayTerm                                     As ADODB.Recordset
Dim rsBanks                                       As ADODB.Recordset
Dim rsCustomer                                    As ADODB.Recordset
Dim rsInvoiceType                                 As ADODB.Recordset
Dim kcnt, Jcnt                                    As Integer
Attribute kcnt.VB_VarUserMemId = 1073938447
Attribute Jcnt.VB_VarUserMemId = 1073938447
Dim AddorEdit                                     As String
Attribute AddorEdit.VB_VarUserMemId = 1073938449
Dim SearchBy                                      As String

Dim TOTDEBIT                                      As Double
Attribute TOTDEBIT.VB_VarUserMemId = 1073938451
Dim TOTCREDIT                                     As Double
Dim TOTTAX                                        As Double
Dim OUTBALANCE                                    As Double
Dim TOTAL_AR_AMOUNT                               As Double
Dim TOTALPVAMOUNT                                 As Double
Dim COMP_SJ_OUTPUT_TAX                            As Double
Dim PrevJType, PrevJNo                            As String
Attribute PrevJType.VB_VarUserMemId = 1073938458
Attribute PrevJNo.VB_VarUserMemId = 1073938458
Dim PrevInvoiceType                               As String
Attribute PrevInvoiceType.VB_VarUserMemId = 1073938460
Dim PrevInvoiceNo                                 As String
Dim PrevPV_VoucherNo                              As String
Dim xEntityClass                                  As String
Dim WithEvents frmNewEntity                       As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim xJOURNALTYPE                                  As String

Function SetCOBAcctNo(XXX As String) As String
    Dim rsCOBAcctName                             As ADODB.Recordset
    Set rsCOBAcctName = New ADODB.Recordset
    Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Description = '" & XXX & "' and Description IS NOT NULL")
    If Not rsCOBAcctName.EOF And Not rsCOBAcctName.BOF Then
        SetCOBAcctNo = Null2String(rsCOBAcctName!ACCTCODE)
    End If
End Function

Function Setacctname(VVV As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,Description from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        Setacctname = UCase(Null2String(rsChartAccount2!Description))
    Else
        Setacctname = ""
    End If
End Function

Function SetAcctType(VVV As Variant) As String
    Dim rsChartAccount2                           As ADODB.Recordset
    Set rsChartAccount2 = New ADODB.Recordset
    rsChartAccount2.Open "Select AcctCode,AcctType from AMIS_ChartAccount where AcctCode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsChartAccount2.EOF And Not rsChartAccount2.BOF Then
        SetAcctType = SetDebitCredit(Null2String(rsChartAccount2!ACCTTYPE))
    Else
        SetAcctType = ""
    End If
End Function

Function SetBankCode(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    rsBanks.Open "Select bankcode,bankname from ALL_Banks where bankname = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankCode = Null2String(rsBanks!bankcode)
    Else
        SetBankCode = ""
    End If
End Function

Function SetBankName(VVV As Variant)
    Set rsBanks = New ADODB.Recordset
    rsBanks.Open "Select bankcode,bankname from ALL_Banks where bankcode = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsBanks.EOF And Not rsBanks.BOF Then
        SetBankName = Null2String(rsBanks!BankName)
    Else
        SetBankName = ""
    End If
End Function

Function SetCustomerCode(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,AcctName from ALL_CUSTMASTER_AMIS where AcctName = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = Null2String(rsCustomer!CUSTCODE)
    Else
        SetCustomerCode = ""
    End If
End Function

Function SetCustomerName(CCC As Variant)
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,AcctName from ALL_CUSTMASTER_AMIS where custcode = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!AcctName)
    Else
        SetCustomerName = ""
    End If
End Function

Function SetDebitCredit(VVV As Variant) As String
    Dim rsAccountType                             As ADODB.Recordset
    Set rsAccountType = New ADODB.Recordset
    rsAccountType.Open "Select Code,DebitCredit from AMIS_Acctype where Code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAccountType.EOF And Not rsAccountType.BOF Then
        If xJOURNALTYPE = "CDJ" Then
            If txtAcct_Name.Text = "ACCOUNTS PAYABLE - TRADE" Then SetDebitCredit = "D"
        ElseIf xJOURNALTYPE = "CRJ" Then
            If txtAcct_Name.Text = "ACCOUNTS RECEIVABLE - TRADE" Then SetDebitCredit = "C"
        Else
            SetDebitCredit = Null2String(rsAccountType!DebitCredit)
        End If
    Else
        SetDebitCredit = ""
    End If
End Function

Function SetInvCode(INV As Variant)
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invtype = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvCode = Null2String(rsInvoiceType!InvCode)
    Else
        SetInvCode = ""
    End If
End Function

Function SetInvType(INV As Variant)
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invcode = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvType = Null2String(rsInvoiceType!INVTYPE)
    Else
        SetInvType = ""
    End If
End Function

Function SetPayCode(VVV As Variant)
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayCode = Null2String(rsPayTerm!pay_Code)
    Else
        SetPayCode = ""
    End If
End Function

Function SetPayDesc(VVV As Variant) As String
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayDesc = Null2String(rsPayTerm!pay_desc)
    Else
        SetPayDesc = ""
    End If
End Function

Function SetPayNoDays(VVV As Variant) As Integer
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_Desc,no_days from ALL_PayTerm where pay_Desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayNoDays = Null2String(rsPayTerm!no_Days)
    Else
        SetPayNoDays = 0
    End If
End Function

Function SetVendorAddress(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,address from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    'rsVENDOR.Open "Select Code,Address from ALL_ENTITY where code = " & N2Str2Null(VVV) & " AND ENTITYCODE=" & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorAddress = Null2String(rsVENDOR!Address)
    Else
        SetVendorAddress = ""
    End If
End Function

Function SetVendorAddressNew(VVV As Variant, XXX As Variant)
    Set rsVENDOR = New ADODB.Recordset
    'rsVENDOR.Open "Select code,address from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsVENDOR.Open "Select Code,Address from ALL_ENTITY where code = " & N2Str2Null(VVV) & " AND ENTITYCODE=" & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorAddressNew = Null2String(rsVENDOR!Address)
    Else
        SetVendorAddressNew = ""
    End If
End Function

Function SetVendorCode(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where nameofvendor = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorCode = Null2String(rsVENDOR!Code)
    Else
        SetVendorCode = ""
    End If
End Function

Function SetVendorName(VVV As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,nameofvendor from ALL_Vendor where code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function SetVendorNameNew(VVV As Variant, XXX As Variant)
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "Select code,ACCOUNTNAME from ALL_ENTITY where code = " & N2Str2Null(VVV) & " AND ENTITYCODE = " & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorNameNew = Null2String(rsVENDOR!AccountName)
    Else
        SetVendorNameNew = ""
    End If
End Function

Function StoreGJEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,JNo,acct_code,acct_name,debit,jitemno,credit,tax from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labGJID.Caption = rsJournal_Det!ID
        txtGJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboGJAccountNo.Text = Null2String(rsJournal_Det!Acct_code)
        txtGJAccountName.Text = Null2String(rsJournal_Det!acct_Name)
        txtGJDebit.Text = N2Str2Zero(rsJournal_Det!DEBIT)
        txtGJCredit.Text = N2Str2Zero(rsJournal_Det!CREDIT)
        StoreGJParticulars Null2String(rsJournal_Det!JNo), Null2String(rsJournal_Det!jitemno)
    End If
End Function

Function StoreGJParticulars(ByVal JNo As Variant, ByVal ItemNo As Variant)
    Set rsJV_detail = New ADODB.Recordset
    rsJV_detail.Open "select JNo,ItemNo,Particulars from AMIS_JV_Detail where JNo = " & N2Str2Null(JNo) & " and ItemNo = " & N2Str2Null(ItemNo), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJV_detail.EOF And Not rsJV_detail.BOF Then
        txtGJAccountParticulars.Text = Null2String(rsJV_detail!Particulars)
    End If
End Function

Function StoreJournalEntry(ByVal ID As Variant)
    Set rsJournal_Det = New ADODB.Recordset
    rsJournal_Det.Open "select id,acct_code,acct_name,debit,jitemno,credit,tax,grossamt,netamt from AMIS_Journal_Det where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        labDetID.Caption = rsJournal_Det!ID
        labPartNo.Caption = Null2String(rsJournal_Det!Acct_code)
        txtJItemNo.Text = Null2String(rsJournal_Det!jitemno)
        cboAcct_Code.Text = Null2String(rsJournal_Det!Acct_code)
        txtAcct_Name.Text = Null2String(rsJournal_Det!acct_Name)
        txtDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!DEBIT))
        txtCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!CREDIT))
        txtTax.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!tax))
        txtGrossAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!grossamt))
        txtNetAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_Det!netamt))
    End If
End Function

Function StorePVEntry(ByVal ID As Variant)
'If JOURNALTYPE = "VPJ" Then
    Set rsPV_Detail = New ADODB.Recordset
    rsPV_Detail.Open "select * from AMIS_PV_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPV_Detail.EOF And Not rsPV_Detail.BOF Then
        labPVID.Caption = rsPV_Detail!ID
        txtPVItemNo.Text = Null2String(rsPV_Detail!ItemNo)
        txtPO_No.Text = Null2String(rsPV_Detail!po_no)
        txtMRR_No.Text = Null2String(rsPV_Detail!MRR_No)
        txtINV_No.Text = Null2String(rsPV_Detail!INV_NO)
        txtProd_No.Text = Null2String(rsPV_Detail!Prod_No)
        txtPVAmount.Text = N2Str2Zero(rsPV_Detail!amount)
    End If
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   Set rsCRJ_Detail = New ADODB.Recordset
    '       rsCRJ_Detail.Open "select * from AMIS_CRJ_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    '   If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
    '      labPVID.Caption = rsCRJ_Detail!ID
    '      txtPVItemNo.Text = Null2String(rsCRJ_Detail!ItemNo)
    '      txtPO_No.Text = txtVoucherNo.Text
    '      txtPO_No.Enabled = False
    '      txtMRR_No.Text = Null2String(rsCRJ_Detail!InvoiceType)
    '      txtINV_No.Text = Null2String(rsCRJ_Detail!InvoiceNo)
    '      txtProd_No.Text = Null2String(rsCRJ_Detail!InvoiceDate)
    '      txtPVAmount.Text = N2Str2Zero(rsCRJ_Detail!INVOICEAMOUNT)
    '      PrevInvoiceType = Null2String(rsCRJ_Detail!InvoiceType)
    '      PrevInvoiceNo = Null2String(rsCRJ_Detail!InvoiceNo)
    '   End If
    'Else
    '   Set rsCV_Detail = New ADODB.Recordset
    '       rsCV_Detail.Open "select * from AMIS_CV_Detail where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    '   If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
    '      labPVID.Caption = rsCV_Detail!ID
    '      txtPVItemNo.Text = Null2String(rsCV_Detail!ItemNo)
    '      txtPO_No.Text = txtVoucherNo.Text
    '      txtPO_No.Enabled = False
    '      txtMRR_No.Text = Null2String(rsCV_Detail!PV_VoucherNo)
    '      PrevPV_VoucherNo = Null2String(rsCV_Detail!PV_VoucherNo)
    '      txtINV_No.Text = Null2String(rsCV_Detail!DocDate)
    '      txtProd_No.Text = Null2String(rsCV_Detail!DueDate)
    '      txtPVAmount.Text = N2Str2Zero(rsCV_Detail!AMOUNT)
    '   End If
    'End If
End Function

Sub BringToFront()
    cmdAddJournal.ZOrder 0
    cmdAddJournal.Visible = True
    fraAddJournal.ZOrder 0
    fraAddJournal.Visible = True
    fraAddJournal.Enabled = True
End Sub

Sub BringToFrontGJ()
    cmdGJEntry.ZOrder 0
    cmdGJEntry.Visible = True
    picGJEntry.ZOrder 0
    picGJEntry.Visible = True
    picGJEntry.Enabled = True
End Sub

Sub BringToFrontPV()
    cmdPV_Entry.ZOrder 0
    cmdPV_Entry.Visible = True
    picPV_Entry.ZOrder 0
    picPV_Entry.Visible = True
    picPV_Entry.Enabled = True
End Sub

Sub BringToFrontTemplates()
    cmdTemplates.ZOrder 0
    picTemplates.ZOrder 0
    FillTemplates
End Sub

Sub FillCboAcctName()
    Dim rsCOBAcctName                             As ADODB.Recordset
    Set rsCOBAcctName = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where (Titles IN ('2101') OR (TITLES IN ('2102','2107') AND ISNULL(IS_SCHEDULE_ACCNT,0) <> 1)) OR ACCTCODE IN ('11-02015-00','11-02017-00','11-02018-00') and Description IS NOT NULL Order by AcctCode asc")
    Else
        Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Titles IN ('2101','2102','2107') AND IS_SCHEDULE_ACCNT=1 and Description IS NOT NULL Order by AcctCode asc")
    End If
    If Not rsCOBAcctName.EOF And Not rsCOBAcctName.BOF Then
        rsCOBAcctName.MoveFirst: cboCOBAcctName.Clear
        Do While Not rsCOBAcctName.EOF
            cboCOBAcctName.AddItem Null2String(rsCOBAcctName!Description)
            rsCOBAcctName.MoveNext
        Loop
    End If
End Sub

Sub FillDetails()
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE
    Dim PV_ITEMNO                                 As Integer
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        txtCOBAcctNo.Text = Null2String(rsJournal_Det!Acct_code)
        cboCOBAcctName.Text = Null2String(rsJournal_Det!acct_Name)
    Else
        txtCOBAcctNo.Text = ""
        cboCOBAcctName.Text = ""
        cmdPost.Enabled = False
    End If
    Jcnt = 0
    TOTALPVAMOUNT = 0
    txtTotalPV_Amount.Text = ZERO
    If xJOURNALTYPE = "VPJ" Then
        lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
        Set rsPV_Detail = New ADODB.Recordset
        Set rsPV_Detail = gconDMIS.Execute("select * from AMIS_PV_Detail where JTYPE = 'VPJ' AND VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
        If Not rsPV_Detail.EOF And Not rsPV_Detail.BOF Then
            Screen.MousePointer = 11
            rsPV_Detail.MoveFirst
            Do While Not rsPV_Detail.EOF
                Jcnt = Jcnt + 1
                If Null2String(rsPV_Detail!ItemNo) = "" Then PV_ITEMNO = Jcnt Else PV_ITEMNO = Null2String(rsPV_Detail!ItemNo)
                lstPV_Detail.ListItems.Add Jcnt, , Format(PV_ITEMNO, "0000")
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 1, , Null2String(rsPV_Detail!po_no)
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 2, , Null2String(rsPV_Detail!MRR_No)
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 3, , Null2String(rsPV_Detail!INV_NO)
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 4, , Null2String(rsPV_Detail!Prod_No)
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsPV_Detail!amount))
                lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 6, , rsPV_Detail!ID
                TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsPV_Detail!amount))
                rsPV_Detail.MoveNext
            Loop
            lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
            txtTotalPV_Amount.Text = TOTALPVAMOUNT
            Screen.MousePointer = 0
        End If
    End If
    'If JOURNALTYPE = "CDJ" Then
    '   lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
    '   Set rsCV_Detail = New ADODB.Recordset
    '   Set rsCV_Detail = gconDMIS.Execute("select * from AMIS_CV_Detail where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
    '   If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
    '      Screen.MousePointer = 11
    '      rsCV_Detail.MoveFirst
    '      Do While Not rsCV_Detail.EOF
    '         Jcnt = Jcnt + 1
    '         If Null2String(rsCV_Detail!ItemNo) = "" Then PV_ITEMNO = Jcnt Else PV_ITEMNO = Null2String(rsCV_Detail!ItemNo)
    '         lstPV_Detail.ListItems.Add Jcnt, , Format(PV_ITEMNO, "0000")
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 1, , Null2String(rsCV_Detail!PV_VoucherNo)
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 2, , Null2String(rsCV_Detail!DocDate)
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 3, , Null2String(rsCV_Detail!DueDate)
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsCV_Detail!AMOUNT))
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsCV_Detail!AMOUNT))
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 6, , rsCV_Detail!ID
    '         TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsCV_Detail!AMOUNT))
    '         rsCV_Detail.MoveNext
    '      Loop
    '      lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
    '      txtTotalPV_Amount.Text = TOTALPVAMOUNT
    '      Screen.MousePointer = 0
    '   End If
    'End If
    'If JOURNALTYPE = "CRJ" Then
    '   lstPV_Detail.ColumnHeaders(2).Width = lstPV_Detail.ColumnHeaders(2).Width + lstPV_Detail.ColumnHeaders(5).Width
    '   lstPV_Detail.ColumnHeaders(5).Width = 1
    '   lstPV_Detail.Sorted = False: lstPV_Detail.ListItems.Clear
    '   Set rsCRJ_Detail = New ADODB.Recordset
    '   Set rsCRJ_Detail = gconDMIS.Execute("select * from AMIS_CRJ_Detail where VoucherNo = " & N2Str2Null(txtVoucherNo.Text) & " order by itemno asc")
    '   If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
    '      Screen.MousePointer = 11
    '      rsCRJ_Detail.MoveFirst
    '      Do While Not rsCRJ_Detail.EOF
    '         Jcnt = Jcnt + 1
    '         If Null2String(rsCRJ_Detail!ItemNo) = "" Then PV_ITEMNO = Jcnt Else PV_ITEMNO = Null2String(rsCRJ_Detail!ItemNo)
    '         lstPV_Detail.ListItems.Add Jcnt, , Format(PV_ITEMNO, "0000")
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 1, , SetInvType(Null2String(rsCRJ_Detail!InvoiceType))
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 2, , Null2String(rsCRJ_Detail!InvoiceNo)
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 3, , Null2String(rsCRJ_Detail!InvoiceDate)
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsCRJ_Detail!INVOICEAMOUNT))
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 5, , ToDoubleNumber(N2Str2Zero(rsCRJ_Detail!INVOICEAMOUNT))
    '         lstPV_Detail.ListItems(Jcnt).ListSubItems.Add 6, , rsCRJ_Detail!ID
    '         TOTALPVAMOUNT = TOTALPVAMOUNT + NumericVal(N2Str2Zero(rsCRJ_Detail!INVOICEAMOUNT))
    '         rsCRJ_Detail.MoveNext
    '      Loop
    '      lstPV_Detail.Sorted = True: lstPV_Detail.Refresh
    '      txtTotalPV_Amount.Text = TOTALPVAMOUNT
    '      If TOTAL_AR_AMOUNT <> TOTALPVAMOUNT Then
    '         picChat.Visible = True
    '      Else
    '         picChat.Visible = False
    '      End If
    '      Screen.MousePointer = 0
    '   End If
    'End If
    'Else
    '   txtGJTotDebit.Text = ZERO: txtGJTotCredit.Text = ZERO: txtGJOutBalance.Text = ZERO
    '   lstGJ.Sorted = False: lstGJ.ListItems.Clear
    '   Set rsJOURNAL_DET = New ADODB.Recordset
    '   Set rsJOURNAL_DET = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & JOURNALTYPE & "' order by jitemno asc")
    '   If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
    '      Screen.MousePointer = 11
    '      rsJournal_Det.MoveFirst
    '      Do While Not rsJournal_Det.EOF
    '         kcnt = kcnt + 1
    '         If Null2String(rsJOURNAL_DET!jitemno) = "" Then J_ITemNo = kcnt Else J_ITemNo = Null2String(rsJOURNAL_DET!jitemno)
    '         lstGJ.ListItems.Add kcnt, , Format(J_ITemNo, "0000")
    '         lstGJ.ListItems(kcnt).ListSubItems.Add 1, , Null2String(rsJOURNAL_DET!acct_code)
    '         lstGJ.ListItems(kcnt).ListSubItems.Add 2, , Null2String(rsJOURNAL_DET!acct_name)
    '         lstGJ.ListItems(kcnt).ListSubItems.Add 3, , ToDoubleNumber(N2Str2Zero(rsJOURNAL_DET!Debit))
    '         lstGJ.ListItems(kcnt).ListSubItems.Add 4, , ToDoubleNumber(N2Str2Zero(rsJOURNAL_DET!Credit))
    '         lstGJ.ListItems(kcnt).ListSubItems.Add 5, , rsJOURNAL_DET!ID
    '         TOTDEBIT = TOTDEBIT + NumericVal(N2Str2Zero(rsJOURNAL_DET!Debit))
    '         TOTCREDIT = TOTCREDIT + NumericVal(N2Str2Zero(rsJOURNAL_DET!Credit))
    '         TOTTAX = TOTTAX + NumericVal(N2Str2Zero(rsJOURNAL_DET!TAX))
    '         rsJournal_Det.MoveNext
    '      Loop
    '      lstGJ.Sorted = True: lstGJ.Refresh
    '      OUTBALANCE = TOTDEBIT - TOTCREDIT
    '      txtGJTotDebit.Text = ToDoubleNumber(TOTDEBIT)
    '      txtGJTotCredit.Text = ToDoubleNumber(TOTCREDIT)
    '      txtGJOutBalance.Text = ToDoubleNumber(Abs(OUTBALANCE))
    '      Screen.MousePointer = 0
    '   End If
    'End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsChartAccount2                           As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccount2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If SearchBy = "NAME" Then
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,upper(Description),Accttype,ID from AMIS_ChartAccount where description like'" & XXX & "%' order by acctcode asc")
    Else
        Set rsChartAccount2 = gconDMIS.Execute("select acctcode,UPPER(Description),Accttype,ID from AMIS_ChartAccount where acctcode like'" & XXX & "%' order by acctcode asc")
    End If
    If Not (rsChartAccount2.EOF And rsChartAccount2.BOF) Then
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccount2
        lstAccounts.Refresh
        lstAccounts.Enabled = True
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Sub FillSearchTemplates(XXX As String)
    Dim rsTemplate_Header                         As ADODB.Recordset
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & xJOURNALTYPE & "' AND description like '" & XXX & "%' order by description asc")
    If Not (rsTemplate_Header.EOF And rsTemplate_Header.BOF) Then
        Listview_Loadval Me.lstTemplates.ListItems, rsTemplate_Header
        lstTemplates.Refresh
        lstTemplates.Enabled = True
    Else
        lstTemplates.Enabled = False
    End If
End Sub

Sub FillTemplates()
    Dim rsTemplate_Header                         As ADODB.Recordset
    lstTemplates.Sorted = False: lstTemplates.ListItems.Clear
    Set rsTemplate_Header = New ADODB.Recordset
    Set rsTemplate_Header = gconDMIS.Execute("select Description,templatecode from AMIS_Template_Header where Jtype = '" & xJOURNALTYPE & "' order by description asc")
    If Not (rsTemplate_Header.EOF And rsTemplate_Header.BOF) Then
        lstTemplates.Enabled = True
        Listview_Loadval Me.lstTemplates.ListItems, rsTemplate_Header
        lstTemplates.Refresh
    Else
        lstTemplates.Enabled = False
    End If
End Sub

Sub FindDupJNo(DDD As String)
    rsJournal_HD.Bookmark = rsFind(rsJournal_HD.Clone, "jno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub InitCbo()
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "select nameofvendor from ALL_Vendor order by nameofvendor asc", gconDMIS    '
    'rsVENDOR.Open "select AccountName from ALL_Entity order by nameofvendor asc", gconDMIS
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        rsVENDOR.MoveFirst
        cboNameofVendor.Clear
        Do While Not rsVENDOR.EOF
            cboNameofVendor.AddItem Null2String(rsVENDOR!nameofvendor)
            rsVENDOR.MoveNext
        Loop
    End If
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "select pay_desc from ALL_PayTerm order by pay_desc asc", gconDMIS
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        rsPayTerm.MoveFirst
        cboPayType.Clear
        Do While Not rsPayTerm.EOF
            cboPayType.AddItem Null2String(rsPayTerm!pay_desc)
            rsPayTerm.MoveNext
        Loop
    End If
    FillCboAcctName
End Sub

Sub InitGJ()
    txtGJItemNo.Text = Format(kcnt + 1, "0000")
    cboGJAccountNo.Text = ""
    txtGJAccountName.Text = ""
    txtGJDebit.Text = ZERO
    txtGJCredit.Text = ZERO
    txtGJAccountParticulars.Text = "Pls. Type Your Remarks Here..."
    txtSearch.Text = ""
End Sub

Sub InitGrid()
'If JOURNALTYPE = "CDJ" Then
'   With lstPV_Detail
'        .ColumnHeaders(2).Text = "PV Number"
'        .ColumnHeaders(2).Width = 2900
'        .ColumnHeaders(3).Text = "Doc. Date"
'        .ColumnHeaders(3).Width = 2000
'        .ColumnHeaders(4).Text = "Due Date"
'        .ColumnHeaders(4).Alignment = lvwColumnLeft
'        .ColumnHeaders(4).Width = 2000
'        .ColumnHeaders(5).Width = 1
'        txtMRR_No.MaxLength = 6
'   End With
'End If
End Sub

Sub InitJournal()
    txtJItemNo.Text = Format(kcnt + 1, "0000")
    cboAcct_Code.Text = ""
    txtAcct_Name.Text = ""
    txtDebit.Text = ZERO
    txtCredit.Text = ZERO
    txtTax.Text = ZERO
    txtGrossAmt.Text = ZERO
    txtNetAmt.Text = ZERO
    txtSearch.Text = ""
End Sub

Sub initMemvars()

    Call Get_VoucherNo
    Dim rsJDate                                   As ADODB.Recordset
    Set rsJDate = New ADODB.Recordset
    rsJDate.Open "Select * from ALL_PROFILE", gconDMIS, adOpenKeyset
    If Not rsJDate.EOF And Not rsJDate.BOF Then
        txtJDate.Text = Null2Date(rsJDate!Cut_Off_Date)
    End If

    'Accounts Payable Module'
    txtCode.Text = ""
    txtAddress.Caption = "":
    txtInvoiceDate.Text = LOGDATE
    txtDueDate.Text = LOGDATE:
    txtBankCode.Text = ""
    txtRemarks.Text = "Pls Type Your Message Here!"
    '---------------------------'
    'Cash Disbursement Module'
    txtCheckNo.Text = "": txtCheckDate.Text = LOGDATE: txtPayCode.Text = ""
    txtNameofVendor.Text = ""
    cboNameofVendor.Text = ""
    txtTotDebit.Text = ZERO: txtTotCredit.Text = ZERO
    txtAmountToPay.Text = ZERO: txtOutBalance.Text = ZERO
    txtParticulars2.Locked = False
    txtParticulars.Text = "Pls Type Your Message Here!"
    txtParticulars2.Text = "Pls Type Your Message Here!"
    If COMPANY_CODE = "HMH" Then
        picNameofVendor.Visible = True
    Else
        picNameofVendor.Visible = False
    End If
    '---------------------------'
    'Accounts Receivable Module'
    txtCustCode.Text = ""
    cboCustName.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate2.Text = LOGDATE
    txtInvoiceAmt.Text = ZERO
    txtRefNo.Text = ""
    txtRefDate.Text = LOGDATE
    txtRemarks2.Text = "Pls Type Your Message Here!"
    '---------------------------'

    txtTotalPV_Amount.Text = ZERO
    labPosted.Caption = ""
    labPosted.Visible = False
    labOutBalance.Visible = False
    txtOutBalance.Visible = False
    picRefCDJ.Visible = False
    InitGrid
    SendToBack

    txtCOBAcctNo.Text = ""
    cboCOBAcctName.Text = ""
End Sub

Private Sub Get_VoucherNo()
    Dim rsJournal_HDDup                           As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
End Sub
Sub InitPV_Detail()
    txtPVItemNo.Text = Format(Jcnt + 1, "0000")
    txtMRR_No.Text = ""
    'If JOURNALTYPE = "VPJ" Then
    txtPO_No.Text = "": txtINV_No.Text = "": txtProd_No.Text = ""
    txtPVAmount.Text = ZERO
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   txtPO_No.Text = txtVoucherNo.Text: txtINV_No.Text = ""
    '   txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
    '   txtPVAmount.Text = ZERO
    'Else
    '   labPV1.Caption = "Voucher No": txtPO_No.Text = txtVoucherNo.Text: txtPO_No.Enabled = False
    '   labPV2.Caption = "PV Voucher No.": labPV3.Caption = "Doc. Date": labPV4.Caption = "Due Date"
    '   txtINV_No.Text = LOGDATE: txtINV_No.Format = "dd-mmm-yy"
    '   txtProd_No.Text = LOGDATE: txtProd_No.Format = "dd-mmm-yy"
    '   txtPVAmount.Text = ZERO
    'End If
End Sub

Sub InsertAccountEntries(XXX As Variant)
    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET  As Double
    Dim J_STATUS, J_JITEMNO                       As String
    Dim rsTemplate_Details                        As ADODB.Recordset
    Set rsTemplate_Details = New ADODB.Recordset
    Set rsTemplate_Details = gconDMIS.Execute("Select * from AMIS_Template_Details Where TemplateCode = " & XXX)
    If Not rsTemplate_Details.EOF And Not rsTemplate_Details.BOF Then
        rsTemplate_Details.MoveFirst
        Screen.MousePointer = 11
        Do While Not rsTemplate_Details.EOF
            kcnt = kcnt + 1
            J_JDATE = N2Date2Null(txtJDate.Text)
            J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
            J_JTYPE = N2Str2Null(xJOURNALTYPE)
            J_JNO = N2Str2Null(txtJNo.Text)
            J_JITEMNO = N2Str2Null(Format(kcnt, "0000"))
            J_ACCT_CODE = N2Str2Null(rsTemplate_Details!AccountCode)
            J_ACCT_NAME = N2Str2Null(rsTemplate_Details!Description)
            J_DEBIT = 0: J_CREDIT = 0: J_TAX = 0: J_GROSS = 0: J_NET = 0
            J_STATUS = "'N'"
            gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                             "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status)" & _
                             " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                             ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                             ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ")"
            rsTemplate_Details.MoveNext
        Loop
        StoreMemVars
        FillDetails

        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then: lstDetails.SetFocus

        Screen.MousePointer = 0
    End If
End Sub

Sub OkAccount()
    fraFindAccount.Visible = False: cmdFindAccount.Visible = False
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Then
        cboGJAccountNo.Text = labAccountCode.Caption
        txtGJAccountName.Text = Setacctname(labAccountCode.Caption)
        If cboGJAccountNo.Text <> "" Then
            If SetAcctType(cboGJAccountNo.Text) = "C" Then
                On Error Resume Next
                txtGJCredit.SetFocus
            Else
                On Error Resume Next
                txtGJDebit.SetFocus
            End If
        End If
    Else
        cboAcct_Code.Text = labAccountCode.Caption
        If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "SJ" Then
            On Error Resume Next
            txtGrossAmt.SetFocus
        Else
            If cboAcct_Code.Text <> "" Then
                If SetAcctType(cboAcct_Code.Text) = "C" Then
                    On Error Resume Next
                    txtCredit.SetFocus
                Else
                    On Error Resume Next
                    txtDebit.SetFocus
                End If
            End If
        End If
    End If
    cmdFindAccount.ZOrder 1
    fraFindAccount.ZOrder 1
End Sub

Sub OkAccountSetCursor()
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Then
        If cboGJAccountNo.Text <> "" Then
            If SetAcctType(cboGJAccountNo.Text) = "C" Then
                On Error Resume Next
                txtGJCredit.SetFocus
            Else
                On Error Resume Next
                txtGJDebit.SetFocus
            End If
        End If
    Else
        If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "SJ" Then
            On Error Resume Next
            txtGrossAmt.SetFocus
        Else
            If cboAcct_Code.Text <> "" Then
                If SetAcctType(cboAcct_Code.Text) = "C" Then
                    On Error Resume Next
                    txtCredit.SetFocus
                Else
                    On Error Resume Next
                    txtDebit.SetFocus
                End If
            End If
        End If
    End If
End Sub

Sub rsRefresh()
    If xJOURNALTYPE = "VPJ" Then Me.Caption = "VENDOR OPENING BALANCE - DATA ENTRY"
    'If JOURNALTYPE = "APJ" Then Me.Caption = "ACCOUNTS PAYABLE JOURNAL ENTRY"
    'If JOURNALTYPE = "CDJ" Then Me.Caption = "CASH DISBURSEMENT JOURNAL ENTRY"
    'If JOURNALTYPE = "SJ" Then Me.Caption = "SALES JOURNAL ENTRY"
    'If JOURNALTYPE = "CRJ" Then Me.Caption = "CASH RECEIPTS JOURNAL ENTRY"
    'If JOURNALTYPE = "GJ" Then Me.Caption = "GENERAL JOURNAL DATA ENTRY"
    'If JOURNALTYPE = "ADJ" Then Me.Caption = "AUDIT ADJUSTMENTS DATA ENTRY"
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by VoucherNo asc", gconDMIS, adOpenKeyset
End Sub

Sub SearchVoucherNo(XXX As String)
    If XXX <> "" Then
        On Error GoTo ErrorCode
        rsJournal_HD.Bookmark = rsFind(rsJournal_HD.Clone, "voucherno", XXX).Bookmark
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.Number = 3021 Then
        MsgBox "Can't find " & XXX, vbOKOnly + vbExclamation, "Not Found"
        Resume Next
    End If
End Sub

Sub SendToBack()
    cmdAddJournal.ZOrder 1
    cmdAddJournal.Visible = False
    fraAddJournal.ZOrder 1
    fraAddJournal.Visible = False
    fraAddJournal.Enabled = False
    fraFindAccount.ZOrder 1
    cmdFindAccount.ZOrder 1
    fraFindAccount.Visible = False
    cmdFindAccount.Visible = False
    cmdShowPostRange.Visible = False
    picShowPostRange.Visible = False
    cmdPrinting.ZOrder 1
    picPrinting.ZOrder 1
End Sub

Sub SendToBackGJ()
    cmdGJEntry.ZOrder 1
    cmdGJEntry.Visible = False
    picGJEntry.ZOrder 1
    picGJEntry.Visible = False
    picGJEntry.Enabled = False
End Sub

Sub SendToBackPV()
    cmdPV_Entry.ZOrder 1
    cmdPV_Entry.Visible = False
    picPV_Entry.ZOrder 1
    picPV_Entry.Visible = False
    picPV_Entry.Enabled = False
End Sub

Sub SendToBackTemplates()
    cmdTemplates.ZOrder 1
    picTemplates.ZOrder 1
End Sub

Sub StoreMemVars()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        labID.Caption = rsJournal_HD!ID
        txtJNo.Text = Null2String(rsJournal_HD!JNo)
        txtVoucherNo.Text = Null2String(rsJournal_HD!VOUCHERNO)
        txtJDate.Text = Format(Null2String(rsJournal_HD!JDate), "DD-MMM-YY")
        txtInvoiceDate.Text = Format(Null2String(rsJournal_HD!invoicedate), "DD-MMM-YY")
        txtDueDate.Text = Format(Null2String(rsJournal_HD!duedate), "DD-MMM-YY")
        txtPayCode.Text = Null2String(rsJournal_HD!paytype)
        txtTerms.Text = Null2String(rsJournal_HD!TERMS)
        cboPayType.Text = SetPayDesc(Null2String(rsJournal_HD!paytype))
        txtCode.Text = Null2String(rsJournal_HD!VendorCode)
        cboNameofVendor.Text = SetVendorName(txtCode.Text)
        txtNameofVendor.Text = SetVendorNameNew(txtCode.Text, Null2String(rsJournal_HD!Entity_Class))
        CURRENT_VENDORCODE = Null2String(rsJournal_HD!VendorCode)
        If COMPANY_CODE = "HMH" Then
            txtAddress.Caption = SetVendorAddressNew(txtCode.Text, Null2String(rsJournal_HD!Entity_Class))
        Else
            txtAddress.Caption = SetVendorAddress(txtCode.Text)
        End If
        cboBankName.Text = SetBankName(Null2String(rsJournal_HD!bankcode))
        xEntityClass = Null2String(rsJournal_HD!Entity_Class)
        Set rsCRJ_Detail = New ADODB.Recordset
        Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail Where JTYPE='VPJ' AND PV_VoucherNo = " & N2Str2Null(txtVoucherNo.Text))
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            picRefCDJ.ZOrder 0: picRefCDJ.Visible = True
            RefCDJ.Caption = "Ref CDJ# " & Null2String(rsCRJ_Detail!VOUCHERNO)
        Else
            picRefCDJ.ZOrder 1: picRefCDJ.Visible = False
            RefCDJ.Caption = ""
        End If
        txtBankCode.Text = Null2String(rsJournal_HD!bankcode)
        txtCheckNo.Text = Null2String(rsJournal_HD!CheckNo)
        txtCheckDate.Text = Null2String(rsJournal_HD!CheckDate)
        txtParticulars.Text = Null2String(rsJournal_HD!remarks)
        txtParticulars2.Text = Null2String(rsJournal_HD!remarks)
        txtTotDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!DEBIT))
        txtTotCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!CREDIT))
        txtOutBalance.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!OUTBALANCE))
        txtAmountToPay.Text = ToDoubleNumber(NumericVal(rsJournal_HD!amounttopay))
        txtRemarks.Text = Null2String(rsJournal_HD!remarks)
        txtRemarks2.Text = Null2String(rsJournal_HD!remarks)
        If Null2String(rsJournal_HD!Status) = "C" Then
            labPosted.Visible = True: labPosted.Caption = "*** CANCELLED ***"
            cmdEdit.Enabled = False: cmdCancelCO.Enabled = False: cmdPost.Enabled = False
            cmdUnPost.Enabled = False: cmdPrint.Enabled = False
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            labPosted.Visible = True: labPosted.Caption = "*** POSTED ***"
            cmdEdit.Enabled = False: cmdPost.Enabled = False
            cmdCancelCO.Enabled = False:                   'cmdPrint.Enabled = True
            cmdUnPost.Enabled = True
        Else
            labPosted.Caption = "": labPosted.Visible = False
            cmdEdit.Enabled = True: cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = True: cmdPost.Enabled = True
            cmdPrint.Enabled = False
        End If
        DoEvents
        FillDetails
    Else
        MsgBox "No Such Record!": If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub StoreSearch(XXX As Variant)
    rsRefresh
    rsJournal_HD.Find "VoucherNo = " & N2Str2Null(XXX)
    StoreMemVars
End Sub

Private Sub cboAcct_Code_Change()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cboAcct_Code_Click()
    txtAcct_Name.Text = Setacctname(cboAcct_Code.Text)
End Sub

Private Sub cboBankName_Click()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cboBankName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboBankName_LostFocus()
    txtBankCode.Text = SetBankCode(cboBankName.Text)
End Sub

Private Sub cboBankName2_Click()
    txtBankCode.Text = SetBankCode(cboBankName2.Text)
End Sub

Private Sub cboBankName2_LostFocus()
    txtBankCode.Text = SetBankCode(cboBankName2.Text)
End Sub

Private Sub cboCOBAcctName_Click()
    txtCOBAcctNo.Text = SetCOBAcctNo(cboCOBAcctName.Text)
End Sub

Private Sub cboCustName_Change()
    txtCustCode.Text = SetCustomerCode(cboCustName.Text)
End Sub

Private Sub cboCustName_Click()
    txtCustCode.Text = SetCustomerCode(cboCustName.Text)
End Sub

Private Sub cboCustName_GotFocus()
    VBComBoBoxDroppedDown cboCustName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
End Sub

Private Sub txtNameofVendor_Change()
    txtCode.Text = SetVendorCode(txtNameofVendor.Text)
    txtAddress.Caption = SetVendorAddress(txtCode.Text)
End Sub

Private Sub cboNameofVendor_Click()
    txtCode.Text = SetVendorCode(cboNameofVendor.Text)
    txtAddress.Caption = SetVendorAddress(txtCode.Text)
End Sub

Private Sub cboNameofVendor_GotFocus()
    VBComBoBoxDroppedDown cboNameofVendor
End Sub

Private Sub cboNameofVendor_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub cboPayType_Change()
    txtPayCode.Text = SetPayCode(cboPayType.Text)
    On Error Resume Next
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayType_Click()
    txtPayCode.Text = SetPayCode(cboPayType.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub cboPayType_LostFocus()
    txtPayCode.Text = SetPayCode(cboPayType.Text)
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

'Upating Code       : AXP-0713200714:20
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "VENDOR OPENING BALANCE") = False Then Exit Sub

    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    Dim rsProfile                                 As ADODB.Recordset
    Dim AccountingMonth, AccountingYear           As Integer
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        AccountingMonth = rsProfile!PERIODMONTH
        AccountingYear = rsProfile!PERIODYEAR
    End If
    Dim rsDetails                                 As ADODB.Recordset
    Set rsDetails = New ADODB.Recordset
    Set rsDetails = gconDMIS.Execute("Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit, VoucherNo from AMIS_Journal_Det Where jtype = '" & xJOURNALTYPE & "' and Month(Jdate) = " & AccountingMonth & " and Year(Jdate) = " & AccountingYear & " and Status <> 'C' group by VoucherNo order by VoucherNo asc")
    If Not rsDetails.EOF And Not rsDetails.EOF Then
        Screen.MousePointer = 11
        Do While Not rsDetails.EOF
            If Round(rsDetails!TotalDebit, 2) <> Round(rsDetails!Totalcredit, 2) Then
                Screen.MousePointer = 0
                MsgBox "Warning: " & xJOURNALTYPE & "-" & rsDetails!VOUCHERNO & " is still not balance or has zero details" & vbCrLf & _
                       "              Adding Other Entries is not Allowed!", vbCritical, "Unbalanced Entry"
                Exit Sub
            End If
            rsDetails.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    SendToBack
    initMemvars
    FillDetails
    lstDetails.Enabled = False
    On Error Resume Next
    txtVoucherNo.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdAddAccount_Click()
    Screen.MousePointer = 11
    REFRESH_ACCOUNT = True
    frmAMISFILESChartOfAccount.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmdAddJournal_Click()
    SendToBack
    cmdAddJournal.Visible = True: cmdAddJournal.ZOrder 0
    fraAddJournal.Visible = True: fraAddJournal.ZOrder 0
    fraAddJournal.Enabled = True: cmdJournalDelete.Visible = False
    AddorEdit = "ADD"
    InitJournal
    On Error Resume Next
    cboAcct_Code.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDetails.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdCancelCO_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_CancelEntry", "VENDOR OPENING BALANCE") = False Then Exit Sub

    If MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Cancel Journal") = vbYes Then
        Screen.MousePointer = 11
        gconDMIS.Execute "update AMIS_Journal_HD set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HSB" Then
            If xJOURNALTYPE = "VPJ" Then
                With FrmCancelTransaction
                    .lblTransaction_type = xJOURNALTYPE
                    .LblTransactionNo = txtVoucherNo.Text
                    FrmCancelTransaction.Show
                End With
            End If
        End If
        
        Dim xAPVOUCHERNO                              As String
        xAPVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
        If xJOURNALTYPE = "VPJ" Then
            gconDMIS.Execute "DELETE FROM AMIS_AP WHERE VOUCHERNO = " & N2Str2Null(xAPVOUCHERNO) & ""
        End If
    
        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        Screen.MousePointer = 0
    End If
    LogAudit "C", "VENDOR A/P OPENING", txtVoucherNo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "VENDOR OPENING BALANCE") = False Then Exit Sub
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "VENDOR OPENING BALANCE") = False Then Exit Sub

    AddorEdit = "EDIT"
    PrevJType = UCase(xJOURNALTYPE)
    PrevJNo = Format(txtJNo.Text, "000000")
    lstDetails.Enabled = False
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_HD!ID
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Then txtParticulars2.Locked = False
    On Error Resume Next
    txtVoucherNo.SetFocus
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
'If JOURNALTYPE = "APJ" Then
    frmAMISSearchVPJ.Show vbModal
    'ElseIf JOURNALTYPE = "CDJ" Then
    '   frmAMISSearchCDJ.Show vbModal
    'ElseIf JOURNALTYPE = "SJ" Then
    '   frmAMISSearchSJ.Show vbModal
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   frmAMISSearchCRJ.Show vbModal
    'Else
    '   frmAMISSearchGJ.Show vbModal
    'End If
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdFirst_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MoveFirst
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdGJCancel_Click()
    SendToBackGJ
    StoreMemVars
End Sub

Private Sub cmdGJDelete_Click()
    If labGJID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Journal_Det where id = " & labGJID.Caption
    End If
    Dim cnt                                       As Integer
    Dim rsJournalDup                              As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemNo,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update AMIS_Journal_Det set JItemNo = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            rsJournalDup.MoveNext
        Loop
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    cmdGJCancel.Value = True
    If lstGJ.ListItems.Count > 0 And lstGJ.Enabled = True Then: lstGJ.SetFocus
End Sub

Private Sub cmdGJEntry_Click()
    SendToBackGJ
    cmdGJEntry.Visible = True: cmdGJEntry.ZOrder 0
    picGJEntry.Visible = True: picGJEntry.ZOrder 0
    picGJEntry.Enabled = True: cmdGJDelete.Visible = False
    AddorEdit = "ADD"
    InitGJ
    On Error Resume Next
    cboGJAccountNo.SetFocus
End Sub

Private Sub cmdGJSave_Click()
    On Error GoTo ErrorCode
    If cboGJAccountNo.Text = "" Then
        MsgBox "Account Code must have a value", vbInformation, "Error Encountered!"
        Exit Sub
    End If
    If AddorEdit = "ADD" Then
        Dim rsJournal_DetClone                    As ADODB.Recordset
        Set rsJournal_DetClone = New ADODB.Recordset
        rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(xJOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
        If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
            MsgBox "Account Code already used in this transaction", vbInformation, "Error in Part Number Validation"
            Exit Sub
        End If
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX                  As Double
    Dim J_STATUS, J_JITEMNO                       As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_JITEMNO = N2Str2Null(txtGJItemNo.Text)
    J_ACCT_CODE = N2Str2Null(cboGJAccountNo.Text)
    J_ACCT_NAME = N2Str2Null(txtGJAccountName.Text)
    J_DEBIT = Round(NumericVal(txtGJDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtGJCredit.Text), 2)
    J_TAX = Round(NumericVal(txtTax.Text), 2)
    J_STATUS = "'N'"

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        If txtGJAccountParticulars.Text <> "" And txtGJAccountParticulars.Text <> "Pls Type Your Remarks Here!" Then
            gconDMIS.Execute "insert into AMIS_JV_Detail " & _
                             "(JNo,VoucherNo,itemno,Particulars,status)" & _
                             " values (" & J_JNO & ", " & J_VOUCHERNO & ", " & J_JITEMNO & _
                             ", " & N2Str2Null(txtGJAccountParticulars.Text) & _
                             ", " & J_STATUS & ")"
        End If
        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,status)" & _
                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                         ", " & J_CREDIT & ", " & J_TAX & ", " & J_STATUS & ")"
    Else
        gconDMIS.Execute "update AMIS_Journal_Det set" & _
                         " jdate = " & J_JDATE & "," & _
                         " voucherno = " & J_VOUCHERNO & "," & _
                         " jtype = " & J_JTYPE & "," & _
                         " jno = " & J_JNO & "," & _
                         " jitemno = " & J_JITEMNO & "," & _
                         " acct_code = " & J_ACCT_CODE & "," & _
                         " acct_name = " & J_ACCT_NAME & "," & _
                         " debit = " & J_DEBIT & "," & _
                         " credit = " & J_CREDIT & "," & _
                         " tax = " & J_TAX & "," & _
                         " status = " & J_STATUS & _
                         " where id = " & labGJID.Caption
        gconDMIS.Execute "update AMIS_JV_Detail set" & _
                         " Particulars = " & N2Str2Null(txtGJAccountParticulars.Text) & _
                         " where JNo = " & J_JNO & " and ItemNo = " & J_JITEMNO
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then cmdGJEntry_Click Else cmdGJCancel_Click
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub cmdJournalCancel_Click()
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdJournalDelete_Click()
    If labDetID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If MsgBox("Delete This Journal, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
        gconDMIS.Execute "delete from AMIS_Journal_Det where id = " & labDetID.Caption
    End If
    Dim cnt                                       As Integer
    Dim rsJournalDup                              As ADODB.Recordset
    Set rsJournalDup = New ADODB.Recordset
    rsJournalDup.Open "select id,JItemno,JType,VoucherNo from AMIS_Journal_Det where JType = " & N2Str2Null(xJOURNALTYPE) & " and VoucherNo = " & N2Str2Null(rsJournal_HD!VOUCHERNO) & " order by ID asc", gconDMIS
    If Not rsJournalDup.EOF And Not rsJournalDup.BOF Then
        rsJournalDup.MoveFirst
        cnt = 0
        Do While Not rsJournalDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update AMIS_Journal_Det set JItemno = " & Format(cnt, "0000") & " where id = " & rsJournalDup!ID
            rsJournalDup.MoveNext
        Loop
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    cmdJournalCancel.Value = True
    If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then: lstDetails.SetFocus
End Sub

Private Sub cmdJournalSave_Click()
    On Error GoTo ErrorCode
    If cboAcct_Code.Text = "" Then
        MsgBox "Account Code must have a value", vbInformation, "Error Encountered!"
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsJournal_DetClone                    As ADODB.Recordset
        Set rsJournal_DetClone = New ADODB.Recordset
        rsJournal_DetClone.Open "select JType,JNo,JItemno,Acct_code from AMIS_Journal_Det where Acct_Code = " & N2Str2Null(cboAcct_Code.Text) & " and Jtype = " & N2Str2Null(xJOURNALTYPE) & " and Jno =" & N2Str2Null(txtJNo.Text) & " order by Jitemno asc", gconDMIS
        If Not rsJournal_DetClone.EOF And Not rsJournal_DetClone.BOF Then
            MsgBox "Account Code already used in this transaction", vbInformation, "Error in Part Number Validation"
            Exit Sub
        End If
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_ACCT_CODE, J_ACCT_NAME           As String
    Dim J_DEBIT, J_CREDIT, J_TAX, J_GROSS, J_NET  As Double
    Dim J_STATUS, J_JITEMNO                       As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_JITEMNO = N2Str2Null(Format(txtJItemNo.Text, "0000"))
    J_ACCT_CODE = N2Str2Null(cboAcct_Code.Text)
    J_ACCT_NAME = N2Str2Null(txtAcct_Name.Text)
    J_DEBIT = Round(NumericVal(txtDebit.Text), 2)
    J_CREDIT = Round(NumericVal(txtCredit.Text), 2)
    J_TAX = Round(NumericVal(txtTax.Text), 2)
    J_GROSS = Round(NumericVal(txtGrossAmt.Text), 2)
    J_NET = Round(NumericVal(txtNetAmt.Text), 2)
    J_STATUS = "'N'"

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "insert into AMIS_Journal_Det " & _
                         "(jdate,voucherno,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax,grossamt,netamt,status,USERCODE,LASTUPDATE)" & _
                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & _
                         ", " & J_JNO & ", " & J_JITEMNO & ", " & J_ACCT_CODE & ", " & J_ACCT_NAME & ", " & J_DEBIT & _
                         ", " & J_CREDIT & ", " & J_TAX & "," & J_GROSS & "," & J_NET & ", " & J_STATUS & ",'" & LOGCODE & "','" & LOGDATE & "')"
    Else
        gconDMIS.Execute "update AMIS_Journal_Det set" & _
                         " jdate = " & J_JDATE & "," & _
                         " voucherno = " & J_VOUCHERNO & "," & _
                         " jtype = " & J_JTYPE & "," & _
                         " jno = " & J_JNO & "," & _
                         " jitemno = " & J_JITEMNO & "," & _
                         " acct_code = " & J_ACCT_CODE & "," & _
                         " acct_name = " & J_ACCT_NAME & "," & _
                         " debit = " & J_DEBIT & "," & _
                         " credit = " & J_CREDIT & "," & _
                         " tax = " & J_TAX & "," & _
                         " grossamt = " & J_GROSS & "," & _
                         " netamt = " & J_NET & "," & _
                         " USERCODE = '" & LOGCODE & "'," & _
                         " LASTUPDATE = '" & LOGDATE & "'," & _
                         " status = " & J_STATUS & _
                         " where id = " & labDetID.Caption
    End If
    FillDetails
    gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     " debit = " & TOTDEBIT & "," & _
                     " credit = " & TOTCREDIT & "," & _
                     " tax = " & TOTTAX & "," & _
                     " outbalance = " & OUTBALANCE & _
                     " where id = " & labID.Caption
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then cmdAddJournal_Click Else cmdJournalCancel_Click
    If AddorEdit = "EDIT" Then
        If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then: lstDetails.SetFocus
    End If
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200714:20
Private Sub cmdLast_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MoveLast
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdNext_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MoveNext
    If rsJournal_HD.EOF Then
        rsJournal_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdOkPrint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optPrintVoucher.Value = True Then
        Screen.MousePointer = 11
        picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
        ShowReport "CashDisbursement", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "CASH DISBURSEMENT JOURNAL PRINTOUT", LOGDATE, False
        Screen.MousePointer = 0
    Else
        If optSECBANK.Value = True Then
            If MsgBox("Please Insert Security Bank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Security Bank Check
                ShowReport "SecurityBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
        If optPRUDBANK.Value = True Then
            If MsgBox("Please Insert Prudential Bank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Prudential Bank Check
                ShowReport "PrudentialBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
        If optCHINBANK.Value = True Then
            If MsgBox("Please Insert Chinabank Check...", vbOKCancel + vbInformation, "Press Ok To Continue Printing") = vbOK Then
                picPrinting.ZOrder 1: cmdPrinting.ZOrder 1
                'Print Chinabank Check
                ShowReport "ChinaBankCheck", "Checks", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "", LOGDATE, False
            End If
        End If
    End If
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdPost_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Post", "VENDOR OPENING BALANCE") = False Then Exit Sub

    Screen.MousePointer = 11
    'PostJournal:
    gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute "UPDATE AMIS_PV_DETAIL SET STATUS='P' WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & ""

    Call GET_AP_VOUCHERNO

    rsRefresh
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
    LogAudit "P", "VENDOR A/P OPENING", txtVoucherNo
    Exit Sub
ErrorCode:
    'SaveLogFile
    ShowVBError
End Sub

Private Sub cmdPostRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtToVNo.Text < txtFromVNo.Text Then
        MsgBox "Error: Invalid Voucher No. Range", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    txtFromVNo.Text = Format(txtFromVNo.Text, "000000")
    txtToVNo.Text = Format(txtToVNo.Text, "000000")
    Dim rsCheckVouchers, rsCheckUnBalancedVouchers As ADODB.Recordset
    Set rsCheckVouchers = New ADODB.Recordset
    Set rsCheckVouchers = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD where Jtype = '" & xJOURNALTYPE & "' AND VoucherNo = '" & txtToVNo.Text & "'")
    If rsCheckVouchers.EOF And rsCheckVouchers.BOF Then
        MsgBox "Error: Voucher No. Range Exceeds Current Records Available.", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    Dim KIM, JOY, YZA                             As Integer
    Screen.MousePointer = 11
    JOY = 0
    YZA = NumericVal(txtToVNo.Text) - NumericVal(txtFromVNo.Text)
    picShowPostRange.Enabled = False
    For KIM = txtFromVNo.Text To txtToVNo.Text
        Set rsCheckVouchers = New ADODB.Recordset
        Set rsCheckVouchers = gconDMIS.Execute("Select JType,VoucherNo,Debit,Credit,Status from AMIS_Journal_HD Where JType = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000")))
        If Not rsCheckVouchers.EOF And Not rsCheckVouchers.BOF Then
            Set rsCheckUnBalancedVouchers = New ADODB.Recordset
            Set rsCheckUnBalancedVouchers = gconDMIS.Execute("Select SUM(DEBIT) as TotalDebit, SUM(CREDIT) as TotalCredit from AMIS_Journal_Det Where jtype = '" & xJOURNALTYPE & "' and Status <> 'C' and VoucherNo = " & N2Str2Null(Format(KIM, "000000")))
            If Round(rsCheckUnBalancedVouchers!TotalDebit, 2) <> Round(rsCheckUnBalancedVouchers!Totalcredit, 2) Then
                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
            Else
                If Null2String(rsCheckVouchers!Status) = "N" Then
                    If N2Str2Zero(rsCheckVouchers!DEBIT) = N2Str2Zero(rsCheckVouchers!CREDIT) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!jtype) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
                    End If
                ElseIf Null2String(rsCheckVouchers!Status) = "C" Then
                    gconDMIS.Execute "update AMIS_Journal_HD set status = 'C' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    gconDMIS.Execute "update AMIS_Journal_Det set status = 'C' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                Else
                    If N2Str2Zero(rsCheckVouchers!DEBIT) = N2Str2Zero(rsCheckVouchers!CREDIT) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!jtype) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
                    End If
                End If
            End If
        End If
        JOY = JOY + 1
        If YZA <> 0 Then prgPostRange.Value = (JOY / YZA) * 100
        DoEvents
    Next
    cmdShowPostRange.Visible = False: picShowPostRange.Visible = False
    rsRefresh
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdPrevious_Click()
    On Error GoTo ErrorCode:

    rsJournal_HD.MovePrevious
    If rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "VENDOR OPENING BALANCE") = False Then Exit Sub
    If xJOURNALTYPE = "GJ" Then ShowReport "GeneralJournal", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "GENERAL JOURNAL PRINTOUT", LOGDATE, False
    If xJOURNALTYPE = "APJ" Then ShowReport "AccountsPayable", "Vouchers", "{Journal_Hd.jno} = '" & txtJNo.Text & "'", "ACCOUNTS PAYABLE JOURNAL PRINTOUT", LOGDATE, False
    If xJOURNALTYPE = "CDJ" Then cmdPrinting.ZOrder 0: picPrinting.ZOrder 0
    LogAudit "V", "VENDOR A/P OPENING", xJOURNALTYPE & "-" & txtVoucherNo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPV_Entry_Click()
    SendToBackPV
    BringToFrontPV
    AddorEdit = "ADD"
    cmdPVDelete.Visible = False
    InitPV_Detail
    On Error Resume Next
    'If JOURNALTYPE = "APJ" Then txtPO_No.SetFocus Else txtMRR_No.SetFocus
    If xJOURNALTYPE = "VPJ" Then
        On Error Resume Next
        txtPO_No.SetFocus
    Else
        On Error Resume Next
        txtMRR_No.SetFocus
    End If
End Sub

Private Sub cmdPVCancel_Click()
    SendToBackPV
    StoreMemVars
End Sub

'Upating Code       : AXP-0713200714:23
Private Sub cmdPVDelete_Click()
    On Error GoTo ErrorCode:

    If labPVID.Caption = "" Then
        MsgBox "Nothing to delete!", vbInformation, "Ma man."
        Exit Sub
    End If
    If xJOURNALTYPE = "VPJ" Then
        If MsgBox("Delete This Vendor Opening PV Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete PV Opening Entry") = vbYes Then
            gconDMIS.Execute "delete from AMIS_PV_Detail where id = " & labPVID.Caption
        End If
    ElseIf xJOURNALTYPE = "CRJ" Then
        If MsgBox("Delete This CRJ Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            gconDMIS.Execute "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype = 'SJ'"
            gconDMIS.Execute "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype = 'CSJ'"
            gconDMIS.Execute "delete from AMIS_CRJ_Detail where id = " & labPVID.Caption
        End If
    Else
        If MsgBox("Delete This CV Detail, Are you Sure?", vbQuestion + vbYesNo, "Delete Journal Entry") = vbYes Then
            gconDMIS.Execute "update AMIS_Journal_HD set PaidStatus = 'N' where VoucherNo = '" & PrevPV_VoucherNo & "' and Jtype='APJ'"
            gconDMIS.Execute "delete from AMIS_CV_Detail where id = " & labPVID.Caption
        End If
    End If
    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO    As String
    Dim J_JVOUCHERNO                              As String
    Dim PV_AMOUNT                                 As Double
    Dim PV_STATUS, PV_ITEMNO                      As String

    J_JVOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    PV_ITEMNO = N2Str2Null(Format(txtPVItemNo.Text, "0000"))
    PV_PONO = N2Str2Null(txtPO_No.Text)
    PV_MRRNO = N2Str2Null(txtMRR_No.Text)                  ' TYPE
    PV_INVNO = N2Str2Null(txtINV_No.Text)                  'NO
    PV_PRODNO = N2Str2Null(txtProd_No.Text)                ' DATE
    PV_AMOUNT = NumericVal(txtPVAmount.Text)               'AMT
    PV_STATUS = "'N'"

    If xJOURNALTYPE = "CDJ" Then
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'")
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " PaidStatus = 'N' " & "," & _
                             " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                             " Balance = (Balance + " & PV_AMOUNT & ") - (AmountPaid - " & PV_AMOUNT & ")" & _
                             " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
        End If
    End If
    If xJOURNALTYPE = "CRJ" Then
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReceiveStatus = 'N' " & "," & _
                             " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                             " Balance = (Balance + " & PV_AMOUNT & ") - (AmountPaid - " & PV_AMOUNT & ")" & _
                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        End If
        Set rsCheckJournal_HD = New ADODB.Recordset
        Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
        If Not rsCheckJournal_HD.EOF And Not rsCheckJournal_HD.BOF Then
            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                             " ReceiveStatus = 'N' " & "," & _
                             " AmountPaid = AmountPaid - " & PV_AMOUNT & "," & _
                             " Balance = (Balance + " & PV_AMOUNT & ") - (AmountPaid - " & PV_AMOUNT & ")" & _
                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
        End If
    End If
    FillDetails
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    cmdPVCancel.Value = True
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPVSave_Click()
    On Error GoTo ErrorCode

    If AddorEdit = "ADD" Then
        Dim rsPV_DetailClone                      As ADODB.Recordset
        Set rsPV_DetailClone = New ADODB.Recordset
        rsPV_DetailClone.Open "select * from AMIS_PV_Detail where JTYPE = 'VPJ' AND PO_NO = " & N2Str2Null(txtPO_No.Text) & " and MRR_NO = " & N2Str2Null(txtMRR_No.Text) & " and INV_NO = " & N2Str2Null(txtINV_No.Text), gconDMIS
        If Not rsPV_DetailClone.EOF And Not rsPV_DetailClone.BOF Then
            MsgBox "PO Number : " & txtPO_No.Text & " with MRR Number : " & txtMRR_No.Text & " and Invoice Number : " & txtINV_No.Text & " already used in this transaction", vbInformation, "Error in PO Number, MRR Number, Invoice Number Validation"
            Exit Sub
        End If
    End If

    Dim PV_PONO, PV_MRRNO, PV_INVNO, PV_PRODNO    As String
    Dim J_JVOUCHERNO, J_JDATE                     As String
    Dim PV_AMOUNT                                 As Double
    Dim PV_STATUS, PV_ITEMNO                      As String

    J_JDATE = N2Date2Null(txtJDate.Text)
    J_JVOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    PV_ITEMNO = N2Str2Null(Format(txtPVItemNo.Text, "0000"))
    PV_PONO = N2Str2Null(txtPO_No.Text)
    PV_MRRNO = N2Str2Null(txtMRR_No.Text)                  ' TYPE
    PV_INVNO = N2Str2Null(txtINV_No.Text)                  ' NO
    PV_PRODNO = N2Str2Null(txtProd_No.Text)                ' DATE
    PV_AMOUNT = NumericVal(txtPVAmount.Text)               ' AMOUNT
    PV_STATUS = "'N'"

    Screen.MousePointer = 11
    If AddorEdit = "ADD" Then
        'Dim rsJournal_HD_APJ As ADODB.Recordset
        'Dim rsPV_Detail_APJ As ADODB.Recordset
        'Set rsJournal_HD_APJ = New ADODB.Recordset
        'Set rsJournal_HD_APJ = gconDMIS.Execute("Select VoucherNo,VendorCode from AMIS_Journal_HD Where Jtype = 'APJ' and VendorCode = '" & txtCode.Text & "' Order By VoucherNo Asc")
        'If Not rsJournal_HD_APJ.EOF And Not rsJournal_HD_APJ.BOF Then
        '   Do While Not rsJournal_HD_APJ.EOF
        '      Set rsPV_Detail_APJ = New ADODB.Recordset
        '      Set rsPV_Detail_APJ = gconDMIS.Execute("Select * from AMIS_PV_Detail Where (Inv_No = " & PV_INVNO & " OR Prod_No = " & PV_PRODNO & ") AND VoucherNo = " & N2Str2Null(rsJournal_HD_APJ!VoucherNo))
        '      If Not rsPV_Detail_APJ.EOF And Not rsPV_Detail_APJ.BOF Then
        '         Screen.MousePointer = 0
        '         MsgBox "Invoice No or Prod No Already Used in PV Number - " & Null2String(rsPV_Detail_APJ!VoucherNo)
        '         Exit Sub
        '      Else
        '         rsJournal_HD_APJ.MoveNext
        '      End If
        '   Loop
        'End If
        If xJOURNALTYPE = "VPJ" Then
            gconDMIS.Execute "insert into AMIS_PV_Detail " & _
                             "(JDATE,JTYPE,VoucherNo,itemno,PO_No,MRR_No,INV_No,PROD_No,Amount,status)" & _
                             " values (" & J_JDATE & ",'VPJ'," & J_JVOUCHERNO & ", " & PV_ITEMNO & ", " & PV_PONO & _
                             ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
                             ", " & PV_STATUS & ")"
        End If
        'If JOURNALTYPE = "CDJ" Then
        '   gconDMIS.Execute "insert into AMIS_CV_Detail " & _
            '                   "(VoucherNo,itemno,PV_VoucherNo,DocDate,DueDate,Amount,status)" & _
            '                   " values (" & J_JVOUCHERNO & ", " & PV_ITEMNO & _
            '                   ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
            '                   ", " & PV_STATUS & ")"
        '   gconDMIS.Execute "update AMIS_Journal_HD set PaidStatus = 'N' where VoucherNo = '" & PrevPV_VoucherNo & "' and Jtype='APJ'"
        '   Set rsCheckJournal_HD = New ADODB.Recordset
        '   Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'")
        '   If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '      If N2Str2Zero(rsCheckJournal_HD!AmountToPay) <= PV_AMOUNT Then
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " PaidStatus = 'Y'," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " [Balance] = " & NumericVal(N2Str2Zero(rsCheckJournal_HD!AmountToPay) - PV_AMOUNT) & _
                  '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
        '      Else
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " PaidStatus = 'N'," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " [Balance] = " & NumericVal(N2Str2Zero(rsCheckJournal_HD!AmountToPay) - PV_AMOUNT) & _
                  '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
        '      End If
        '   End If
        'End If
        'If JOURNALTYPE = "CRJ" Then
        '   gconDMIS.Execute "insert into AMIS_CRJ_Detail " & _
            '                    "(VoucherNo,itemno,INVOICETYPE,INVOICENO,INVOICEDATE,INVOICEAMOUNT,status)" & _
            '                    " values (" & J_JVOUCHERNO & ", " & PV_ITEMNO & _
            '                    ", " & PV_MRRNO & ", " & PV_INVNO & ", " & PV_PRODNO & ", " & PV_AMOUNT & _
            '                    ", " & PV_STATUS & ")"
        '   Set rsCheckJournal_HD = New ADODB.Recordset
        '   Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
        '   If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '      If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " ReceiveStatus = 'Y' " & "," & _
                  '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                  '                          " Balance = Balance - " & PV_AMOUNT & _
                  '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '      Else
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " ReceiveStatus = 'N' " & "," & _
                  '                          " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                  '                          " Balance = Balance - " & PV_AMOUNT & _
                  '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '      End If
        '   Else
        '      Set rsCheckJournal_HD = New ADODB.Recordset
        '      Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'")
        '      If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '         If N2Str2Zero(rsCheckJournal_HD!InvoiceAmt) <= PV_AMOUNT Then
        '            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     '                             " ReceiveStatus = 'Y' " & "," & _
                     '                             " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                     '                             " Balance = Balance - " & PV_AMOUNT & _
                     '                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
        '         Else
        '            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     '                             " ReceiveStatus = 'N' " & "," & _
                     '                             " AmountPaid = AmountPaid + " & PV_AMOUNT & "," & _
                     '                             " Balance = Balance - " & PV_AMOUNT & _
                     '                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'CSJ'"
        '         End If
        '      End If
        '   End If
        'End If
    Else
        If xJOURNALTYPE = "VPJ" Then
            gconDMIS.Execute "update AMIS_PV_Detail set" & _
                             " VoucherNo = " & J_JVOUCHERNO & "," & _
                             " JDATE = " & J_JDATE & "," & _
                             " itemno = " & PV_ITEMNO & "," & _
                             " PO_No = " & PV_PONO & "," & _
                             " MRR_No = " & PV_MRRNO & "," & _
                             " INV_No = " & PV_INVNO & "," & _
                             " PROD_No = " & PV_PRODNO & "," & _
                             " Amount = " & PV_AMOUNT & "," & _
                             " status = " & PV_STATUS & _
                             " where id = " & labPVID.Caption
        End If
        'If JOURNALTYPE = "CDJ" Then
        '   gconDMIS.Execute "update AMIS_CV_Detail set" & _
            '                   " VoucherNo = " & J_JVOUCHERNO & "," & _
            '                   " itemno = " & PV_ITEMNO & "," & _
            '                   " PV_VoucherNo = " & PV_MRRNO & "," & _
            '                   " DocDate = " & PV_INVNO & "," & _
            '                   " DueDate = " & PV_PRODNO & "," & _
            '                   " Amount = " & PV_AMOUNT & "," & _
            '                   " status = " & PV_STATUS & _
            '                   " where id = " & labPVID.Caption
        '   Set rsCheckJournal_HD = New ADODB.Recordset
        '   Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'")
        '   If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '      If N2Str2Zero(rsCheckJournal_HD!AmountToPay) <= PV_AMOUNT Then
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " PaidStatus = 'Y'," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " [Balance] = [Balance] - " & PV_AMOUNT & _
                  '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
        '      Else
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " PaidStatus = 'N' " & "," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " Balance = Balance - " & PV_AMOUNT & _
                  '                          " where VoucherNo = " & PV_MRRNO & " and Jtype = 'APJ'"
        '      End If
        '   End If
        'End If
        'If JOURNALTYPE = "CRJ" Then
        '   gconDMIS.Execute "update AMIS_CRJ_Detail set" & _
            '                    " VoucherNo = " & J_JVOUCHERNO & "," & _
            '                    " itemno = " & PV_ITEMNO & "," & _
            '                    " INVOICETYPE = " & PV_MRRNO & "," & _
            '                    " INVOICENO = " & PV_INVNO & "," & _
            '                    " INVOICEDATE = " & PV_PRODNO & "," & _
            '                    " INVOICEAMOUNT = " & PV_AMOUNT & "," & _
            '                    " status = " & PV_STATUS & _
            '                    " where id = " & labPVID.Caption
        '   gconDMIS.Execute "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype='SJ'"
        '   gconDMIS.Execute "update AMIS_Journal_HD set ReceiveStatus = 'N' where InvoiceType = '" & PrevInvoiceType & "' and InvoiceNo = '" & PrevInvoiceNo & "' and Jtype='CSJ'"
        '   Set rsCheckJournal_HD = New ADODB.Recordset
        '   Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
        '   If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '      If N2Str2Zero(rsCheckJournal_HD!Balance) <= PV_AMOUNT Then
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " ReceiveStatus = 'Y' " & "," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " Balance = Balance - " & PV_AMOUNT & _
                  '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '      Else
        '         gconDMIS.Execute "update AMIS_Journal_HD set" & _
                  '                          " ReceiveStatus = 'N' " & "," & _
                  '                          " AmountPaid = " & PV_AMOUNT & "," & _
                  '                          " Balance = Balance - " & PV_AMOUNT & _
                  '                          " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '      End If
        '   Else
        '      Set rsCheckJournal_HD = New ADODB.Recordset
        '      Set rsCheckJournal_HD = gconDMIS.Execute("Select * from AMIS_Journal_HD Where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'")
        '      If Not rsCheckJournal_Hd.EOF And Not rsCheckJournal_Hd.BOF Then
        '         If N2Str2Zero(rsCheckJournal_HD!Balance) <= PV_AMOUNT Then
        '            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     '                             " ReceiveStatus = 'Y' " & "," & _
                     '                             " AmountPaid = " & PV_AMOUNT & "," & _
                     '                             " Balance = Balance - " & PV_AMOUNT & _
                     '                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '         Else
        '            gconDMIS.Execute "update AMIS_Journal_HD set" & _
                     '                             " ReceiveStatus = 'N' " & "," & _
                     '                             " AmountPaid = " & PV_AMOUNT & "," & _
                     '                             " Balance = Balance - " & PV_AMOUNT & _
                     '                             " where InvoiceType = " & PV_MRRNO & " and InvoiceNo = " & PV_INVNO & " and Jtype = 'SJ'"
        '         End If
        '      End If
        '   End If
        'End If
    End If
    FillDetails
    rsRefresh
    On Error Resume Next
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    If AddorEdit = "ADD" Then cmdPV_Entry_Click Else cmdPVCancel.Value = True
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup                                 As ADODB.Recordset

    If IsNull(txtJNo.Text) = True Then
        MsgBox "Journal No. must not be empty"
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' and jno = '" & txtJNo.Text & "' order by jtype,jno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                '                MsgBox "Journal No. already exist!"
                '                Exit Sub
                Call Get_VoucherNo
            End If
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where invoiceno = '" & txtInvoiceNo.Text & "' and invoicedate = '" & CDate(txtInvoiceDate2.Text) & "' and invoicetype = '" & cboInvoiceType.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgBox "Invoice Transaction already Encoded!"
                Exit Sub
            End If
        Else
            Set rsfindDup = New ADODB.Recordset
            Set rsfindDup = gconDMIS.Execute("SELECT JTYPE,JNO,VOUCHERNO,ID FROM AMIS_JOURNAL_HD WHERE JTYPE = '" & xJOURNALTYPE & "' AND VOUCHERNO = " & N2Str2Null(txtVoucherNo) & "")
            If Not (rsfindDup.BOF And rsfindDup.EOF) Then
                If rsfindDup!ID <> labID Then
                    MsgBox "Voucher No. already exist", vbExclamation, "Info"
                    txtVoucherNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "Error"
        Exit Sub
    End If
    'Set rsProfile = New ADODB.Recordset
    'Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_PROFILE")
    'If Not rsProfile.EOF And Not rsProfile.BOF Then
    '   If Year(txtJDate.Text) = rsProfile!periodyear Then
    '      If Month(txtJDate.Text) <> rsProfile!periodmonth Then
    '         MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
    '         Exit Sub
    '      End If
    '   Else
    '      MsgBox "Warning: Journal Date is not valid in Accounting Period!", vbCritical, "Error!"
    '      Exit Sub
    '   End If
    'End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE             As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE As String
    Dim J_DEBIT, J_CREDIT, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_CHECKNO                       As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE       As String
    Dim J_INVOICETYPE, J_INVOICENO                As String
    Dim J_CHECKDATE, J_BANKCODE                   As String
    Dim J_REFNO, J_REFDATE                        As String
    Dim J_TERMS, J_DEALER                         As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS             As String
    Dim J_TAXRATECODE                             As String
    Dim J_ENTITYCLASS                             As String
    Dim J_TAXBASE                                 As Double
    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_TAXRATECODE = "NULL"
    J_TAXBASE = 0
    'If JOURNALTYPE = "VPJ" Then
    J_INVOICEDATE = N2Str2Null(txtInvoiceDate.Text)
    J_TAXRATECODE = N2Str2Null(Trim(Left(cboATCTAXRATE.Text, 2)))
    J_TAXBASE = NumericVal(txtTaxBaseAmount.Text)
    'If txtPayCode.Text = "CSH" Then
    '   J_BALANCE = 0
    '   J_AMOUNTPAID = NumericVal(txtAmountToPay.Text)
    'Else
    J_BALANCE = NumericVal(txtAmountToPay.Text)
    J_AMOUNTPAID = 0
    'End If
    'ElseIf JOURNALTYPE = "SJ" Then
    '   J_INVOICEDATE = N2Str2Null(txtInvoiceDate2.Text)
    '   J_BALANCE = NumericVal(txtInvoiceAmt.Text)
    '   J_AMOUNTPAID = 0
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   J_INVOICEDATE = N2Str2Null(txtInvoiceDate2.Text)
    '   J_BALANCE = 0
    '   J_AMOUNTPAID = 0
    'ElseIf JOURNALTYPE = "CDJ" Then
    '   J_INVOICEDATE = "NULL"
    '   J_BALANCE = 0
    '   J_AMOUNTPAID = NumericVal(txtAmountToPay.Text)
    'Else
    '   J_INVOICEDATE = "NULL"
    '   J_BALANCE = 0
    '   J_AMOUNTPAID = 0
    'End If
    J_DUEDATE = N2Str2Null(txtDueDate.Text)
    'If JOURNALTYPE = "CRJ" Then
    '   J_PAYTYPE = N2Str2Null(cboInvoiceType.Text)
    'Else
    J_PAYTYPE = N2Str2Null(txtPayCode.Text)
    'End If
    J_JNO = N2Str2Null(txtJNo.Text)
    J_DEBIT = NumericVal(txtTotDebit.Text)
    J_CREDIT = NumericVal(txtTotCredit.Text)
    J_OUTBALANCE = NumericVal(txtOutBalance.Text)
    J_AMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
    J_STATUS = "'N'"

    J_CHECKNO = N2Str2Null(txtCheckNo.Text)
    'If JOURNALTYPE = "SJ" Then
    '   J_TERMS = N2Str2Null(txtTerms.Text)
    '   J_DEALER = N2Str2Null(txtDealer.Text)
    'Else
    J_TERMS = "NULL"
    J_DEALER = "NULL"
    'End If
    'If JOURNALTYPE = "CDJ" Or JOURNALTYPE = "CRJ" Then
    '   J_CHECKDATE = N2Str2Null(txtCheckDate.Text)
    'Else
    J_CHECKDATE = "NULL"
    'End If
    J_BANKCODE = N2Str2Null(txtBankCode.Text)

    'If JOURNALTYPE = "VPJ" Or JOURNALTYPE = "CDJ" Then
    J_CUSTOMERCODE = "'999999'"
    J_VENDORCODE = N2Str2Null(txtCode.Text)
    'Else
    '   J_VENDORCODE = "'999999'"
    '   If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Then
    '      J_CUSTOMERCODE = "'999999'"
    '   Else
    '      J_CUSTOMERCODE = N2Str2Null(txtCustCode.Text)
    '   End If
    'End If
    J_INVOICETYPE = N2Str2Null(SetInvCode(cboInvoiceType.Text))
    'If JOURNALTYPE = "CRJ" Then
    '   If chkNonVat.Value = 1 Then
    '      J_INVOICENO = N2Str2Null("NV" & Format(txtInvoiceNo.Text, "000000"))
    '   Else
    '      J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
    '   End If
    'Else
    J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
    'End If
    J_INVOICEAMT = NumericVal(txtInvoiceAmt.Text)
    J_REFNO = N2Str2Null(txtRefNo.Text)
    J_REFDATE = N2Date2Null(txtRefDate.Text)
    J_ENTITYCLASS = N2Str2Null(xEntityClass)
    'If JOURNALTYPE = "APJ" Then
    If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks.Text))
    'ElseIf JOURNALTYPE = "CDJ" Then
    '   If Trim(txtParticulars.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars.Text))
    'ElseIf JOURNALTYPE = "SJ" Then
    '   If Trim(txtRemarks2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks2.Text))
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   If Trim(txtRemarks2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks2.Text))
    'Else
    '   If Trim(txtParticulars2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtParticulars2.Text))
    'End If
    J_PAIDSTATUS = "'N'"
    J_RECEIVESTATUS = "'N'"
    If txtCode = "" Then
        MsgBox "Please select Vendor", vbInformation, "Message"
        txtNameofVendor.SetFocus
        Exit Sub
    ElseIf txtPayCode = "" Then
        MsgBox "Please select Payment Type", vbInformation, "Message"
        cboPayType.SetFocus
        Exit Sub
    ElseIf txtAmountToPay = "" Or txtAmountToPay.Text = 0 Then
        MsgBox "Amount to pay is empty. Please check.", vbInformation, "Message"
        txtAmountToPay.SetFocus
        Exit Sub
    ElseIf cboCOBAcctName.Text = "" Then
        MsgBox "Please select Account Name", vbInformation, "Message"
        cboCOBAcctName.SetFocus
        Exit Sub
    End If
    If AddorEdit = "ADD" Then
        Dim rsJournal_HDDup                       As ADODB.Recordset
        Set rsJournal_HDDup = New ADODB.Recordset
        Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
        If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
        J_JNO = N2Str2Null(txtJNo.Text)
        gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                         " (jdate,voucherno,jtype,vendorcode,customercode,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,CheckDate,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE,ENTITY_CLASS)" & _
                         " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                         ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_CHECKDATE & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "'," & J_ENTITYCLASS & ")"
        LogAudit "A", "VENDOR A/P OPENING", txtVoucherNo
    Else
        gconDMIS.Execute "update AMIS_Journal_HD set" & _
                         " jdate = " & J_JDATE & "," & _
                         " voucherno = " & J_VOUCHERNO & "," & _
                         " jtype = " & J_JTYPE & "," & _
                         " vendorcode = " & J_VENDORCODE & "," & _
                         " customercode = " & J_CUSTOMERCODE & "," & _
                         " invoicedate = " & J_INVOICEDATE & "," & _
                         " invoicetype = " & J_INVOICETYPE & "," & _
                         " invoiceno = " & J_INVOICENO & "," & _
                         " invoiceamt = " & J_INVOICEAMT & "," & _
                         " duedate = " & J_DUEDATE & "," & _
                         " paytype = " & J_PAYTYPE & "," & _
                         " refno = " & J_REFNO & "," & _
                         " refdate = " & J_REFDATE & ", terms = " & J_TERMS & ", dealer = " & J_DEALER & "," & _
                         " amounttopay = " & J_AMOUNTTOPAY & ", Balance = " & J_BALANCE & ", AmountPaid = " & J_AMOUNTPAID & "," & _
                         " jno = " & J_JNO & "," & _
                         " debit = " & J_DEBIT & "," & _
                         " credit = " & J_CREDIT & "," & _
                         " outbalance = " & J_OUTBALANCE & "," & _
                         " CheckNo = " & J_CHECKNO & ", " & _
                         " CheckDate = " & J_CHECKDATE & ", " & _
                         " BankCode = " & J_BANKCODE & ", " & _
                         " status = " & J_STATUS & ", PaidStatus = " & J_PAIDSTATUS & ", ReceiveStatus = " & J_RECEIVESTATUS & "," & _
                         " remarks = " & J_REMARKS & ", USERCODE = '" & LOGCODE & "', LASTUPDATE = '" & LOGDATE & "',ENTITY_CLASS=" & J_ENTITYCLASS & "" & _
                    " where id = " & labID.Caption
        gconDMIS.Execute "update AMIS_Journal_Det set" & _
                         " jtype = " & J_JTYPE & "," & _
                         " VOUCHERNO = " & J_VOUCHERNO & "," & _
                         " jdate = " & J_JDATE & "," & _
                         " USERCODE = '" & LOGCODE & "'," & _
                         " LASTUPDATE = '" & LOGDATE & "'," & _
                         " jno = " & J_JNO & _
                         " where jtype = '" & PrevJType & "' and jno = '" & PrevJNo & "'"
        LogAudit "E", "VENDOR A/P OPENING", txtVoucherNo
    End If
    If Trim(txtCOBAcctNo.Text) <> "" Then
        Dim rsCOB_Journal_Det                     As ADODB.Recordset
        Set rsCOB_Journal_Det = New ADODB.Recordset
        Set rsCOB_Journal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det Where Jtype = 'VPJ' and JNO = " & J_JNO)
        If Not rsCOB_Journal_Det.EOF And Not rsCOB_Journal_Det.BOF Then
            gconDMIS.Execute "UPDATE AMIS_JOURNAL_DET SET" & _
                             " JITEMNO = '0001'," & _
                             " JTYPE = " & J_JTYPE & "," & _
                             " JDATE = " & J_JDATE & "," & _
                             " USERCODE = '" & LOGCODE & "'," & _
                             " LASTUPDATE = '" & LOGDATE & "'," & _
                             " ACCT_CODE = " & N2Str2Null(txtCOBAcctNo.Text) & "," & _
                             " ACCT_NAME = " & N2Str2Null(cboCOBAcctName.Text) & "," & _
                             " JNO = " & J_JNO & _
                             " WHERE JTYPE = 'VPJ' AND JNO = '" & PrevJNo & "'"
        Else
            gconDMIS.Execute "INSERT INTO AMIS_JOURNAL_DET (JITEMNO,JTYPE,JDATE,VOUCHERNO,JNO,ACCT_CODE,ACCT_NAME)" & _
                             " VALUES ('0001','VPJ'," & J_JDATE & "," & J_VOUCHERNO & "," & J_JNO & "," & N2Str2Null(txtCOBAcctNo.Text) & "," & N2Str2Null(cboCOBAcctName.Text) & ")"
        End If
    End If
    rsRefresh
    rsJournal_HD.Find "jno = " & J_JNO
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
    ShowVBError
End Sub

Private Sub cmdSelect_Click()
    Set frmNewEntity = New frmEntity
    frmNewEntity.Show 1
End Sub

'Upating Code       : AXP-0713200714:21
Private Sub cmdUnPost_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_UnPost", "VENDOR OPENING BALANCE") = False Then Exit Sub
    If xJOURNALTYPE = "VPJ" Then
        Dim rsCV_Detail                           As ADODB.Recordset
        Set rsCV_Detail = New ADODB.Recordset
        Set rsCV_Detail = gconDMIS.Execute("Select * from AMIS_CV_Detail where PV_VoucherNo = '" & txtVoucherNo.Text & "' and JType = 'VPJ' and status <> 'C'")
        If Not rsCV_Detail.EOF And Not rsCV_Detail.BOF Then
            MsgBox "Warning: This AP Journal is already link to Cash Disbursement Voucher No. " & Null2String(rsCV_Detail!VOUCHERNO) & vbCrLf & _
                   "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
            Exit Sub
        End If
    End If
    Screen.MousePointer = 11
    gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
    gconDMIS.Execute "UPDATE AMIS_PV_DETAIL SET STATUS='N' WHERE JTYPE='" & xJOURNALTYPE & "' AND VOUCHERNO =" & N2Str2Null(txtVoucherNo.Text) & ""

    Dim xAPVOUCHERNO                              As String
    xAPVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
    If xJOURNALTYPE = "VPJ" Then
        gconDMIS.Execute "DELETE FROM AMIS_AP WHERE VOUCHERNO = " & N2Str2Null(xAPVOUCHERNO) & ""
    End If

    rsRefresh
    rsJournal_HD.Find "id = " & labID.Caption
    StoreMemVars
    Screen.MousePointer = 0
    LogAudit "U", "VENDOR A/P OPENING", txtVoucherNo
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub FillGrid()
    Dim rsChartAccount2                           As ADODB.Recordset
    lstAccounts.Sorted = False: lstAccounts.ListItems.Clear
    Set rsChartAccount2 = New ADODB.Recordset
    Set rsChartAccount2 = gconDMIS.Execute("select acctcode,UPPER(Description),Accttype,ID from AMIS_ChartAccount order by acctcode asc")
    If Not (rsChartAccount2.EOF And rsChartAccount2.BOF) Then
        lstAccounts.Enabled = True
        Listview_Loadval Me.lstAccounts.ListItems, rsChartAccount2
        lstAccounts.Refresh
    Else
        lstAccounts.Enabled = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        'If Me.ActiveControl.Name = "cboAcct_Code" And cboAcct_Code.Text = "" Then
        '   fraFindAccount.Visible = True
        '   cmdFindAccount.Visible = True
        '   cmdFindAccount.ZOrder 0
        '   fraFindAccount.ZOrder 0
        '   fraFindAccount.Enabled = True
        '   DoEvents
        '   txtSearch.SetFocus
        'ElseIf Me.ActiveControl.Name = "cboGJAccountNo" And cboGJAccountNo.Text = "" Then
        '   fraFindAccount.Visible = True
        '   cmdFindAccount.Visible = True
        '   cmdFindAccount.ZOrder 0
        '   fraFindAccount.ZOrder 0
        '   fraFindAccount.Enabled = True
        '   DoEvents
        '   txtSearch.SetFocus
        'ElseIf Me.ActiveControl.Name = "cboAccount" Then
        '   OkAccount
        'ElseIf Me.ActiveControl.Name = "txtPO_No" And txtPO_No.Text = "" Then
        '   txtPO_No.SetFocus
        'ElseIf Me.ActiveControl.Name = "txtCredit" And SetAcctType(cboAcct_Code.Text) = "C" And Val(txtCredit.Text) <= 0 And Val(txtDebit.Text) <= 0 Then
        '   txtCredit.SetFocus
        'ElseIf Me.ActiveControl.Name = "txtDebit" And SetAcctType(cboAcct_Code.Text) = "D" And Val(txtDebit.Text) <= 0 And Val(txtCredit.Text) <= 0 Then
        '   txtDebit.SetFocus
        'ElseIf Me.ActiveControl.Name = "txtGJCredit" And SetAcctType(cboGJAccountNo.Text) = "C" And Val(txtGJCredit.Text) <= 0 And Val(txtGJDebit.Text) <= 0 Then
        '   txtGJCredit.SetFocus
        'ElseIf Me.ActiveControl.Name = "txtGJDebit" And SetAcctType(cboGJAccountNo.Text) = "D" And Val(txtGJDebit.Text) <= 0 And Val(txtGJCredit.Text) <= 0 Then
        '   txtGJDebit.SetFocus
        'ElseIf Me.ActiveControl.Name = "txtGrossAmt" And NumericVal(txtGrossAmt.Text) <= 0 Then
        '   txtGrossAmt.SetFocus
        'Else
        MoveKeyPress KeyCode
        'End If
    Case vbKeyEscape
        'If fraFindAccount.Visible = True Then
        '   If Me.ActiveControl.Name = "txtSearch" Then
        '      SendToBack
        '      SendToBackPV
        '      SendToBackGJ
        '      SendToBackTemplates
        '      StoreMemvars
        '   Else
        '      txtSearch.SetFocus
        '   End If
        'Else
        '   If Picture1.Visible = True Then
        '      If Me.ActiveControl.Name = "txtSearchTemplates" Then
        '         SendToBack
        SendToBackPV
        '         SendToBackGJ
        '         SendToBackTemplates
        StoreMemVars
        '      ElseIf Me.ActiveControl.Name = "lstTemplates" Then
        '         txtSearchTemplates.SetFocus
        '      Else
        '         SendToBack
        '         SendToBackPV
        '         SendToBackGJ
        '         SendToBackTemplates
        '         StoreMemvars
        '      End If
        '   End If
        'End If
    Case vbKeyF3
        'If Picture1.Visible = True Then
        '   If Null2String(rsJOURNAL_HD!Status) = "C" Then
        '      MsgBox "Journals are Already Cancelled" & vbCrLf & _
               '             "and cannot be Change", vbInformation, "Edit Not Allowed!"
        '   ElseIf Null2String(rsJOURNAL_HD!Status) = "P" Then
        '      MsgBox "Journals are Already Posted" & vbCrLf & _
               '             "and cannot be Change", vbInformation, "Edit Not Allowed!"
        '   Else
        '      JournalTAB.Tab = 0
        '      If JOURNALTYPE = "GJ" Or JOURNALTYPE = "OPB" Or JOURNALTYPE = "ADJ" Then
        '         cmdGJEntry_Click
        '      Else
        '         cmdAddJournal_Click
        '      End If
        '   End If
        'End If
    Case vbKeyF4
        'If JOURNALTYPE <> "SJ" Then
        'If Picture1.Visible = True Then
        'If Null2String(rsJOURNAL_HD!Status) = "C" Then
        '   MsgBox "Journals are Already Cancelled" & vbCrLf & _
            '          "and cannot be Change", vbInformation, "Edit Not Allowed!"
        'ElseIf Null2String(rsJOURNAL_HD!Status) = "P" Then
        '   MsgBox "Journals are Already Posted" & vbCrLf & _
            '          "and cannot be Change", vbInformation, "Edit Not Allowed!"
        'Else
        JournalTAB.Tab = 1
        cmdPV_Entry_Click
        'End If
        'End If
        'End If
    Case vbKeyF5
        cmdPost.Value = True
    Case vbKeyF6
        cmdUnPost.Value = True
    Case vbKeyF7
        cmdCancelCO.Value = True
    Case vbKeyF8
        'If SearchBy = "NAME" Then
        '   SearchBy = "CODE": fraFindAccount.Caption = "Search Accounts by Account Code"
        'Else
        '   SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
        'End If
    Case vbKeyF9
        'If Picture1.Visible = True Then
        '   If Null2String(rsJOURNAL_HD!Status) = "C" Then
        '      MsgBox "Journals are Already Cancelled" & vbCrLf & _
               '             "and cannot be Change", vbInformation, "Edit Not Allowed!"
        '   ElseIf Null2String(rsJOURNAL_HD!Status) = "P" Then
        '      MsgBox "Journals are Already Posted" & vbCrLf & _
               '             "and cannot be Change", vbInformation, "Edit Not Allowed!"
        '   Else
        '      JournalTAB.Tab = 0
        '      fraFindAccount.ZOrder 1: cmdFindAccount.ZOrder 1
        '      fraFindAccount.Visible = False: cmdFindAccount.Visible = False: BringToFrontTemplates
        '      txtSearchTemplates.SetFocus
        '   End If
        'End If
    Case vbKeyF11
        'SendToBack
        'SendToBackPV
        'SendToBackGJ
        'SendToBackTemplates
        'cmdShowPostRange.Visible = True: picShowPostRange.Visible = True
        'picShowPostRange.Enabled = True
        'cmdShowPostRange.ZOrder 0: picShowPostRange.ZOrder 0
        'On Error Resume Next
        'txtFromVNo.SetFocus
    Case vbKeyF12
        If Null2String(rsJournal_HD!Status) = "C" Then
            If MsgBox("Are you sure you want to Un-Cancel this Transaction?", vbQuestion + vbYesNo, "Un-Cancel Journal") = vbYes Then
                Screen.MousePointer = 11
                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                rsRefresh
                rsJournal_HD.Find "id = " & labID.Caption
                StoreMemVars
                Screen.MousePointer = 0
            End If
        End If
        If Null2String(rsJournal_HD!Status) = "P" Then
            If Function_Access(LOGID, "Acess_Post", "VENDOR OPENING BALANCE") = False Then Exit Sub

            If MsgBox("Are you sure you want to Un-Post this Transaction?", vbQuestion + vbYesNo, "Un-Post Journal") = vbYes Then
                Screen.MousePointer = 11
                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
                rsRefresh
                rsJournal_HD.Find "id = " & labID.Caption
                StoreMemVars
                Screen.MousePointer = 0
            End If
        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
    If Shift = 2 Then
        'If KeyCode = vbKeyA Then cmdAddAccount_Click
        'If KeyCode = vbKeyJ Then
        '   If JournalTAB.Tab = 1 Then JournalTAB.Tab = 0
        'End If
        'If KeyCode = vbKeyD Then
        '   If JournalTAB.Tab = 0 Then JournalTAB.Tab = 1
        'End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBack: SendToBackPV: SendToBackGJ: SendToBackTemplates
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    picPayables.Top = 1230
    picDisbursement.Top = 1230
    picReceivable.Top = 420
    Frame1.Top = 90
    'If JOURNALTYPE = "VPJ" Then
    chkNonVat.Visible = False
    Me.Caption = "VENDOR OPENING BALANCE - DATA ENTRY"
    labSupplierPayTo = "Supplier Code"
    picGJ.Visible = False: picPayables.Visible = True: picPayables.ZOrder 0: picPayables.Enabled = True
    picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
    picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
    labPV1.Caption = "PO Number": labPV2.Caption = "MRR Number"
    labPV3.Caption = "Invoice Number": labPV4.Caption = "Product Number"
    labTax.Caption = "Input Tax": RefCRJ.Visible = False
    cboATCTAXRATE.Clear
    cboATCTAXRATE.AddItem " 1 %"
    cboATCTAXRATE.AddItem " 2 %"
    cboATCTAXRATE.AddItem " 5 %"
    cboATCTAXRATE.AddItem "10 %"
    'ElseIf JOURNALTYPE = "CDJ" Then
    '   chkNonVat.Visible = False
    '   fraComp.Visible = False
    '   Me.Caption = "CASH DISBURSEMENT JOURNAL DATA ENTRY"
    '   labSupplierPayTo = "Pay To": RefCRJ.Visible = False
    '   picGJ.Visible = False: labDueDate.Visible = False: txtDueDate.Visible = False
    '   picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
    '   picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
    '   picDisbursement.Visible = True: picDisbursement.ZOrder 0: picDisbursement.Enabled = True
    'ElseIf JOURNALTYPE = "SJ" Then
    '   chkNonVat.Visible = False
    '   JournalTAB.TabEnabled(1) = False: labBankName.Visible = False: cboBankName2.Visible = False
    '   'labParticulars.Top = 960: 'txtRemarks2.Top = 930: txtRemarks2.Height = 1125
    '   Me.Caption = "SALES JOURNAL DATA ENTRY"
    '   labSupplierPayTo = "Supplier Code"
    '   labType.Caption = "Invoice Type": LabNo.Caption = "Invoice No."
    '   labDate.Caption = "Invoice Date": labAmt.Caption = "Invoice Amt."
    '   picGJ.Visible = False: RefCRJ.Visible = True
    '   picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
    '   picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
    '   picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
    '   labTax.Caption = "Output Tax"
    'ElseIf JOURNALTYPE = "CRJ" Then
    '   chkNonVat.Visible = True
    '   txtInvoiceNo.Left = 2040
    '   txtInvoiceNo.Width = 975
    '   fraComp.Visible = False
    '   Me.Caption = "CASH RECEIPTS JOURNAL DATA ENTRY"
    '   picGJ.Visible = False: RefCRJ.Visible = False
    '   labType.Caption = "Payment Type": LabNo.Caption = "O.R. No."
    '   labDate.Caption = "O.R. Date": labAmt.Caption = "O.R. Amount": labTerms.Visible = False
    '   picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
    '   picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
    '   picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
    '   labPV1.Caption = "Voucher No": txtPO_No.Enabled = False
    '   labPV2.Caption = "Invoice Type": labPV3.Caption = "Invoice No.": labPV4.Caption = "Invoice Date"
    '   lstPV_Detail.ColumnHeaders(2).Text = "Invoice Type"
    '   lstPV_Detail.ColumnHeaders(3).Text = "Invoice No."
    '   lstPV_Detail.ColumnHeaders(4).Text = "Invoice Date"
    '   lstPV_Detail.ColumnHeaders(5).Text = "Invoice Amt."
    'ElseIf JOURNALTYPE = "GJ" Then
    '   chkNonVat.Visible = False
    '   fraComp.Visible = False: RefCRJ.Visible = False
    '   Me.Caption = "GENERAL JOURNAL DATA ENTRY"
    '   picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
    '   labOutBalance.Visible = False: txtOutBalance.Visible = False
    '   picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
    '   picPayables.Enabled = False: picDisbursement.Enabled = False
    '   txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    'ElseIf JOURNALTYPE = "ADJ" Then
    '   chkNonVat.Visible = False
    '   fraComp.Visible = False: RefCRJ.Visible = False
    '   Me.Caption = "AUDIT ADJUSTMENTS DATA ENTRY"
    '   picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
    '   labOutBalance.Visible = False: txtOutBalance.Visible = False
    '   picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
    '   picPayables.Enabled = False: picDisbursement.Enabled = False
    '   txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    'ElseIf JOURNALTYPE = "OPB" Then
    '   chkNonVat.Visible = False
    '   fraComp.Visible = False: RefCRJ.Visible = False
    '   Label3.Caption = "Ref. No.": Label5.Caption = "Ref. Date"
    '   Me.Caption = "OPENING BALANCES"
    '   picGJ.Visible = True: picGJ.ZOrder 0: txtParticulars2.Locked = True
    '   labOutBalance.Visible = False: txtOutBalance.Visible = False
    '   picReceivable.Visible = False: picReceivable.ZOrder 1: picReceivable.Enabled = False
    '   picPayables.Enabled = False: picDisbursement.Enabled = False
    '   txtCode.Enabled = False: cboNameofVendor.Enabled = False: txtDueDate.Enabled = False
    'End If
InitGrid:     InitCbo: initMemvars: txtSearch.Text = "": txtSearchTemplates.Text = ""

    rsRefresh
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveLast
    End If
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub JournalTAB_Click(PreviousTab As Integer)
    If Picture1.Visible = True Then
        If JournalTAB.Tab = 0 Then
            If lstDetails.ListItems.Count > 0 And lstDetails.Enabled = True Then: lstDetails.SetFocus
        End If
        If JournalTAB.Tab = 1 Then
            If lstPV_Detail.ListItems.Count > 0 And lstPV_Detail.Enabled = True Then: lstPV_Detail.SetFocus
        End If
    End If
End Sub

Private Sub lstAccounts_DblClick()
    labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
    OkAccount
End Sub

Private Sub lstAccounts_GotFocus()
    On Error Resume Next
    labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
End Sub

Private Sub lstAccounts_ItemClick(ByVal Item As MSComctlLib.ListItem)
    labAccountCode.Caption = Item: cboAcct_Code.Text = Item
End Sub

Private Sub lstAccounts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        labAccountCode.Caption = lstAccounts.SelectedItem: cboAcct_Code.Text = lstAccounts.SelectedItem
        OkAccount
    End If
End Sub

Private Sub lstDetails_DblClick()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            If kcnt > 0 Then
                AddorEdit = "EDIT"
                cmdJournalDelete.Visible = True
                BringToFront
                StoreJournalEntry (lstDetails.SelectedItem.SubItems(5))
                On Error Resume Next
                txtGrossAmt.SetFocus
                OkAccountSetCursor
            End If
        End If
    End If
End Sub

Private Sub lstDetails_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstDetails_DblClick
        If Me.ActiveControl.Name = "txtDebit" Then SendKeys MOVEUP
    End If
End Sub

Private Sub lstDetails_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If kcnt > 0 Then
            If Null2String(rsJournal_HD!Status) = "C" Then
                MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            ElseIf Null2String(rsJournal_HD!Status) = "P" Then
                MsgBox "Journals are Already Posted" & vbCrLf & _
                       "and cannot be Change", vbInformation, "Edit Not Allowed!"
            Else
                AddorEdit = "EDIT"
                cmdJournalDelete.Visible = True
                BringToFront
                StoreJournalEntry (lstDetails.SelectedItem.SubItems(5))
                cmdJournalDelete_Click
            End If
        End If
    End If
End Sub

Private Sub lstGJ_DblClick()
    If Null2String(rsJournal_HD!Status) = "C" Then
        MsgBox "Transactions are Already Cancelled" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
    ElseIf Null2String(rsJournal_HD!Status) = "P" Then
        MsgBox "Journals are Already Posted" & vbCrLf & _
               "and cannot be Change", vbInformation, "Edit Not Allowed!"
    Else
        AddorEdit = "EDIT"
        cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
    End If
End Sub

Private Sub lstGJ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstGJ_DblClick
End Sub

Private Sub lstGJ_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        AddorEdit = "EDIT"
        cmdGJDelete.Visible = True
        BringToFrontGJ
        On Error Resume Next
        StoreGJEntry (lstGJ.SelectedItem.SubItems(5))
        cmdGJDelete_Click
    End If
End Sub

Private Sub lstPV_Detail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPV_Detail
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

Private Sub lstPV_Detail_DblClick()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        If Null2String(rsJournal_HD!Status) = "C" Then
            MsgBox "Transactions are Already Cancelled" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            MsgBox "Journals are Already Posted" & vbCrLf & _
                   "and cannot be Change", vbInformation, "Edit Not Allowed!"
        Else
            If Jcnt > 0 Then
                AddorEdit = "EDIT"
                cmdPVDelete.Visible = True
                BringToFrontPV
                StorePVEntry (lstPV_Detail.SelectedItem.SubItems(6))
            End If
        End If
    End If
End Sub

Private Sub lstPV_Detail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lstPV_Detail_DblClick
End Sub

Private Sub lstTemplates_DblClick()
    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub lstTemplates_KeyPress(KeyAscii As Integer)
    SendToBack
    SendToBackPV
    SendToBackGJ
    SendToBackTemplates
    On Error Resume Next
    If KeyAscii = 13 Then InsertAccountEntries lstTemplates.SelectedItem.SubItems(1)
End Sub

Private Sub optPrintCheck_Click()
    If optPrintCheck.Value = True Then
        picPrintCheck.Enabled = True
    Else
        picPrintCheck.Enabled = False
    End If
End Sub

Private Sub optPrintVoucher_Click()
    picPrintCheck.Enabled = False
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then labPosted.Visible = False Else labPosted.Visible = True
    End If
End Sub

Private Sub txtAmountToPay_GotFocus()
    If Val(txtAmountToPay.Text) = 0 Then txtAmountToPay.Text = "" Else txtAmountToPay.Text = NumericVal(txtAmountToPay.Text)
End Sub

Private Sub txtAmountToPay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtAmountToPay_LostFocus()
'If txtAmountToPay.Text = "" Then txtAmountToPay.Text = "0.00" Else txtAmountToPay.Text = Format(txtAmountToPay.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtBankCode_Change()
    If xJOURNALTYPE = "APJ" Or xJOURNALTYPE = "CDJ" Then cboBankName.Text = SetBankName(txtBankCode.Text)
    If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Then cboBankName2.Text = SetBankName(txtBankCode.Text)
End Sub

Private Sub txtBankCode_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtCheckDate_GotFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCheckDate_LostFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
End Sub

'Private Sub txtCode_Change()
'cboNameofVendor.Text = SetVendorName(txtCode.Text)
'txtAddress.Caption = SetVendorAddress(txtCode.Text)
'End Sub

Private Sub txtCredit_GotFocus()
    If NumericVal(txtDebit.Text) = 0 Then
        If Val(txtCredit.Text) = 0 Then
            If NumericVal(txtNetAmt.Text) > 0 Then
                txtDebit.Text = ZERO
                txtCredit.Text = NumericVal(txtNetAmt.Text)
            Else
                If OUTBALANCE > 0 And TOTDEBIT > 0 Then
                    txtCredit.Text = OUTBALANCE
                    txtDebit.Text = ZERO
                Else
                    txtCredit.Text = ""
                End If
            End If
        Else
            txtCredit.Text = NumericVal(txtCredit.Text)
        End If
    Else
        txtCredit.Text = ZERO
    End If
End Sub

Private Sub txtCredit_LostFocus()
    If txtCredit.Text = "" Then txtCredit.Text = 0
End Sub

'Private Sub txtCustCode_Change()
'cboCustName.Text = SetCustomerName(txtCustCode.Text)
'End Sub

Private Sub txtDebit_GotFocus()
    If NumericVal(txtCredit.Text) = 0 Then
        If NumericVal(txtDebit.Text) = 0 Then
            If NumericVal(txtNetAmt.Text) > 0 Then
                txtDebit.Text = NumericVal(txtNetAmt.Text)
            Else
                If txtAcct_Name.Text = "OUTPUT TAX" And xJOURNALTYPE = "SJ" Then
                    txtDebit.Text = ZERO: txtCredit.Text = OUTBALANCE
                Else
                    If OUTBALANCE > 0 And TOTCREDIT > 0 Then
                        txtCredit.Text = ZERO: txtDebit.Text = OUTBALANCE
                    Else
                        txtDebit.Text = ""
                    End If
                End If
            End If
        Else
            txtDebit.Text = NumericVal(txtDebit.Text)
        End If
    Else
        txtDebit.Text = ZERO
    End If
End Sub

Private Sub txtDebit_LostFocus()
    If txtDebit.Text = "" Then txtDebit.Text = 0
End Sub

Private Sub txtGJAccountName_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtGJAccountParticulars_GotFocus()
    If txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!" Then txtGJAccountParticulars.Text = ""
End Sub

Private Sub txtGJAccountParticulars_LostFocus()
    If txtGJAccountParticulars.Text = "" Then txtGJAccountParticulars.Text = "Pls Type Your Remarks Here!"
End Sub

Private Sub txtGJCredit_GotFocus()
'If Val(txtGJCredit.Text) = 0 Then
'   If OUTBALANCE > 0 And TOTDEBIT > 0 Then
'      txtGJCredit.Text = OUTBALANCE
'      txtGJDebit.Text = ZERO
'   Else
'      txtGJCredit.Text = ""
'   End If
'Else
'   txtGJCredit.Text = NumericVal(txtGJCredit.Text)
'End If
    If NumericVal(txtGJDebit.Text) = 0 Then
        If Val(txtGJCredit.Text) = 0 Then
            If OUTBALANCE > 0 And TOTDEBIT > 0 Then
                txtGJCredit.Text = OUTBALANCE
                txtGJDebit.Text = ZERO
            Else
                txtGJCredit.Text = ""
            End If
        Else
            txtGJCredit.Text = NumericVal(txtGJCredit.Text)
        End If
    Else
        txtGJCredit.Text = ZERO
    End If
End Sub

Private Sub txtGJCredit_LostFocus()
    If txtGJCredit.Text = "" Then txtGJCredit.Text = 0
End Sub

Private Sub txtGJDebit_GotFocus()
'If NumericVal(txtGJDebit.Text) = 0 Then
'   If OUTBALANCE > 0 And TOTCREDIT > 0 Then
'      txtGJCredit.Text = ZERO
'      txtGJDebit.Text = OUTBALANCE
'   Else
'      txtGJDebit.Text = ""
'   End If
'Else
'   txtGJDebit.Text = NumericVal(txtGJDebit.Text)
'End If
    If NumericVal(txtGJCredit.Text) = 0 Then
        If NumericVal(txtGJDebit.Text) = 0 Then
            If OUTBALANCE > 0 And TOTCREDIT > 0 Then
                txtGJCredit.Text = ZERO: txtGJDebit.Text = OUTBALANCE
            Else
                txtGJDebit.Text = ""
            End If
        Else
            txtGJDebit.Text = NumericVal(txtGJDebit.Text)
        End If
    Else
        txtGJDebit.Text = ZERO
    End If
End Sub

Private Sub txtGJDebit_LostFocus()
    If txtGJDebit.Text = "" Then txtGJDebit.Text = 0
End Sub

Private Sub txtGrossAmt_Change()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtTax.Text = Round((NumericVal(txtGrossAmt.Text) / 1.12) * 0.12, 2)
        txtNetAmt.Text = NumericVal(txtGrossAmt.Text) - NumericVal(txtTax.Text)
    Else
        txtTax.Text = 0: txtNetAmt.Text = 0
    End If
End Sub

Private Sub txtGrossAmt_GotFocus()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtGrossAmt.Text = NumericVal(txtGrossAmt.Text)
    Else
        txtGrossAmt.Text = ""
    End If
End Sub

Private Sub txtGrossAmt_LostFocus()
    If NumericVal(txtGrossAmt.Text) > 0 Then
        txtGrossAmt.Text = ToDoubleNumber(txtGrossAmt.Text)
    End If
End Sub

Private Sub txtINV_No_GotFocus()
    If xJOURNALTYPE = "CDJ" Then
        If txtMRR_No.Text = "" Then On Error Resume Next: txtMRR_No.SetFocus
    End If
End Sub

Private Sub txtINV_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtInvoiceAmt_GotFocus()
    txtInvoiceAmt.Text = NumericVal(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtInvoiceAmt_LostFocus()
    txtInvoiceAmt.Text = ToDoubleNumber(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceDate_Change()
    If IsDate(txtInvoiceDate.Text) = True Then
        txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
    End If
End Sub

Private Sub txtInvoiceDate_GotFocus()
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate2_GotFocus()
    txtInvoiceDate2.Text = Format(txtInvoiceDate2.Text, "MM-DD-YYYY")
End Sub

Private Sub txtInvoiceDate2_LostFocus()
    If txtInvoiceDate2.Text <> "" Then
        If IsDate(txtInvoiceDate2.Text) = True Then
            txtInvoiceDate2.Text = Format(txtInvoiceDate2.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid Invoice Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtInvoiceDate.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtJDate_GotFocus()
    txtJDate.Text = Format(txtJDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtJDate_LostFocus()
    txtJDate.Text = Format(txtJDate.Text, "DD-MMM-YY")
    If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Then
        On Error Resume Next
        cboCustName.SetFocus
    Else
        On Error Resume Next
        txtParticulars2.SetFocus
    End If
End Sub

Private Sub txtMRR_No_Change()
    If xJOURNALTYPE = "CDJ" Then
        Set rsJournal_HD2 = New ADODB.Recordset
        Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and (JType = 'VPJ' OR JType = 'APJ')")
        If Not rsJournal_HD2.EOF And Not rsJournal_HD2.BOF Then
            txtINV_No.Text = Null2String(rsJournal_HD2!JDate)
            txtProd_No.Text = Null2String(rsJournal_HD2!duedate)
            txtPVAmount.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD2!amounttopay))
        Else
            txtINV_No.Text = ""
            txtProd_No.Text = ""
            txtPVAmount.Text = ZERO
        End If
    End If
End Sub

Private Sub txtMRR_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If xJOURNALTYPE = "CDJ" Then
        If KeyAscii = 13 Then
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchAPJ2.Show vbModal
        End If
    End If
    If xJOURNALTYPE = "CRJ" Then
        If KeyAscii = 13 Then
            SEARCH_TAB = 0
            If Trim(txtMRR_No.Text) = "" Then frmAMISSearchSJ2.Show vbModal
        End If
    End If
End Sub

Private Sub txtParticulars_GotFocus()
    If txtParticulars.Text = "Pls Type Your Message Here!" Then txtParticulars.Text = ""
End Sub

Private Sub txtParticulars_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtParticulars.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtParticulars_LostFocus()
    If txtParticulars.Text = "" Then txtParticulars.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtParticulars2_GotFocus()
    If txtParticulars2.Text = "Pls Type Your Message Here!" Then txtParticulars2.Text = ""
End Sub

Private Sub txtParticulars2_LostFocus()
    If txtParticulars2.Text = "" Then txtParticulars2.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtPayCode_Change()
    cboPayType.Text = SetPayDesc(txtPayCode.Text)
End Sub

Private Sub txtPayCode_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPayCode_LostFocus()
    cboPayType.Text = SetPayDesc(txtPayCode.Text)
End Sub

Private Sub txtPO_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtProd_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtPVAmount_GotFocus()
    If NumericVal(txtPVAmount.Text) = 0 Then txtPVAmount.Text = ""
End Sub

Private Sub txtPVAmount_LostFocus()
    If NumericVal(txtPVAmount.Text) > 0 Then txtPVAmount.Text = ToDoubleNumber(txtPVAmount.Text)
End Sub

Private Sub txtRefDate_GotFocus()
    txtRefDate.Text = Format(txtRefDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtRefDate_LostFocus()
    If txtRefDate.Text <> "" Then
        If IsDate(txtRefDate.Text) = True Then
            txtRefDate.Text = Format(txtRefDate.Text, "DD-MMM-YY")
        Else
            MsgBoxXP "Invalid Reference Date!", "Error", XP_OKOnly, msg_Exclamation
            On Error Resume Next
            txtRefDate.SetFocus
            Exit Sub
        End If
    End If
    If xJOURNALTYPE = "CRJ" Then
        On Error Resume Next
        cboBankName2.SetFocus
    End If
End Sub

Private Sub txtRefNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtRemarks.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtRemarks_LostFocus()
    If txtRemarks.Text = "" Then txtRemarks.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtRemarks2_GotFocus()
    If txtRemarks2.Text = "Pls Type Your Message Here!" Then txtRemarks2.Text = ""
End Sub

Private Sub txtRemarks2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If txtRemarks2.Text = "" Then
            SendKeys "+{TAB}^{HOME}+{END}"
        End If
    End If
End Sub

Private Sub txtRemarks2_LostFocus()
    If txtRemarks2.Text = "" Then txtRemarks2.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtSearch.Text)
    End If
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstAccounts.ListItems.Count > 0 And lstAccounts.Enabled = True Then: lstAccounts.SetFocus
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSearchTemplates_Change()
    If Trim(txtSearchTemplates.Text) = "" Then
        FillTemplates
    Else
        FillSearchTemplates (txtSearchTemplates.Text)
    End If
End Sub

Private Sub txtSearchTemplates_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        If lstTemplates.ListItems.Count > 0 And lstTemplates.Enabled = True Then: lstTemplates.SetFocus
    End If
End Sub

Private Sub txtTax_Change()
    txtNetAmt.Text = ToDoubleNumber(NumericVal(txtGrossAmt.Text) - NumericVal(txtTax.Text))
End Sub

Private Sub txtTax_GotFocus()
    If Val(txtTax.Text) = 0 Then txtTax.Text = ""
End Sub

Private Sub txtTax_LostFocus()
    If SetAcctType(cboAcct_Code.Text) = "C" Then
        txtDebit.Text = ZERO
        txtCredit.Text = ToDoubleNumber(txtNetAmt.Text)
        On Error Resume Next
        txtCredit.SetFocus
    Else
        If xJOURNALTYPE = "SJ" And txtAcct_Name.Text = "OUTPUT TAX" Then
            txtDebit.Text = ZERO
            txtCredit.Text = ToDoubleNumber(txtNetAmt.Text)
            On Error Resume Next
            txtCredit.SetFocus
        Else
            txtCredit.Text = ZERO
            txtDebit.Text = ToDoubleNumber(txtNetAmt.Text)
            On Error Resume Next
            txtDebit.SetFocus
        End If
    End If
End Sub

Sub GET_AP_VOUCHERNO()
    Dim rsAP_VOUCHER                              As ADODB.Recordset
    Dim xVOUCHERNO                                As String
    Dim xJdate                                    As String
    Dim xDUEDATE                                  As String
    Dim xJType                                    As String
    Dim XCustomerCode                             As String
    Dim xCUST_NAME                                As String
    Dim xINVOICENO                                As String
    Dim xInvoiceType                              As String
    Dim xInvoicedate                              As String
    Dim xAMOUNT_TO_PAY                            As Double
    Dim xAMOUNT_PAID                              As Double
    Dim xACCT_CODE                                As String
    Dim xLAST_UPDATED                             As String
    Dim xBAL                                      As Double

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsAP_VOUCHER = New ADODB.Recordset
    rsAP_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,HD.AMOUNTTOPAY,ACCT_CODE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07','11-02') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsAP_VOUCHER.EOF And Not rsAP_VOUCHER.BOF Then
        xVOUCHERNO = N2Str2Null(Null2String(rsAP_VOUCHER!jtype) & "-" & Null2String(rsAP_VOUCHER!VOUCHERNO))
        xJdate = N2Str2Null(Null2String(rsAP_VOUCHER!JDate))
        xJType = N2Str2Null(Null2String(rsAP_VOUCHER!jtype))
        xDUEDATE = N2Str2Null(Null2String(rsAP_VOUCHER!duedate))

        If xJOURNALTYPE = "VPJ" Then
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!VendorCode))
            xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAP_VOUCHER!VendorCode)))
        Else
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!CustomerCode))
            xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAP_VOUCHER!CustomerCode)))
        End If

        xINVOICENO = N2Str2Null(Null2String(rsAP_VOUCHER!INVOICENO))
        xInvoiceType = N2Str2Null(Null2String(rsAP_VOUCHER!InvoiceType))
        xInvoicedate = N2Str2Null(Null2String(rsAP_VOUCHER!invoicedate))
        'xAMOUNT_TO_PAY = GET_AP_AMOUNT(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!jtype), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xAMOUNT_TO_PAY = NumericVal(rsAP_VOUCHER!amounttopay)
        xAMOUNT_PAID = 0
        xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
        xACCT_CODE = N2Str2Null(Null2String(rsAP_VOUCHER!Acct_code))
        xLAST_UPDATED = N2Str2Null(LOGDATE)

        SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                        "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xINVOICENO & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
        gconDMIS.Execute SQL_STATEMENT
    End If
    Set rsAP_VOUCHER = Nothing
End Sub

Function GET_CUST_NAME(xCUSCODE As String) As String
    Dim rsGET_CUST_NAME                           As ADODB.Recordset
    Set rsGET_CUST_NAME = New ADODB.Recordset
    rsGET_CUST_NAME.Open "SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE RTRIM(LTRIM(CUSCDE)) = '" & RTrim(LTrim(xCUSCODE)) & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CUST_NAME.EOF And Not rsGET_CUST_NAME.BOF Then
        GET_CUST_NAME = Null2String(rsGET_CUST_NAME!AcctName)
    Else
        GET_CUST_NAME = ""
    End If
    Set rsGET_CUST_NAME = Nothing
End Function

Function GET_VEN_NAME(xVENCODE As String) As String
    Dim rsGET_VEN_NAME                            As ADODB.Recordset
    Set rsGET_VEN_NAME = New ADODB.Recordset
    rsGET_VEN_NAME.Open "SELECT NAMEOFVENDOR FROM ALL_VENDOR WHERE  RTRIM(LTRIM(CODE)) = " & N2Str2Null(xVENCODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_VEN_NAME.EOF And Not rsGET_VEN_NAME.BOF Then
        GET_VEN_NAME = Null2String(rsGET_VEN_NAME!nameofvendor)
    Else
        GET_VEN_NAME = ""
    End If
    Set rsGET_VEN_NAME = Nothing
End Function

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    txtCode.Text = strCode
    txtNameofVendor.Text = strAccountName
    xEntityClass = strEntityClass
    txtAddress.Caption = SetVendorAddressNew(strCode, strEntityClass)
End Sub

Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
        KeyAscii = 0
        MsgBox "Invalid Character", vbExclamation, "Check Entry"
        txtVoucherNo.SetFocus
    End Select
End Sub

Sub LoadJournal(XXX As String)
    xJOURNALTYPE = XXX
End Sub
