VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{205EA659-0BC9-4F44-85D9-FBC10C8940C1}#1.0#0"; "wizDigit.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCMISOREntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Official Receipt Data Entry [With VAT]"
   ClientHeight    =   8940
   ClientLeft      =   810
   ClientTop       =   3285
   ClientWidth     =   12420
   ForeColor       =   &H00F5F5F5&
   Icon            =   "OREntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   12420
   Begin VB.PictureBox picOptions 
      Height          =   1515
      Left            =   9270
      ScaleHeight     =   1455
      ScaleWidth      =   1530
      TabIndex        =   68
      Top             =   6330
      Width           =   1590
      Begin VB.CommandButton cmdRecoverOR 
         Caption         =   "Recover Cancelled Official Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1650
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Recover Cancelled Official Receipt"
         Top             =   300
         Width           =   1425
      End
      Begin VB.CommandButton cmdCancelOR 
         Caption         =   "Cancel Official Receipt"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   60
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":102C
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":1336
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Cancel Official Receipt"
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "OPTIONS"
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
         Height          =   255
         Left            =   -240
         TabIndex        =   97
         Top             =   0
         Width           =   2025
      End
   End
   Begin VB.PictureBox picORType 
      BackColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   6000
      ScaleHeight     =   2475
      ScaleWidth      =   3915
      TabIndex        =   146
      Top             =   3600
      Visible         =   0   'False
      Width           =   3975
      Begin VB.OptionButton OptPR 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PROVISIONAL RECEIPTS"
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
         Height          =   555
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   1800
         Width           =   3465
      End
      Begin VB.OptionButton optService 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OFFICIAL RECEIPTS"
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
         Height          =   555
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   1200
         Width           =   3465
      End
      Begin VB.OptionButton optGoods 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ACKNOWLEDGMENT RECEIPTS"
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
         Height          =   555
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   600
         Width           =   3465
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
         Height          =   405
         Left            =   0
         TabIndex        =   149
         Top             =   -30
         Width           =   4245
         _Version        =   655364
         _ExtentX        =   7488
         _ExtentY        =   714
         _StockProps     =   14
         Caption         =   "SELECT OR TYPE"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.74
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         GradientColorLight=   8388608
         GradientColorDark=   16711680
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   3420
      ScaleHeight     =   3945
      ScaleWidth      =   8895
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   8925
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   795
         Left            =   2670
         ScaleHeight     =   795
         ScaleWidth      =   3075
         TabIndex        =   54
         Top             =   960
         Width           =   3075
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1590
            TabIndex        =   57
            Text            =   "0.00"
            Top             =   420
            Width           =   1395
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   1590
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   30
            Width           =   1395
         End
         Begin VB.Label Label14 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Balance :"
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
            Left            =   30
            TabIndex        =   58
            Top             =   450
            Width           =   1545
         End
         Begin VB.Label Label13 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Amount :"
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
            Left            =   30
            TabIndex        =   56
            Top             =   60
            Width           =   1545
         End
      End
      Begin VB.TextBox Payment 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   400
         Left            =   7320
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   2160
         Width           =   1400
      End
      Begin VB.CommandButton cmdDetails 
         Caption         =   "Command2"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7380
         TabIndex        =   99
         Top             =   4080
         Width           =   615
      End
      Begin VB.CommandButton cmdInsurance 
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
         Left            =   5760
         TabIndex        =   198
         Top             =   1800
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtReference1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3930
         MaxLength       =   8
         TabIndex        =   124
         Top             =   4080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox cboInvoiceType 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   390
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   960
         Width           =   1365
      End
      Begin VB.CheckBox chkCreditCardTrans 
         BackColor       =   &H00C0C0C0&
         Caption         =   "This is a Credit Card Transaction"
         Height          =   285
         Left            =   6000
         TabIndex        =   100
         Top             =   330
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   120
         MouseIcon       =   "OREntry.frx":17D8
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":192A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
         Width           =   705
      End
      Begin VB.CommandButton cmdCardPayment 
         Caption         =   "Compute for Card Tax and Discount"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6000
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":1C55
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Save changes"
         Top             =   630
         Visible         =   0   'False
         Width           =   2745
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5490
         TabIndex        =   53
         Top             =   540
         Width           =   285
      End
      Begin VB.ComboBox cboTranType 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   390
         ItemData        =   "OREntry.frx":1F5F
         Left            =   1200
         List            =   "OREntry.frx":1F61
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   510
         Width           =   4245
      End
      Begin VB.TextBox txtReference 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1440
         Width           =   1365
      End
      Begin VB.ComboBox cboBranch 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   2235
      End
      Begin VB.ComboBox cboPaidFor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   1200
         TabIndex        =   9
         Text            =   "cboPaidFor"
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7320
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1140
         Width           =   1400
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7320
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   1560
         Width           =   1400
      End
      Begin VB.CommandButton cmdTranCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   8040
         MouseIcon       =   "OREntry.frx":1F63
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":20B5
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   705
      End
      Begin VB.CommandButton cmdTranSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   7320
         MouseIcon       =   "OREntry.frx":23F3
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":2545
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   705
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "*"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         TabIndex        =   72
         Top             =   510
         Width           =   285
      End
      Begin VB.TextBox txtDescript 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   1215
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2160
         Width           =   4545
      End
      Begin VB.Label labReference 
         Caption         =   "Label10"
         Height          =   345
         Left            =   3000
         TabIndex        =   73
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label lbltotalAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payment:"
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
         Height          =   525
         Left            =   5760
         TabIndex        =   113
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblView 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - View Bank Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         TabIndex        =   125
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblDetID 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1110
         TabIndex        =   127
         Top             =   3660
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblReference1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Reference No."
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
         Left            =   1170
         TabIndex        =   126
         Top             =   4110
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   90
         TabIndex        =   104
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice"
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
         Left            =   90
         TabIndex        =   102
         Top             =   960
         Width           =   1005
      End
      Begin XtremeShortcutBar.ShortcutCaption labStatusMode 
         Height          =   285
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "System is Adding/Editing OR Detail"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.76
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   8421504
         GradientColorDark=   16711680
      End
      Begin VB.Label labCUSCODE 
         Caption         =   "Label21"
         Height          =   195
         Left            =   2370
         TabIndex        =   61
         Top             =   1830
         Width           =   1305
      End
      Begin VB.Label labDetID 
         Caption         =   "Label21"
         Height          =   135
         Left            =   2040
         TabIndex        =   60
         Top             =   2370
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label labDocDate 
         BackStyle       =   0  'Transparent
         Caption         =   "[DOC DATE]"
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
         Height          =   225
         Left            =   30
         TabIndex        =   59
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tran Type :"
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
         Left            =   -60
         TabIndex        =   34
         Top             =   570
         Width           =   1155
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Application :"
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
         Left            =   0
         TabIndex        =   33
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code :"
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
         Left            =   9090
         TabIndex        =   32
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "EWT"
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
         Left            =   6120
         TabIndex        =   31
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Charges"
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
         Left            =   5880
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for :"
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
         Left            =   60
         TabIndex        =   29
         Top             =   1830
         Width           =   1125
      End
      Begin VB.Label lblVendorName 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         TabIndex        =   199
         Top             =   3120
         Width           =   4065
      End
      Begin VB.Label labRef 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Left            =   90
         TabIndex        =   35
         Top             =   1470
         Width           =   1035
      End
   End
   Begin VB.PictureBox picORPayment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   4320
      ScaleHeight     =   3510
      ScaleWidth      =   7125
      TabIndex        =   151
      Top             =   3240
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox txtChattelAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   195
         Text            =   "0.00"
         Top             =   2040
         Width           =   1530
      End
      Begin VB.TextBox txtChattelBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   194
         Text            =   "0.00"
         Top             =   2040
         Width           =   1485
      End
      Begin VB.TextBox txtChattelPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5220
         TabIndex        =   193
         Text            =   "0.00"
         Top             =   2040
         Width           =   1755
      End
      Begin VB.CheckBox chkChattel 
         BackColor       =   &H80000004&
         Height          =   240
         Left            =   4905
         TabIndex        =   192
         Top             =   2085
         Width           =   240
      End
      Begin VB.TextBox txtDownPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5220
         TabIndex        =   174
         Text            =   "0.00"
         Top             =   600
         Width           =   1755
      End
      Begin VB.PictureBox Picture10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1710
         ScaleHeight     =   390
         ScaleWidth      =   1500
         TabIndex        =   172
         Top             =   2640
         Width           =   1530
         Begin VB.Label lblInvoiceNo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "000000"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   60
            TabIndex        =   173
            Top             =   0
            Width           =   1410
         End
      End
      Begin VB.CheckBox chkTPL 
         BackColor       =   &H80000004&
         Height          =   240
         Left            =   4905
         TabIndex        =   171
         Top             =   1725
         Width           =   240
      End
      Begin VB.CheckBox chkInsurance 
         BackColor       =   &H80000004&
         Height          =   240
         Left            =   4905
         TabIndex        =   170
         Top             =   1365
         Width           =   240
      End
      Begin VB.CheckBox chkLTORegFee 
         BackColor       =   &H80000004&
         Height          =   240
         Left            =   4905
         TabIndex        =   169
         Top             =   1005
         Width           =   285
      End
      Begin VB.CheckBox chkDownPayment 
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   4905
         TabIndex        =   168
         Top             =   645
         Width           =   240
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3930
         MaxLength       =   8
         TabIndex        =   167
         Top             =   4080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00973640&
         Height          =   330
         Left            =   9960
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   840
         Width           =   2235
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Enabled         =   0   'False
         Height          =   195
         Left            =   7380
         TabIndex        =   165
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtTPLPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5220
         TabIndex        =   164
         Text            =   "0.00"
         Top             =   1680
         Width           =   1755
      End
      Begin VB.TextBox txtOtherBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   163
         Text            =   "0.00"
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox txtTPLAmout 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   162
         Text            =   "0.00"
         Top             =   1680
         Width           =   1530
      End
      Begin VB.TextBox txtInsPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5220
         TabIndex        =   161
         Text            =   "0.00"
         Top             =   1320
         Width           =   1755
      End
      Begin VB.TextBox txtInsBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   160
         Text            =   "0.00"
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txtInsAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   159
         Text            =   "0.00"
         Top             =   1320
         Width           =   1530
      End
      Begin VB.TextBox txtLTOPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   5220
         TabIndex        =   158
         Text            =   "0.00"
         Top             =   960
         Width           =   1755
      End
      Begin VB.TextBox txtLTOBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   157
         Text            =   "0.00"
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtLTOAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   156
         Text            =   "0.00"
         Top             =   960
         Width           =   1530
      End
      Begin VB.TextBox txtDownBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3285
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   155
         Text            =   "0.00"
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtDownAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   154
         Text            =   "0.00"
         Top             =   600
         Width           =   1530
      End
      Begin VB.CommandButton cmdCancelPayment 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6240
         MouseIcon       =   "OREntry.frx":2895
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":29E7
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   2580
         Width           =   735
      End
      Begin VB.CommandButton cmdSaveORDetail 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5520
         MouseIcon       =   "OREntry.frx":2D25
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":2E77
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   2580
         Width           =   735
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sales Price:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   450
         TabIndex        =   200
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   405
         TabIndex        =   197
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chattel:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   810
         TabIndex        =   196
         Top             =   2085
         Width           =   810
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   1110
         TabIndex        =   185
         Top             =   3660
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Card Reference No."
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
         Left            =   1170
         TabIndex        =   184
         Top             =   4110
         Visible         =   0   'False
         Width           =   2280
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   0
         TabIndex        =   183
         Top             =   0
         Width           =   7155
         _Version        =   655364
         _ExtentX        =   12621
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Invoice Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   16711680
         GradientColorDark=   4194304
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code :"
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
         Left            =   9090
         TabIndex        =   182
         Top             =   870
         Width           =   1125
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TPL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1185
         TabIndex        =   181
         Top             =   1725
         Width           =   450
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LTO Reg Fee:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   285
         TabIndex        =   180
         Top             =   1005
         Width           =   1350
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   585
         TabIndex        =   179
         Top             =   1365
         Width           =   1050
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Downpayment:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   178
         Top             =   645
         Width           =   1485
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   177
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3330
         TabIndex        =   176
         Top             =   375
         Width           =   1500
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amount"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1755
         TabIndex        =   175
         Top             =   375
         Width           =   1500
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   6855
      Left            =   30
      TabIndex        =   47
      Top             =   0
      Width           =   3285
      Begin VB.OptionButton optORNo 
         Caption         =   "OR No."
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
         TabIndex        =   50
         Top             =   390
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton optCustName 
         Caption         =   "Customer Name"
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
         TabIndex        =   49
         Top             =   630
         Width           =   1905
      End
      Begin VB.TextBox textSearch 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   60
         MaxLength       =   35
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   960
         Width           =   3135
      End
      Begin MSComctlLib.ListView lstOFF_HD 
         Height          =   5415
         Left            =   30
         TabIndex        =   51
         Top             =   1350
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   9551
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
         MouseIcon       =   "OREntry.frx":31C7
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR Number."
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label20 
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
         TabIndex        =   52
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   3390
      ScaleHeight     =   1215
      ScaleWidth      =   8925
      TabIndex        =   20
      Top             =   90
      Width           =   8955
      Begin wizDigits.wizDigit wizDigit1 
         Height          =   1215
         Left            =   -480
         TabIndex        =   21
         Top             =   0
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   2143
      End
   End
   Begin VB.PictureBox picDetail 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   60
      ScaleHeight     =   855
      ScaleWidth      =   3225
      TabIndex        =   69
      Top             =   6900
      Width           =   3225
      Begin VB.CommandButton cmdInvoices 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Customers &Invoices"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   30
         TabIndex        =   132
         ToolTipText     =   "Invoice App. Detail"
         Top             =   585
         Width           =   3150
      End
      Begin VB.CommandButton cmdInvoiceDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Invoice App. Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   71
         ToolTipText     =   "Invoice App. Detail"
         Top             =   285
         Width           =   3150
      End
      Begin VB.CommandButton cmdORDetail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "O.R. Payment Detail"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   30
         TabIndex        =   70
         ToolTipText     =   "O.R. Payment Detail"
         Top             =   0
         Width           =   3150
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   4485
      TabIndex        =   76
      Top             =   7890
      Width           =   4515
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Payment"
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
         Left            =   2280
         TabIndex        =   80
         Top             =   270
         Width           =   2055
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Apply Deposit"
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
         Left            =   2280
         TabIndex        =   79
         Top             =   30
         Width           =   2055
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Edit OR Detail"
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
         Left            =   150
         TabIndex        =   78
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "F2 - Add OR Detail"
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
         Left            =   150
         TabIndex        =   77
         Top             =   30
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   885
      Left            =   3390
      ScaleHeight     =   885
      ScaleWidth      =   8955
      TabIndex        =   36
      Top             =   6885
      Width           =   8955
      Begin VB.TextBox txtBranch 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   6540
         TabIndex        =   46
         Top             =   60
         Width           =   2325
      End
      Begin VB.TextBox txtPaidFor 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1320
         TabIndex        =   45
         Top             =   60
         Width           =   3855
      End
      Begin VB.TextBox txtPaymentAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7440
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   480
         Width           =   1400
      End
      Begin VB.TextBox txtDiscountAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1320
         TabIndex        =   38
         Text            =   "0.00"
         Top             =   450
         Width           =   1400
      End
      Begin VB.TextBox txtTaxAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   4260
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   450
         Width           =   1400
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Amt :"
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
         Left            =   6000
         TabIndex        =   44
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code :"
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
         Left            =   5250
         TabIndex        =   43
         Top             =   90
         Width           =   1245
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amt :"
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
         Left            =   3090
         TabIndex        =   42
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Amt :"
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
         Left            =   60
         TabIndex        =   41
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment for :"
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
         Left            =   90
         TabIndex        =   40
         Top             =   60
         Width           =   1185
      End
   End
   Begin VB.PictureBox picOR 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   3390
      ScaleHeight     =   1935
      ScaleWidth      =   8955
      TabIndex        =   22
      Top             =   1410
      Width           =   8955
      Begin VB.TextBox txtPRDate 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   7140
         MaxLength       =   11
         TabIndex        =   207
         Top             =   1440
         Width           =   1635
      End
      Begin VB.TextBox txtFao 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3840
         MaxLength       =   100
         TabIndex        =   204
         Top             =   1440
         Width           =   2325
      End
      Begin VB.TextBox txtPRNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2100
         MaxLength       =   11
         TabIndex        =   203
         Top             =   1440
         Width           =   1275
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   98
         Top             =   1035
         Width           =   360
      End
      Begin VB.TextBox txtOR_NUM 
         BackColor       =   &H00F1F6F5&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2100
         MaxLength       =   8
         TabIndex        =   150
         Top             =   60
         Width           =   2355
      End
      Begin VB.TextBox cboCUSNAME 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   360
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   1020
         Width           =   4215
      End
      Begin VB.CommandButton cmdVarious 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3000
         TabIndex        =   74
         Top             =   1050
         Width           =   285
      End
      Begin Crystal.CrystalReport rptChat 
         Left            =   8500
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.Timer Timer1 
         Interval        =   300
         Left            =   8460
         Top             =   570
      End
      Begin VB.TextBox txtOR_DATE 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2100
         MaxLength       =   11
         TabIndex        =   2
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox txtCUSCDE 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   3
         Text            =   "BBBBBB"
         Top             =   1020
         Width           =   945
      End
      Begin MSMask.MaskEdBox txtVNF 
         Height          =   525
         Left            =   4800
         TabIndex        =   1
         Top             =   60
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   15857397
         ForeColor       =   0
         Enabled         =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtOR_NUM1 
         Height          =   525
         Left            =   2100
         TabIndex        =   0
         Top             =   60
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   15857397
         ForeColor       =   0
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblPRDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PR Date:"
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
         Left            =   6120
         TabIndex        =   208
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblPRNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PR No.:"
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
         Left            =   1080
         TabIndex        =   205
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label lblFao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "FAO:"
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
         Left            =   3240
         TabIndex        =   206
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label lblInvoiceNo1 
         BackColor       =   &H00E0E0E0&
         Height          =   420
         Index           =   1
         Left            =   4455
         TabIndex        =   191
         Top             =   585
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label paymenttype 
         BackColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   1
         Left            =   4770
         TabIndex        =   190
         Top             =   720
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label labStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "*** Cancelled OR ***"
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
         Height          =   315
         Left            =   5640
         TabIndex        =   62
         Top             =   180
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   " Date:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   660
         Width           =   1905
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code-Name :"
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
         Left            =   -30
         TabIndex        =   25
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label lblReceipt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Official Receipt No. :"
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
         Height          =   255
         Left            =   -210
         TabIndex        =   24
         Top             =   240
         Width           =   2235
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4560
         TabIndex        =   23
         Top             =   90
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label AckReceipts 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Acknowledgment Receipt:"
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
         Height          =   495
         Left            =   0
         TabIndex        =   114
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.TextBox txtSO_NO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   7155
      TabIndex        =   186
      Top             =   2040
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtCustype 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   6885
      TabIndex        =   187
      Top             =   2025
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDeposits 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3675
      Left            =   4680
      ScaleHeight     =   3645
      ScaleWidth      =   6435
      TabIndex        =   133
      Top             =   3000
      Visible         =   0   'False
      Width           =   6465
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6180
         TabIndex        =   134
         Top             =   0
         Width           =   255
      End
      Begin MSComctlLib.ListView lvDeposits 
         Height          =   3240
         Left            =   45
         TabIndex        =   135
         Top             =   360
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   5715
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Customer Code"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "OR Date"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "OR No."
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Applied"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Deposits"
            Object.Width           =   5293
         EndProperty
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   2955
         Left            =   30
         ScaleHeight     =   2895
         ScaleWidth      =   6345
         TabIndex        =   136
         Top             =   330
         Width           =   6405
      End
      Begin VB.Label lblDepositID 
         Height          =   375
         Left            =   3360
         TabIndex        =   137
         Top             =   3120
         Width           =   1635
      End
      Begin XtremeShortcutBar.ShortcutCaption sc3 
         Height          =   285
         Left            =   0
         TabIndex        =   138
         Top             =   0
         Width           =   6435
         _Version        =   655364
         _ExtentX        =   11351
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Unapplied Customer Deposit/s"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   -2147483629
         GradientColorDark=   12582912
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   4620
      ScaleHeight     =   870
      ScaleWidth      =   7695
      TabIndex        =   81
      Top             =   7920
      Width           =   7695
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
         Left            =   6945
         MouseIcon       =   "OREntry.frx":3329
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":347B
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   6255
         MouseIcon       =   "OREntry.frx":37E1
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":3933
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Option"
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
         Left            =   5565
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":3C99
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":3DEB
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "View Options"
         Top             =   30
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
         Left            =   4875
         MouseIcon       =   "OREntry.frx":4151
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":42A3
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
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
         Left            =   4185
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "OREntry.frx":45FF
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":4751
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "Post this Transaction"
         Top             =   30
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
         Left            =   3490
         MouseIcon       =   "OREntry.frx":4A76
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":4BC8
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   2805
         MouseIcon       =   "OREntry.frx":4EDB
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":502D
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "Move to Last Record"
         Top             =   30
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
         MouseIcon       =   "OREntry.frx":537D
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":54CF
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Move to First Record"
         Top             =   30
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
         Left            =   1425
         MouseIcon       =   "OREntry.frx":582D
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":597F
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Find a Record"
         Top             =   30
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
         Left            =   735
         MouseIcon       =   "OREntry.frx":5C79
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":5DCB
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Move to Next Record"
         Top             =   30
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
         Left            =   45
         MouseIcon       =   "OREntry.frx":6123
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":6275
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   10800
      ScaleHeight     =   855
      ScaleWidth      =   1500
      TabIndex        =   93
      Top             =   7935
      Width           =   1500
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
         Left            =   750
         MouseIcon       =   "OREntry.frx":65D4
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":6726
         Style           =   1  'Graphical
         TabIndex        =   94
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
         Left            =   60
         MouseIcon       =   "OREntry.frx":6A64
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":6BB6
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picPayment 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   4440
      ScaleHeight     =   1575
      ScaleWidth      =   6825
      TabIndex        =   63
      Top             =   1170
      Visible         =   0   'False
      Width           =   6855
      Begin VB.OptionButton optCANCEL 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   5070
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCARD 
         Caption         =   "CARD"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCHECK 
         Caption         =   "CHECK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   540
         Width           =   1545
      End
      Begin VB.OptionButton optCASH 
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   540
         Width           =   1545
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   96
         Top             =   0
         Width           =   6840
         _Version        =   655364
         _ExtentX        =   12065
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "  Select Type of Payment..."
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorLight=   16711680
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
   End
   Begin VB.PictureBox picCreditCard 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3420
      ScaleHeight     =   4305
      ScaleWidth      =   8895
      TabIndex        =   105
      Top             =   3000
      Visible         =   0   'False
      Width           =   8925
      Begin MSComctlLib.ListView lvPayments 
         Height          =   2655
         Left            =   120
         TabIndex        =   108
         ToolTipText     =   "Doub"
         Top             =   1200
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "OR No."
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Code"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer Name"
            Object.Width           =   6598
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Reference No"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "OR Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdCardCancel 
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
         Left            =   8070
         MouseIcon       =   "OREntry.frx":6F06
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":7058
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   360
         Width           =   705
      End
      Begin VB.CommandButton cmdCardSave 
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
         Left            =   7380
         MouseIcon       =   "OREntry.frx":7396
         MousePointer    =   99  'Custom
         Picture         =   "OREntry.frx":74E8
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   360
         Width           =   705
      End
      Begin VB.CheckBox chkSelect 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   150
         TabIndex        =   139
         Top             =   3900
         Width           =   195
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Payment Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2700
         TabIndex        =   110
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OR Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   109
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select / Deselect All "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   4440
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   5160
         TabIndex        =   106
         Top             =   360
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   8145
         TabIndex        =   111
         Top             =   720
         Visible         =   0   'False
         Width           =   8145
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   375
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   30
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   405
            Left            =   840
            TabIndex        =   201
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   148832257
            CurrentDate     =   38216
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   405
            Left            =   3060
            TabIndex        =   202
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   148832257
            CurrentDate     =   38216
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   60
            Width           =   675
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "To :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   255
            Left            =   2610
            TabIndex        =   115
            Top             =   60
            Width           =   435
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   8475
         TabIndex        =   117
         Top             =   720
         Width           =   8475
         Begin VB.TextBox txtReference2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2595
            MaxLength       =   8
            TabIndex        =   118
            Top             =   60
            Width           =   1815
         End
         Begin VB.Label lblReference2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter OR Number"
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
            Left            =   900
            TabIndex        =   119
            Top             =   120
            Width           =   1500
         End
      End
      Begin VB.PictureBox picCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   90
         ScaleHeight     =   495
         ScaleWidth      =   6585
         TabIndex        =   128
         Top             =   720
         Visible         =   0   'False
         Width           =   6585
         Begin VB.TextBox txtCustomer 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2580
            MaxLength       =   50
            TabIndex        =   129
            Top             =   30
            Width           =   3465
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name:"
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
            Left            =   930
            TabIndex        =   130
            Top             =   120
            Width           =   1470
         End
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Select All"
         Height          =   195
         Left            =   390
         TabIndex        =   140
         Top             =   3930
         Width           =   660
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6240
         TabIndex        =   123
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6990
         TabIndex        =   122
         Top             =   3930
         Width           =   1695
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   121
         Top             =   0
         Width           =   8925
         _Version        =   655364
         _ExtentX        =   15743
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Bank Payment        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   8421504
         GradientColorDark=   8421504
      End
      Begin VB.Label lblRefNo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1680
         TabIndex        =   120
         Top             =   4080
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.PictureBox picORDetails 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   3390
      ScaleHeight     =   3705
      ScaleWidth      =   8985
      TabIndex        =   27
      Top             =   3270
      Width           =   8985
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   3495
         Left            =   60
         TabIndex        =   4
         Top             =   90
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Type    |    Ref. #       |    Application                                |   AR                  | Balance           ||||||"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "OREntry.frx":7838
      End
   End
   Begin VB.Label lblInvoiceNo1 
      BackColor       =   &H00E0E0E0&
      Height          =   420
      Index           =   0
      Left            =   7830
      TabIndex        =   189
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label paymenttype 
      BackColor       =   &H00E0E0E0&
      Height          =   240
      Index           =   0
      Left            =   8145
      TabIndex        =   188
      Top             =   2115
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label labDetails 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTED: CASH RECEIPTS JOURNAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   1830
      TabIndex        =   145
      Top             =   8535
      Width           =   2835
   End
   Begin VB.Label labCRJNo 
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
      TabIndex        =   144
      Top             =   8460
      Width           =   1005
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4F4CD&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " CRJ #:"
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
      Left            =   60
      TabIndex        =   143
      Top             =   8460
      Width           =   795
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   6150
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   6480
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      Height          =   315
      Left            =   1890
      Top             =   8460
      Width           =   2685
   End
End
Attribute VB_Name = "frmCMISOREntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOFF_HD                                                        As ADODB.Recordset
Dim rsOFF_DT                                                        As ADODB.Recordset
Dim TOTAL_AR_AMOUNT                                                 As Double
Dim AddorEdit, PrevOR_NUM                                           As String
Dim On_Update                                                       As Boolean
Attribute On_Update.VB_VarUserMemId = 1073938437
Dim ChangeORNum                                                     As Boolean
Public LocalAcess                                                   As String
Dim rsINVOICEDUp                                                    As ADODB.Recordset
Dim rsCustomerDeposit                                               As ADODB.Recordset
Dim FIRST_LOAD                                                      As Boolean
Dim vtrantype                                                       As String

Dim tmpTotal                                                        As Double
Dim vDetails                                                        As Boolean
Dim ApplyDeposits                                                   As Boolean
Dim vDeposits                                                       As Double
Dim vCustype                                                        As String
Dim vPaymentType                                                    As String
Dim WithEvents frmNewEntity                                         As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim vENTITY                                                         As String
Dim vINSTPL                                                         As String

Dim vTerm                                                           As Boolean
Dim xTTLINV                                                         As Double
Dim rsDet_ID                                                        As ADODB.Recordset
Dim TRAN_INVOICE_TYPE                                               As String
Dim OR_REFERENCE                                                    As String
Dim Bal                                                             As Double
Dim rsDepositcheck                                  As ADODB.Recordset

Function SetCustomerCode(XXX As Variant)
    Dim rsCustomer                                                  As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("SELECT CusCde FROM ALL_CUSMAS WHERE CusNam = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = rsCustomer!CUSCDE
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerName(XXX As Variant)
    Dim rsCustomer                                                  As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("SELECT CusNam FROM ALL_CUSMAS WHERE CusCde = '" & XXX & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = rsCustomer!CusNam
    End If
    Set rsCustomer = Nothing
End Function

Function SetCheckClassCode(XXX As Variant)
    Dim rsSBOOK                                                     As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'F' AND DescName = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!Code
    End If
End Function

Function SetTranType(XXX As Variant)
    Dim rsType                                                      As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'A' AND Code = '" & XXX & "'")
    If Not rsType.EOF And Not rsType.BOF Then
        SetTranType = rsType!DESCNAME
    End If
End Function

Function SetTranTypeCode(XXX As Variant)
    Dim rsType                                                      As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'A' AND DescName = '" & XXX & "'")
    If Not rsType.EOF And Not rsType.BOF Then
        SetTranTypeCode = rsType!Code
    End If
End Function

Function SetBranch(XXX As Variant)
    Dim rsBranch                                                    As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'C' AND Code = '" & XXX & "'")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        SetBranch = rsBranch!DESCNAME
    End If
End Function

Function SetBranchCode(XXX As Variant)
    Dim rsBranch                                                    As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'C' AND DescName = '" & XXX & "'")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        SetBranchCode = rsBranch!Code
    End If
End Function

Function SetPaidFor(XXX As Variant)
    Dim rsPayment                                                   As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'D' AND Code = '" & XXX & "'")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        SetPaidFor = Null2String(rsPayment!DESCNAME)
    End If
End Function

Function SetPaidForCode(XXX As Variant)
    Dim rsPayment                                                   As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("SELECT Code FROM CMIS_SBOOK WHERE Book = 'D' AND DescName = '" & XXX & "'")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        SetPaidForCode = Null2String(UCase(rsPayment!Code))
    End If
End Function

Sub RefreshDisplay()
    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & N2Str2Null(OR_NUMBER_GLOBAL)
    StoreMemVars
End Sub

Sub Save_CASH_Payment()
    If COMPANY_CODE = "HGC" Then
        Set rsINVOICEDUp = New ADODB.Recordset
        Set rsINVOICEDUp = gconDMIS.Execute("SELECT * FROM CMIS_ORS WHERE ORNO = '" & Format(txtOR_NUM.Text, "000000") & "'")
        If Not rsINVOICEDUp.EOF And Not rsINVOICEDUp.BOF Then
            If Trim(Null2String(rsINVOICEDUp!Status)) = "P" Then
                MsgSpeechBox "OR Number already exist!"
                Exit Sub
            ElseIf Trim(Null2String(rsINVOICEDUp!Status)) = "C" Then
                MsgSpeechBox "OR Number already used and was Cancelled!"
                Exit Sub
            End If
        End If
    End If
    gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET" & _
                      " VAT = " & VAT_OR & "," & _
                      " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                      " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                      " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                      " PAIDNA = 1, STATUS='P'" & _
                      " WHERE OR_NUM = " & N2Str2Null(txtOR_NUM.Text))
    If COMPANY_CODE = "HGC" Then
        gconDMIS.Execute "INSERT INTO CMIS_ORS (ORNO,ORDATE,STATUS) VALUES('" & txtOR_NUM.Text & "','" & CDate(txtOR_DATE.Text) & "','P')"
    End If

    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & txtOR_NUM.Text
    StoreMemVars
End Sub

Sub Save_CHECK_Payment()
    If COMPANY_CODE = "HGC" Then
        Set rsINVOICEDUp = New ADODB.Recordset
        Set rsINVOICEDUp = gconDMIS.Execute("SELECT * FROM CMIS_ORS WHERE ORNO = '" & Format(txtOR_NUM.Text, "000000") & "'")
        If Not rsINVOICEDUp.EOF And Not rsINVOICEDUp.BOF Then
            If Null2String(rsINVOICEDUp!Status) = "P" Then
                MsgSpeechBox "OR Number already exist!"
                Exit Sub
            ElseIf Null2String(rsINVOICEDUp!Status) = "C" Then
                MsgSpeechBox "OR Number already used and was Cancelled!"
                Exit Sub
            End If
        End If
    End If
    gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET" & _
                      " VAT = " & VAT_OR & "," & _
                      " BAYADAMT = " & NumericVal(AMOUNT_TENDERED) & "," & _
                      " SUKLI = " & NumericVal(CHANGE_DUE) & "," & _
                      " CASHAMOUNT = " & NumericVal(RECEIPTS_AMOUNT) & "," & _
                      " PAIDNA = 1, STATUS='P'" & _
                      " WHERE OR_NUM = " & N2Str2Null(txtOR_NUM.Text))
End Sub

Sub SetCustomer()
    Call FillCustomer
    
    Dim rsCustomer                                                  As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("SELECT * FROM ALL_Customer WHERE CusCde = '" & txtCUSCDE.Text & "'")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        cboCUSNAME.Text = Null2String(rsCustomer!AcctName)
    End If
End Sub

Sub rsRefresh()
    Set rsOFF_HD = New ADODB.Recordset
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD = gconDMIS.Execute("SELECT * from CMIS_Off_hd WHERE VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE asc, OR_NUM asc")
    'Else
    '   Set rsOFF_HD = gconDMIS.Execute("SELECT * from CMIS_Off_hd WHERE VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE asc, OR_NUM asc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE VAT = 1 ORDER BY OR_DATE ASC, OR_NUM ASC")
    Else
        Set rsOFF_HD = gconDMIS.Execute("SELECT * FROM CMIS_Off_hd WHERE VAT = 0 ORDER BY OR_DATE ASC, OR_NUM ASC")
    End If
    If FIRST_LOAD = True Then
        If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then rsOFF_HD.MoveLast
    End If
End Sub

Sub StoreMemVars()
    cmdEdit.Enabled = True
    cmdPost.Enabled = True
    cmdOptions.Enabled = True
    cmdPrint.Enabled = False
    
    'JJE 05/20/2016
     If COMPANY_CODE = "MGS" Then
        If OR_VAT_NONVAT = "VAT" Then
            picORType.Height = 2535
            OptPR.Visible = True
        Else
            picORType.Height = 1935
            OptPR.Visible = False
        End If
     Else
        picORType.Height = 1935
        OptPR.Visible = False
        lblReceipt.Visible = True
        AckReceipts.Visible = False
    End If
   'JJE
    
     'JJE 08/29/2015
    If COLLECTION_RECEIPTS_CR = True Then
        optGoods.Caption = "COLLECTION RECEIPTS"
    Else
        optGoods.Caption = "ACKNOWLEDGMENT RECEIPTS"
    End If
    'JJE
    
    If Not rsOFF_HD.EOF And Not rsOFF_HD.BOF Then
        labid.Caption = rsOFF_HD!Id
        
        If OR_OPTION = True Then
            txtOR_NUM.Text = Null2String(rsOFF_HD!OR_NUM)
        Else
            If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Then
                txtOR_NUM.Text = Null2String(rsOFF_HD!OR_NUM)
            Else
                txtOR_NUM.Text = Format(Left(Null2String(rsOFF_HD!OR_NUM), 6), "000000")
            End If
        End If
       
        'JJE
        If COLLECTION_RECEIPTS_CR = True And Left(Null2String(rsOFF_HD!OR_NUM), 2) = "CR" Then
            AckReceipts.Caption = "Collection Receipt:"
            AckReceipts.Visible = True
            lblReceipt.Visible = False
        ElseIf COLLECTION_RECEIPTS_CR = True And Left(Null2String(rsOFF_HD!OR_NUM), 2) = "PR" Then
            AckReceipts.Caption = "Provisional Receipts:"
            AckReceipts.Visible = True
            lblReceipt.Visible = False
        ElseIf OFFICIAL_RECEIPT_OR = True And Left(Null2String(rsOFF_HD!OR_NUM), 2) = "OR" Then
            AckReceipts.Visible = False
            lblReceipt.Visible = True
        Else
            If Left(Null2String(rsOFF_HD!OR_NUM), 3) = "SOA" Then
                AckReceipts.Caption = "SOA number:"
                AckReceipts.Visible = True
                lblReceipt.Visible = False
            ElseIf Left(Null2String(rsOFF_HD!OR_NUM), 1) = "S" Then
                AckReceipts.Visible = False
                lblReceipt.Visible = True
            Else
                AckReceipts.Caption = "Acknowledgment Receipt:"
                AckReceipts.Visible = True
                lblReceipt.Visible = False
            End If
        End If
        'JJE
        
        txtOR_DATE.Text = Null2String(rsOFF_HD!OR_DATE)
        txtCUSCDE.Text = Null2String(rsOFF_HD!CUSCDE)
        cboCUSNAME.Text = SetCustomerName(Null2String(rsOFF_HD!CUSCDE))
        labCRJNo.Caption = GetCRJNo(rsOFF_HD!OR_NUM, "CI")
        
        If COMPANY_CODE = "DJM" Then
            txtPRNo.Text = Null2String(rsOFF_HD!PR_NO)
            txtFao.Text = Null2String(rsOFF_HD!FAO)
            txtPRDate = Null2String(rsOFF_HD!PR_DATE)
        End If
        
        If N2Str2Zero(rsOFF_HD!CashAmount) > 0 Then
            MODE_OF_PAYMENT = "CASH"
        End If
        If N2Str2Zero(rsOFF_HD!CardAmount) > 0 Then
            MODE_OF_PAYMENT = "CARD"
        End If
        If N2Str2Zero(rsOFF_HD!CHKAMOUNT) > 0 Then
            MODE_OF_PAYMENT = "CHECK"
        End If
        RECEIPTS_AMOUNT = N2Str2Zero(rsOFF_HD!OR_AMT)
        
        If rsOFF_HD!Cancel = True Then
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdOptions.Enabled = False
            cmdPrint.Enabled = False
        End If
        If rsOFF_HD!Paidna = True Then
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdOptions.Enabled = False
            cmdPrint.Enabled = True
        Else
            If rsOFF_HD!Cancel = True Then
                cmdOptions.Enabled = False
            Else
                cmdOptions.Enabled = True
            End If
        End If
        StoreDetails
    Else
        'MsgBox "No Such Record!", vbInformation, "Message"
        MessagePop InfoFriend, "Message", "No Such Record"
        'JJE
        AckReceipts.Visible = False
        lblReceipt.Visible = False
        'JJE
        cmdAdd.Value = True
    End If
End Sub

Sub StoreDetails()
    Dim i                                                           As Integer
    Dim vDeposit                                                    As Double
    
    TOTAL_AR_AMOUNT = 0
    InitGrid
    
    Dim rsOFF_Payment                                               As ADODB.Recordset
    Set rsOFF_DT = New ADODB.Recordset
    Set rsOFF_DT = gconDMIS.Execute("SELECT * FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND OR_Num = " & N2Str2Null(rsOFF_HD!OR_NUM) & " ORDER BY [ID] ASC")
    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
        rsOFF_DT.MoveFirst
        i = 0
        Do While Not rsOFF_DT.EOF
            i = i + 1
            If Null2String(rsOFF_DT!TranType) = "RO" Then
                TRAN_INVOICE_TYPE = "SI"
            Else
                TRAN_INVOICE_TYPE = Null2String(rsOFF_DT!TranType)
            End If
            
            'JJE Deposit Applied
            If Null2String(rsOFF_DT!DESCRIPT) <> "DEPOSIT APPLIED" Then
                Set rsOFF_Payment = gconDMIS.Execute("SELECT AMOUNT FROM CMIS_DEPOSITDT WHERE OR_NUM = " & N2Str2Null(rsOFF_DT!OR_NUM) & " AND INVOICENO = " & N2Str2Null(rsOFF_DT!REFERENCE) & "")
            Else
                Set rsOFF_Payment = gconDMIS.Execute("SELECT AMOUNT FROM CMIS_DEPOSITDT WHERE OR_NUM = " & N2Str2Null(rsOFF_DT!OR_NUM) & " AND INVOICENO = " & N2Str2Null(rsOFF_DT!INVOICENO) & "")
            End If
            If Not rsOFF_Payment.EOF And Not rsOFF_Payment.BOF Then
                vDeposit = N2Str2Zero(rsOFF_Payment!amount)
            End If
            
            'JJE 5/17/2013 (3:28PM)
            Dim TotalBalance                                        As ADODB.Recordset
            Set TotalBalance = New ADODB.Recordset
            If Null2String(rsOFF_DT!DESCRIPT) <> "DEPOSIT APPLIED" Then
                'JRE 06/29/2016 To correct the balance to be showed especially in VI
'                Set TotalBalance = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE,PAIDFOR FROM CMIS_Off_Dt WHERE Reference = " & N2Str2Null(rsOFF_DT!REFERENCE) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " AND isnull(PAIDNA,0) = 1 GROUP BY REFERENCE,TRANTYPE,PAIDFOR")
                Set TotalBalance = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE,PAIDFOR FROM CMIS_Off_Dt WHERE Reference = " & N2Str2Null(rsOFF_DT!REFERENCE) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " AND isnull(PAIDNA,0) = 1 GROUP BY REFERENCE,TRANTYPE,PAIDFOR ORDER BY PAIDFOR DESC")
            Else
                Set TotalBalance = gconDMIS.Execute("SELECT ROUND(SUM(isnull(PAYMENT,0) + isnull(TAX,0)),2) AS MGA_BAYAD,TRANTYPE,REFERENCE,PAIDFOR FROM CMIS_Off_Dt WHERE Reference = " & N2Str2Null(rsOFF_DT!REFERENCE) & " GROUP BY REFERENCE,TRANTYPE,PAIDFOR")
            End If
            If Not TotalBalance.EOF And Not TotalBalance.BOF Then
                If ((rsOFF_DT!amount) > (TotalBalance!MGA_BAYAD + vDeposit)) Or ((rsOFF_DT!amount) = (TotalBalance!MGA_BAYAD + vDeposit)) Then
                    Bal = ((rsOFF_DT!amount) - ((TotalBalance!MGA_BAYAD) + vDeposit))
                ElseIf (rsOFF_DT!amount) < (TotalBalance!MGA_BAYAD + vDeposit) Then
                    Bal = 0
                End If
                If UCase(TotalBalance!PAIDFOR) = "412S" Or UCase(TotalBalance!PAIDFOR) = "412P" Or UCase(TotalBalance!PAIDFOR) = "412V" Then
                    If Null2String(rsOFF_DT!DESCRIPT) = "DEPOSIT APPLIED" Then
                        grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!REFERENCE) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & 0 & _
                                      Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                    Else
                        grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!INVOICENO) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & Bal & _
                                  Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                    End If
                Else
                    If Null2String(rsOFF_DT!DESCRIPT) = "DEPOSIT APPLIED" Then
                        grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!REFERENCE) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & 0 & _
                                      Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                    Else
                        grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!INVOICENO) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & Bal & _
                                      Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                    End If
                End If
            Else
                If Null2String(rsOFF_DT!DESCRIPT) = "DEPOSIT APPLIED" Then
                        grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!REFERENCE) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & 0 & _
                                      Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                Else
                    grdDetails.AddItem TRAN_INVOICE_TYPE & Chr(9) & Null2String(rsOFF_DT!INVOICENO) & Chr(9) & Null2String(rsOFF_DT!DESCRIPT) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsOFF_DT!amount)) & Chr(9) & Bal & _
                                  Chr(9) & Null2String(rsOFF_DT!PAIDFOR) & Chr(9) & Null2String(rsOFF_DT!BRANCH) & Chr(9) & Null2String(rsOFF_DT!DISCOUNT) & Chr(9) & Null2String(rsOFF_DT!tax) & Chr(9) & Null2String(rsOFF_DT!Payment) & Chr(9) & rsOFF_DT!Id
                End If
            End If
            
            If i = 1 Then grdDetails.RemoveItem 1
            If Null2String(rsOFF_DT!DESCRIPT) <> "DEPOSIT APPLIED" Then
                TOTAL_AR_AMOUNT = TOTAL_AR_AMOUNT + ToDoubleNumber(rsOFF_DT!Payment)
                wizDigit1.TextValue = ZERO
                wizDigit1.TextValue = ToDoubleNumber(TOTAL_AR_AMOUNT)
                txtPaymentAmt.Text = ToDoubleNumber(TOTAL_AR_AMOUNT)
                Payment.Locked = False
            Else
                Payment.Locked = True
            End If
            'JJE
            rsOFF_DT.MoveNext
        Loop
        
        On Error Resume Next
        grdDetails.Col = 10
        ShowGridDetails grdDetails.Text
        vDetails = True
    Else
        vDetails = False
        wizDigit1.TextValue = ZERO
        txtPaidFor.Text = ""
        txtBranch.Text = ""
        
        txtDiscountAmt.Text = "0.00"
        txtTaxAmt.Text = "0.00"
        txtPaymentAmt.Text = "0.00"
    End If
End Sub

Sub ShowGridDetails(XXX As Long)
    Dim rsOFF_Details                                               As ADODB.Recordset
    Set rsOFF_Details = New ADODB.Recordset
    Set rsOFF_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_Dt WHERE ID = " & XXX)
    vPaymentType = ""
    If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
        txtPaidFor.Text = SetPaidFor(Null2String(rsOFF_Details!PAIDFOR))
        xPAIDFOR = Null2String(rsOFF_Details!PAIDFOR)
        txtBranch.Text = SetBranch(Null2String(rsOFF_Details!BRANCH))
        txtDiscountAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!DISCOUNT))
        txtTaxAmt.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!tax))
        lblDetID.Caption = Null2String(rsOFF_Details!OR_NUM)
        vREFERENCENO = Null2String(rsOFF_Details!ReferenceNo)
        txtPaymentAmt = Null2String(rsOFF_Details!Payment)
    Else
        txtPaidFor.Text = ""
        txtBranch.Text = ""
        
        txtDiscountAmt.Text = "0.00"
        txtTaxAmt.Text = "0.00"
        txtPaymentAmt.Text = "0.00"
    End If
End Sub

Sub StoreGridDetails(XXX As Long)
    Dim rsOFF_Details                                               As ADODB.Recordset
    Dim rsOFF_Det                                                   As ADODB.Recordset
    Dim rsBalanceAmount                                             As ADODB.Recordset
    Dim DepApply_Ref                                                As String
    
    Set rsOFF_Det = New ADODB.Recordset
    Set rsOFF_Det = gconDMIS.Execute("SELECT * FROM CMIS_Off_Dt WHERE ID = " & XXX)
    If Not rsOFF_Det.EOF And Not rsOFF_Det.BOF Then
        If rsOFF_Det!DESCRIPT = "DEPOSIT APPLIED" Then
            DepApply_Ref = rsOFF_Det!INVOICENO
            Set rsOFF_Details = New ADODB.Recordset
            Set rsOFF_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_Dt WHERE INVOICENO = '" & DepApply_Ref & "' and OR_NUM = '" & txtOR_NUM & "' and DESCRIPT <> 'DEPOSIT APPLIED'")
        Else
            Set rsOFF_Details = New ADODB.Recordset
            Set rsOFF_Details = gconDMIS.Execute("SELECT * FROM CMIS_Off_Dt WHERE ID = " & XXX)
        End If
        If Not rsOFF_Details.EOF And Not rsOFF_Details.BOF Then
            AddorEdit = "EDIT"
            labStatusMode.Caption = "System is Editing OR Detail..."
            cmdRefresh.Enabled = False
            cmdTranDelete.Visible = True
            labDetID.Caption = rsOFF_Details!Id
            labCUSCODE.Caption = Null2String(rsOFF_Details!CUSCDE)
            cboTranType.Text = SetTranType(Null2String(rsOFF_Details!TranType))
            
            If Null2String(rsOFF_Details!TranType) = "OTH" Then
                txtReference.Text = Null2String(rsOFF_Details!ReferenceNo)
                txtReference.Enabled = False
            Else
                txtReference.Text = Null2String(rsOFF_Details!INVOICENO)
                txtReference.Enabled = True
            End If
            
            labReference.Caption = Null2String(rsOFF_Details!INVOICENO)
            txtDescript.Text = Null2String(rsOFF_Details!DESCRIPT)
            lblRefNo.Caption = Null2String(rsOFF_Details!ReferenceNo)
            vPaymentType = Null2String(rsOFF_Details!paymenttype)
            If Null2String(rsOFF_Details!PAIDFOR) <> "" Then
                cboPaidFor.Text = SetPaidFor(Null2String(rsOFF_Details!PAIDFOR))
            Else
                cboPaidFor.ListIndex = -1
            End If
            If Null2String(rsOFF_Details!BRANCH) <> "" Then
                cboBranch.Text = SetBranch(Null2String(rsOFF_Details!BRANCH))
            Else
                cboBranch.ListIndex = -1
            End If
            'JJE
            txtAmount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!amount))
            Set rsBalanceAmount = New ADODB.Recordset
            Set rsBalanceAmount = gconDMIS.Execute("SELECT SUM(PAYMENT + TAX) AS TOTALPAY FROM CMIS_Off_Dt WHERE REFERENCENO = '" & txtReference & "' and trantype = '" & rsOFF_Details!TranType & "' AND PAIDNA = 1 AND ID <> " & XXX & " ")
            If ToDoubleNumber(N2Str2Zero(rsOFF_Details!amount)) <= 0 Then
                txtBalance.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!amount))
            Else
                txtBalance.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!amount)) - ToDoubleNumber(N2Str2Zero(rsBalanceAmount!TOTALPAY))
                txtBalance.Text = ToDoubleNumber(txtBalance)
            End If
            
            txtDiscount.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!DISCOUNT))
            txtTax.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!tax))
            Payment.Text = ToDoubleNumber(N2Str2Zero(rsOFF_Details!Payment))
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
            
            lblDetID.Caption = Null2String(rsOFF_Details!OR_NUM)
            If DEPOSITAPPLIED(rsOFF_Details!OR_NUM) = True Then
                txtReference.Enabled = False
                cboPaidFor.Enabled = False
                cboInvoiceType.Enabled = False
                cboTranType.Enabled = False
            End If
            'JJE
        End If
    End If
End Sub

Sub InitGridMemvars()
    AddorEdit = "ADD"
    cmdRefresh.Enabled = True
    cmdTranDelete.Visible = False
    
     'JJE 08/29/2015
    If OR_OPTION = True Then
        'Nothing
    Else '... JRE Service invoice is only allowed in OR with VAT
        If COMPANY_CODE = "CMC" Then
            If OR_VAT_NONVAT = "VAT" Then
                cboTranType.ListIndex = 5
            Else
                cboTranType.ListIndex = -1
            End If
        Else
            cboTranType.ListIndex = -1
        End If
    End If '... JRE
    'JJE
    
    cboTranType.Enabled = True
    labDocDate.Caption = "[DOC DATE]"
    labCUSCODE.Caption = "V00009"
    txtReference.Text = ""
    txtDescript.Text = ""
    cboPaidFor.ListIndex = -1
    cboBranch.ListIndex = -1

    txtAmount.Text = "0.00"
    txtBalance.Text = "0.00"
    txtDiscount.Text = "0.00"
    txtTax.Text = "0.00"
    Payment.Text = "0.00"
    lblVendorName.Caption = ""

    txtReference.Enabled = False
    txtDescript.Enabled = False
    cboPaidFor.Enabled = False
    cboBranch.Enabled = False
    txtDiscount.Enabled = False
    txtTax.Enabled = True
    Payment.Enabled = False
    Payment.Locked = False
    On Error Resume Next
    cboTranType.SetFocus
End Sub

Sub initMemvars()
    txtOR_NUM.Text = ""
    txtOR_DATE.Text = LOGDATE
    txtCUSCDE.Text = ""
    cboCUSNAME = ""
    txtDiscountAmt.Text = ZERO
    txtTaxAmt.Text = ZERO
    txtPaymentAmt.Text = ZERO
    wizDigit1.TextValue = ZERO
    labCRJNo.Caption = ""
    labDetails.Caption = ""
    txtPRNo.Text = ""
    txtFao.Text = ""
    txtPRDate.Text = ""
    ApplyDeposits = False
    InitGrid
End Sub

Sub InitGrid()
    cleargrid grdDetails
    grdDetails.FormatString = "  Type    |    Ref. #       |    Application                                |   AR                  | Balance           "
    grdDetails.ColWidth(5) = 1: grdDetails.ColWidth(6) = 1: grdDetails.ColWidth(7) = 1: grdDetails.ColWidth(8) = 1: grdDetails.ColWidth(9) = 1: grdDetails.ColWidth(10) = 1
End Sub

Sub FillCustomer()
    Exit Sub

    'TO DO
    'TAKE OUT THIS PROCEDURE AND USE ONLY THE SELECT BUTTON
    Dim rsCustomer                                                  As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("SELECT CUSNAM FROM ALL_CUSMAS WHERE CUSNAM <> '' AND CUSNAM IS NOT NULL ORDER BY CUSNAM ASC")
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        Combo_Loadval cboCUSNAME, rsCustomer
    End If
End Sub

Sub FillType()
    Dim rsType                                                      As ADODB.Recordset
    Set rsType = New ADODB.Recordset
    Set rsType = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'A' ORDER BY DescName ASC")
    If Not rsType.EOF And Not rsType.BOF Then
        Combo_Loadval cboTranType, rsType
    End If
End Sub

Sub FillBranch()
    Dim rsBranch                                                    As ADODB.Recordset
    Set rsBranch = New ADODB.Recordset
    Set rsBranch = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'C' ORDER BY DescName ASC")
    If Not rsBranch.EOF And Not rsBranch.BOF Then
        Combo_Loadval cboBranch, rsBranch
    End If
End Sub

Sub FillPayment()
    Dim rsPayment                                                   As ADODB.Recordset
    Set rsPayment = New ADODB.Recordset
    Set rsPayment = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'D' ORDER BY DescName ASC")
    If Not rsPayment.EOF And Not rsPayment.BOF Then
        Combo_Loadval cboPaidFor, rsPayment
    End If
End Sub

Sub FillGrid()
    
    lstOFF_HD.Sorted = False
    lstOFF_HD.ListItems.Clear
    lstOFF_HD.Enabled = False
    
    Dim rsOFF_HD2                                                   As ADODB.Recordset
    Set rsOFF_HD2 = New ADODB.Recordset
    If OR_VAT_NONVAT = "VAT" Then
        'Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM,ID from CMIS_Off_hd WHERE VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE desc, OR_NUM desc")
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM,ID FROM CMIS_Off_hd WHERE VAT = 1 ORDER BY OR_DATE DESC, OR_NUM DESC")
    Else
        'Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM,ID from CMIS_Off_hd WHERE VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE desc, OR_NUM desc")
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM,ID FROM CMIS_Off_hd WHERE VAT = 0 ORDER BY OR_DATE DESC, OR_NUM DESC")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If
    Set rsOFF_HD2 = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    
    lstOFF_HD.Sorted = False
    lstOFF_HD.ListItems.Clear
    lstOFF_HD.Enabled = False
    XXX = Repleys(XXX)
    
    Dim rsOFF_HD2                                                   As ADODB.Recordset
    Set rsOFF_HD2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM, ID from CMIS_Off_hd WHERE VAT = 1 AND OR_NUM like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE Desc, OR_NUM desc")
    'Else
    '   Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM, ID from CMIS_Off_hd WHERE VAT = 0 AND OR_NUM like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY OR_DATE Desc, OR_NUM desc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM, ID FROM CMIS_Off_hd WHERE VAT = 1 AND OR_NUM LIKE '" & ReplaceQuote(XXX) & "%' ORDER BY OR_DATE DESC, OR_NUM DESC")
    Else
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT OR_NUM, ID FROM CMIS_Off_hd WHERE VAT = 0 AND OR_NUM LIKE '" & ReplaceQuote(XXX) & "%' ORDER BY OR_DATE DESC, OR_NUM DESC")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If
    Set rsOFF_HD2 = Nothing
End Sub

Sub FillGrid2()
    
    lstOFF_HD.Sorted = False
    lstOFF_HD.ListItems.Clear
    lstOFF_HD.Enabled = False
    
    Dim rsOFF_HD2                                                   As ADODB.Recordset
    Set rsOFF_HD2 = New ADODB.Recordset
    If OR_VAT_NONVAT = "VAT" Then
        'Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME,ID from CMIS_Off_hd WHERE VAT = 1 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY CUSNAME asc")
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME,ID FROM CMIS_Off_hd WHERE VAT = 1 ORDER BY CUSNAME ASC")
    Else
        'Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME,ID from CMIS_Off_hd WHERE VAT = 0 and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY CUSNAME asc")
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME,ID FROM CMIS_Off_hd WHERE VAT = 0 ORDER BY CUSNAME ASC")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If
    Set rsOFF_HD2 = Nothing
End Sub

Sub FillSearchGrid2(XXX As String)
    
    lstOFF_HD.Enabled = False
    lstOFF_HD.Sorted = False
    lstOFF_HD.ListItems.Clear
    
    Dim rsOFF_HD2                                                   As ADODB.Recordset
    Set rsOFF_HD2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    'If OR_VAT_NONVAT = "VAT" Then
    '   Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME, ID from CMIS_Off_hd WHERE VAT = 1 and CUSNAME like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY CUSNAME asc")
    'Else
    '   Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME, ID from CMIS_Off_hd WHERE VAT = 0 and CUSNAME like '"  &   XXX  &  "%' and month(or_date) = " & PERIODMONTH & " and year(or_date) = " & PERIODYEAR & " ORDER BY CUSNAME asc")
    'End If
    If OR_VAT_NONVAT = "VAT" Then
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME, ID FROM CMIS_Off_hd WHERE VAT = 1 AND CUSNAME LIKE '" & XXX & "%' ORDER BY CUSNAME ASC")
    Else
        Set rsOFF_HD2 = gconDMIS.Execute("SELECT CUSNAME, ID FROM CMIS_Off_hd WHERE VAT = 0 AND CUSNAME LIKE '" & XXX & "%' ORDER BY CUSNAME ASC")
    End If
    If Not (rsOFF_HD2.EOF And rsOFF_HD2.BOF) Then
        lstOFF_HD.Enabled = True
        Listview_Loadval Me.lstOFF_HD.ListItems, rsOFF_HD2
        lstOFF_HD.Refresh
        lstOFF_HD.Enabled = True
    Else
        lstOFF_HD.Enabled = False
    End If
    Set rsOFF_HD2 = Nothing
End Sub

Private Sub cboBranch_GotFocus()
    VBComBoBoxDroppedDown cboBranch
End Sub

Private Sub cboPaidFor_Click()
    txtDescript.Text = cboPaidFor.Text
    If SetPaidForCode(cboPaidFor.Text) = "427" Then
        If CheckIfBank(txtCUSCDE.Text) = True Then
            picCreditCard.Visible = True
            txtReference2.Text = ""
            txtCustomer.Text = ""
            picCreditCard.ZOrder 0
            Option1.Value = True
            chkCreditCardTrans.Value = 0
            chkCreditCardTrans.Enabled = False
            Option1_Click
            chkSelect.Value = 0
        Else
            MsgBox "For BANK use only", vbInformation, "Payment Received from Bank"
            cboPaidFor.ListIndex = -1
            Exit Sub
        End If
    Else
        chkCreditCardTrans.Enabled = True
    End If
    
    'MITSUBISHI ONLY
    vENTITY = ""
    lblVendorName.Caption = ""
     
    'JJE Tagging for Insurance and LTO AP Entity 01/30/2013 3:16PM
    If Tagging(cboPaidFor.Text) = True Then
       cmdInsurance.Visible = True
    Else
       cmdInsurance.Visible = False
    End If
    'JJE
End Sub

'JJE Tagging of SCHEDULE ACCOUNT 01/30/2013 3:16PM
Function Tagging(XXX As Variant)
    Dim rsTag                                                       As ADODB.Recordset
    Dim Chartcode                                                   As String
    Set rsTag = New ADODB.Recordset
    Set rsTag = gconDMIS.Execute("SELECT Chartcodes FROM CMIS_Sbook WHERE Book = 'D' AND DescName = '" & XXX & "'")
    If Not rsTag.EOF And Not rsTag.BOF Then
        Chartcode = Null2String(rsTag!CHARTCODES)
        Dim Trantype1                                               As String
        Dim Trantype2                                               As String
        Set rsTag = gconDMIS.Execute("SELECT * FROM AMIS_Chartaccount WHERE AcctCode = '" & Chartcode & "'")
        If Not rsTag.EOF And Not rsTag.BOF Then
            Trantype1 = Null2String(rsTag!Trantype1)
            Trantype2 = Null2String(rsTag!Trantype2)
            If Trantype1 = "INSURANCE" And Trantype2 = "AP" Then
                Tagging = True
            ElseIf Trantype1 = "LTO" And Trantype2 = "AP" Then
                Tagging = True
            Else
                Tagging = False
            End If
        End If
    End If
End Function
'JJE

Private Sub cboPaidFor_GotFocus()
    VBComBoBoxDroppedDown cboPaidFor
End Sub

Private Sub cboPaidFor_KeyPress(KeyAscii As Integer)
    txtDescript.Text = ""
    If KeyAscii = 13 Then
        txtDescript.Text = cboPaidFor.Text
        If SetPaidForCode(cboPaidFor.Text) = "427" Then
            If CheckIfBank(txtCUSCDE.Text) = True Then
                picCreditCard.Visible = True
                txtReference1.Text = "NULL"
                txtReference2.Text = ""
                txtCustomer.Text = ""
                picCreditCard.ZOrder 0
                Option1.Value = True
                chkCreditCardTrans.Value = 0
                chkCreditCardTrans.Enabled = False
            Else
                MsgBox "For BANK use only", vbInformation, "Payment Received from Bank"
                cboPaidFor.ListIndex = -1
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub cboPaidFor_LostFocus()
    If SetTranTypeCode(cboTranType.Text) = "OTH" Then
        Dim rsPayment                                               As ADODB.Recordset
        Set rsPayment = New ADODB.Recordset
        Set rsPayment = gconDMIS.Execute("SELECT DescName FROM CMIS_SBOOK WHERE Book = 'D' AND DescName = '" & cboPaidFor.Text & "'")
        If Not rsPayment.EOF And Not rsPayment.BOF Then
            'nothing
        Else
            MsgBox "Please select from the list." & Chr(13) & "If not available it should be added in Other Transaction Masterfile", vbInformation, "System Message"
            On Error Resume Next
            cboPaidFor.SetFocus
        End If
    End If

    If SetPaidForCode(cboPaidFor.Text) = "427" Then
        If CheckIfBank(txtCUSCDE.Text) = True Then
            picCreditCard.Visible = True
            txtReference2.Text = ""
            txtCustomer.Text = ""
            picCreditCard.ZOrder 0
            Option1.Value = True
            chkCreditCardTrans.Value = 0
            chkCreditCardTrans.Enabled = False
            Option1_Click
        Else
            MsgBox "For BANK use only", vbInformation, "Payment Received from Bank"
            cboPaidFor.ListIndex = -1
            Exit Sub
        End If
    Else
        chkCreditCardTrans.Enabled = True
    End If
End Sub

Private Sub cboTranType_Click()
    TRANTYPE_VALIDATION
    txtReference.Enabled = True
    If SetTranTypeCode(cboTranType.Text) = "PI" Or SetTranTypeCode(cboTranType.Text) = "SI" Or SetTranTypeCode(cboTranType.Text) = "AI" Or SetTranTypeCode(cboTranType.Text) = "MI" Then
        cboInvoiceType.Enabled = True
        chkCreditCardTrans.Enabled = True
    Else
        cboInvoiceType.Enabled = False
        chkCreditCardTrans.Value = 0
        chkCreditCardTrans.Enabled = False
    End If
End Sub

Function TRANTYPE_VALIDATION()
    'JJE 08/29/2015 OR with Prefixes (Invoice type validation)
    If OR_VAT_NONVAT = "VAT" Then
        If OR_OPTION = True Then
            If (Left(frmCMISOREntry.txtOR_NUM, 2) = "CR" Or Left(frmCMISOREntry.txtOR_NUM, 2) = "PR") Then
                If UCase(cboTranType.Text) = "SERVICE INVOICE" Then
                    MsgBox "Service Invoice is not allowed in Collection Receipts", vbOKOnly, "INVALID INVOICE TYPE"
                    cboTranType.ListIndex = -1
                    cboTranType.Enabled = True
                    VBComBoBoxDroppedDown cboTranType
                    Exit Function
                End If
            ElseIf Left(frmCMISOREntry.txtOR_NUM, 2) = "OR" Then
                If COMPANY_CODE = "DJM" Or COMPANY_CODE = "MGS" Then
                    If UCase(cboTranType.Text) <> "SERVICE INVOICE" And UCase(cboTranType.Text) <> "OTHER TRANSACTION" Then
                        MsgBox "Official Receipts only allows Service Invoices and Other Trasaction", vbOKOnly, "INVALID INVOICE TYPE"
                        cboTranType.ListIndex = 5
                        cboTranType.Enabled = True
                        VBComBoBoxDroppedDown cboTranType
                        Exit Function
                    End If
                Else
                    If UCase(cboTranType.Text) <> "SERVICE INVOICE" Then
                        MsgBox "Official Receipts only allows Service Invoices", vbOKOnly, "INVALID INVOICE TYPE"
                        cboTranType.ListIndex = 5
                        cboTranType.Enabled = True
                        VBComBoBoxDroppedDown cboTranType
                        Exit Function
                    End If
                End If
            End If
'        Else 'JRE 05/06/2016 Service Invoice is only allowed in Official Receipt - CMC
'            If COMPANY_CODE = "CMC" Then
'                If UCase(cboTranType.Text) = "SERVICE INVOICE" Or UCase(cboTranType.Text) = "OTHER TRANSACTION" Then
'                    cboTranType.Enabled = True
'                    Exit Function
'                Else
'                    MsgBox "Official Receipts only allows Service Invoices and Other Transactions", vbOKOnly, "INVALID INVOICE TYPE"
'                    cboTranType.ListIndex = 5
'                    cboTranType.Enabled = True
'                    VBComBoBoxDroppedDown cboTranType
'                    Exit Function
'                End If
'            End If
'        End If
'    Else
'        If OR_VAT_NONVAT = "NON-VAT" Then
'            If UCase(cboTranType.Text) = "SERVICE INVOICE" Then
'                MsgBox "Service Invoice is only allowed in Official Receipt", vbOKOnly, "INVALID INVOICE TYPE"
'                cboTranType.ListIndex = -1
'                cboTranType.Enabled = True
'                Exit Function
'            End If
        End If
    End If
    'JJE
End Function

Private Sub cboTranType_GotFocus()
    VBComBoBoxDroppedDown cboTranType
End Sub

Private Sub cboTranType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Trim(cboTranType.Text) = "" Then VBComBoBoxDroppedDown cboTranType
    End If
End Sub

Private Sub cboTranType_LostFocus()
    cboTranType.Enabled = False
    If SetTranTypeCode(cboTranType.Text) = "OTH" Then
        Dim rsJoy                                                   As ADODB.Recordset
        Set rsJoy = New ADODB.Recordset
        'JJE 01/08/2013 11:36 To generate next reference number
        Set rsJoy = gconDMIS.Execute("SELECT MAX(CAST(RIGHT(REFERENCENO,8) AS int)) AS REFERENCENO FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND trantype = 'OTH' AND LEN(Reference) = '8'")
        If Not rsJoy.EOF And Not rsJoy.BOF Then
            txtReference.Text = Format(N2Str2Zero(rsJoy!ReferenceNo) + 1, "00000000")
        Else
            txtReference.Text = "00000001"
        End If
        'JJE
        
        txtReference.Enabled = False
        cmdTranSave.Enabled = True
        cboPaidFor.Enabled = True
        cboBranch.Enabled = True
        txtDescript.Enabled = True
        txtDiscount.Enabled = True
        txtTax.Enabled = True
        Payment.Enabled = True
        'Call txtReference_KeyDown(vbKeyReturn, 0)
        On Error Resume Next
        cboPaidFor.SetFocus
    End If
End Sub

Private Sub chkCreditCardTrans_Click()
    If chkCreditCardTrans.Value = 1 Then
        'cmdCardPayment.Enabled = True
        txtReference1.Text = GetReferenceNo
        txtReference1.Locked = True
        lblReference1.Visible = True
        txtReference1.Visible = True
    Else
        cmdCardPayment.Enabled = False
        txtReference1.Visible = False
        lblReference1.Visible = False
        txtReference1.Text = ""
        txtDiscount.Text = "0.00"
        txtTax.Text = "0.00"
    End If
End Sub

Private Sub chkSelect_Click()
    Dim iCount                                                      As Integer
    If lblTotal.Caption <> 0 Then lblTotal.Caption = "0.00"
    If chkSelect.Value = 1 Then
        For iCount = 1 To lvPayments.ListItems.Count
            lvPayments.ListItems(iCount).Checked = True
            lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(lvPayments.ListItems(iCount).SubItems(3)))
        Next
    Else
        'UPDATED BY: ROWEL DE QUIROZ
        'DATE : MARCH 3 2011
        'DESCRIPTION :
        For iCount = 1 To lvPayments.ListItems.Count
            lvPayments.ListItems(iCount).Checked = False
            lblTotal.Caption = "0.00"
        Next
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", LocalAcess) = False Then Exit Sub
    
    On_Update = True
    AddorEdit = "ADD"
    grdDetails.Enabled = False
    picOR.Enabled = True
    
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    cmdORDetail.Enabled = False
    cmdInvoiceDetail.Enabled = False
    cmdInvoices.Enabled = True
    initMemvars
    
    'JJE with prefix
    'If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then **FOR APPROVAL**
    
    If OR_OPTION = True Then
        If COMPANY_CODE = "DJM" And OR_VAT_NONVAT = "NON-VAT" Then
            txtOR_NUM.Text = GetLASTOR("SOA")
            picORType.ZOrder 1
            picORType.Visible = False
            cmdSelect.Enabled = True
        Else
            picORType.ZOrder 0
            picORType.Visible = True
            If picORType.Visible = True Then
                optGoods.Value = True
                optService.Value = False
                optGoods.SetFocus
            End If
        End If
    '... JRE 06/08/2016 auto increment of OR number for CMC
    Else
        If COMPANY_CODE = "CMC" Then
            If OR_VAT_NONVAT = "VAT" Then
                txtOR_NUM.Text = GetLASTOR("V")
                picORType.ZOrder 1
                picORType.Visible = False
                cmdSelect.Enabled = True
                txtOR_NUM.Enabled = True
            Else
                txtOR_NUM.Text = GetLASTOR("NV")
                picORType.ZOrder 1
                picORType.Visible = False
                cmdSelect.Enabled = True
                txtOR_NUM.Enabled = True
            End If
    '...
        Else
            picORType.ZOrder 1
            picORType.Visible = False
            cmdSelect.Enabled = True
        End If
    End If
    'JJE
    On Error Resume Next
End Sub

Private Sub cmdCancel_Click()
    On_Update = False
    picOR.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    grdDetails.Enabled = True
    fraDetails.Enabled = True
    cmdORDetail.Enabled = True
    cmdInvoiceDetail.Enabled = True
    cmdInvoices.Enabled = False
    StoreMemVars
    picORType.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdCancelOR_Click()
    If CheckORCutOff(txtOR_NUM) = True Then
        MsgBox "Cancel is not allowed. Cut Off has been processed.", vbInformation, "Message"
        picOptions.Visible = False
        Exit Sub
    Else
        If MsgBox("Cancel this O.R. Entries, Are you Sure?", vbQuestion + vbYesNo, "Confirm Cancelation") = vbYes Then
            'JRE 11162016 - USER WILL NOT BE ABLE TO CANCEL AN OR IF OR DETAIL IS NOT YET DELETED
            Dim rsCheckDetails                                      As ADODB.Recordset
            
            Set rsCheckDetails = New ADODB.Recordset
            rsCheckDetails.Open "SELECT * FROM CMIS_OFF_DT WHERE OR_NUM = '" & txtOR_NUM.Text & "'", gconDMIS, adOpenKeyset
            If Not rsCheckDetails.EOF And Not rsCheckDetails.BOF Then
                MsgBox "Delete OR detail/s first.", vbCritical, "Information"
            Else
                If COMPANY_CODE = "HGC" Then
                    gconDMIS.Execute "UPDATE CMIS_ORS SET Status = 'C',CANCELLEDDATE = '" & CDate(LOGDATE) & "' WHERE ORNO = '" & txtOR_NUM.Text & "'"
                    'Update By BTT:06/05/2008
    
                    SQL_STATEMENT = "UPDATE CMIS_OFF_HD SET dateCancel='" & CDate(LOGDATE) & "' WHERE OR_NUM = '" & txtOR_NUM.Text & "'"
                    gconDMIS.Execute SQL_STATEMENT
                    
                    If OR_VAT_NONVAT = "VAT" Then
                        NEW_LogAudit "C", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", "OR NO: " & Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                    Else
                        NEW_LogAudit "C", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", "OR NO: " & Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                    End If
                End If
    
                SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Cancel = 1 WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "C", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                Else
                    NEW_LogAudit "C", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                End If
    
                SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET payment = 0, Cancel = 1 WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "CC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                Else
                    NEW_LogAudit "CC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, labid, "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                End If
            End If
            rsRefresh
            rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
            picOptions.Visible = False
            StoreMemVars
        End If
    End If
End Sub

Private Sub cmdCardCancel_Click()
    lvPayments.ListItems.Clear
    picCreditCard.ZOrder 1
    picCreditCard.Visible = False
    lblView.Visible = True
End Sub

Private Sub cmdCardPayment_Click()
    'txtDiscount.Text = ToDoubleNumber(NumericVal(txtAmount.Text) * 0.05)
    'txtTax.Text = ToDoubleNumber(NumericVal(txtAmount.Text) * 0.005)
End Sub

Private Sub cmdCardSave_Click()
    'UPDATED BY: ROWEL DE QUIROZ
    'DATE : MARCH 04 2011
    'DESCRIPTION :

    Dim vtrantype                                                   As String
    Dim vOR_NUM                                                     As String
    Dim vInvoiceno                                                  As String
    Dim vDescript                                                   As String
    Dim vinvoicetype                                                As String
    Dim vReference                                                  As String
    Dim vREFERENCENO                                                As String
    Dim vCUSCDE                                                     As String
    Dim vBalance                                                    As String
    Dim vDOCDTE                                                     As String
    Dim vORDATE                                                     As String
    Dim vPAYMENT                                                    As String
    Dim vDISCOUNT                                                   As String
    Dim vTAX                                                        As String
    Dim vPaidFor                                                    As String
    Dim vAmount                                                     As String
    Dim vBRANCH                                                     As String
    Dim vOVER                                                       As String
    Dim vCUTDATE                                                    As String
    Dim vBankCharges                                                As Double
    Dim vEWT                                                        As Double
    Dim vTotal                                                      As Double
    Dim IS_VAT                                                      As Integer
    Dim iCount                                                      As Integer
    Dim C                                                           As Integer
    Dim i                                                           As Integer
    Dim X                                                           As Integer
    Dim SQL_STATEMENT                                               As ADODB.Recordset

    If SetTranTypeCode(cboTranType.Text) = "PI" Or SetTranTypeCode(cboTranType.Text) = "AI" Or SetTranTypeCode(cboTranType.Text) = "MI" Or SetTranTypeCode(cboTranType.Text) = "SI" Or SetTranTypeCode(cboTranType.Text) = "VI" Or SetTranTypeCode(cboTranType.Text) = "UI" Then
        vinvoicetype = N2Str2Null(cboInvoiceType.Text)
    Else
        vinvoicetype = "NULL"
    End If

    vOR_NUM = N2Str2Null(txtOR_NUM.Text)
    vtrantype = N2Str2Null(SetTranTypeCode(cboTranType.Text))
    vREFERENCENO = N2Str2Null(txtReference1.Text)
    vCUSCDE = N2Str2Null(txtCUSCDE.Text)

    If SetTranTypeCode(cboTranType.Text) <> "RO" Then
        vReference = N2Str2Null(txtReference.Text)
        vInvoiceno = N2Str2Null(txtReference.Text)
    Else
        If labRef.Caption = "Ref. '" Then
            vReference = N2Str2Null(txtReference.Text)
            vInvoiceno = N2Str2Null(labReference.Caption)
        Else
            vReference = N2Str2Null(labReference.Caption)
            vInvoiceno = N2Str2Null(txtReference.Text)
        End If
    End If

    If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
        vCUSCDE = N2Str2Null(txtCUSCDE.Text)
    Else
        vCUSCDE = N2Str2Null(labCUSCODE.Caption)

    End If

    '===================================================================================
    i = 0
    For C = 1 To lvPayments.ListItems.Count
        If lvPayments.ListItems(C).Checked = True Then
            i = i + 1
        End If
    Next C
        If i <> 0 Then
            'nothing
        Else
            MsgBox "No record selected.", vbCritical + vbOKOnly
            Exit Sub
        End If

    '===================================================================================

    For X = 1 To lvPayments.ListItems.Count
        If lvPayments.ListItems(X).Checked = True Then
            vCUTDATE = "NULL"
            vOR_NUM = txtOR_NUM.Text
            vinvoicetype = cboInvoiceType.Text
            vtrantype = N2Str2Null(SetTranTypeCode(cboTranType.Text))
            vReference = N2Str2Null(txtReference.Text)
            vCUSCDE = lvPayments.ListItems(X).SubItems(1)
            vDescript = cboPaidFor.Text
            vBalance = NumericVal(txtBalance.Text)
            vAmount = NumericVal(lvPayments.ListItems(X).SubItems(3))
            vInvoiceno = N2Str2Null(Mid(lvPayments.ListItems(X), 1, 500))
            vDOCDTE = "NULL"
            vORDATE = lvPayments.ListItems(X).SubItems(5)
            vPAYMENT = NumericVal(lvPayments.ListItems(X).SubItems(3))
            'vDISCOUNT = NumericVal(txtDiscount.Text)
            'vTAX = NumericVal(txtTax.Text)
            vPaidFor = N2Str2Null(SetPaidForCode(cboPaidFor.Text))
            vBRANCH = N2Str2Null(SetBranchCode(cboBranch.Text))
            vOVER = NumericVal(NumericVal(Payment.Text) - NumericVal(txtBalance.Text))
            If OR_VAT_NONVAT = "VAT" Then IS_VAT = 1 Else IS_VAT = 0
            vREFERENCENO = N2Str2Null(lvPayments.SelectedItem.SubItems(4))

            Dim rsCardCompany                                       As ADODB.Recordset
            Set rsCardCompany = New ADODB.Recordset
            If COMPANY_CODE = "DJM" Then
                rsCardCompany.Open "SELECT * FROM CMIS_CardCompany WHERE CUSCDE = '" & txtCUSCDE.Text & "'", gconDMIS, adOpenKeyset
            Else
                rsCardCompany.Open "SELECT * FROM CMIS_CardBank WHERE CUSCDE = '" & txtCUSCDE.Text & "'", gconDMIS, adOpenKeyset
            End If
            If Not rsCardCompany.EOF And Not rsCardCompany.BOF Then
                vBankCharges = NumericVal(rsCardCompany!BankCharges) / 100
                vEWT = NumericVal(rsCardCompany!EWT) / 100
                vTotal = 1 - (vBankCharges + vEWT)
            End If
            
            vPAYMENT = Format(NumericVal(lvPayments.ListItems(X).SubItems(3)) - (NumericVal(lvPayments.ListItems(X).SubItems(3)) * vBankCharges + (NumericVal(lvPayments.ListItems(X).SubItems(3)) * vEWT)), "#,###,##0.00")
            vPAYMENT = NumericVal(vPAYMENT)
            Payment = vPAYMENT
            txtDiscount.Text = ToDoubleNumber(lvPayments.ListItems(X).SubItems(3)) * vBankCharges
            txtTax.Text = ToDoubleNumber(lvPayments.ListItems(X).SubItems(3)) * vEWT
            
            vDISCOUNT = txtDiscount.Text
            vTAX = txtTax.Text

            gconDMIS.Execute "INSERT INTO CMIS_Off_Dt " & _
                             "(CUTDATE,OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],VAT,REFERENCENO)" & _
                             " VALUES (" & vCUTDATE & "," & vOR_NUM & ",'" & vinvoicetype & "'," & vtrantype & "," & vReference & "," & vInvoiceno & ",'" & vCUSCDE & "','" & vDescript & "','" & vBalance & "','" & vAmount & "'," & vDOCDTE & ",'" & vORDATE & "','" & vPAYMENT & "','" & vDISCOUNT & "','" & vTAX & "'," & vPaidFor & "," & vBRANCH & ",'" & vOVER & "','" & IS_VAT & "'," & vREFERENCENO & ")"
            lvDeposits.ListItems.Clear
        End If

        If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
            If NumericVal(vPAYMENT) > NumericVal(vBalance) Then
                MsgBox "The Payment Amount is Greater than Balance Amount", vbInformation, "Message"
                If MsgBox("Accept Over Payment?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                Else
                    Exit Sub
                End If
            End If
        End If

        If labDocDate.Caption = "[DOC DATE]" Then vDOCDTE = "NULL" Else vDOCDTE = N2Date2Null(labDocDate.Caption)
        vORDATE = N2Str2Null(txtOR_DATE.Text)
        If OR_VAT_NONVAT = "VAT" Then IS_VAT = 1 Else IS_VAT = 0
        '====
        If CheckIfBank(txtCUSCDE.Text) = True Then
            gconDMIS.Execute "UPDATE CMIS_Off_Hd SET PAIDBY ='Y' WHERE OR_NUM = '" & vOR_NUM2 & "'"
        End If

        If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
            vinvoicetype = SetPaidForCode(cboPaidFor.Text)
            Select Case vinvoicetype
                Case "412P"
                    vinvoicetype = "'PI'"
                Case "412S"
                    vinvoicetype = "'SI'"
                Case "412V"
                    vinvoicetype = "'VI'"
            End Select
        End If
    Next X

    rsRefresh
    rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
    StoreMemVars
    ShowSuccessFullyUpdated
    cmdTranCancel.Value = True
    Exit Sub
End Sub

Private Sub cmdDetails_Click()
    On_Update = True
    cmdDetails.Enabled = False
    cmdDetails.ZOrder 0
    cmdDetails.Visible = True
    picDetails.ZOrder 0
    picDetails.Visible = True
    'UPDATE BY   : MJP 09032008 05:41 PM
    'DESCRIPTION : TO LIMIT THE USER ON CLICKING THE NAVIGATION BUTTON WHILE ADDING A DETAILS
    Picture1.Enabled = False
    fraDetails.Enabled = False
    chkCreditCardTrans.Value = 0
    'SetTranTypeCode(cboTranType.Text) = "Vehicle Invoice"
    'UPDATE BY   : MJP 09032008 05:41 PM

    InitGridMemvars
    If TranType <> "" Then
        cboTranType.Text = TranType
    End If
    AddorEdit = "ADD"
    labStatusMode.Caption = "System is Adding OR Detail..."
    cmdTranSave.Enabled = False
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", LocalAcess) = False Then Exit Sub
    'JJE MITSUBISHI ONLY
    If vDetails = True Then
        MsgBox "Please delete first the OR detail(s)", vbCritical, "Official Receipt"
        Exit Sub
    End If
    'JJE
    On_Update = True
    AddorEdit = "EDIT"
    PrevOR_NUM = txtOR_NUM.Text
    grdDetails.Enabled = False
    picOR.Enabled = True
    cmdSelect.Enabled = True
    
    'JJE Disabled editing of OR NUMBER
    'If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Then
        txtOR_NUM.Enabled = False
    End If
    'JJE
    
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    cmdORDetail.Enabled = False
    cmdInvoiceDetail.Enabled = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsOFF_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdInsurance_Click()
    SelectEntity = "Vendor"
    vINSTPL = "INS"
    Set frmNewEntity = New frmEntity
    Call frmNewEntity.LOADJOURNAL("CMIS")
    frmNewEntity.Show 1
End Sub

Private Sub cmdInvoiceDetail_Click()
    grdDetails.Col = 0
    INVOICE_DETAIL_TYPE = grdDetails.Text
    grdDetails.Col = 1
    INVOICE_DETAIL_TRANNO = grdDetails.Text
    frmInvoiceAppDetail.Show vbModal
End Sub

Private Sub cmdLast_Click()
    rsOFF_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsOFF_HD.MoveNext
    If rsOFF_HD.EOF Then
        rsOFF_HD.MoveLast
        ShowLastRecordMsg
        ' MsgBox "Last of Record!", vbInformation, "Info"
        '   MessagePop NaviBegin, "End of Record", "Last Record", 1000
    End If
    StoreMemVars
End Sub

Private Sub cmdOK_Click()
    If Option2.Value = True Then
        Dim xList                                                   As ListItem
        Dim rsCMIS_OFF_HD                                           As ADODB.Recordset
        Set rsCMIS_OFF_HD = New ADODB.Recordset
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3'AND OR_Date >= '" & dtFrom & "' AND OR_Date <= '" & dtTo & "' AND (Paidby IS NULL OR paidby = 'N') ORDER BY OR_Date")
        If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
            lvPayments.ListItems.Clear
            lblTotal = "0.00"
            Do While Not rsCMIS_OFF_HD.EOF
                Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
                xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
                xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!cusname)
                xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
                xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
                xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
                tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
                lblTotal = Format(tmpTotal, "#,###,##0.00")
                rsCMIS_OFF_HD.MoveNext
            Loop
        Else
            MessagePop RecNotFound, "No record to view", "No Record"
        End If
    End If
End Sub

Private Sub cmdOptions_Click()
    If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub
    picOptions.Visible = True
    picOptions.ZOrder 0
End Sub

Private Sub cmdORDetail_Click()
    OR_NUMBER_GLOBAL = txtOR_NUM.Text
    frmORPaymentDetail.Show vbModal
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:
    Dim rsCheckTransaction                                  As New ADODB.Recordset
    If vDetails = 0 Then
        MsgBox "No information to be posted!", vbInformation, "Message"
    Else
        If CheckPostedOR(txtOR_NUM) = True Then
            MsgBox "Transaction has been posted.", vbInformation, "Posted"
        Else
            picPayment.ZOrder 0
            picPayment.Visible = True
            optCASH.Value = True
            optCHECK.Value = False
            optCANCEL.Value = False
            If CheckIfBank(txtCUSCDE.Text) = True Then
                'JRE 09202016 - FOR THE BANK CUSTOMERS TO BE ABLE TO PAY VIA CARD PAYMENT
                Set rsCheckTransaction = New ADODB.Recordset
                rsCheckTransaction.Open "SELECT * FROM CMIS_OFF_DT WHERE OR_NUM = '" & txtOR_NUM.Text & "'", gconDMIS, adOpenKeyset
                If rsCheckTransaction!TranType = "OTH" Then
                    If rsCheckTransaction!PAIDFOR = "427" Then
                        optCARD.Enabled = False
                    Else
                        optCARD.Enabled = True
                    End If
                Else
                    optCARD.Enabled = True
                End If
            Else
                optCARD.Enabled = True
            End If
            On Error Resume Next
            optCASH.SetFocus
        End If
    End If
    'LogAudit "P", "OFFICIAL RECEIPT", txtOR_NUM
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    rsOFF_HD.MovePrevious
    If rsOFF_HD.BOF Then
        rsOFF_HD.MoveFirst
        ShowFirstRecordMsg
        'MsgBox "First of Record!", vbInformation, "Info"
        'MessagePop NaviBegin, "Beginning of Record", "First Record", 1000
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", LocalAcess) = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:
    Dim xlApp                                                       As Excel.Application
    Dim xlBook                                                      As Excel.Workbook
    Dim xlSheet                                                     As Excel.Worksheet
    Dim rsTMP                                                       As New ADODB.Recordset
    Dim XCNT                                                        As Integer
    'UPDATED BY: JUN
    'DATE UPDATED: 02-10-2009
    rptChat.Reset
    'UPDATED BY: JUN

    If MsgBox("Print Official Receipt now?", vbQuestion + vbYesNo) = vbYes Then
        If COMPANY_CODE = "HEI" Then
            XCNT = 3
            Set xlApp = CreateObject("Excel.Application")
            Set xlBook = xlApp.Workbooks.Open(CMIS_REPORT_PATH & "\OR.xlt")
            Set xlSheet = xlBook.Worksheets(1)

            Set rsTMP = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_NUM = " & N2Str2Null(txtOR_NUM) & "")
            If Not (rsTMP.BOF And rsTMP.EOF) Then
                Do While Not rsTMP.EOF
                    xlSheet.Cells(XCNT, "A") = Null2String(rsTMP!TranType) & Null2String(rsTMP!REFERENCE)
                    xlSheet.Cells(XCNT, "B") = Format(NumericVal(rsTMP!Payment), MAXIMUM_DIGIT)

                    XCNT = XCNT + 1
                    rsTMP.MoveNext
                Loop
            End If

            XCNT = 9
            xlSheet.Cells(XCNT, "i") = txtOR_DATE

            XCNT = 10
            xlSheet.Cells(XCNT, "F") = cboCUSNAME

            XCNT = 16
            xlSheet.Cells(XCNT, "F") = NumToText(NumericVal(rsOFF_HD!OR_AMT))

            If rsOFF_HD!TOF = 1 Then XCNT = 18
            If rsOFF_HD!TOF = 2 Then XCNT = 19
            If rsOFF_HD!TOF = 3 Then XCNT = 20
            xlSheet.Cells(XCNT, "B") = Format(NumericVal(rsOFF_HD!OR_AMT), MAXIMUM_DIGIT)

            xlApp.Windows.Item(1).Caption = "Official Receipt"
            xlApp.Visible = True
            Set xlApp = Nothing
        Else
            If OR_VAT_NONVAT = "VAT" Then
                If CR_OR_PRINTING = True Then
                    If Left(txtOR_NUM, 2) = "CR" Then
                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                    End If
                Else
                    'JJE Official Receipt of DSSC and DGI for service and goods 12/18/2012
                    If COMPANY_CODE = "DSSC" Then
                        If Left(txtOR_NUM.Text, 1) = "S" Then
                            PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                        Else
                            PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                        End If
                    ElseIf COMPANY_CODE = "DGI" Then
                        If Left(txtOR_NUM.Text, 1) = "S" Then
                            Dim RSTRANTYPE                                                     As ADODB.Recordset
                            Dim TYPE_SI                                                        As Double
                            Dim TYPE_OTH                                                       As Double
                            
                            TYPE_SI = 0
                            TYPE_OTH = 0
                            
                            Set RSTRANTYPE = New ADODB.Recordset
                            Set RSTRANTYPE = gconDMIS.Execute("Select TRANTYPE from cmis_off_dt where or_num = '" & txtOR_NUM.Text & "'")
                            Do While Not RSTRANTYPE.EOF
                                If (RSTRANTYPE!TranType) = "SI" Then
                                    TYPE_SI = TYPE_SI + 1
                                ElseIf (RSTRANTYPE!TranType) = "OTH" Then
                                    TYPE_OTH = TYPE_OTH + 1
                                End If
                                RSTRANTYPE.MoveNext
                            Loop
                            
                            If TYPE_SI = 0 And TYPE_OTH >= 1 Then
                                PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptOTH.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                            Else
                                PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                            End If
                        Else
                            PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                        End If
                    'JJE Official Receipt of JABEZ for service and goods with REPRINTED on twice printing
    '                ElseIf COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
    '                    SaveReprintInformation OR_VAT_NONVAT, MODULENAME, txtOR_NUM.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
    '                    rptChat.Formulas(0) = "REPRINT = '" & REPRINT_CAPTION & "'"
    '                    If Left(txtOR_NUM.Text, 1) = "S" Then
    '                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
    '                    Else
    '                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 1" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
    '                    End If
                    'JJE
                    ElseIf COMPANY_CODE = "CMC" Then
                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                    'JJE
                    Else
                        If Left(txtOR_NUM.Text, 1) = "G" Then
                            PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 1" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                        Else
                            PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 1" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                        End If
                    End If
                End If
            Else
                'JJE Update Report DSSC for service and goods 12/18/2012
                If COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Then
                    If Left(txtOR_NUM.Text, 1) = "S" Then
                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 0" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                    Else
                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 0" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                    End If
                'JJE Official Receipt of JABEZ for service and goods REPRINTED
'                ElseIf COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
'                    SaveReprintInformation OR_VAT_NONVAT, MODULENAME, txtOR_NUM.Text, "Null", LOGDATE, LOGNAME, False: If CANCEL_ANS = "NO" Then Exit Sub
'                    rptChat.Formulas(0) = "REPRINT = '" & REPRINT_CAPTION & "'"
'                    If Left(txtOR_NUM.Text, 1) = "S" Then
'                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 0" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
'                    Else
'                        PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptGoods.rpt", "{OFF_HD.VAT} = 0" & " AND {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
'                    End If
                'JJE
                ElseIf COMPANY_CODE = "CMC" Then
                    PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceiptNonVat.rpt", "{OFF_HD.VAT} = 0" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                ElseIf COMPANY_CODE = "DJM" Then
                    PrintSQLReport rptChat, CMIS_REPORT_PATH & "SOA.rpt", "{OFF_HD.VAT} = 0" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                Else
                    PrintSQLReport rptChat, CMIS_REPORT_PATH & "OfficialReceipt.rpt", "{OFF_HD.VAT} = 0" & " and {OFF_HD.OR_NUM} = '" & txtOR_NUM.Text & "'", DMIS_REPORT_Connection, 1
                End If
            End If
            'JJE
        End If
    End If

    If OR_VAT_NONVAT = "VAT" Then
        NEW_LogAudit "V", "TRANSACTION O.R. WITH VAT", "", labid, "", "OR NO: " & txtOR_NUM, "VAT", ""
    Else
        NEW_LogAudit "V", "TRANSACTION O.R. WITHOUT VAT", "", labid, "", "OR NO: " & txtOR_NUM, "NON VAT", ""
    End If
    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdRecoverOR_Click()
    If MsgBox("Recover this O.R. Entries? Are you Sure?", vbQuestion + vbYesNo, "Confirm Recovery") = vbYes Then
        If CheckIfCancel(txtOR_NUM) = True Then
            SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET Cancel = 0 WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "RC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            Else
                NEW_LogAudit "RC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
            End If
            '=================================================

            SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET Cancel = 0 WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
            
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "RC", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            Else
                NEW_LogAudit "RC", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            End If

            SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET payment = " & RECEIPTS_AMOUNT & " WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
            gconDMIS.Execute SQL_STATEMENT

            If MODE_OF_PAYMENT = "CASH" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CASH = CASH + " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf MODE_OF_PAYMENT = "CHECK" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " [CHECK] = [CHECK] + " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            ElseIf MODE_OF_PAYMENT = "CARD" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CARD = CARD + " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
            End If
        Else
            MsgBox "OR No. is not cancelled.", vbInformation, "O.R. Recovery"
        End If

        '        If OR_VAT_NONVAT = "VAT" Then
        '            LogAudit "R", "OFFICIAL RECEIPT DATA ENTRY [VAT]", "OR NO.: " & txtOR_NUM
        '        Else
        '            LogAudit "R", "OFFICIAL RECEIPT DATA ENTRY [NON VAT]", "OR NO.: " & txtOR_NUM
        '        End If

        rsRefresh
        rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
        picOptions.Visible = False
        StoreMemVars
    End If
End Sub

Private Sub cmdRef_Click()
    If labRef.Caption = "Ref. '" Then
        labRef.Caption = "Inv. '"
    Else
        labRef.Caption = "Ref. '"
    End If
    On Error Resume Next
    txtReference.SetFocus
End Sub

Private Sub cmdRefresh_Click()
    InitGridMemvars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    
    Dim varOR_NUM, varOR_DATE, varCUSCDE, varCUSNAME                As String
    Dim IS_VAT                                                      As Integer
    Dim rsCheckORNUM                                                As ADODB.Recordset

    If txtCUSCDE.Text = "" Then
        MsgBox "Customer not yet added in Customer Master File..."
        Exit Sub
    ElseIf txtOR_NUM.Text = "" Then
        MessagePop InfoFriend, "OR Number", "OR Number is blank"
        Exit Sub
    End If
    
    varOR_NUM = N2Str2Null(Format(txtOR_NUM.Text, "000000"))
    varOR_DATE = N2Str2Null(txtOR_DATE.Text)
    varCUSCDE = N2Str2Null(RTrim(LTrim(txtCUSCDE.Text)))
    varCUSNAME = N2Str2Null(cboCUSNAME.Text)

    If OR_VAT_NONVAT = "VAT" Then
        IS_VAT = 1
    Else
        IS_VAT = 0
    End If
    
'    'JJE 12/10/2013
'    If COMPANY_CODE = "CMC" Then
'        If IS_VAT = 1 Then
'            If (txtOR_NUM < 76001 Or txtOR_NUM > 97000) Then
'                MsgBox "OR Number is not allowed! You can only use between 076001 and 097000 for OR VAT", vbCritical + vbOKOnly, "Invalid OR No."
'                Exit Sub
'            End If
'        ElseIf IS_VAT = 0 Then
'            If (txtOR_NUM < 1 Or txtOR_NUM > 15000) Then
'                MsgBox "OR Number is not allowed! You can only use between 000001 and 015000 for OR NONVAT", vbCritical + vbOKOnly, "Invalid OR No."
'                Exit Sub
'            End If
'        End If
'    End If
'    'JJE

    If AddorEdit = "ADD" Then
        If VALID_COMPANY_CODE_FORHAI = True Then
            Set rsCheckORNUM = gconDMIS.Execute("SELECT OR_NUM FROM CMIS_Off_hd WHERE VAT = " & IS_VAT & " AND OR_NUM = " & varOR_NUM)
        Else
            Set rsCheckORNUM = gconDMIS.Execute("SELECT OR_NUM FROM CMIS_Off_hd WHERE OR_NUM = " & varOR_NUM)
        End If
        If Not rsCheckORNUM.EOF And Not rsCheckORNUM.BOF Then
            Screen.MousePointer = 0
            MsgBox "OR Number is already used! Pls. use valid OR number...", vbCritical + vbOKOnly, "Invalid OR No."
            On Error Resume Next
            
            txtOR_NUM.SetFocus
            txtOR_NUM.SelLength = Len(txtOR_NUM)
            Exit Sub
        End If
    Else
        If varOR_NUM <> N2Str2Null(rsOFF_HD!OR_NUM) Then
            'If PrevOR_NUM <> txtOR_NUM.Text Then
            Set rsCheckORNUM = gconDMIS.Execute("SELECT OR_NUM FROM CMIS_Off_hd WHERE VAT = " & IS_VAT & " AND OR_NUM = " & varOR_NUM)
            If Not rsCheckORNUM.EOF And Not rsCheckORNUM.BOF Then
                Screen.MousePointer = 0
                MsgBox "OR Number already used! Pls. use valid OR number...", vbCritical + vbOKOnly, "Invalid OR No."
                On Error Resume Next
                
                txtOR_NUM.SetFocus
                txtOR_NUM.SelLength = Len(txtOR_NUM)
                Exit Sub
            End If
            'End If
        End If
    End If

    If AddorEdit = "ADD" Then
        If COMPANY_CODE = "DJM" Then
            SQL_STATEMENT = "INSERT INTO CMIS_Off_hd " & _
                            "(OR_NUM,OR_DATE,CUSCDE,CUSNAME,DATECREATE,TIMECREATE,VAT,STATUS,PR_NO,FAO,PR_DATE) " & _
                            "VALUES (" & varOR_NUM & "," & varOR_DATE & "," & varCUSCDE & "," & varCUSNAME & ",'" & LOGDATE & "','" & Time & "'," & IS_VAT & ",'N','" & txtPRNo.Text & "','" & txtFao.Text & "','" & txtPRDate.Text & "')"
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = "INSERT INTO CMIS_Off_hd " & _
                "(OR_NUM,OR_DATE,CUSCDE,CUSNAME,DATECREATE,TIMECREATE,VAT,STATUS) " & _
                "VALUES (" & varOR_NUM & "," & varOR_DATE & "," & varCUSCDE & "," & varCUSNAME & ",'" & LOGDATE & "','" & Time & "'," & IS_VAT & ",'N')"
            gconDMIS.Execute SQL_STATEMENT
        End If

        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "A", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        Else
            NEW_LogAudit "A", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        End If

        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET" & _
                        " OR_NUM = " & N2Str2Null(varOR_NUM) & "," & _
                        " VAT = " & VAT_OR & _
                        " WHERE VAT = " & IS_VAT & " AND OR_NUM = " & N2Str2Null(PrevOR_NUM)
        gconDMIS.Execute SQL_STATEMENT

        If OR_VAT_NONVAT = "VAT" Then
            If NumericVal(FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd")) > 0 Then NEW_LogAudit "EE", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_Dt")
        Else
            If NumericVal(FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd")) > 0 Then NEW_LogAudit "EE", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_Dt")
        End If
        
        If COMPANY_CODE = "DJM" Then
            SQL_STATEMENT = " UPDATE CMIS_Off_Hd SET" & _
                            " VAT = " & VAT_OR & "," & _
                            " OR_NUM = " & N2Str2Null(varOR_NUM) & "," & _
                            " OR_DATE = " & N2Str2Null(varOR_DATE) & "," & _
                            " CUSCDE = " & N2Str2Null(varCUSCDE) & "," & _
                            " CUSNAME = " & N2Str2Null(varCUSNAME) & "," & _
                            " PR_NO = " & N2Str2Null(txtPRNo.Text) & "," & _
                            " FAO = " & N2Str2Null(txtFao.Text) & "," & _
                            " PR_DATE = " & N2Str2Null(txtPRDate.Text) & _
                            " WHERE VAT = " & IS_VAT & " AND OR_NUM = " & N2Str2Null(PrevOR_NUM)
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = " UPDATE CMIS_Off_Hd SET" & _
                            " VAT = " & VAT_OR & "," & _
                            " OR_NUM = " & N2Str2Null(varOR_NUM) & "," & _
                            " OR_DATE = " & N2Str2Null(varOR_DATE) & "," & _
                            " CUSCDE = " & N2Str2Null(varCUSCDE) & "," & _
                            " CUSNAME = " & N2Str2Null(varCUSNAME) & " " & _
                            " WHERE VAT = " & IS_VAT & " AND OR_NUM = " & N2Str2Null(PrevOR_NUM)
            gconDMIS.Execute SQL_STATEMENT
        End If
        
        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "E", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        Else
            NEW_LogAudit "E", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(varOR_NUM, "OR_NUM", "CMIS_Off_hd"), "", Null2String(varOR_NUM), OR_VAT_NONVAT, ""
        End If

        ShowSuccessFullyUpdated
    End If

    rsRefresh
    rsOFF_HD.Find "OR_NUM = " & varOR_NUM
    cmdCancel.Value = True
    cmdSelect.Enabled = False
    FillGrid
    Screen.MousePointer = 0
    If AddorEdit = "ADD" Then cmdDetails_Click
    Exit Sub

ErrorCode:
    Screen.MousePointer = 0
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    SelectCustomer = "Customer"
    frmCustomerSearch1.Show 1
End Sub

Private Sub cmdTranCancel_Click()
'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:
    On_Update = False
    
    'JJE Cancelled Applied Deposit will be Unapplied (12/14/2012 4:12PM)
    ApplyDeposits = False
    picDetails.ZOrder 1: picDetails.Visible = False
    cmdDetails.ZOrder 1: cmdDetails.Visible = False
    'UPDATE BY   : MJP 09032008 05:41 PM
    'DESCRIPTION : TO LIMIT THE USER ON CLICKING THE NAVIGATION BUTTON WHILE ADDING A DETAILS
    fraDetails.Enabled = True
    Picture1.Enabled = True
    'UPDATE BY   : MJP 09032008 05:41 PM
    StoreMemVars
    picCreditCard.ZOrder 1: picCreditCard.Visible = False
    picDeposits.ZOrder 1: picDeposits.Visible = False
    Exit Sub
    
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdTranDelete_Click()
    On Error GoTo ErrorCode:
    If MsgQuestionBox("Delete This Entry, Are you Sure?", "Delete OR Entry") = True Then
        'JJE check if Customer Deposit is already applied
        Dim rsDeposits                                      As ADODB.Recordset
        Set rsDeposits = gconDMIS.Execute("SELECT OR_Num FROM CMIS_OFF_DT WHERE OR_Num IN (SELECT OR_Num FROM CMIS_DEPOSITS WHERE OR_Num = '" & lblDetID.Caption & "' AND Applied = 'Y')")
        If Not rsDeposits.EOF And Not rsDeposits.BOF Then
            MessagePop InfoWarning, "Applied Payment", "Customer deposit cannot be deleted!"
        Else
            If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
                SQL_STATEMENT = "DELETE FROM CMIS_Off_Dt WHERE OR_NUM = '" & txtOR_NUM & "' AND reference = '" & txtReference & "'"
            Else
                SQL_STATEMENT = "DELETE FROM CMIS_Off_Dt WHERE OR_NUM = '" & txtOR_NUM & "' AND referenceno = '" & txtReference & "'"
            End If
            gconDMIS.Execute SQL_STATEMENT
            
            gconDMIS.Execute ("DELETE FROM CMIS_OFF_DT WHERE OR_Num ='" & lblDetID.Caption & "' and invoiceno = '" & txtReference & "' and descript = 'DEPOSIT APPLIED'")
            gconDMIS.Execute ("DELETE FROM CMIS_Deposits WHERE OR_Num ='" & lblDetID.Caption & "'")
            
            'JJE Delete Customer Deposit Applied and Update Deposit Table set Applied to 'N' (01/08/2013 2:41PM)
            Dim Dep_ID                                              As String
            Dim rsDep_ID                                            As ADODB.Recordset
            
            Set rsDep_ID = New ADODB.Recordset
            Set rsDep_ID = gconDMIS.Execute("SELECT * FROM cmis_depositdt WHERE or_num = '" & txtOR_NUM & "'")
            If Not rsDep_ID.EOF And Not rsDep_ID.BOF Then
                Dep_ID = Null2String(rsDep_ID!Deposit_id)
                
                gconDMIS.Execute ("DELETE FROM CMIS_DEPOSITDT WHERE INVOICENO = '" & txtReference.Text & "'")

                Set rsDep_ID = gconDMIS.Execute("SELECT * FROM CMIS_Depositdt WHERE Deposit_id = " & Dep_ID & "")
                If rsDep_ID.EOF And rsDep_ID.BOF Then
                    gconDMIS.Execute ("UPDATE CMIS_Deposits SET Applied ='N' WHERE ID ='" & Dep_ID & "'")
                End If
            End If
            'JJE
            
            gconDMIS.Execute "UPDATE CMIS_Off_Hd SET PAIDBY = 'N' WHERE ReferenceNo = '" & lblRefNo & "'"
            gconDMIS.Execute "UPDATE CMIS_Off_Hd SET OR_AMT=NULL,BAYADAMT=NULL,CASHAMOUNT=NULL,CHKAMOUNT=NULL,TOF=NULL,ReferenceNo=NULL,Bank=NULL WHERE OR_NUM = '" & txtOR_NUM & "'"
           
            '=================================================
            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "XX", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            Else
                NEW_LogAudit "XX", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            End If
            '=================================================
            ShowDeletedMsg
        End If
    End If

    rsRefresh
    On Error Resume Next
    rsOFF_HD.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
    Exit Sub
    
ErrorCode:
    ShowVBError
    'MsgBox Error
End Sub

Private Sub cmdTranSave_Click()
    On Error GoTo ErrorCode
    Dim str_MSG                                                     As String
    
    If cboTranType.Text = "" Then
        MsgBox "Transaction cannot be save", vbCritical, "Error!"
        Exit Sub
    ElseIf txtReference.Text = "" Then
        MsgBox "Transaction cannot be save", vbCritical, "Error!"
        Exit Sub
    ElseIf SetTranTypeCode(cboTranType.Text) = "OTH" Then
        If cboPaidFor.Text = "" Then
            MsgBox "Field cannot be empty. Please select.", vbCritical, "Error!"
            cboPaidFor.SetFocus
            Exit Sub
        Else
'            If CHECKIFSCHED(GetChartCodes(cboPaidFor.Text)) = True Then
'                If lblVendorName.Caption = "" Then
'                    MsgBox "Please SELECT specific vendor for this schedule account.", vbCritical, "Schedule Account"
'                    cmdInsurance.SetFocus
'                    Exit Sub
'                End If
'            End If
            'JJE TAGGING OF SCHEDULE ACCOUNT VALIDATION V.2.2
            If vENTITY = "" Then
                If Tagging(cboPaidFor) = True Then
                    MsgBox "Please select specific Vendor for this Schedule Account.", vbCritical, "Schedule Account"
                    cmdInsurance.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    str_MSG = "Error in Saving @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telephone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    
    If SaveTransaction = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "OR Details")
        MsgBox str_MSG, vbCritical, "Saving Error "
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0
    Exit Sub

ErrorCode:
    ShowVBError
    'MsgBox Error
End Sub

Function SaveTransaction() As Boolean
    On Error GoTo ErrorCode
    
    Dim vOR_NUM                                                     As String
    Dim vSUB_OR_NUM                                                 As String
    Dim vReference                                                  As String
    Dim vInvoiceno                                                  As String
    Dim vCUSCDE                                                     As String
    Dim varCUSCDE                                                   As String
    Dim vDescript                                                   As String
    Dim vDOCDTE                                                     As String
    Dim vORDATE                                                     As String
    Dim vPaidFor                                                    As String
    Dim vBRANCH                                                     As String
    Dim vBalance                                                    As String
    Dim vAmount                                                     As String
    Dim vPAYMENT                                                    As String
    Dim vDISCOUNT                                                   As String
    Dim vTAX                                                        As Double
    Dim IS_VAT                                                      As Integer
    Dim vOVER                                                       As Double

    Dim vLTORegFee                                                  As String
    Dim vInsuranceFee                                               As String
    Dim vChattelFee                                                 As String
    Dim vOthers                                                     As String
    Dim vDownAmount                                                 As String
    Dim vLTORegFeeAmount                                            As String
    Dim vInsuranceFeeAmount                                         As String
    Dim vChattelFeeAmount                                           As String
    Dim vOthersAmount                                               As String
    Dim vDownFee                                                    As String
    Dim vOVER2                                                      As String
    Dim vOVERIns                                                    As String
    Dim vOVERLTO                                                    As String
    Dim vOVERTPL                                                    As String
    Dim vLTOBal                                                     As String
    Dim vOthersBal                                                  As String
    Dim vInsBal                                                     As String
    Dim vDownBal                                                    As String
    Dim vChattelBal                                                 As String
    Dim vOVERCHATTEL                                                As String
    
    '***************************************************************************
    'updating code:     jaa - 11202008      - save trantype for PI,SI,MI,AI only
    Dim vinvoicetype                                                As String
    If SetTranTypeCode(cboTranType.Text) = "PI" Or SetTranTypeCode(cboTranType.Text) = "AI" Or SetTranTypeCode(cboTranType.Text) = "MI" Or SetTranTypeCode(cboTranType.Text) = "SI" Or SetTranTypeCode(cboTranType.Text) = "VI" Or SetTranTypeCode(cboTranType.Text) = "UI" Then
        vinvoicetype = N2Str2Null(cboInvoiceType.Text)
    Else
        vinvoicetype = "NULL"
    End If
    '***************************************************************************

    vOR_NUM = N2Str2Null(txtOR_NUM.Text)
    vSUB_OR_NUM = N2Str2Null(txtOR_NUM.Text)
    vtrantype = N2Str2Null(SetTranTypeCode(cboTranType.Text))
    varCUSCDE = N2Str2Null(txtCUSCDE.Text)
    
    If SetTranTypeCode(cboTranType.Text) = "OTH" Then
        If SetPaidForCode(cboPaidFor.Text) = "427" Then
            vREFERENCENO = N2Str2Null(txtReference.Text)
            vReference = labReference.Caption
            vInvoiceno = labReference.Caption
        Else
            vREFERENCENO = N2Str2Null(txtReference.Text)
            vReference = N2Str2Null(txtReference.Text)
            vInvoiceno = N2Str2Null(txtReference.Text)
        End If
    Else
        vREFERENCENO = N2Str2Null(vREFERENCENO)
        vReference = N2Str2Null(txtReference.Text)
        vInvoiceno = N2Str2Null(txtReference.Text)
    End If
    

    'JJE Check the type of Customer Deposit by Code
    If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
        vCUSCDE = N2Str2Null(txtCUSCDE.Text)
    Else
        vCUSCDE = N2Str2Null(txtCUSCDE.Text)
'        vCUSCDE = N2Str2Null(labCUSCODE.Caption)

    End If
    
    vDescript = N2Str2Null(txtDescript.Text)
'    vBalance = NumericVal(txtBalance.Text)
    vAmount = NumericVal(txtAmount.Text)
    
    'JJE
    If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
        If txtBalance.Text > 0 Then
            If N2Str2Zero(Replace(txtBalance, ",", "")) > N2Str2Zero(Replace(Payment, ",", "")) Then 'Partial Payment
                vPAYMENT = N2Str2Zero(Replace(Payment, ",", ""))
            ElseIf N2Str2Zero(Replace(txtBalance, ",", "")) < N2Str2Zero(Replace(Payment, ",", "")) Then 'Over Payment
                vPAYMENT = N2Str2Zero(Replace(Payment, ",", ""))
            Else
                vPAYMENT = Round(NumericVal(txtBalance) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
                wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
            End If
        Else
            vPAYMENT = Round(NumericVal(txtAmount) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
        End If
    Else
        vPAYMENT = Round(NumericVal(Payment.Text), 2)
    End If
    'JJE
    
    vDISCOUNT = NumericVal(txtDiscount.Text)
    vTAX = NumericVal(txtTax.Text)
    
    vPaidFor = N2Str2Null(SetPaidForCode(cboPaidFor.Text))
    vBRANCH = N2Str2Null(SetBranchCode(cboBranch.Text))
    
    vOthers = NumericVal(txtTPLPay.Text)
    vDownFee = NumericVal(txtDownPay.Text)
    vLTORegFee = NumericVal(txtLTOPay.Text)
    vInsuranceFee = NumericVal(txtInsPay.Text)
    vChattelFee = NumericVal(txtChattelPay.Text)
    
    vOthersAmount = NumericVal(txtTPLAmout.Text)
    vDownAmount = NumericVal(txtDownAmount.Text)
    vLTORegFeeAmount = NumericVal(txtLTOAmount.Text)
    vInsuranceFeeAmount = NumericVal(txtInsAmount.Text)
    vChattelFeeAmount = NumericVal(txtChattelAmount.Text)
    
    vInsBal = NumericVal(txtInsBal.Text)
    vLTOBal = NumericVal(txtLTOBal.Text)
    vDownBal = NumericVal(txtDownBal.Text)
    vOthersBal = NumericVal(txtOtherBal.Text)
    vChattelBal = NumericVal(txtChattelBal.Text)
    
    vOVER = NumericVal(NumericVal(txtDownPay.Text) - NumericVal(txtDownBal.Text))
    vOVERTPL = NumericVal(NumericVal(txtTPLPay.Text) - NumericVal(txtOtherBal.Text))
    vOVERIns = NumericVal(NumericVal(txtInsPay.Text) - NumericVal(txtInsBal.Text))
    vOVERLTO = NumericVal(NumericVal(txtLTOPay.Text) - NumericVal(txtLTOBal.Text))
    vOVERCHATTEL = NumericVal(NumericVal(txtChattelPay.Text) - NumericVal(txtChattelBal.Text))
    vOVER2 = NumericVal(NumericVal(Payment.Text) - NumericVal(txtBalance.Text))
    
    'JJE 01/08/2013 10:10AM
    If vDeposits <= 0 Then
        If vPAYMENT <= 0 Then
            MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
            SaveTransaction = True
            Exit Function
            Payment.SetFocus
        End If
    End If
    'JJE
    
    If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
        If NumericVal(vPAYMENT) > NumericVal(txtBalance.Text) Then
            MsgBox "The Payment is greater than Balance Amount", vbInformation, "Message"
            If MsgBox("Accept Over Payment?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            Else
                SaveTransaction = True
                Exit Function
            End If
        End If
    End If
    
    If labDocDate.Caption = "[DOC DATE]" Then vDOCDTE = "NULL" Else vDOCDTE = N2Date2Null(labDocDate.Caption)
    vORDATE = N2Str2Null(txtOR_DATE.Text)
    If OR_VAT_NONVAT = "VAT" Then IS_VAT = 1 Else IS_VAT = 0
    If AddorEdit = "ADD" Then
        'Updated: ACL 05292009
        Dim rsCardCompany                                           As ADODB.Recordset
        Dim vBankCharges                                            As Double
        Dim vEWT                                                    As Double
        Dim vTotal                                                  As Double
    
        Dim vLTO                                                    As String
        Dim vIns                                                    As String
        Dim vOthers1                                                As String
        Dim vDown                                                   As String
        Dim vChattel                                                As String
        Dim vNetsales                                               As String
        Dim vDesc                                                   As String
        '*initialize payment type
        vLTO = "L"
        vIns = "I"
        vOthers1 = "O"
        vChattel = "C"
        
        If vTerm = True Then
            vDown = "N"
            vDesc = "Vehicle price"
        Else
            vDown = "D"
            vDesc = "Downpayment"
        End If

        Set rsCardCompany = New ADODB.Recordset
        rsCardCompany.Open "SELECT * FROM CMIS_CardBank WHERE CUSCDE = '" & txtCUSCDE.Text & "'", gconDMIS, adOpenKeyset
        If Not rsCardCompany.EOF And Not rsCardCompany.BOF Then
            vBankCharges = NumericVal(rsCardCompany!BankCharges) / 100
            vEWT = NumericVal(rsCardCompany!EWT) / 100
            vTotal = 1 - (vBankCharges + vEWT)
        End If

        If lvPayments.ListItems.Count <> 0 Then
            If SetPaidForCode(cboPaidFor.Text) = "427" Then
                vREFERENCENO = N2Str2Null(lvPayments.SelectedItem.SubItems(4))
                vReference = N2Str2Null(lvPayments.SelectedItem.Text)
                vInvoiceno = N2Str2Null(lvPayments.SelectedItem.Text)
                vCUSCDE = N2Str2Null(lvPayments.SelectedItem.SubItems(1))
                xTTLINV = xTTLINV
                
                vPAYMENT = Round(NumericVal(lvPayments.SelectedItem.SubItems(3) * vTotal), 2)
                
                If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
                    If txtBalance.Text > 0 Then
                        If (ToDoubleNumber(txtBalance.Text) > ToDoubleNumber(Payment.Text)) Then
                            vPAYMENT = Payment
                        Else
                            vPAYMENT = Round(NumericVal(txtBalance) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
                            wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
                        End If
                    Else
                        vPAYMENT = Round(NumericVal(txtAmount) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
                        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
                    End If
                Else
                    vPAYMENT = Round(NumericVal(Payment.Text), 2)
                End If
                
                vOR_NUM2 = N2Str2Null(lvPayments.SelectedItem.Text)
                vPAYMENT = Round(vPAYMENT, 2)
            End If
        End If
                
        'JJE For vehicle
        Dim rsCountRecord                                           As ADODB.Recordset
        If SetTranTypeCode(cboTranType.Text) = "VI" And vCustype <> "B" Then
            If COMPANY_CODE <> "DJM" Then
                If vDownFee <= 0 And vLTORegFee <= 0 And vInsuranceFee <= 0 And vOthers <= 0 And vChattelFee <= 0 Then
                    MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
                    Payment.SetFocus
                    Exit Function
                End If
        
                If SetTranTypeCode(cboTranType.Text) <> "OTH" Then
                    If NumericVal(vPAYMENT) > NumericVal(txtBalance.Text) Then
                        MsgBox "The Payment Amount is Greater than balance Amount", vbInformation, "Message"
                        If MsgBox("Accept Over Payment?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
                        Else
                            SaveTransaction = True
                            Exit Function
                        End If
                    End If
                End If
        
                If chkDownPayment.Value = 1 Then
                    vDescript = "Payment for" + vDesc
                    vDescript = N2Str2Null(vDescript)
                    vPaidFor = "'VII'"
                    vDown = N2Str2Null(vDown)
                        If txtDownPay.Text = "0.00" Or NumericVal(vDownFee) > NumericVal(vDownBal) Then
                            MsgBox "Kindly check the Payment Amount.", vbExclamation, "Invalid Payment"
                            txtDownPay.SetFocus
                            Exit Function
                        End If
                    
                        Set rsCountRecord = gconDMIS.Execute("SELECT COUNT(PAYMENTTYPE) AS totl FROM CMIS_Off_Dt WHERE reference = " & vInvoiceno & " AND PAYMENTTYPE = " & N2Str2Null(vDown) & " AND OR_NUM = " & vOR_NUM)
                        If Not rsCountRecord.EOF And Not rsCountRecord.BOF Then
                            If rsCountRecord.Fields("totl").Value > 0 Then
                                MessagePop RecLocekd, "System Info", "Cannot Process your Request. Duplicate Payment of " + vDesc
                                Call cmdTranCancel_Click
                                grdDetails.Enabled = True
                                Exit Function
                            Else
                                'JRE 07/20/16 If deposit applied is true
                                If ApplyDeposits = True Then
                                Set rsDepositcheck = New ADODB.Recordset
                                Set rsDepositcheck = gconDMIS.Execute("SELECT ISNULL(AMOUNT,0) AS AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO = '" & txtReference.Text & "' AND INVOICETYPE = '" & SetTranTypeCode(cboTranType.Text) & "' AND OR_NUM = " & vOR_NUM & "")
                                If rsDepositcheck.EOF And rsDepositcheck.BOF Then
                                    If Payment.Text < (N2Str2Zero(vBalance) - N2Str2Zero(vDeposits)) Then
                                        MsgBox "The Payment Amount is Less than balance Amount", vbInformation, "Message"
                                        Payment.Text = vBalance - vDeposits
                                        Payment.SetFocus
                                        SaveTransaction = True
                                        Exit Function
                                    Else
                                        xTTLINV = xTTLINV
                                       'JJE Insert invoice Detail
                                       'JRE 07/20/16 To display the correct amount paid by the customer
                                        Dim PaidAmount                      As Double
                                        PaidAmount = NumericVal(vDownFee - vDeposits)
                                        SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                                        "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                                                        " VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & N2Str2Null(vDescript) & "," & vDownBal & "," & vDownAmount & "," & vDOCDTE & "," & vORDATE & "," & PaidAmount & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & vDown & ")"
                                        'JRE
                                        gconDMIS.Execute SQL_STATEMENT
                                        'JJE Insert Detail of Deposit Applied into invoice
                                        SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                                        "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAIDFOR,BRANCH,ORIGINAL_D,VAT) " & _
                                                        "VALUES (" & vOR_NUM & "," & vinvoicetype & ",'OTH','" & OR_REFERENCE & "'," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & ", 'DEPOSIT APPLIED'," & vDeposits & "," & vDOCDTE & "," & vORDATE & "," & vPaidFor & "," & vBRANCH & "," & vDOCDTE & "," & IS_VAT & ")"
                                        gconDMIS.Execute SQL_STATEMENT
                                        xTTLINV = 0
            
                                        'JJE Insert Deposit Applied information to CMIS_DEPOSITDT
                                        gconDMIS.Execute ("INSERT INTO CMIS_DEPOSITDT(OR_NUM,INVOICENO,INVOICETYPE,AMOUNT,DEPOSIT_ID,PAYMENTFOR) " & _
                                                        "VALUES ('" & txtOR_NUM.Text & "','" & txtReference.Text & "','" & SetTranTypeCode(cboTranType.Text) & "','" & vDeposits & "','" & lblDepositID.Caption & "','" & vPaymentType & "')")
                                        lvDeposits.ListItems.Clear
            
                                        'gconDMIS.Execute ("UPDATE CMIS_Deposits SET Applied ='Y' WHERE ID ='" & lblDepositID.Caption & "'")
                                    End If
                                End If
                            Else
                                xTTLINV = xTTLINV
                                'Insert Invoice Detail
                                SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                                "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                                                " VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & N2Str2Null(vDescript) & "," & vDownBal & "," & vDownAmount & "," & vDOCDTE & "," & vORDATE & "," & vDownFee & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & vDown & ")"
                                gconDMIS.Execute SQL_STATEMENT
                            End If
                        End If
                    End If
                    rsCountRecord.Close
                End If
                '*******************
                If chkLTORegFee.Value = 1 Then
                    vLTO = N2Str2Null(vLTO)
                    vDescript = "Payment for LTO"
                    vDescript = N2Str2Null(vDescript)
                    vPaidFor = "414"
                    If txtLTOPay = "0.00" Or NumericVal(vLTORegFee) > NumericVal(vLTOBal) Then
                        MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
                        txtLTOPay.SetFocus
                        Exit Function
                    End If
                
                    Set rsCountRecord = gconDMIS.Execute("SELECT COUNT(PAYMENTTYPE) AS totl FROM CMIS_Off_Dt WHERE reference = " & vInvoiceno & " AND PAYMENTTYPE = " & N2Str2Null(vLTO) & " AND OR_NUM = " & vOR_NUM)
                    If Not rsCountRecord.EOF And Not rsCountRecord.BOF Then
                        If rsCountRecord.Fields("totl").Value > 0 Then
                            MessagePop RecLocekd, "System Info", "Cannot Process your Request. Duplicate Payment of LTO Registration fee "
                            grdDetails.Enabled = True
                            Call cmdTranCancel_Click
                            Exit Function
                        Else
                            'JRE 08012016 - REMOVE THE VALUE OF INVOICETYPE AND CHANGE THE TRANTYPE TO OTH
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                            "(OR_NUM,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                            " VALUES (" & vOR_NUM & ",'OTH'," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vLTOBal & "," & vLTORegFeeAmount & "," & vDOCDTE & "," & vORDATE & "," & vLTORegFee & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVERLTO & "," & vDOCDTE & "," & IS_VAT & "," & vLTO & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                    End If
                    rsCountRecord.Close
                End If
                
                '*******************
                If chkInsurance.Value = 1 Then
                    vIns = N2Str2Null(vIns)
                    vDescript = "Payment for Insurance"
                    vDescript = N2Str2Null(vDescript)
                    vPaidFor = "413"
                    If txtInsPay.Text = "0.00" Or NumericVal(vInsuranceFee) > NumericVal(vInsBal) Then
                        MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
                        txtInsPay.SetFocus
                        Exit Function
                    End If
                  
                    Set rsCountRecord = gconDMIS.Execute("SELECT COUNT(PAYMENTTYPE) AS totl FROM CMIS_Off_Dt WHERE reference = " & vInvoiceno & " AND PAYMENTTYPE = " & N2Str2Null(vIns) & " AND OR_NUM = " & vOR_NUM)
                    If Not rsCountRecord.EOF And Not rsCountRecord.BOF Then
                        If rsCountRecord.Fields("totl").Value > 0 Then
                            MessagePop RecLocekd, "System Info", "Cannot Process your Request. Duplicate Payment of Insurance fee"
                            Call cmdTranCancel_Click
                            grdDetails.Enabled = True
                            Exit Function
                        Else
                            'JRE 08012016 - REMOVE THE VALUE OF INVOICETYPE AND CHANGE THE TRANTYPE TO OTH
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                            "(OR_NUM,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                            " VALUES (" & vOR_NUM & ",'OTH'," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vInsBal & "," & vInsuranceFeeAmount & "," & vDOCDTE & "," & vORDATE & "," & vInsuranceFee & "," & vDISCOUNT & "," & vTAX & "," & "413" & "," & vBRANCH & "," & vOVERIns & "," & vDOCDTE & "," & IS_VAT & "," & vIns & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                    End If
                    rsCountRecord.Close
                End If
                '******************
                If chkTPL.Value = 1 Then
                    vOthers1 = N2Str2Null(vOthers1)
                    vDescript = "Payment for TPL"
                    vDescript = N2Str2Null(vDescript)
                    vPaidFor = "416"
                    If txtTPLPay.Text = "0.00" Or NumericVal(vOthers) > NumericVal(vOthersBal) Then
                        MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
                        txtTPLPay.SetFocus
                        Exit Function
                    End If
                    
                    Set rsCountRecord = gconDMIS.Execute("SELECT COUNT(PAYMENTTYPE) AS totl FROM CMIS_Off_Dt WHERE reference = " & vInvoiceno & " AND PAYMENTTYPE = " & N2Str2Null(vOthers1) & " AND OR_NUM = " & vOR_NUM)
                    If Not rsCountRecord.EOF And Not rsCountRecord.BOF Then
                        If rsCountRecord.Fields("totl").Value > 0 Then
                            MessagePop RecLocekd, "System Info", "Cannot Process your Request. Duplicate Payment of TPL fee"
                            Call cmdTranCancel_Click
                            grdDetails.Enabled = True
                            Exit Function
                        Else
                            'JRE 08012016 - REMOVE THE VALUE OF INVOICETYPE AND CHANGE THE TRANTYPE TO OTH
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                            "(OR_NUM,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                            " VALUES (" & vOR_NUM & ",'OTH'," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vOthersBal & "," & vOthersAmount & "," & vDOCDTE & "," & vORDATE & "," & vOthers & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVERTPL & "," & vDOCDTE & "," & IS_VAT & "," & vOthers1 & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                    End If
                    rsCountRecord.Close
                End If
                '******************
                If chkChattel.Value = 1 Then
                    vChattel = N2Str2Null(vChattel)
                    vDescript = "Payment for Chattel"
                    vDescript = N2Str2Null(vDescript)
                    vPaidFor = "435"
                    If txtChattelPay.Text = "0.00" Or NumericVal(vChattelFee) > NumericVal(vChattelBal) Then
                        MsgBox "Kindly check the Payment Amount.", vbInformation, "Invalid Payment"
                        txtChattelPay.SetFocus
                        Exit Function
                    End If
                    
                    Set rsCountRecord = gconDMIS.Execute("SELECT COUNT(PAYMENTTYPE) AS totl FROM CMIS_Off_Dt WHERE reference = " & vInvoiceno & " AND PAYMENTTYPE = " & N2Str2Null(vChattel) & " AND OR_NUM = " & vOR_NUM)
                    If Not rsCountRecord.EOF And Not rsCountRecord.BOF Then
                        If rsCountRecord.Fields("totl").Value > 0 Then
                            MessagePop RecLocekd, "System Info", "Cannot Process your Request. Duplicate Payment of Chattel fee"
                            Call cmdTranCancel_Click
                            grdDetails.Enabled = True
                            Exit Function
                        Else
                            'JRE 08012016 - REMOVE THE VALUE OF INVOICETYPE AND CHANGE THE TRANTYPE TO OTH
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                            "(OR_NUM,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,BALANCE,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,PAYMENTTYPE)" & _
                            " VALUES (" & vOR_NUM & ",'OTH'," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vChattelBal & "," & vChattelFeeAmount & "," & vDOCDTE & "," & vORDATE & "," & vChattelFee & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVERCHATTEL & "," & vDOCDTE & "," & IS_VAT & "," & vChattel & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                    End If
                    rsCountRecord.Close
                End If
            Else
                If ApplyDeposits = True Then
                    Set rsDepositcheck = New ADODB.Recordset
                    Set rsDepositcheck = gconDMIS.Execute("SELECT ISNULL(AMOUNT,0) AS AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO = '" & txtReference.Text & "' AND INVOICETYPE = '" & SetTranTypeCode(cboTranType.Text) & "' AND OR_NUM = " & vOR_NUM & "")
                    If rsDepositcheck.EOF And rsDepositcheck.BOF Then
                        If Payment.Text < (N2Str2Zero(vBalance) - N2Str2Zero(vDeposits)) Then
                            MsgBox "The Payment Amount is Less than balance Amount", vbInformation, "Message"
                            Payment.Text = vBalance - vDeposits
                            Payment.SetFocus
                            SaveTransaction = True
                            Exit Function
                        Else
                            xTTLINV = xTTLINV
                            'JJE Insert invoice Detail
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                            "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,ENTITY,XTTLINVAMT)" & _
                                            " VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vAmount & "," & vDOCDTE & "," & vORDATE & "," & vPAYMENT & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & N2Str2Null(vENTITY) & ",'" & xTTLINV & "')"
                            gconDMIS.Execute SQL_STATEMENT
                            
                            'JJE Insert Detail of Deposit Applied into invoice
                            SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                            "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAIDFOR,BRANCH,ORIGINAL_D,VAT) " & _
                                            "VALUES (" & vOR_NUM & "," & vinvoicetype & ",'OTH','" & OR_REFERENCE & "'," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & ", 'DEPOSIT APPLIED'," & vDeposits & "," & vDOCDTE & "," & vORDATE & "," & vPaidFor & "," & vBRANCH & "," & vDOCDTE & "," & IS_VAT & ")"
                            gconDMIS.Execute SQL_STATEMENT
                            xTTLINV = 0
                            
                            'JJE Insert Deposit Applied information to CMIS_DEPOSITDT
                            gconDMIS.Execute ("INSERT INTO CMIS_DEPOSITDT(OR_NUM,INVOICENO,INVOICETYPE,AMOUNT,DEPOSIT_ID,PAYMENTFOR) " & _
                                              "VALUES ('" & txtOR_NUM.Text & "','" & txtReference.Text & "','" & SetTranTypeCode(cboTranType.Text) & "','" & vDeposits & "','" & lblDepositID.Caption & "','" & vPaymentType & "')")
                            lvDeposits.ListItems.Clear
                        
                            'gconDMIS.Execute ("UPDATE CMIS_Deposits SET Applied ='Y' WHERE ID ='" & lblDepositID.Caption & "'")
                        End If
                    End If
                Else
                    xTTLINV = xTTLINV
                    'Insert Invoice Detail
                    SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                    "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,ENTITY,XTTLINVAMT) " & _
                                    "VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vAmount & "," & vDOCDTE & "," & vORDATE & "," & Round(vPAYMENT, 2) & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & N2Str2Null(vENTITY) & ",'" & xTTLINV & "')"
                    gconDMIS.Execute SQL_STATEMENT
                    xTTLINV = 0
                    lvDeposits.ListItems.Clear
                    
                    'JJE For Customer Deposit
                    If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
                        Set rsDet_ID = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_Num = " & N2Str2Null(txtOR_NUM.Text) & "")
                        If Not rsDet_ID.EOF And Not rsDet_ID.BOF Then
                            SQL_STATEMENT = "INSERT INTO CMIS_Deposits " & _
                                            "(CusCde,ORDate,OR_Num,Amount,Applied,PaidFor)" & _
                                            " VALUES (" & varCUSCDE & "," & vORDATE & "," & vOR_NUM & ", " & vPAYMENT & ", 'N'," & vPaidFor & ")"
                            gconDMIS.Execute SQL_STATEMENT
                        End If
                        Set rsDet_ID = Nothing
                    End If
                End If
            
                'BANK FOR CREDIT CARD TRANSACTION
                If vOR_NUM2 = "" Then
                    vOR_NUM2 = vOR_NUM
                End If
        
                gconDMIS.Execute ("INSERT INTO CMIS_TranList " & _
                              "(VAT,TRANTYPE,REFERENCE,DOCDTE)" & _
                              " VALUES (" & VAT_OR & "," & vtrantype & "," & vReference & "," & vDOCDTE & ")")

                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "AA", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                Else
                    NEW_LogAudit "AA", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                End If
                ShowSuccessFullyAdded
            End If
        Else
            'JJE Deposit is applied (01/08/2013)
            If ApplyDeposits = True Then
                Set rsDepositcheck = New ADODB.Recordset
                Set rsDepositcheck = gconDMIS.Execute("SELECT ISNULL(AMOUNT,0) AS AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO = '" & txtReference.Text & "' AND INVOICETYPE = '" & SetTranTypeCode(cboTranType.Text) & "' AND OR_NUM = " & vOR_NUM & "")
                If rsDepositcheck.EOF And rsDepositcheck.BOF Then
                    If Payment.Text < (N2Str2Zero(vBalance) - N2Str2Zero(vDeposits)) Then
                        MsgBox "The Payment Amount is Less than balance Amount", vbInformation, "Message"
                        Payment.Text = vBalance - vDeposits
                        Payment.SetFocus
                        SaveTransaction = True
                        Exit Function
                    Else
                        xTTLINV = xTTLINV
                        'JJE Insert invoice Detail
                        SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                        "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,ENTITY,XTTLINVAMT)" & _
                                        " VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vAmount & "," & vDOCDTE & "," & vORDATE & "," & vPAYMENT & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & N2Str2Null(vENTITY) & ",'" & xTTLINV & "')"
                        gconDMIS.Execute SQL_STATEMENT
                        
                        'JJE Insert Detail of Deposit Applied into invoice
                        SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                        "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAIDFOR,BRANCH,ORIGINAL_D,VAT) " & _
                                        "VALUES (" & vOR_NUM & "," & vinvoicetype & ",'OTH','" & OR_REFERENCE & "'," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & ", 'DEPOSIT APPLIED'," & vDeposits & "," & vDOCDTE & "," & vORDATE & "," & vPaidFor & "," & vBRANCH & "," & vDOCDTE & "," & IS_VAT & ")"
                        gconDMIS.Execute SQL_STATEMENT
                        xTTLINV = 0
                        
                        'JJE Insert Deposit Applied information to CMIS_DEPOSITDT
                        gconDMIS.Execute ("INSERT INTO CMIS_DEPOSITDT(OR_NUM,INVOICENO,INVOICETYPE,AMOUNT,DEPOSIT_ID,PAYMENTFOR) " & _
                                          "VALUES ('" & txtOR_NUM.Text & "','" & txtReference.Text & "','" & SetTranTypeCode(cboTranType.Text) & "','" & vDeposits & "','" & lblDepositID.Caption & "','" & vPaymentType & "')")
                        lvDeposits.ListItems.Clear
                    
                        'gconDMIS.Execute ("UPDATE CMIS_Deposits SET Applied ='Y' WHERE ID ='" & lblDepositID.Caption & "'")
                    End If
                End If
            Else
                xTTLINV = xTTLINV
                'Insert Invoice Detail
                SQL_STATEMENT = "INSERT INTO CMIS_Off_Dt " & _
                                "(OR_NUM,INVOICETYPE,TRANTYPE,REFERENCE,REFERENCENO,INVOICENO,CUSCDE,DESCRIPT,AMOUNT,DOCDTE,ORDATE,PAYMENT,DISCOUNT,TAX,PAIDFOR,BRANCH,[OVER],ORIGINAL_D,VAT,ENTITY,XTTLINVAMT) " & _
                                "VALUES (" & vOR_NUM & "," & vinvoicetype & "," & vtrantype & "," & vReference & "," & vREFERENCENO & "," & vInvoiceno & "," & vCUSCDE & "," & vDescript & "," & vAmount & "," & vDOCDTE & "," & vORDATE & "," & Round(vPAYMENT, 2) & "," & vDISCOUNT & "," & vTAX & "," & vPaidFor & "," & vBRANCH & "," & vOVER & "," & vDOCDTE & "," & IS_VAT & "," & N2Str2Null(vENTITY) & ",'" & xTTLINV & "')"
                gconDMIS.Execute SQL_STATEMENT
                xTTLINV = 0
                lvDeposits.ListItems.Clear
                
                'JJE For Customer Deposit
                If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
                    Set rsDet_ID = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_Num = " & N2Str2Null(txtOR_NUM.Text) & "")
                    If Not rsDet_ID.EOF And Not rsDet_ID.BOF Then
                        SQL_STATEMENT = "INSERT INTO CMIS_Deposits " & _
                                        "(CusCde,ORDate,OR_Num,Amount,Applied,PaidFor)" & _
                                        " VALUES (" & varCUSCDE & "," & vORDATE & "," & vOR_NUM & ", " & vPAYMENT & ", 'N'," & vPaidFor & ")"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    Set rsDet_ID = Nothing
                End If
            End If
            
            'BANK FOR CREDIT CARD TRANSACTION
            If vOR_NUM2 = "" Then
                vOR_NUM2 = vOR_NUM
            End If
        
            gconDMIS.Execute ("INSERT INTO CMIS_TranList " & _
                              "(VAT,TRANTYPE,REFERENCE,DOCDTE)" & _
                              " VALUES (" & VAT_OR & "," & vtrantype & "," & vReference & "," & vDOCDTE & ")")

            If OR_VAT_NONVAT = "VAT" Then
                NEW_LogAudit "AA", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            Else
                NEW_LogAudit "AA", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
            End If
            ShowSuccessFullyAdded
        End If
    Else
        If Payment.Text = "0.00" Then
            MsgBox "Cannot accept zero payment", vbCritical
            Exit Function
        End If
    
        If SetTranTypeCode(cboTranType.Text) = "VI" And vCustype <> "B" And COMPANY_CODE <> "DJM" Then
            Dim rsBalance                                           As New ADODB.Recordset
            Dim balance                                             As String
            Dim num                                                 As Long
            Dim num2                                                As Long
            
            Set rsBalance = gconDMIS.Execute("SELECT ROUND(SUM(balance),2) AS Balance FROM CMIS_OFF_DT WHERE OR_NUM = " & N2Str2Null(txtOR_NUM.Text) & " AND invoiceno = " & N2Str2Null(txtReference.Text) & " AND PAYMENTTYPE = " & N2Str2Null(vPaymentType))
                balance = NumericVal(rsBalance!balance)
                num = Val(balance)
                num2 = Val(vPAYMENT)
                If (num2 > num) Then
                    MsgBox "Your Payment is Greater than the previous Balance!!", vbCritical
                    Exit Function
                End If
                
            SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET " & _
                            " VAT = " & IS_VAT & "," & _
                            " INVOICETYPE = " & vinvoicetype & "," & _
                            " TRANTYPE = " & N2Str2Null(vtrantype) & "," & _
                            " REFERENCE = " & N2Str2Null(vReference) & "," & _
                            " REFERENCENO = " & N2Str2Null(vREFERENCENO) & "," & _
                            " INVOICENO = " & N2Str2Null(vInvoiceno) & "," & _
                            " CUSCDE = " & N2Str2Null(vCUSCDE) & "," & _
                            " DESCRIPT = " & N2Str2Null(vDescript) & "," & _
                            " AMOUNT = " & N2Str2Zero(vAmount) & "," & _
                            " DOCDTE = " & N2Str2Null(vDOCDTE) & "," & _
                            " ORDATE = " & N2Str2Null(vORDATE) & "," & _
                            " PAYMENT = " & N2Str2Zero(vPAYMENT) & "," & _
                            " DISCOUNT = " & N2Str2Zero(vDISCOUNT) & "," & _
                            " TAX = " & N2Str2Null(vTAX) & "," & _
                            " PAIDFOR = " & N2Str2Null(vPaidFor) & "," & _
                            " BRANCH = " & N2Str2Null(vBRANCH) & "," & _
                            " ENTITY = " & N2Str2Null(vENTITY) & "," & _
                            " [OVER] = " & N2Str2Null(vOVER) & _
                            " WHERE ID = " & labDetID.Caption
                            
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET " & _
                            " VAT = " & IS_VAT & "," & _
                            " INVOICETYPE = " & vinvoicetype & "," & _
                            " TRANTYPE = " & N2Str2Null(vtrantype) & "," & _
                            " REFERENCE = " & N2Str2Null(vReference) & "," & _
                            " REFERENCENO = " & N2Str2Null(vREFERENCENO) & "," & _
                            " INVOICENO = " & N2Str2Null(vInvoiceno) & "," & _
                            " CUSCDE = " & N2Str2Null(vCUSCDE) & "," & _
                            " DESCRIPT = " & N2Str2Null(vDescript) & "," & _
                            " AMOUNT = " & N2Str2Zero(vAmount) & "," & _
                            " DOCDTE = " & N2Str2Null(vDOCDTE) & "," & _
                            " ORDATE = " & N2Str2Null(vORDATE) & "," & _
                            " PAYMENT = " & N2Str2Zero(vPAYMENT) & "," & _
                            " DISCOUNT = " & N2Str2Zero(vDISCOUNT) & "," & _
                            " TAX = " & N2Str2Null(vTAX) & "," & _
                            " PAIDFOR = " & N2Str2Null(vPaidFor) & "," & _
                            " BRANCH = " & N2Str2Null(vBRANCH) & "," & _
                            " ENTITY = " & N2Str2Null(vENTITY) & "," & _
                            " [OVER] = " & N2Str2Null(vOVER) & _
                            " WHERE ID = " & labDetID.Caption
            gconDMIS.Execute SQL_STATEMENT
            
            'JJE Update Deposit info
            If SetPaidForCode(cboPaidFor.Text) = "412P" Or SetPaidForCode(cboPaidFor.Text) = "412S" Or SetPaidForCode(cboPaidFor.Text) = "412V" Then
                Set rsDet_ID = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_Num = " & N2Str2Null(txtOR_NUM.Text) & "")
                If Not rsDet_ID.EOF And Not rsDet_ID.BOF Then
                    SQL_STATEMENT = "UPDATE CMIS_Deposits SET " & _
                                    " CUSCDE = " & N2Str2Null(vCUSCDE) & "," & _
                                    " ORDATE = " & N2Str2Null(vORDATE) & "," & _
                                    " AMOUNT = " & N2Str2Zero(vPAYMENT) & "," & _
                                    " PAIDFOR = " & N2Str2Null(vPaidFor) & "," & _
                                    " INVOICETYPE = " & vinvoicetype & "" & _
                                    " WHERE OR_NUM = " & N2Str2Null(txtOR_NUM.Text) & ""
                    gconDMIS.Execute SQL_STATEMENT
                End If
                Set rsDet_ID = Nothing
            End If
        End If
        'JJE
        
        'JJE Update Tranlist
        SQL_STATEMENT = "UPDATE CMIS_TranList SET " & _
                        " VAT = " & VAT_OR & "," & _
                        " TRANTYPE = " & N2Str2Null(vtrantype) & "," & _
                        " REFERENCE = " & N2Str2Null(vReference) & "," & _
                        " DOCDTE = " & N2Str2Null(vDOCDTE) & "" & _
                        " WHERE REFERENCE = " & N2Str2Null(vReference) & ""
        gconDMIS.Execute SQL_STATEMENT
        
        If OR_VAT_NONVAT = "VAT" Then
            NEW_LogAudit "EE", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        Else
            NEW_LogAudit "EE", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", cboTranType & ": " & Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
        End If
        '=================================================
        ShowSuccessFullyUpdated
    End If
    
    rsRefresh
    'JJE 01/09/2013 5:54PM
    picDetails.Enabled = True
    grdDetails.Enabled = True
    rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
    
    picDetails.ZOrder 1: picDetails.Visible = False
    cmdDetails.ZOrder 1: cmdDetails.Visible = False
    fraDetails.Enabled = True
    Picture1.Enabled = True
    StoreMemVars
    picCreditCard.ZOrder 1: picCreditCard.Visible = False
    picDeposits.ZOrder 1: picDeposits.Visible = False
    'JJE
    
    SaveTransaction = True
    Exit Function

ErrorCode:
    SaveTransaction = False
    'MsgBox error
End Function

Private Sub cmdVarious_Click()
'UPDATED BY AXP-061920071101
'frmALLCustomer.Show vbModal
''FillCustomer
'If CURRENT_CUST_CODE <> "" Then
'    txtCUSCDE.Text = CURRENT_CUST_CODE
'    cboCUSNAME.Text = SetCustomerName(txtCUSCDE.Text)
'End If
End Sub

Private Sub cmdInvoices_Click()
    If txtOR_NUM.Text = "" Then
        MessagePop InfoFriend, "OR Number", "OR Number is blank"
        Exit Sub
    Else
        frmCMIS_viewInvoiceDetail.Show 1
    End If
End Sub

Private Sub Command4_Click()
    picDeposits.Visible = False
    cmdTranSave.Enabled = True
End Sub

Private Sub Form_Activate()
'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If picDetails.Visible = True Then
            If picDeposits.Visible = True Then
                picDetails.Enabled = True
                picDeposits.Visible = False
                picDeposits.ZOrder 1
            ElseIf picCreditCard.Visible = True Then
                picCreditCard.Visible = False
                picCreditCard.ZOrder 1
            Else
                picDetails.Enabled = True
                fraDetails.Enabled = True
                On_Update = False
                picDetails.ZOrder 1
                picDetails.Visible = False

                cmdDetails.ZOrder 1
                cmdDetails.Visible = False
                picDeposits.Visible = False
                picCreditCard.Visible = False
                picCreditCard.ZOrder 1

                Picture1.Enabled = True
                On Error Resume Next
                grdDetails.SetFocus
            End If
        End If
        
        'JJE Customer Deposit update
        If picDeposits.Visible = True Then
            picDeposits.Visible = False
            picDeposits.ZOrder 1
            ApplyDeposits = False
        End If
        'JJE
        
        If picOptions.Visible = True Then
            picOptions.Visible = False
            picOptions.ZOrder 1
        End If
        'If picDetail.Visible = True Then
        '   picDetail.Visible = False
        'End If
    Case vbKeyF2
        If Null2Bool(rsOFF_HD!Paidna) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
            lblView.Visible = False
            If Picture1.Visible = True Then cmdDetails_Click
        End If
    Case vbKeyF3
        grdDetails_DblClick
        'Case vbKeyF7
        '     picDetail.ZOrder 0
        '     picDetail.Visible = True
        '     cmdORDetail.SetFocus
    Case vbKeyF4
        If grdDetails.Text <> "" Then
            picDeposits.Visible = True
            picDeposits.ZOrder 0
            'Call Unapplied_Deposits(txtCUSCDE.Text)
            Call Applied_Deposits(txtCUSCDE.Text)
        End If
    Case vbKeyF5
        If SetPaidForCode(cboPaidFor.Text) = "427" Then
            If CheckIfBank(txtCUSCDE.Text) = True Then
                picCreditCard.Visible = True
                picCreditCard.ZOrder 0
                lblView.Visible = False
            End If
        End If
    Case vbKeyF6
        If picDetails.Visible = True Then
            picDeposits.Visible = True
            picDeposits.ZOrder 0
            cmdTranSave.Enabled = False
            Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
        End If
    Case vbKeyF7
        If picOR.Enabled = True Then Call cmdSelect_Click
    Case vbKeyF8
        If picDetails.Visible = False Then cmdPost_Click
    Case vbKeyF11
        Shell "calc.exe"
    Case vbKeyF12
        If CheckPostedOR(txtOR_NUM.Text) = False Then
            MsgBox "Transaction is not yet posted.", vbInformation, "Message"
            Exit Sub
        Else
            If MsgBox("Unpost this O.R. Entries, Are you Sure?", vbQuestion + vbYesNo, "Unpost Transaction") = vbYes Then
                If CheckORCutOff(txtOR_NUM) = True Then
                    MsgBox "Unposting is not allowed. Cut Off has been processed.", vbInformation, "Message"
                    Exit Sub
                ElseIf CheckIfImportedinAMIS(txtOR_NUM) = True Then
                    MsgBox "Unposting is not allowed. Already imported in accounting.", vbInformation, "Message"
                    Exit Sub
                ElseIf CheckAppliedDeposit(txtOR_NUM) = True Then
                    MessagePop InfoWarning, "Unposting is not Allowed", "Customer Deposit is already Applied!"
                    Exit Sub
                Else
                    'DESCRIPTION: DEDUCT FROM CMIS_CASH_POS IF TRANSACTION IS UNPOSTED
                    '             POSTED NOT YET DEPOSITED
                    If CheckDeposited(txtOR_NUM) = False Then
                        Call UnPost_CashPos
                    Else
                        'DESCRIPTION: POSTED AND DEPOSITED
                        Call UnPost_CashPos
                        'DESCRIPTION: UNPOSTING OF TRANSACTION, SET DEPOSIT=FALSE
                        If COMPANY_CODE = M_COMPANY_CODE Then
                            gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit1 = 0 WHERE OR_NUM = '" & txtOR_NUM & "'")
                            gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit2 = 0 WHERE OR_NUM = '" & txtOR_NUM & "'")
                            gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit3 = 0 WHERE OR_NUM = '" & txtOR_NUM & "'")
                        Else
                            gconDMIS.Execute ("UPDATE CMIS_Off_Hd SET Deposit = 0 WHERE OR_NUM = '" & txtOR_NUM & "'")
                        End If
                        'DESCRIPTION: DELETE FROM BANKDEPOSIT AND CASH POSITION IF CUT OFF IS NOT YET PROCESS
                        gconDMIS.Execute ("DELETE FROM CMIS_BankDepo WHERE OR_NUM = " & N2Str2Null(txtOR_NUM))
                    End If
                End If
                '================================================
                'UPDATING CODE:     JAA - 08272008   11:00PM
                
                'JJE
                If rsOFF_HD!TOF = 1 Then
                    SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET OR_amt = 0,Bayadamt = 0,SUKLI = 0,TOF = null,TAX = 0,Discount = 0,Paidna = 0,STATUS = 'N',CASHAMOUNT = 0 WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                ElseIf rsOFF_HD!TOF = 2 Then
                    SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET OR_amt = 0,Bayadamt = 0,SUKLI = 0,TOF = null,TAX = 0,Discount = 0,Paidna = 0,STATUS = 'N',CHKAMOUNT = 0,TSEKE = null,CHECKDATE = null,BANKCODE = null,tseklase = null,BANKBRANCH = null WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                ElseIf rsOFF_HD!TOF = 3 Then
                    SQL_STATEMENT = "UPDATE CMIS_Off_Hd SET OR_amt = 0,Bayadamt = 0,SUKLI = 0,TOF = null,TAX = 0,Discount = 0,Paidna = 0,STATUS = 'N',CARDAMOUNT = 0,CARDNUMBER = NULL,CARDDATE = null,CARDBNKCDE = null,BANKBRANCH = null WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                End If
                gconDMIS.Execute SQL_STATEMENT
                'JJE
                
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "U", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_hd"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                Else
                    NEW_LogAudit "U", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_hd"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, ""
                End If
                
                SQL_STATEMENT = "UPDATE CMIS_Off_Dt SET paidna = 0, STATUS='N' WHERE VAT = " & VAT_OR & " AND OR_NUM = '" & txtOR_NUM.Text & "'"
                gconDMIS.Execute SQL_STATEMENT
                
                If CheckIfBank(txtCUSCDE.Text) = True Then
                    Getdetailsinfo
                    gconDMIS.Execute "UPDATE CMIS_Off_Hd SET PAIDBY = 'N' WHERE OR_NUM = '" & vOR_NUM2 & "'"
                End If
                
                If OR_VAT_NONVAT = "VAT" Then
                    NEW_LogAudit "UU", "TRANSACTION O.R. WITH VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                Else
                    NEW_LogAudit "UU", "TRANSACTION O.R. WITHOUT VAT", SQL_STATEMENT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_HD"), "", Null2String(txtOR_NUM), OR_VAT_NONVAT, FindTransactionID(N2Str2Null(txtOR_NUM), "OR_NUM", "CMIS_Off_Dt")
                End If
                '=================================================

                rsRefresh
                rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
                StoreMemVars
            End If
        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
    
    If Shift = 1 Then
        If KeyCode = vbKeyF1 Then
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            If OR_VAT_NONVAT = "VAT" Then
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION O.R. WITH VAT)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "TRANSACTION O.R. WITH VAT")
            Else
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION O.R. WITHOUT VAT)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "TRANSACTION O.R. WITHOUT VAT")
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Dim rsProfile                                                   As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("SELECT * FROM ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing
    CenterMe frmMain, Me, 1
    picOptions.Visible = False
    
    'JJE 01/08/2016
    If COMPANY_CODE = "DJM" Then
        txtPRNo.Enabled = True
        txtFao.Enabled = True
        txtPRDate.Enabled = True
    Else
        txtPRNo.Enabled = False
        txtFao.Enabled = False
        txtPRDate.Enabled = False
    End If
    
    
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    'JJE Disabled editing of OR NUMBER
    'If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
    If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then
        txtOR_NUM.Locked = True
    End If
    'JJE
    
    If COMPANY_CODE = "DJM" And OR_VAT_NONVAT = "NON-VAT" Then
        txtOR_NUM.Font.Size = 16.5
        txtOR_NUM.MaxLength = 9
        txtDescript.Height = 1215
    Else
        txtOR_NUM.Font.Size = 18
        txtOR_NUM.MaxLength = 8
        txtDescript.Height = 615
    End If
    
    picOR.Enabled = False
    FillGrid
    initMemvars
    textSearch.Text = ""
    FillCustomer
    FillType
    FillBranch
    FillPayment
    FillInvoiceType
    On_Update = False
    
    If OR_VAT_NONVAT = "VAT" Then
        VAT_OR = 1
    Else
        VAT_OR = 0
    End If
    
    FIRST_LOAD = True
    rsRefresh
    FIRST_LOAD = False
    StoreMemVars
    picPayment.Top = 3120
    dtFrom = LOGDATE
    dtTo = LOGDATE
    ChangeORNum = False
    Screen.MousePointer = 0
    cmdSelect.Enabled = False
End Sub

Sub FillInvoiceType()
    cboInvoiceType.Clear
    cboInvoiceType.AddItem "CSH"
    cboInvoiceType.AddItem "CHG"
    cboInvoiceType.AddItem "DR"
    cboInvoiceType.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsOFF_HD = Nothing
    Set rsOFF_DT = Nothing
    Set frmCMISOREntry = Nothing
    LocalAcess = ""
End Sub

Public Sub frmNewEntity_EntitySELECTed(strCode As String, strAccountName As String, strEntityClass As String)
    vENTITY = strEntityClass + strCode
    lblVendorName.Caption = SetVendorName(vENTITY)
End Sub

Private Sub grdDetails_Click()
    grdDetails.Col = 10
    If grdDetails.Text <> "" Then
        ShowGridDetails grdDetails.Text
    End If
End Sub

Private Sub grdDetails_DblClick()
    If Null2Bool(rsOFF_HD!Paidna) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
        grdDetails.Col = 10
        If grdDetails.Text <> "" Then
            On_Update = True
            cmdDetails.Enabled = False
            cmdDetails.ZOrder 0
            cmdDetails.Visible = True
            picDetails.ZOrder 0
            picDetails.Visible = True
            fraDetails.Enabled = False
            Picture1.Enabled = False
            StoreGridDetails grdDetails.Text
            If SetPaidForCode(cboPaidFor.Text) = "427" Then
                chkCreditCardTrans.Value = 0
                chkCreditCardTrans.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub lvDeposits_DblClick()
'DESCRIPTION: Apply customer deposits
    Dim vInvoiceno                                                  As String
    Dim reply                                                       As String
    Dim ictr                                                        As Integer
    Dim InvoiceAmount                                               As Double
    Dim balance                                                     As Double
    
    vDeposits = 0
    InvoiceAmount = 0
    
    'JJE 12/27/2012 Check the remaining deposit amount for specific transaction type
    Dim rsUnapplied                                                 As ADODB.Recordset
    Set rsUnapplied = New ADODB.Recordset
    
    If Left(cboTranType, 1) = "S" Or Left(cboTranType, 1) = "V" Then
        rsUnapplied.Open "SELECT AMOUNT-APPLIEDAMT AS BALANCE FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & txtCUSCDE.Text & "' AND PAIDNA =1 AND PAIDFOR = '412" & Left(cboTranType, 1) & "'", gconDMIS, adOpenKeyset
    Else
'        rsUnapplied.Open "SELECT AMOUNT-APPLIEDAMT AS BALANCE FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & txtCUSCDE.Text & "' AND PAIDNA =1 ", gconDMIS, adOpenKeyset
        rsUnapplied.Open "SELECT AMOUNT-APPLIEDAMT AS BALANCE FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & txtCUSCDE.Text & "' AND PAIDNA =1 AND PAIDFOR = '412P'", gconDMIS, adOpenKeyset
    End If
    
    If Not rsUnapplied.EOF And Not rsUnapplied.BOF Then
        balance = rsUnapplied!balance
    End If
    Set rsUnapplied = Nothing
    'JJE
    If Not lvDeposits.SelectedItem Is Nothing Then
        reply = MsgBox("Are you sure you want to apply" + vbCrLf + "this customer deposit?", vbQuestion + vbYesNo, "Customer Deposit")
        If reply = vbYes Then
            'JJE 12/27/2012 Check CMIS records if deposit is already applied to invoice
            Dim rsDeposits                                          As ADODB.Recordset
            Set rsDeposits = New ADODB.Recordset
            rsDeposits.Open "SELECT ISNULL(AMOUNT,0) AS AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO='" & txtReference.Text & "' AND INVOICETYPE='" & SetTranTypeCode(cboTranType.Text) & "'", gconDMIS, adOpenForwardOnly
            If Not rsDeposits.EOF And Not rsDeposits.BOF Then
                InvoiceAmount = NumericVal(txtAmount.Text) - rsDeposits!amount
                If InvoiceAmount = 0 Then
                    MsgBox "Please select other invoice.", vbInformation, "Customer Deposit " & SetTranTypeCode(cboTranType.Text) + "-" + txtReference.Text
                    Exit Sub
                End If
            Else
                InvoiceAmount = NumericVal(txtAmount.Text)
            End If
            Set rsDeposits = Nothing
            'JJE
            If InvoiceAmount > 0 Then
                If NumericVal(lvDeposits.SelectedItem.SubItems(3)) > NumericVal(txtAmount.Text) Then
                    vDeposits = Val(InputBox("Customer Deposit is greater than invoice amount." + Chr(13) + "Please enter correct amount.", "Apply Customer Deposit", NumericVal(txtAmount.Text)))
                    If vDeposits > NumericVal(txtAmount.Text) Then
                        MsgBox "Applied Deposit cannot be greater than invoice amount.", vbInformation, "Customer Deposit"
                        Exit Sub
                    End If
                    'txtPayment.Text = NumericVal(txtPayment.Text) - vDeposits
                Else
                    vDeposits = Val(InputBox("Please enter correct amount.", "Apply Customer Deposit", NumericVal(lvDeposits.SelectedItem.SubItems(3))))
                    If vDeposits > NumericVal(txtAmount.Text) Then
                        MsgBox "Applied Deposit cannot be greater than invoice amount.", vbInformation, "Customer Deposit"
                        Exit Sub
                    End If
'                    'txtPayment.Text = ToDoubleNumber(NumericVal(txtPayment.Text) - vDeposits)
                End If
                
                If N2Str2Zero(vDeposits) > 0 Then
                    If N2Str2Zero(vDeposits) > lvDeposits.SelectedItem.SubItems(3) Then
                        MsgBox "Applied Deposit cannot be greater than the actual balance amount.", vbInformation, "Customer Deposit"
                        Exit Sub
                    End If
                End If
            End If
            
            ApplyDeposits = True
            picDeposits.Visible = False
            'JRE 07/20/16 To unload picORPayment form
            picORPayment.Visible = False
            
            If SetTranTypeCode(cboTranType.Text) <> "RO" Then
                vInvoiceno = N2Str2Null(txtReference.Text)
            Else
                If labRef.Caption = "Ref. '" Then
                    vInvoiceno = N2Str2Null(labReference.Caption)
                Else
                    vInvoiceno = N2Str2Null(txtReference.Text)
                End If
            End If

            If picDetails.Visible = False Then
                rsRefresh
                rsOFF_HD.Find "OR_NUM = '" & txtOR_NUM.Text & "'"
                StoreMemVars
            End If
            
            picDetails.Enabled = True
            cmdTranSave.Enabled = True
            OR_REFERENCE = lvDeposits.SelectedItem.SubItems(2)
            ictr = lvDeposits.SelectedItem.Index
            lvDeposits.ListItems.Remove (ictr)
            'lvDeposits.ListItems.ITEM(iCtr).ForeColor = vbRed
            If lvDeposits.ListItems.Count = 0 Then
                picDeposits.Visible = False
            End If
            
            'JJE 12/17/2012 10:10AM
            'Deposit update. If Deposit is applied to Official Receipt then Payment is always = 0
            If vDeposits > 0 Then
                Payment.Text = ToDoubleNumber(NumericVal(Payment.Text) - vDeposits)
                Payment.Locked = True
            End If
            'JJE
        Else
            Exit Sub
        End If
    Else
        MessagePop RecNotFound, "", "No Record Found"
    End If
End Sub

Private Sub lvDeposits_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblDepositID.Caption = lvDeposits.SelectedItem.SubItems(6)
End Sub

Private Sub lvPayments_DblClick()
'DESCRIPTION: Apply payments from Credit Card Company with Bank Charges and Expanded Withheld Tax
    Dim rsCardCompany                                               As ADODB.Recordset
    Dim vBankCharges                                                As Double
    Dim vEWT                                                        As Double
    
    Set rsCardCompany = New ADODB.Recordset
    If COMPANY_CODE = "DJM" Then
        rsCardCompany.Open "SELECT * FROM CMIS_CardCompany WHERE CUSCDE = '" & txtCUSCDE.Text & "'", gconDMIS, adOpenKeyset
    Else
        rsCardCompany.Open "SELECT * FROM CMIS_CardBank WHERE CUSCDE = '" & txtCUSCDE.Text & "'", gconDMIS, adOpenKeyset
    End If
    If Not rsCardCompany.EOF And Not rsCardCompany.BOF Then
        vBankCharges = NumericVal(rsCardCompany!BankCharges) / 100
        vEWT = NumericVal(rsCardCompany!EWT) / 100
    End If

    If Not lvPayments.SelectedItem Is Nothing Then
        txtAmount = Format(NumericVal(lvPayments.SelectedItem.SubItems(3)), "#,###,##0.00")
        txtDiscount.Text = Format(ToDoubleNumber(lvPayments.SelectedItem.SubItems(3)) * vBankCharges, "#,###,##0.00")
        txtTax.Text = Format(ToDoubleNumber(lvPayments.SelectedItem.SubItems(3)) * vEWT, "#,###,##0.00")
        Payment.Text = Format(NumericVal(lvPayments.SelectedItem.SubItems(3)) - (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vBankCharges + (NumericVal(lvPayments.SelectedItem.SubItems(3)) * vEWT)), "#,###,##0.00")
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
        picCreditCard.Visible = False
    Else
        MessagePop RecNotFound, "", "No Record Found"
    End If
End Sub

Private Sub lvPayments_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'UPDATE BY : ROWEL DE QUIROZ
'DATE : MARCH 3 2011
'DESCRPTION:
    Dim RDQ                                                         As Integer

    If lblTotal.Caption <> "0.00" Then lblTotal.Caption = "0.00"
    For RDQ = 1 To lvPayments.ListItems.Count
        If lvPayments.ListItems.Item(RDQ).Checked = True Then
            lblTotal.Caption = ToDoubleNumber(lblTotal.Caption + NumericVal(lvPayments.ListItems.Item(RDQ).SubItems(3)))
        End If
    Next
End Sub

Private Sub optCANCEL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        picPayment.ZOrder 1: picPayment.Visible = False
    End If
End Sub

Private Sub optCARD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CARD"
        picPayment.ZOrder 1: picPayment.Visible = False

        frmCMISCARDPaymentEntry.Show vbModal
    End If
End Sub

Private Sub optCASH_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CASH"
        picPayment.ZOrder 1: picPayment.Visible = False
        
        frmCMISCASHPaymentEntry.Show vbModal
    End If
End Sub

Private Sub optCHECK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OR_NUMBER_GLOBAL = txtOR_NUM.Text
        RECEIPTS_AMOUNT = wizDigit1.TextValue
        MODE_OF_PAYMENT = "CHECK"
        picPayment.ZOrder 1: picPayment.Visible = False
        frmCMISCHECKPaymentEntry.Show vbModal
    End If
End Sub

Private Sub optGoods_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        picORType.Visible = False
        txtOR_NUM.Text = GetLASTOR("G")
        
        If COLLECTION_RECEIPTS_CR = True Then
            AckReceipts.Caption = "Collection Receipts:"
        Else
            AckReceipts.Caption = "ACKNOWLEDGMENT RECEIPTS"
        End If
        
        AckReceipts.Visible = True
        lblReceipt.Visible = False
        cmdSelect.Enabled = True
        'If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then
            'txtOR_NUM.SetFocus N/A
        Else
            txtOR_NUM.SetFocus
        End If
        'JJE
    End If
End Sub

Private Sub Option1_Click()
    Picture3.Visible = True
    Picture5.Visible = False
    picCustomer.Visible = False
    lvPayments.ListItems.Clear
    txtReference2 = ""
    'lvPayments.Checkboxes = False
    txtReference2.SetFocus
    Payment = "0.00"
    CreditCardPayments
End Sub

Private Sub Option2_Click()
    Picture3.Visible = False
    Picture5.Visible = True
    picCustomer.Visible = False
    lvPayments.ListItems.Clear
    'lvPayments.Checkboxes = True
    lblTotal = "0.00"
    Payment = "0.00"
    CreditCardPayments
End Sub

Private Sub Option3_Click()
    Picture3.Visible = False
    Picture5.Visible = False
    picCustomer.Visible = True
    lvPayments.ListItems.Clear
    txtCustomer.Text = ""
    'lvPayments.Checkboxes = True
    lblTotal = "0.00"
    Payment = "0.00"
    txtCustomer.SetFocus
    CreditCardPayments
End Sub

Private Sub OptPR_Keydown(KeyCode As Integer, Shift As Integer)
    'JJE 05/20/2016
    If KeyCode = vbKeyReturn Then
        picORType.Visible = False
        txtOR_NUM.Text = GetLASTOR("V")
        
        AckReceipts.Caption = "Provisional Receipts:"
        
        AckReceipts.Visible = True
        lblReceipt.Visible = False
        cmdSelect.Enabled = True
        txtOR_NUM.SetFocus
    End If
    'JJE
End Sub

Private Sub optService_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        picORType.Visible = False
        txtOR_NUM.Enabled = True
        
        'JJE 01/27/2012 3:08PM
        txtOR_NUM.Text = GetLASTOR("S")
        
        AckReceipts.Visible = False
        lblReceipt.Visible = True
        cmdSelect.Enabled = True
        'If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then ** FOR APPROVAL **
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DJM" Then
            'txtOR_NUM.SetFocus N/A
        Else
            txtOR_NUM.SetFocus
        End If
        'JJE
    End If
End Sub

Private Sub Payment_KeyPress(KeyCode As Integer)
    'JJE
    KeyCode = OnlyNumeric(KeyCode)
    If KeyCode = 13 Then
        cmdTranSave.SetFocus
    End If
    'JJE
End Sub

Private Sub textSearch_Change()
    If optORNo.Value = True Then
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

Private Sub Timer1_Timer()
    If On_Update = False Then
        If rsOFF_HD!Cancel = True Then
            labStatus.Caption = "*** Cancelled OR ***"
            If labStatus.Visible = True Then
                labStatus.Visible = False
            Else
                labStatus.Visible = True
            End If
        Else
            If rsOFF_HD!Paidna = True Then
                labStatus.Caption = "*** PAID OR ***"
                If labStatus.Visible = True Then
                    labStatus.Visible = False
                Else
                    labStatus.Visible = True
                End If
            Else
                labStatus.Visible = False
            End If
        End If
    Else
        labStatus.Visible = False
    End If
End Sub

Private Sub txtBalance_Change()
    Payment.Text = ToDoubleNumber(NumericVal(txtBalance.Text) - (NumericVal(txtDiscount.Text) + NumericVal(txtTax.Text)))
End Sub

Private Sub txtCustomer_Change()
    Dim xList                                                       As ListItem
    Dim rsCMIS_OFF_HD                                               As ADODB.Recordset
    
    Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3' AND (Paidby is null or paidby = 'N') and cardbnkcde = '" & txtCUSCDE & "' AND CusName LIKE '" & txtCustomer.Text & "%' ORDER BY OR_Date")
    If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
        lvPayments.ListItems.Clear
        lblTotal = "0.00"
        Do While Not rsCMIS_OFF_HD.EOF
            Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
            xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
            xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!cusname)
            xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
            xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
            xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
            tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
            lblTotal = Format(tmpTotal, "#,###,##0.00")
            rsCMIS_OFF_HD.MoveNext
        Loop
    Else
        'MsgBox "No customer record to view", vbCritical, "No Record"
        '        MessagePop RecNotFound, "No record to view", "No Record"
    End If
End Sub

Private Sub txtDescript_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtDiscount_Change()
    'JJE
    If txtBalance.Text > 0 Then
        Payment = Round(NumericVal(txtBalance) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
    Else
        Payment = Round(NumericVal(txtAmount) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
    End If
    'JJE
End Sub

Private Sub txtDiscount_GotFocus()
    txtDiscount.Text = ToDoubleNumber(txtDiscount.Text)
End Sub

Private Sub txtDiscount_LostFocus()
    txtDiscount.Text = ToDoubleNumber(txtDiscount.Text)
End Sub

Private Sub txtOR_DATE_GotFocus()
    If IsDate(txtOR_DATE.Text) = True Then txtOR_DATE.Text = Format(txtOR_DATE.Text, "MM/DD/YYYY") Else txtOR_DATE.Text = ""
End Sub

Private Sub txtOR_DATE_LostFocus()
    If IsDate(txtOR_DATE.Text) = True Then
        txtOR_DATE.Text = Format(txtOR_DATE.Text, "DD-MMM-YYYY")
        If CheckORCutOff(txtOR_NUM) = True Then
            If Format(CDate(txtOR_DATE.Text), "mm/dd/yyyy") < Format(LOGDATE, "mm/dd/yyyy") Then
                MsgBox ("OR back date is not allowed!"), vbCritical, "System Message"
                txtOR_DATE.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgBox "Invalid date! Please check...", vbExclamation, "System Message"
        txtOR_DATE.Text = LOGDATE
        txtOR_DATE.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtOR_NUM_Change()
    'JRE 01042017 - TO LIMIT THE NUMBER OF CHARACTER INPUT IN THE OR_NUM TEXTBOX
    If COMPANY_CODE = "CMC" Then
        txtOR_NUM.MaxLength = 6
    Else
    End If
End Sub

Private Sub txtOR_NUM_KeyPress(KeyAscii As Integer)
'    KeyAscii = OnlyNumeric(KeyAscii)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '    If COMPANY_CODE = "DGI" Then
    '        KeyAscii = 0
    '    End If
End Sub

Private Sub txtOR_NUM_LostFocus()

     'JJE
    If OR_VAT_NONVAT = "VAT" Then
        If COLLECTION_RECEIPTS_CR = True Then
            If COMPANY_CODE = "MGS" Then
                If (AckReceipts.Visible = True And AckReceipts.Caption = "Collection Receipts:") Then
                    If Len(txtOR_NUM) >= 3 And IsNumeric(Left(txtOR_NUM.Text, 2)) = False Then
                        txtOR_NUM.Text = Mid(txtOR_NUM.Text, 3, Len(txtOR_NUM.Text) - 2)
                    ElseIf IsNumeric(txtOR_NUM.Text) = True Then
                        txtOR_NUM.Text = Format(Right(txtOR_NUM.Text, 6), "000000")
                    End If
                    txtOR_NUM.Text = "CR" + Format(Right(NumericVal(Mid(txtOR_NUM.Text, 1, Len(txtOR_NUM.Text))), 6), "000000")
                ElseIf (AckReceipts.Visible = True And AckReceipts.Caption = "Provisional Receipts:") Then
                    If Len(txtOR_NUM) >= 3 And IsNumeric(Left(txtOR_NUM.Text, 2)) = False Then
                        txtOR_NUM.Text = Mid(txtOR_NUM.Text, 3, Len(txtOR_NUM.Text) - 2)
                    ElseIf IsNumeric(txtOR_NUM.Text) = True Then
                        txtOR_NUM.Text = Format(Right(txtOR_NUM.Text, 6), "000000")
                    End If
                    txtOR_NUM.Text = "PR" + Format(Right(NumericVal(Mid(txtOR_NUM.Text, 1, Len(txtOR_NUM.Text))), 6), "000000")
                End If
            Else
                If AckReceipts.Visible = True Then
                    If Len(txtOR_NUM) >= 3 And IsNumeric(Left(txtOR_NUM.Text, 2)) = False Then
                        txtOR_NUM.Text = Mid(txtOR_NUM.Text, 3, Len(txtOR_NUM.Text) - 2)
                    ElseIf IsNumeric(txtOR_NUM.Text) = True Then
                        txtOR_NUM.Text = Format(Right(txtOR_NUM.Text, 6), "000000")
                    End If
                    txtOR_NUM.Text = "CR" + Format(Right(NumericVal(Mid(txtOR_NUM.Text, 1, Len(txtOR_NUM.Text))), 6), "000000")
                End If
            End If
        ElseIf OFFICIAL_RECEIPT_OR = True Then
            If lblReceipt.Visible = True Then
                If Len(txtOR_NUM) >= 3 And IsNumeric(Left(txtOR_NUM.Text, 2)) = False Then
                    txtOR_NUM.Text = Mid(txtOR_NUM.Text, 3, Len(txtOR_NUM.Text) - 2)
                ElseIf IsNumeric(txtOR_NUM.Text) = True Then
                    txtOR_NUM.Text = Format(Right(txtOR_NUM.Text, 6), "000000")
                End If
                txtOR_NUM.Text = "OR" + Format(Right(NumericVal(Mid(txtOR_NUM.Text, 1, Len(txtOR_NUM.Text))), 6), "000000")
            End If
        ElseIf COMPANY_CODE = "DSSC" Or COMPANY_CODE = "DGI" Then
            txtOR_NUM.Text = Format(Right(txtOR_NUM.Text, 6), "000000")
        End If
    End If
    'JJE
End Sub

Private Sub Payment_Change()
    'JJE
    wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
    'JJE
End Sub

Private Sub Payment_GotFocus()
    'JJE
    If Payment.Text <> "" Then
        Payment.Text = NumericVal(Payment.Text)
    End If
    'JJE
End Sub

Private Sub Payment_LostFocus()
    'JJE
    wizDigit1.TextValue = ToDoubleNumber(Payment.Text)
    'JJE
End Sub

Private Sub txtReference_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtReference.Text <> "" Then
        If COMPANY_CODE = "MGS" Then 'JRE 05/30/2016 To accomodate service invoice from AR with 5 numbers only
        
        Else
        txtReference.Text = Format(txtReference.Text, "000000")
        End If 'JRE
        cboPaidFor.ListIndex = -1
        If labRef.Caption = "Inv. #" Then txtReference.Text = Format(txtReference.Text, "000000")
            Dim rsOrd_Hd                                    As ADODB.Recordset
            Dim rsOFF_DT                                    As ADODB.Recordset
            
            If SetTranTypeCode(cboTranType.Text) = "SI" Then
                Dim rsREPOR                                 As ADODB.Recordset
                Set rsREPOR = New ADODB.Recordset
                If labRef.Caption = "Ref. '" Then
                    'UPDATE: to accommodate payment which invoices are from opening balance of accounting -ACL
                    'JRE 08162016 - REMOVE THE ACCOMMODATION OF PAYMENT FROM OPENING BALANCE OF ACCOUNTING ACCORDING TO SJR
'                    Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt FROM CSMS_REPOR WHERE Rep_or = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCUSCDE.Text) & " " & _
'                                                   "UNION SELECT CUSTOMERCODE AS ACCT_NO,NULL AS REP_OR,CUSTOMERNAME AS NIYM,INVOICEAMT AS AMOUNT,INVOICEDATE AS DTE_COMP,INVOICENO AS INVOICE,NULL AS L_AMTVALUE,NULL AS P_AMTVALUE,NULL AS A_AMTVALUE,NULL AS M_AMTVALUE,INVOICEAMT AS RO_AMOUNT,NULL AS INSAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='SI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt FROM CSMS_REPOR WHERE Rep_or = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCUSCDE.Text) & " ")
                Else
'                    Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt FROM CSMS_REPOR WHERE invoice = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCUSCDE.Text) & " " & _
'                                                   "UNION SELECT CUSTOMERCODE AS ACCT_NO,NULL AS REP_OR,CUSTOMERNAME AS NIYM,INVOICEAMT AS AMOUNT,INVOICEDATE AS DTE_COMP,INVOICENO AS INVOICE,NULL AS L_AMTVALUE,NULL AS P_AMTVALUE,NULL AS A_AMTVALUE,NULL AS M_AMTVALUE,INVOICEAMT AS RO_AMOUNT,NULL AS INSAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='SI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,L_AmtValue,P_AmtValue,A_AmtValue,M_AmtValue,RO_Amount,Insamt FROM CSMS_REPOR WHERE invoice = " & N2Str2Null(txtReference.Text) & " AND ACCT_NO =" & N2Str2Null(txtCUSCDE.Text) & " ")
                End If
                If Not rsREPOR.EOF And Not rsREPOR.BOF Then
                    If labRef.Caption = "Ref. '" Then
                        labReference.Caption = Null2String(rsREPOR!INVOICE)
                    Else
                        labReference.Caption = Null2String(rsREPOR!REP_OR)
                    End If
                    txtDescript.Text = Null2String(rsREPOR!niym)
                    txtAmount.Text = ToDoubleNumber(NumericVal(rsREPOR!RO_Amount) - NumericVal(rsREPOR!INSAMT))
                    labDocDate.Caption = Null2Date(rsREPOR!dte_comp)
                    labCUSCODE.Caption = Null2String(rsREPOR!ACCT_NO)
                    xTTLINV = ToDoubleNumber(txtAmount.Text)
                    Set rsOFF_DT = New ADODB.Recordset
                                    Set rsOFF_DT = gconDMIS.Execute("SELECT isnull(SUM(PAYMENT),0) as MGA_BAYAD from CMIS_Off_Dt WHERE (trantype = 'RO' OR trantype = 'SI') AND Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCUSCDE.Text) & " AND LEFT(OR_NUM,3) <> 'SOA'")
                                    If (rsOFF_DT!MGA_BAYAD) > 0 Then
                                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!MGA_BAYAD))
                                        Call BalanceCash(cboInvoiceType, txtReference)
                                    Else
                                        txtBalance.Text = ToDoubleNumber(txtAmount.Text)
                                    End If
                    'JJE 01/03/2013
                    
                    Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE (trantype = 'RO' OR trantype = 'SI') AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                    If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                        Set rsCustomerDeposit = New ADODB.Recordset
                        rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                        If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        End If
                    Else
                        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                    End If
                    'JJE
                    Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                Else
                    'UPDATE: to accommodate payment which invoices are from opening balance of accounting -ACL
                    If labRef.Caption = "Ref. '" Then
                        'JRE 08162016 - REMOVE THE ACCOMMODATION OF PAYMENT FROM OPENING BALANCE OF ACCOUNTING ACCORDING TO SJR
'                        Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR WHERE Rep_or = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCUSCDE.Text) & " " & _
'                                                       "UNION SELECT CUSTOMERCODE AS ACCT_NO,NULL AS REP_OR,CUSTOMERNAME AS NIYM,INVOICEAMT AS AMOUNT,INVOICEDATE AS DTE_COMP,INVOICENO AS INVOICE,NULL AS L_AMTVALUE,NULL AS P_AMTVALUE,NULL AS A_AMTVALUE,NULL AS M_AMTVALUE,NULL AS RO_AMOUNT,NULL AS INSAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='SI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                        Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR WHERE Rep_or = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCUSCDE.Text) & " ")
                    Else
'                        Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR WHERE invoice = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCUSCDE.Text) & " " & _
'                               "UNION SELECT CUSTOMERCODE AS ACCT_NO,NULL AS REP_OR,CUSTOMERNAME AS NIYM,INVOICEAMT AS AMOUNT,INVOICEDATE AS DTE_COMP,INVOICENO AS INVOICE,NULL AS L_AMTVALUE,NULL AS P_AMTVALUE,NULL AS A_AMTVALUE,NULL AS M_AMTVALUE,NULL AS RO_AMOUNT,NULL AS INSAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='SI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                        Set rsREPOR = gconDMIS.Execute("SELECT Acct_No,rep_or,niym,amount,dte_comp,invoice,PartLabor,PartParts,PartMaterials,PartAccessories,Participat,InsCde from CSMS_REPOR WHERE invoice = " & N2Str2Null(txtReference.Text) & " AND Participat =" & N2Str2Null(txtCUSCDE.Text) & " ")
                    End If
                    If Not rsREPOR.EOF And Not rsREPOR.BOF Then
                        If labRef.Caption = "Ref. '" Then
                            labReference.Caption = Null2String(rsREPOR!INVOICE)
                        Else
                            labReference.Caption = Null2String(rsREPOR!REP_OR)
                        End If
                        txtDescript.Text = Null2String(rsREPOR!InsCde)
                        txtAmount.Text = ToDoubleNumber(N2Str2Zero(rsREPOR!PARTLABOR) + N2Str2Zero(rsREPOR!PARTPARTS) + N2Str2Zero(rsREPOR!PARTMATERIALS) + N2Str2Zero(rsREPOR!PARTACCESSORIES))
                        labDocDate.Caption = Null2Date(rsREPOR!dte_comp)
                        labCUSCODE.Caption = Null2String(rsREPOR!Participat)
'                        Set rsOFF_DT = New ADODB.Recordset
'                                            Set rsOFF_DT = gconDMIS.Execute("SELECT SUM(PAYMENT) as MGA_BAYAD from CMIS_Off_Dt WHERE (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' and Reference = " & N2Str2Null(txtReference.Text) & " and CusCde = " & N2Str2Null(txtCUSCDE.Text) & "")
'                                            If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
'                                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - N2Str2Zero(rsOFF_DT!Mga_Bayad))
'                                                Call BalanceCash(cboInvoiceType, txtReference)
'                                            Else
'                                                txtBalance.Text = ToDoubleNumber(txtAmount.Text)
'                                            End If
'                        'JJE 1 / 3 / 2013
                        
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE (trantype = 'RO' OR trantype = 'SI') AND INVOICETYPE='CSH' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " AND isnull(Cancel,0) <> 1 and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        'MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        On Error Resume Next
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        
            If SetTranTypeCode(cboTranType.Text) = "PI" Then
                Set rsOrd_Hd = New ADODB.Recordset
                '************************************************
                'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
                If cboInvoiceType = "CSH" Then
                    'Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT from PMIS_vw_ISS_HISTORY WHERE TYPE = 'P' AND trantype = 'CSH' and tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    'UPDATE: to accommodate payment which invoices are from opening balance of accounting - ACL
                    'JRE 08162016 - REMOVE THE ACCOMMODATION OF PAYMENT FROM OPENING BALANCE OF ACCOUNTING ACCORDING TO SJR
'                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'P' AND trantype = 'CSH' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text) & " " & _
'                                                    "UNION SELECT CUSTOMERCODE AS CUSTCODE,INVOICEDATE AS TRANDATE,INVOICENO,CUSTOMERNAME AS CUSTNAME,INVOICEAMT AS NETINVAMT from AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='PI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'P' AND trantype = 'CSH' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        
                        Set rsOFF_DT = New ADODB.Recordset
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' AND TranType = 'PI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) as DEPOSIT_AMOUNT from CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " and INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " and APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        cmdDetails_Click
                        Exit Sub
                    End If
                ElseIf cboInvoiceType = "CHG" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    'UPDATE: to accommodate payment which invoices are from opening balance of accounting
                    'JRE 08162016 - REMOVE THE ACCOMMODATION OF PAYMENT FROM OPENING BALANCE OF ACCOUNTING ACCORDING TO SJR
'                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY WHERE TYPE = 'P' AND trantype = 'CHG' AND tranno = " & N2Str2Null(txtReference.Text) & " and custcode = " & N2Str2Null(txtCUSCDE.Text) & " " & _
'                                                    "UNION SELECT CUSTOMERCODE AS CUSTCODE,INVOICEDATE AS TRANDATE,INVOICENO,CUSTOMERNAME AS CUSTNAME,INVOICEAMT AS NETINVAMT from AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICETYPE='PI' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt from PMIS_vw_ISS_HISTORY WHERE TYPE = 'P' AND trantype = 'CHG' AND tranno = " & N2Str2Null(txtReference.Text) & " and custcode = " & N2Str2Null(txtCUSCDE.Text) & " ")
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' AND TranType = 'PI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) as DEPOSIT_AMOUNT from CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " and INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " and APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                Else
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt FROM PMIS_vw_ISS_HISTORY WHERE [TYPE] = 'P' AND trantype = 'DR' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' AND TranType = 'PI' and Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) as DEPOSIT_AMOUNT from CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " and INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " and APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
                '************************************************
            End If
            
            If SetTranTypeCode(cboTranType.Text) = "AI" Then
                Set rsOrd_Hd = New ADODB.Recordset
                '************************************************
                'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
                If cboInvoiceType = "CSH" Then
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'A' AND trantype = 'CSH' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' AND TranType = 'AI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT from CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                ElseIf cboInvoiceType = "CHG" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'A' AND trantype = 'CHG' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " and INVOICETYPE = 'CHG' AND TranType = 'AI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                Else
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'A' AND trantype = 'DR' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' AND TranType = 'AI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        
            If SetTranTypeCode(cboTranType.Text) = "MI" Then
                Set rsOrd_Hd = New ADODB.Recordset
                '************************************************
                'updating code:     jaa - 11202008          - check surely if the transaction is CSH or CHG transaction
                If cboInvoiceType = "CSH" Then
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'M' AND trantype = 'CSH' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        xTTLINV = ToDoubleNumber(rsOrd_Hd!ttlinvAMT)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'CSH' AND TranType = 'MI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                ElseIf cboInvoiceType = "CHG" Then
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'M' AND trantype = 'CHG' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        xTTLINV = ToDoubleNumber(rsOrd_Hd!ttlinvAMT)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'CHG' AND TranType = 'MI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                        Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                Else
                    Set rsOrd_Hd = New ADODB.Recordset
                    Set rsOrd_Hd = gconDMIS.Execute("SELECT custcode,trandate,tranno,custname,netinvamt,ttlinvAMT FROM PMIS_vw_ISS_HISTORY WHERE TYPE = 'M' AND trantype = 'DR' AND tranno = " & N2Str2Null(txtReference.Text) & " AND custcode = " & N2Str2Null(txtCUSCDE.Text))
                    If Not rsOrd_Hd.EOF And Not rsOrd_Hd.BOF Then
                        txtDescript.Text = Null2String(rsOrd_Hd!custname)
                        txtAmount.Text = ToDoubleNumber(rsOrd_Hd!NETINVAMT)
                        labDocDate.Caption = Null2Date(rsOrd_Hd!trandate)
                        labCUSCODE.Caption = Null2String(rsOrd_Hd!custcode)
                        xTTLINV = ToDoubleNumber(rsOrd_Hd!ttlinvAMT)
                        Set rsOFF_DT = New ADODB.Recordset
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ROUND(SUM(PAYMENT + TAX),2) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND INVOICETYPE = 'DR' AND TranType = 'MI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND CusCde = " & N2Str2Null(txtCUSCDE.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            Set rsCustomerDeposit = New ADODB.Recordset
                            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                                Call BalanceCash(cboInvoiceType, txtReference)
                            End If
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                    Else
                        MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
            End If

            If SetTranTypeCode(cboTranType.Text) = "VI" Then
                Dim rsPurchAgree                                    As ADODB.Recordset
                Dim rsCheckFinance                                  As ADODB.Recordset
                Set rsPurchAgree = New ADODB.Recordset

                'UPDATED BY:LJDM
                'MAY 13 FRIDAY 2011.
                'CUSTOMER DOWNPAYMENT INSTEAD TOTAL AMOUNT WITH FINANCED CORPORATION
                Set rsPurchAgree = gconDMIS.Execute("SELECT * FROM ALL_Customer WHERE cuscde=" & N2Str2Null(txtCUSCDE.Text))
                vCustype = rsPurchAgree!CUSTYPE
                If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
                    If vCustype = "C" Then
                        Set rsCheckFinance = gconDMIS.Execute("SELECT code from SMIS_FINCOM WHERE code = " & N2Str2Null(rsPurchAgree!CUSCDE))
                        If Not rsCheckFinance.EOF And Not rsCheckFinance.BOF Then
                            CompanyFinancing (rsPurchAgree!CUSCDE)
                        Else
                            PersonalPayment
                        End If
                    Else
                        PersonalPayment
                    End If
                    If PersonalPayment = False And CompanyFinancing(rsPurchAgree!CUSCDE) = False Then
                        MsgBox "Transaction Reference not found in that Transaction Type", vbOKOnly, "Invalid Reference Number"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
                rsPurchAgree.Close
            End If

'            If COMPANY_CODE = "HMH" Then
                If SetTranTypeCode(cboTranType.Text) = "UI" Then
                    Dim rsJOURNALHD                         As ADODB.Recordset
                    Set rsJOURNALHD = New ADODB.Recordset
                    rsJOURNALHD.Open "SELECT CUSTOMERCODE,CUSTOMERNAME,JDATE,INVOICEAMT FROM AMIS_JOURNAL_HD WHERE JTYPE='COB' AND INVOICENO=" & N2Str2Null(txtReference.Text) & " AND CUSTOMERCODE=" & N2Str2Null(txtCUSCDE.Text) & "", gconDMIS, adOpenKeyset
                    If Not rsJOURNALHD.EOF And Not rsJOURNALHD.BOF Then
                        txtDescript.Text = Null2String(rsJOURNALHD!CUSTOMERNAME)
                        txtAmount.Text = ToDoubleNumber(rsJOURNALHD!InvoiceAmt)
                        labDocDate.Caption = Null2Date(rsJOURNALHD!JDATE)
                        labCUSCODE.Caption = Null2String(rsJOURNALHD!CustomerCode)
                        xTTLINV = ToDoubleNumber(rsJOURNALHD!InvoiceAmt)
                        'JJE 01/03/2013
                        Set rsOFF_DT = gconDMIS.Execute("SELECT ISNULL(ROUND(SUM(PAYMENT + TAX),2),0) as MGA_BAYAD from CMIS_Off_Dt WHERE VAT = " & VAT_OR & " AND TranType = 'UI' AND Reference = " & N2Str2Null(txtReference.Text) & " and left(or_num,3) <> 'SOA'")
                        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD)))
                            Call BalanceCash(cboInvoiceType, txtReference)
                        Else
                            txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
                        End If
                        'JJE
                    Else
                        MsgBox "Kindly Check the Invoice number, Invoice Type and Customer code", vbOKOnly, "No Record found"
                        txtReference.SetFocus
                        Exit Sub
                    End If
                End If
'            End If
            
            If Left(txtOR_NUM, 3) = "SOA" Then
                    cmdTranSave.Enabled = True
                    If SetTranTypeCode(cboTranType.Text) = "OTH" Then
                        cboPaidFor.Enabled = True
                        cboBranch.Enabled = True
                    Else
                        cboPaidFor.Enabled = False
                        cboBranch.Enabled = False
                    End If
                    txtDescript.Enabled = True
                    txtDiscount.Enabled = True
                    txtTax.Enabled = True
                    Payment.Enabled = True
            Else
                If SetTranTypeCode(cboTranType.Text) = "SI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "PI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "AI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "MI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "VI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "UI" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "OTH" Then
                    If lvDeposits.ListItems.Count <> 0 Then
                        cmdTranSave.Enabled = False
                    Else
                        cmdTranSave.Enabled = True
                    End If
                    cboPaidFor.Enabled = True
                    cboBranch.Enabled = True
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                ElseIf SetTranTypeCode(cboTranType.Text) = "EST" Then
                    cmdTranSave.Enabled = True
                    cboPaidFor.Enabled = False
                    cboBranch.Enabled = False
                    txtDescript.Enabled = False
                    txtDiscount.Enabled = False
                    txtTax.Enabled = True
                    Payment.Enabled = True
                End If
            End If
        End If
End Sub

Function PersonalPayment() As Boolean

    If COMPANY_CODE <> "DJM" Then
        Dim rsOFF_Det                                                   As New ADODB.Recordset
        'return the amount and balance
        If InitializePayment = False Then
            PersonalPayment = False
            Exit Function
        Else
            PersonalPayment = True
        End If
        'CHECK THE BALANCE
        'CHECKING IS PER PAYMENT
        Set rsOFF_Det = gconDMIS.Execute("SELECT TRANTYPE,REFERENCE,paymenttype,SUM(payment) AS MGA_BAYAD FROM CMIS_Off_Dt WHERE trantype = 'VI' AND Reference = " & N2Str2Null(txtReference.Text) & " and left(or_num,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
        Do While Not rsOFF_Det.EOF And Not rsOFF_Det.BOF
            'INITIALIZE DOWNPAYMENT BALANCE
    '        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT),2) as MGA_BAYAD,PAYMENTTYPE = 'D',TRANTYPE,REFERENCE from CMIS_Off_Dt WHERE trantype = 'VI' and PAYMENTTYPE = 'D'  and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
            
    '        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
             If (rsOFF_Det!paymenttype = "N" Or rsOFF_Det!paymenttype = "D") And rsOFF_Det!MGA_BAYAD <> 0 Then
                Set rsCustomerDeposit = New ADODB.Recordset
                'JRE 07/20/16 Replaced the CMIS_DEPOSITS table by CMIS_DEPOSITDT to get the amount of deposit applied
'                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & "", gconDMIS, adOpenKeyset
                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                     txtDownBal.Text = ToDoubleNumber(NumericVal(txtDownAmount.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                        'Call BalanceCash(cboInvoiceType, txtReference)
                End If
             ElseIf (rsOFF_Det!paymenttype = "L" And rsOFF_Det!MGA_BAYAD <> 0) Then
                Set rsCustomerDeposit = New ADODB.Recordset
'                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & "", gconDMIS, adOpenKeyset
                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                  txtLTOBal.Text = ToDoubleNumber(NumericVal(txtLTOAmount.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD)))
                           'Call BalanceCash(cboInvoiceType, txtReference)
                End If
             ElseIf (rsOFF_Det!paymenttype = "I" And rsOFF_Det!MGA_BAYAD <> 0) Then
                Set rsCustomerDeposit = New ADODB.Recordset
'                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & "", gconDMIS, adOpenKeyset
                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                  txtInsBal.Text = ToDoubleNumber(NumericVal(txtInsAmount.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD)))
                           'Call BalanceCash(cboInvoiceType, txtReference)
                End If
             ElseIf (rsOFF_Det!paymenttype = "O" And rsOFF_Det!MGA_BAYAD <> 0) Then
                Set rsCustomerDeposit = New ADODB.Recordset
'                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & "", gconDMIS, adOpenKeyset
                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                 txtOtherBal.Text = ToDoubleNumber(NumericVal(txtTPLAmout.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD)))
                           'Call BalanceCash(cboInvoiceType, txtReference)
                End If
             ElseIf (rsOFF_Det!paymenttype = "C" And rsOFF_Det!MGA_BAYAD <> 0) Then
                Set rsCustomerDeposit = New ADODB.Recordset
'                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
                rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITDT WHERE INVOICENO=" & N2Str2Null(rsOFF_Det!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_Det!TranType) & "", gconDMIS, adOpenKeyset
                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                'JRE 07/20/2016 Changed from txtChattelAmount to txtChattelBal
'                 txtChattelAmount.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD)))
                 txtChattelBal.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text) - (NumericVal(rsOFF_Det!MGA_BAYAD)))
                           'Call BalanceCash(cboInvoiceType, txtReference)
                End If
             Else
                txtOtherBal.Text = ToDoubleNumber(NumericVal(txtTPLAmout.Text))
                txtChattelBal.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text))
                txtInsBal.Text = ToDoubleNumber(NumericVal(txtInsAmount.Text))
                txtLTOBal.Text = ToDoubleNumber(NumericVal(txtLTOAmount.Text))
                txtDownBal.Text = ToDoubleNumber(NumericVal(txtDownAmount.Text))
             End If
            'INITIALIZE LTO BALANCE
    '        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT),2) as MGA_BAYAD,PAYMENTTYPE= 'L',TRANTYPE,REFERENCE from CMIS_Off_Dt WHERE trantype = 'VI' and PAYMENTTYPE = 'L' and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
    '        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
    '            Set rsCustomerDeposit = New ADODB.Recordset
    '            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
    '            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
    '              txtLTOBal.Text = ToDoubleNumber(NumericVal(txtLTOAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD)))
    '                       'Call BalanceCash(cboInvoiceType, txtReference)
    '            End If
    '        Else
    '            txtLTOBal.Text = ToDoubleNumber(NumericVal(txtLTOAmount.Text))
    '        End If
    '        'INITIALIZE INSURANCE BALANCE
    '        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT),2) as MGA_BAYAD,PAYMENTTYPE = 'I',TRANTYPE,REFERENCE from CMIS_Off_Dt WHERE trantype = 'VI' and PAYMENTTYPE = 'I'  and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
    '        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
    '            Set rsCustomerDeposit = New ADODB.Recordset
    '            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
    '            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
    '                txtInsBal.Text = ToDoubleNumber(NumericVal(txtInsAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD)))
    '                          'Call BalanceCash(cboInvoiceType, txtReference)
    '            End If
    '         Else
    '             txtInsBal.Text = ToDoubleNumber(NumericVal(txtInsAmount.Text))
    '         End If
    '        'INITIALIZE OTHER BALANCE
    '        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT),2) as MGA_BAYAD,PAYMENTTYPE= 'O',TRANTYPE,REFERENCE from CMIS_Off_Dt WHERE trantype = 'VI' and PAYMENTTYPE = 'O'  and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
    '        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
    '            Set rsCustomerDeposit = New ADODB.Recordset
    '            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
    '                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
    '                    txtOtherBal.Text = ToDoubleNumber(NumericVal(txtTPLAmout.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD)))
    '                            'Call BalanceCash(cboInvoiceType, txtReference)
    '                End If
    '         Else
    '            txtOtherBal.Text = ToDoubleNumber(NumericVal(txtTPLAmout.Text))
    '         End If
    '        'INITIALIZE CHATTLE BALANCE
    '        Set rsOFF_DT = gconDMIS.Execute("SELECT round(SUM(PAYMENT),2) as MGA_BAYAD,PAYMENTTYPE= 'C',TRANTYPE,REFERENCE from CMIS_Off_Dt WHERE trantype = 'VI' and PAYMENTTYPE = 'C'  and Reference = " & N2Str2Null(txtReference.Text) & " GROUP BY REFERENCE,TRANTYPE,PAYMENTTYPE")
    '        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
    '            Set rsCustomerDeposit = New ADODB.Recordset
    '            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
    '                If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
    '                    txtChattelAmount.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD)))
    '                            'Call BalanceCash(cboInvoiceType, txtReference)
    '                End If
    '         Else
    '             txtChattelBal.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text))
    '         End If
         rsOFF_Det.MoveNext
         Loop
        'END OF CHECKING
        lblInvoiceNo.Caption = (txtReference.Text)
        'INITIALIZATION OF PAYMENT
        
        'JRE 07/20/2016 For the amount entered not to be included in the TOTAL AMOUNT if it is FREED in SMIS Module
'        txtDownPay.Text = ToDoubleNumber(NumericVal(txtDownBal.Text))
'        txtLTOPay.Text = ToDoubleNumber(NumericVal(txtLTOBal.Text))
'        txtInsPay.Text = ToDoubleNumber(NumericVal(txtInsBal.Text))
'        txtTPLPay.Text = ToDoubleNumber(NumericVal(txtOtherBal.Text))
'        txtChattelPay.Text = ToDoubleNumber(NumericVal(txtChattelBal.Text))
        txtDownPay.Text = ToDoubleNumber(NumericVal(txtDownBal.Text))
        
        Dim rsFreePayment                       As New ADODB.Recordset
        Set rsFreePayment = gconDMIS.Execute("SELECT CODE,FREELTO,FREEINSURANCE,FREETPL,FREECHATTEL FROM SMIS_SALESORDER WHERE CODE = '" & txtCUSCDE.Text & "' AND VI_NO = " & N2Str2Null(txtReference.Text) & "")
        If (rsFreePayment!FreeLTO) = True Then
            txtLTOPay.Text = "FREE"
        Else
            txtLTOPay.Text = ToDoubleNumber(NumericVal(txtLTOBal.Text))
        End If
        
        If (rsFreePayment!FREEINSURANCE) = True Then
            txtInsPay.Text = "FREE"
        Else
            txtInsPay.Text = ToDoubleNumber(NumericVal(txtInsBal.Text))
        End If
        
        If (rsFreePayment!FreeTPL) = True Then
            txtTPLPay.Text = "FREE"
        Else
            txtTPLPay.Text = ToDoubleNumber(NumericVal(txtOtherBal.Text))
        End If
        
        If (rsFreePayment!FreeChattel) = True Then
            txtChattelPay.Text = "FREE"
        Else
            txtChattelPay.Text = ToDoubleNumber(NumericVal(txtChattelBal.Text))
        End If
        'JRE
        
        txtAmount = NumericVal(txtDownAmount) + NumericVal(txtLTOAmount) + NumericVal(txtInsAmount) + NumericVal(txtTPLAmout) + NumericVal(txtChattelAmount)
        txtBalance = NumericVal(txtDownBal) + NumericVal(txtLTOBal) + NumericVal(txtInsBal) + NumericVal(txtOtherBal) + NumericVal(txtChattelBal)
        
        If txtDownBal.Text = "0.00" And txtInsBal.Text = "0.00" And txtLTOBal.Text = "0.00" And txtOtherBal.Text = "0.00" And txtChattelBal.Text = "0.00" Then
           MessagePop Star, "Information", "Balance has been fully paid."
           'StoreMemVars
           Call cmdTranCancel_Click
           Exit Function
        Else
           chkDownPayment.Value = 0
           chkInsurance.Value = 0
           chkTPL.Value = 0
           chkLTORegFee.Value = 0
           chkChattel.Value = 0
           Payment.Text = "0.00"
           picORPayment.Visible = True: picORPayment.ZOrder 0
           picDetails.Enabled = False
           grdDetails.Enabled = False
        End If
        
        Call checkifpaid
        StoreMemVars
    Else
        Dim rsPurchAgree                                                As ADODB.Recordset
        Dim rsBalance                                                   As ADODB.Recordset
        Dim rsDeposit                                                   As ADODB.Recordset
        
        Set rsPurchAgree = New ADODB.Recordset
        Set rsBalance = New ADODB.Recordset
        
        Set rsPurchAgree = gconDMIS.Execute("SELECT CODE,ALL_CUSTOMER.CUSTYPE AS TYPES,SMIS_PURCHAGREE.DEYT,SMIS_PURCHAGREE.DOWNPAYMENT AS DOWNPAYMENT,SMIS_PURCHAGREE.TOTAL AS TOTAL,ALL_CUSTOMER.LASTNAME + ALL_CUSTOMER.FIRSTNAME AS CUSTOMERNAME,SMIS_PURCHAGREE.TERM,SMIS_PURCHAGREE.NETSALESPRICE FROM ALL_CUSTOMER INNER JOIN SMIS_PURCHAGREE ON ALL_CUSTOMER.CUSCDE = SMIS_PURCHAGREE.CODE WHERE SMIS_PURCHAGREE.VI_NO = " & N2Str2Null(txtReference.Text) & " AND SMIS_PURCHAGREE.CODE =" & N2Str2Null(txtCUSCDE.Text))
        Set rsBalance = gconDMIS.Execute("SELECT ROUND(SUM(DT.PAYMENT + TAX),2) AS MGA_BAYAD_MO,INVOICENO FROM CMIS_OFF_DT DT WHERE DT.INVOICENO = " & N2Str2Null(txtReference.Text) & " AND DT.CUSCDE = " & N2Str2Null(txtCUSCDE.Text) & " and Left(OR_NUM,3) <> 'SOA' GROUP BY INVOICENO")
        Set rsDeposit = gconDMIS.Execute("SELECT ISNULL(ROUND(SUM(AMOUNT),2),0) AS MGA_APPLIED_DP FROM CMIS_DEPOSITS WHERE APPLIED ='Y' AND INVOICENO=" & N2Str2Null(txtReference.Text))
        If Not rsPurchAgree.EOF And Not rsPurchAgree.BOF Then
            If Not (rsBalance.EOF And rsBalance.BOF) Then
                If N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP) = N2Str2Zero(rsPurchAgree!NETSALESPRICE) Then
                    MessagePop Star, "Information", "Balance has been fully paid."
                    PersonalPayment = True
                    Exit Function
                End If
            
                If N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP) = N2Str2Zero(rsPurchAgree!DownPayment) Then
                    MessagePop Star, "Information", "Balance has been fully paid."
                    PersonalPayment = True
                    Exit Function
                End If
            End If
            
            If rsPurchAgree!TERM = "COD" Or rsPurchAgree!TERM = "CPO" Then
                If Not rsBalance.EOF And Not rsBalance.BOF Then
                    txtAmount.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE)
                    Payment.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE - N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP))
                    txtBalance.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE - N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP))
                Else
                    txtAmount.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE)
                    Payment.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE)
                    txtBalance.Text = ToDoubleNumber(rsPurchAgree!NETSALESPRICE)
                End If
                
                txtDescript.Text = Null2String(rsPurchAgree!CUSTOMERNAME)
                lblInvoiceNo.Caption = txtReference.Text
                labDocDate.Caption = Null2Date(rsPurchAgree!deyt)
                labCUSCODE.Caption = Null2String(rsPurchAgree!Code)
            Else
                lblInvoiceNo.Caption = txtReference.Text
                txtDescript.Text = Null2String(rsPurchAgree!CUSTOMERNAME)
                
                'JJE
                txtAmount.Text = ToDoubleNumber(rsPurchAgree!DownPayment)
                labDocDate.Caption = Null2Date(rsPurchAgree!deyt)
                labCUSCODE.Caption = Null2String(rsPurchAgree!Code)
                
                If Not rsBalance.EOF And Not rsBalance.BOF Then
                    Payment.Text = ToDoubleNumber(rsPurchAgree!DownPayment - N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP))
                    txtBalance.Text = ToDoubleNumber(rsPurchAgree!DownPayment - N2Str2Zero(rsBalance!MGA_BAYAD_MO) + N2Str2Zero(rsDeposit!MGA_APPLIED_DP))
                Else
                    Payment.Text = ToDoubleNumber(rsPurchAgree!DownPayment)
                    txtBalance.Text = ToDoubleNumber(rsPurchAgree!DownPayment)
                End If
            End If
            
            If Left(txtOR_NUM, 3) <> "SOA" Then
                Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
            End If
            PersonalPayment = True
        Else
            PersonalPayment = False
        End If
    End If
End Function

Function CompanyFinancing(VVV As String) As Boolean
    Dim rsBankFinance                                               As ADODB.Recordset
    Dim rsBalance                                                   As ADODB.Recordset
    
    Set rsBankFinance = New ADODB.Recordset
    Set rsBalance = New ADODB.Recordset
    
    Set rsBankFinance = gconDMIS.Execute("SELECT SMIS_PurchAgree.BalToFinanced AS BalanceFinance,SMIS_PurchAgree.FinancingCo,SMIS_PurchAgree.VI_No as VI,SMIS_PurchAgree.SO_NO from SMIS_PurchAgree INNER JOIN SMIS_SalesOrder ON SMIS_SalesOrder.Code = SMIS_PurchAgree.Code WHERE SMIS_PurchAgree.FinancingCo = " & N2Str2Null(cboCUSNAME.Text) & " AND SMIS_PurchAgree.VI_No=" & N2Str2Null(txtReference.Text))
    If Not rsBankFinance.EOF And Not rsBankFinance.BOF Then
        vCustype = "B"
        'JJE 01/03/2013
        Set rsBalance = gconDMIS.Execute("SELECT ISNULL(ROUND(SUM(PAYMENT + TAX),2),0) AS MGA_BAYAD_MO FROM CMIS_OFF_DT WHERE INVOICENO =" & N2Str2Null(txtReference.Text) & " AND CUSCDE = " & N2Str2Null(txtCUSCDE.Text) & " and Left(OR_NUM,3) <> 'SOA'")
        If Not rsBalance.BOF And Not rsBalance.EOF And rsBalance!MGA_BAYAD_MO <> 0 Then
            If rsBalance!MGA_BAYAD_MO = rsBankFinance!balancefinance Then
                MessagePop Star, "Information", "Balance has been fully paid."
                Exit Function
            Else
                txtBalance.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance - rsBalance!MGA_BAYAD_MO))
                txtAmount.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance))
                Payment.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance - rsBalance!MGA_BAYAD_MO))
                txtSO_NO.Text = Null2String(rsBankFinance!SO_NO)
                labCUSCODE.Caption = Null2String(VVV)
             End If
        Else
            txtBalance.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance))
            txtAmount.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance))
            labCUSCODE.Caption = Null2String(VVV)
            Payment.Text = ToDoubleNumber(NumericVal(rsBankFinance!balancefinance))
        End If
            
        Set rsOFF_DT = New ADODB.Recordset
        Set rsOFF_DT = gconDMIS.Execute("SELECT ISNULL(ROUND(SUM(PAYMENT + TAX),2),0) AS MGA_BAYAD,TRANTYPE,REFERENCE FROM CMIS_Off_Dt INNER JOIN SMIS_FINCOM ON CODE = CUSCDE WHERE trantype = 'VI' AND Reference = " & N2Str2Null(txtReference.Text) & " AND LEFT(OR_NUM,3) <> 'SOA' GROUP BY REFERENCE,TRANTYPE")
        If Not rsOFF_DT.EOF And Not rsOFF_DT.BOF Then
            Set rsCustomerDeposit = New ADODB.Recordset
            rsCustomerDeposit.Open "SELECT ROUND(SUM(AMOUNT),2) AS DEPOSIT_AMOUNT FROM CMIS_DEPOSITS WHERE INVOICENO=" & N2Str2Null(rsOFF_DT!REFERENCE) & " AND INVOICETYPE = " & N2Str2Null(rsOFF_DT!TranType) & " AND APPLIED = 'Y'", gconDMIS, adOpenKeyset
            If Not rsCustomerDeposit.EOF And Not rsCustomerDeposit.BOF Then
                txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text) - (NumericVal(rsOFF_DT!MGA_BAYAD) + NumericVal(rsCustomerDeposit!DEPOSIT_AMOUNT)))
                Call BalanceCash(cboInvoiceType, txtReference)
            End If 'end if rsCustomerDeposit
        End If 'end if rsOFF_DT
        CompanyFinancing = True
    Else
        CompanyFinancing = False
    End If 'end if rsBankFinance
End Function

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtReference_LostFocus()
    If labRef.Caption = "Inv. #" Then txtReference.Text = Format(txtReference.Text, "000000")
End Sub

Private Sub txtReference2_Change()
'DESCRIPTION: Search for OR No. to be linked with CITIBANK Payment
    Dim xList                                                       As ListItem
    Dim rsCMIS_OFF_HD                                               As ADODB.Recordset
    Dim rsValidateOR                                                As ADODB.Recordset
    'If Len(txtReference2) = 8 Then
        Set rsCMIS_OFF_HD = New ADODB.Recordset
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3' AND (Paidby is null or paidby = 'N') AND OR_NUM LIKE '" & txtReference2 & "%' ORDER BY OR_Date")
        If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
            lvPayments.ListItems.Clear
            lblTotal = "0.00"
            Do While Not rsCMIS_OFF_HD.EOF
                Set rsValidateOR = New ADODB.Recordset
                Set rsValidateOR = gconDMIS.Execute("SELECT INVOICENO FROM CMIS_Off_Dt WHERE INVOICENO = '" & rsCMIS_OFF_HD!OR_NUM & "' and PAIDFOR = '427'")
                If Not (rsValidateOR.EOF And rsValidateOR.BOF) Then
                    'Nothing
                rsCMIS_OFF_HD.MoveNext
                Else
                    Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
                    xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
                    xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!cusname)
                    xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
                    xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
                    xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
                    tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
                    lblTotal = Format(tmpTotal, "#,###,##0.00")
                    rsCMIS_OFF_HD.MoveNext
                End If
            Loop
        End If
        Set rsCMIS_OFF_HD = Nothing
    'End If
End Sub

Private Sub txtTax_Change()
    'JJE
    If txtBalance.Text > 0 Then
        Payment = Round(NumericVal(txtBalance) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
    Else
        Payment = Round(NumericVal(txtAmount) - (NumericVal(txtDiscount) + NumericVal(txtTax)), 2)
        wizDigit1.TextValue = ToDoubleNumber(NumericVal(Payment.Text))
    End If
End Sub

Private Sub txtTax_GotFocus()
    'JJE
    txtTax.Text = ToDoubleNumber(txtTax.Text)
    'JJE
End Sub

Private Sub txtTax_KeyPress(KeyAscii As Integer)
    'JJE
     KeyAscii = OnlyNumeric(KeyAscii)
    'JJE
End Sub

'SEARCH MODULE
Private Sub lstOFF_HD_GotFocus()
'On Error Resume Next
    If lstOFF_HD.Enabled = True Then
        rsOFF_HD.MoveFirst
        rsOFF_HD.Find ("ID=" & lstOFF_HD.SelectedItem.SubItems(1))
        StoreMemVars
    End If
End Sub

Private Sub lstOFF_HD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lstOFF_HD.Enabled = True Then
        If optORNo.Value = True Then
            On Error Resume Next
            rsOFF_HD.Bookmark = rsFind(rsOFF_HD.Clone, "OR_NUM", Item).Bookmark
        Else
            On Error Resume Next
            rsOFF_HD.Bookmark = rsFind(rsOFF_HD.Clone, "ID", lstOFF_HD.SelectedItem.SubItems(1)).Bookmark
        End If
        StoreMemVars
    End If
End Sub

Private Sub lstOFF_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOFF_HD
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

Private Sub lstOFF_HD_DblClick()
    If cmdEdit.Enabled = True Then cmdEdit.Value = True
End Sub

Private Sub lstOFF_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstOFF_HD.ListItems.Count > 0 And lstOFF_HD.Enabled = True Then: lstOFF_HD.SetFocus
    End If
End Sub

Private Sub optCustName_Click()
    lstOFF_HD.Enabled = False
    lstOFF_HD.ColumnHeaders(1).Text = "Cust. Name"
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optORNo_Click()
    lstOFF_HD.Enabled = False
    lstOFF_HD.ColumnHeaders(1).Text = "OR No."
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

'Function GETTRANTYPE(vtranno As String) As String
'    Dim rsOrdTranno As ADODB.Recordset
'    Set rsOrdTranno = New ADODB.Recordset
'    Set rsOrdTranno = gconDMIS.Execute("SELECT TRANNO,TRANTYPE FROM PMIS_ORD_HD WHERE TYPE = '" & Mid(vTRANTYPE, 2, 1) & "' AND TRANNO = " & N2Str2Null(vtranno) & "")
'    If Not rsOrdTranno.EOF And Not rsOrdTranno.BOF Then
'        GETTRANTYPE = Null2String(rsOrdTranno!TRANTYPE)
'    Else
'        GETTRANTYPE = "NULL"
'    End If
'    Set rsOrdTranno = Nothing
'End Function

'DESCRIPTION: Generate new ReferenceNo upon posting credit card transaction

Function GetReferenceNo() As String
    Dim rsCMIS_OFF_HD                                               As ADODB.Recordset
    Set rsCMIS_OFF_HD = New ADODB.Recordset
    Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT CAST(ReferenceNo AS int) AS MAX_REFERENCENO FROM CMIS_Off_HD ORDER BY MAX_REFERENCENO DESC")
    If Not rsCMIS_OFF_HD.EOF And Not rsCMIS_OFF_HD.BOF Then
        GetReferenceNo = Format(NumericVal(rsCMIS_OFF_HD!MAX_REFERENCENO) + 1, "00000000")
    Else
        GetReferenceNo = "00000001"
    End If
End Function

Function BalanceCash(xInvoiceType As String, xReference As String)
'DESCRIPTION: Check for Customer Balance
    Dim rsOFF_DTStat                                                As ADODB.Recordset
    Set rsOFF_DTStat = New ADODB.Recordset
    Set rsOFF_DTStat = gconDMIS.Execute("SELECT OR_num,Payment,PaidNa FROM CMIS_OFF_DT WHERE Reference = " & N2Str2Null(xReference) & " and left(or_num,3) <> 'SOA'")

    If Not rsOFF_DTStat.EOF And Not rsOFF_DTStat.BOF Then
        If txtBalance.Text <= 0 And rsOFF_DTStat!Paidna = True Then
            cmdTranCancel.Value = True
            MessagePop Star, "Information", "Balance has been fully paid."
        ElseIf txtBalance.Text > 0 And rsOFF_DTStat!Paidna = False Then
'            Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
            MessagePop InfoWarning, "Information", "Payment has been made but not yet POSTED."
            If Null2Bool(rsOFF_HD!Paidna) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
                If Picture1.Visible = True Then cmdDetails_Click
                chkCreditCardTrans.Value = 0
            End If
        End If
    End If
    Set rsOFF_DTStat = Nothing
End Function

Sub BalanceCharge(xInvoiceType As String, xReference As String)
    Dim rsOFF_DTStat                                                As ADODB.Recordset
    Set rsOFF_DTStat = New ADODB.Recordset
    Set rsOFF_DTStat = gconDMIS.Execute("SELECT PaidNa FROM CMIS_OFF_DT WHERE INVOICETYPE = " & N2Str2Null(xInvoiceType) & " AND Reference = " & N2Str2Null(xReference))
    If Not rsOFF_DTStat.EOF And Not rsOFF_DTStat.BOF Then
        If txtBalance.Text <= 0 And rsOFF_DTStat!Paidna = True Then
            cmdTranCancel.Value = True
            MessagePop Star, "Information", "Balance has been fully paid."
        ElseIf txtBalance.Text <= 0 And rsOFF_HD!Paidna = False Then
            MessagePop InfoWarning, "Information", "Payment has been made but not yet POSTED."
            If Null2Bool(rsOFF_HD!Paidna) = False And Null2Bool(rsOFF_HD!Cancel) = False Then
                If Picture1.Visible = True Then cmdDetails_Click
                chkCreditCardTrans.Value = 0
            End If
        End If
    End If
    Set rsOFF_DTStat = Nothing
End Sub

Sub Unapplied_Deposits(XXX As String)
'DESCRIPTION: List of Customer Deposits
    Dim Trantypecode                                                As String
    Dim xList                                                       As ListItem
    Dim rsUnapplied                                                 As ADODB.Recordset
    Set rsUnapplied = New ADODB.Recordset
    'rsUnapplied.Open "SELECT * FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & XXX & "' AND PAIDNA =1 ", gconDMIS, adOpenKeyset
    
    'JJE 01/03/2013
    If COMPANY_CODE = "DSSC" Then
        If Left(cboTranType, 1) = "S" Then
            rsUnapplied.Open "SELECT * FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & XXX & "' AND PAIDNA =1 AND PAIDFOR = '412S'", gconDMIS, adOpenKeyset
        Else
            rsUnapplied.Open "SELECT * FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & XXX & "' AND PAIDNA =1 AND PAIDFOR in ('412P','412S')", gconDMIS, adOpenKeyset
        End If
    Else
        If Left(cboTranType, 1) = "S" Or Left(cboTranType, 1) = "V" Then
            rsUnapplied.Open "SELECT * FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & XXX & "' AND PAIDNA =1 AND PAIDFOR = '412" & Left(cboTranType, 1) & "'", gconDMIS, adOpenKeyset
        Else
            rsUnapplied.Open "SELECT * FROM (SELECT HD.OR_NUM,HD.STATUS,HD.PAIDNA,DP.ORDATE,ISNULL((SELECT SUM(ISNULL(AMOUNT,0)) AS AMOUNT FROM CMIS_DEPOSITDT WHERE DEPOSIT_ID=DP.ID),0) AS APPLIEDAMT,HD.CUSCDE,DP.APPLIED,DP.ID_DET,DP.ID,DP.PAIDFOR,DP.AMOUNT FROM CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITS DP ON HD.OR_NUM=DP.OR_NUM)T WHERE AMOUNT-APPLIEDAMT>0 AND CUSCDE ='" & XXX & "' AND PAIDNA =1 AND PAIDFOR = '412P'", gconDMIS, adOpenKeyset
        End If
    End If
    
    lvDeposits.ListItems.Clear
    If Not rsUnapplied.EOF And Not rsUnapplied.BOF Then
        picDeposits.Visible = True
        picDeposits.ZOrder 0
        cmdTranSave.Enabled = False
        Do While Not rsUnapplied.EOF
            Set xList = lvDeposits.ListItems.Add(, , Null2String(rsUnapplied!CUSCDE))
            xList.SubItems(1) = Null2Date(Format(rsUnapplied!ORDATE, "mm/dd/yyyy"))
            xList.SubItems(2) = rsUnapplied!OR_NUM
            xList.SubItems(3) = ToDoubleNumber(rsUnapplied!amount - NumericVal(rsUnapplied!APPLIEDAMT))
            xList.SubItems(4) = N2Str2Null(rsUnapplied!Applied)
            xList.SubItems(5) = Null2String(rsUnapplied!ID_Det)
            xList.SubItems(6) = Null2String(rsUnapplied!Id)
            xList.SubItems(7) = SetPaidFor(Null2String(rsUnapplied!PAIDFOR))
            rsUnapplied.MoveNext
        Loop
    End If
    Set rsUnapplied = Nothing
End Sub

Sub CreditCardPayments()
'DESCRIPTION: List all Credit Card Payments
    Dim xList                                                       As ListItem
    Dim rsCMIS_OFF_HD                                               As ADODB.Recordset
    
    If COMPANY_CODE = "HGC" Then
        'UPDATED BY : ROWEL DE QUIROZ
        'DATE : MARCH 3 2011
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3' AND Paidby <> 'Y' AND cardbnkcde = '" & txtCUSCDE & "' and OR_DATE >='2/1/2010' AND OR_NUM NOT IN(SELECT INVOICENO FROM CMIS_Off_Dt WHERE or_num = '" & txtOR_NUM.Text & "' ) ORDER BY OR_Date")
    ElseIf COMPANY_CODE = "DGI" Then
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3' AND (Paidby IS NULL OR PAIDBY = 'N') and cardbnkcde = '" & txtCUSCDE & "' ORDER BY OR_Date")
    Else
        Set rsCMIS_OFF_HD = gconDMIS.Execute("SELECT ReferenceNo,CusCde,CusName,Or_Amt,OR_NUM,OR_Date FROM CMIS_OFF_HD WHERE TOF = '3' AND (Paidby IS NULL OR PAIDBY = 'N') and cardbnkcde = '" & txtCUSCDE & "' ORDER BY OR_Date")
    End If
    If Not (rsCMIS_OFF_HD.EOF And rsCMIS_OFF_HD.BOF) Then
        lvPayments.ListItems.Clear
        lblTotal = "0.00"
        Do While Not rsCMIS_OFF_HD.EOF
            Set xList = lvPayments.ListItems.Add(, , Null2String(rsCMIS_OFF_HD!OR_NUM))
            xList.SubItems(1) = Null2String(rsCMIS_OFF_HD!CUSCDE)
            xList.SubItems(2) = Null2String(rsCMIS_OFF_HD!cusname)
            xList.SubItems(3) = ToDoubleNumber(rsCMIS_OFF_HD!OR_AMT)
            xList.SubItems(4) = Null2String(rsCMIS_OFF_HD!ReferenceNo)
            xList.SubItems(5) = Null2Date(rsCMIS_OFF_HD!OR_DATE)
            tmpTotal = NumericVal(lblTotal) + NumericVal(xList.SubItems(3))
            lblTotal = Format(tmpTotal, "#,###,##0.00")
            rsCMIS_OFF_HD.MoveNext
        Loop
    Else
        MessagePop RecNotFound, "No record to view", "No Record"
    End If
    Set rsCMIS_OFF_HD = Nothing
End Sub

Function CheckIfBank(xCusCde As String) As Boolean
    Dim rsCheckCode                                                 As ADODB.Recordset
    Set rsCheckCode = New ADODB.Recordset
    rsCheckCode.Open "SELECT Cuscde FROM All_Customer_Table WHERE CusCde = " & N2Str2Null(xCusCde) & "", gconDMIS, adOpenForwardOnly
    If Not rsCheckCode.EOF And Not rsCheckCode.BOF Then
        Do While Not rsCheckCode.EOF
            Dim rsCheckBank                                         As ADODB.Recordset
            Set rsCheckBank = New ADODB.Recordset
            rsCheckBank.Open "SELECT CusCde FROM CMIS_CardCompany WHERE CusCde = " & N2Str2Null(rsCheckCode!CUSCDE) & "", gconDMIS, adOpenForwardOnly
            If Not rsCheckBank.EOF And Not rsCheckBank.BOF Then
                CheckIfBank = True
            Else
                CheckIfBank = False
            End If
            rsCheckCode.MoveNext
        Loop
    End If
    Set rsCheckCode = Nothing
    Set rsCheckBank = Nothing
End Function

Function CheckDeposited(xORNUM As String) As Boolean
    Dim rsCheckDeposited                                            As ADODB.Recordset
    Set rsCheckDeposited = New ADODB.Recordset
    rsCheckDeposited.Open "SELECT * FROM CMIS_BANKDEPO WHERE OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckDeposited.EOF And Not rsCheckDeposited.BOF Then
        CheckDeposited = True
    End If
End Function

Function CheckORCutOff(xORNUM As String) As Boolean
    On Error Resume Next
    Dim rsCheckORCutOff                                             As ADODB.Recordset
    Set rsCheckORCutOff = New ADODB.Recordset
    rsCheckORCutOff.Open "SELECT * from CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "' AND CutDate IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsCheckORCutOff.EOF And Not rsCheckORCutOff.BOF Then
        CheckORCutOff = True
    End If
End Function

Function CheckCutOffDate(xORNUM As String) As String
    On Error Resume Next
    Dim rsCheckORCutOff                                             As ADODB.Recordset
    Set rsCheckORCutOff = New ADODB.Recordset
    rsCheckORCutOff.Open "SELECT * FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "' AND CutDate IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsCheckORCutOff.EOF And Not rsCheckORCutOff.BOF Then
        CheckCutOffDate = CDate(rsCheckORCutOff!CUTDATE)
    End If
End Function

Function CheckPostedOR(xORNUM As String) As Boolean
    Dim rsCheckPosted                                               As ADODB.Recordset
    Set rsCheckPosted = New ADODB.Recordset
    If VAT_OR = 1 Then
        rsCheckPosted.Open "SELECT * FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "' AND PAIDNA=1 AND VAT=1", gconDMIS, adOpenKeyset
    Else
        rsCheckPosted.Open "SELECT * FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "' AND PAIDNA=1 AND VAT=0", gconDMIS, adOpenKeyset
    End If
    If Not rsCheckPosted.EOF And Not rsCheckPosted.BOF Then
        CheckPostedOR = True
    End If
End Function

Function CashAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                             As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "SELECT CashAmount,ChkAmount,CardAmount FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CashAmount = NumericVal(rsCheckPayments!CashAmount)
    End If
End Function

Function CheckAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                             As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "SELECT CashAmount,ChkAmount,CardAmount FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CheckAmount = NumericVal(rsCheckPayments!CHKAMOUNT)
    End If
End Function

Function CardAmount(xORNUM As String) As Currency
    Dim rsCheckPayments                                             As ADODB.Recordset
    Set rsCheckPayments = New ADODB.Recordset
    rsCheckPayments.Open "SELECT CashAmount,ChkAmount,CardAmount FROM CMIS_OFF_HD WHERE OR_NUM = '" & xORNUM & "'", gconDMIS, adOpenKeyset
    If Not rsCheckPayments.EOF And Not rsCheckPayments.BOF Then
        CardAmount = NumericVal(rsCheckPayments!CardAmount)
    End If
End Function

Function CheckIfImportedinAMIS(xOR_Num As String) As Boolean
    Dim rsPostedCRJ                                                 As ADODB.Recordset
    Set rsPostedCRJ = New ADODB.Recordset
    rsPostedCRJ.Open "SELECT TOP 1 * FROM AMIS_Journal_HD WHERE JTYPE='CRJ' AND Status <> 'C' AND ISNULL(InvoiceNo,'') ='" & xOR_Num & "'", gconDMIS, adOpenKeyset
    If Not rsPostedCRJ.EOF And Not rsPostedCRJ.BOF Then
        CheckIfImportedinAMIS = True
    End If
End Function

Sub UnPost_CashPos()
    If COMPANY_CODE = "DJM" And OR_VAT_NONVAT = "NON-VAT" Then
        'Do nothing SOA transactions does not have to reflect in Cash Position
    Else
        If MODE_OF_PAYMENT = "CASH" Then
            gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET " & _
                              "CASH = CASH - " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        ElseIf MODE_OF_PAYMENT = "CHECK" Then
            gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET " & _
                              "[CHECK] = [CHECK] - " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        ElseIf MODE_OF_PAYMENT = "CARD" Then
            gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET " & _
                              "CARD = CARD - " & RECEIPTS_AMOUNT & " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
    
        If CheckDeposited(txtOR_NUM) = True Then
            If MODE_OF_PAYMENT = "CASH" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CASH = CASH + " & RECEIPTS_AMOUNT & "," & _
                                  " CASHDEPO = CASHDEPO - " & RECEIPTS_AMOUNT & "" & _
                                  " WHERE CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
            ElseIf MODE_OF_PAYMENT = "CHECK" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " [CHECK] = [CHECK] + " & RECEIPTS_AMOUNT & "," & _
                                  " CHECKDEPO = CHECKDEPO - " & RECEIPTS_AMOUNT & "" & _
                                  " WHERE CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
            ElseIf MODE_OF_PAYMENT = "CARD" Then
                gconDMIS.Execute ("UPDATE CMIS_Cash_Pos SET" & _
                                  " CARD = CARD + " & RECEIPTS_AMOUNT & "," & _
                                  " CARDDEPO = CARDDEPO - " & RECEIPTS_AMOUNT & _
                                  " WHERE CUTDATE = '" & Format(CDate(CURRENT_CUTOFF_DATE), "MM/DD/YYYY") & "'")
            End If
        End If
    End If
End Sub

Function CheckIfCancel(xOR_Num As String) As Boolean
    Dim rsCheckCancel                                               As ADODB.Recordset
    Set rsCheckCancel = New ADODB.Recordset
    rsCheckCancel.Open "SELECT * FROM CMIS_OFF_HD WHERE Cancel=1 AND OR_NUM = '" & xOR_Num & "'", gconDMIS, adOpenKeyset
    If Not rsCheckCancel.EOF And Not rsCheckCancel.BOF Then
        CheckIfCancel = True
    End If
    Set rsCheckCancel = Nothing
End Function

Function CheckAppliedDeposit(xOR_Num As String) As Boolean
    Dim rsDeposit                                                   As ADODB.Recordset
    Set rsDeposit = New ADODB.Recordset
    
    'JJE Updated 01/24/2013 2:32PM
    'Check if Deposit is already applied
    rsDeposit.Open "SELECT * FROM CMIS_DEPOSITS WHERE OR_NUM = '" & xOR_Num & "' and APPLIED = 'Y'", gconDMIS, adOpenKeyset
    If Not rsDeposit.EOF And Not rsDeposit.BOF Then
        CheckAppliedDeposit = True
    End If
    'JJE
    Set rsDeposit = Nothing
End Function

Function GetInvoiceNo(xOR_Num As String) As String
    Dim rsInvoiceNo                                                 As ADODB.Recordset
    Set rsInvoiceNo = New ADODB.Recordset
    rsInvoiceNo.Open "SELECT INVOICENO FROM CMIS_OFF_DT WHERE OR_NUM =" & xOR_Num & "", gconDMIS, adOpenKeyset
    If Not rsInvoiceNo.EOF And Not rsInvoiceNo.BOF Then
        GetInvoiceNo = N2Str2Null(rsInvoiceNo!INVOICENO)
    End If
    Set rsInvoiceNo = Nothing
End Function

Function GetCRJNo(xOR_Num As String, xInvoiceType As String) As String
    Dim rsJOURNALHD                                                 As ADODB.Recordset
    Set rsJOURNALHD = New ADODB.Recordset
    rsJOURNALHD.Open ("SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE INVOICENO='" & xOR_Num & "' AND INVOICETYPE='" & xInvoiceType & "' AND JTYPE='CRJ'"), gconDMIS, adOpenForwardOnly
    If Not rsJOURNALHD.EOF And Not rsJOURNALHD.BOF Then
        GetCRJNo = rsJOURNALHD!VOUCHERNO
        labDetails = "IMPORTED: CASH RECEIPTS JOURNAL"
    Else
        labDetails = ""
    End If
    Set rsJOURNALHD = Nothing
End Function

Function GetLASTOR(XXX As String) As String
    Dim rsOR                                                        As ADODB.Recordset
    Set rsOR = New ADODB.Recordset
    
    'JJE 06/08/2012
    If NEXTNUMBER_CR_OR = True And OR_VAT_NONVAT = "VAT" Then 'Add Prefixes back to 1
        If XXX = "G" Then
            rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where left(or_num,2) = 'CR'", gconDMIS, adOpenForwardOnly
            If Not rsOR.EOF And Not rsOR.BOF Then
                GetLASTOR = "CR" + "" + Format(N2Str2Zero(rsOR!OR_NUM) + 1, "000000")
            Else
                GetLASTOR = "CR" + "" + "000001"
            End If
        'JJE 05/20/2016 FOR MGS
        ElseIf XXX = "V" Then
            rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where left(or_num,2) = 'PR'", gconDMIS, adOpenForwardOnly
            If Not rsOR.EOF And Not rsOR.BOF Then
                GetLASTOR = "PR" + "" + Format(N2Str2Zero(rsOR!OR_NUM) + 1, "000000")
            Else
                GetLASTOR = "PR" + "" + "000001"
            End If
        Else
            rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where left(OR_NUM,2) = 'OR'", gconDMIS, adOpenForwardOnly
            If Not rsOR.EOF And Not rsOR.BOF Then
                GetLASTOR = "OR" + "" + Format(N2Str2Zero(rsOR!OR_NUM) + 1, "000000")
            Else
                GetLASTOR = "OR" + "" + "000001"
            End If
        End If
    Else
        'JJE 11/29/2012 2PM
        If COMPANY_CODE = "DGI" Or COMPANY_CODE = "DSSC" Then
            If XXX <> "S" Then
                rsOR.Open "SELECT REPLICATE('0',6-LEN((case when COUNT(or_num) = 0 then '1' else max(or_num) + 1 end)))+ cast( (case when count(or_num) = 0  then '1' else max(or_num) + 1 end) AS nvarchar(6))as OR_NUM FROM CMIS_OFF_HD WHERE Left(OR_NUM,1) NOT IN ('S','G')", gconDMIS, adOpenForwardOnly
                If Not rsOR.EOF And Not rsOR.BOF Then
                    GetLASTOR = Null2String(rsOR!OR_NUM)
                End If
            Else
                'rsOR.Open "SELECT max (substring(or_num,2,len(or_num))) + 1 as OR_NUM FROM CMIS_OFF_HD WHERE Left(OR_NUM,1) in ('S')", gconDMIS, adOpenForwardOnly
                rsOR.Open "SELECT case when MAX(or_num) IS NULL then 1 when max(or_num)<> 'null' then max (SUBSTRING(ISNULL(or_num,0),2,LEN(ISNULL(or_num,0)))) + 1  end or_num FROM CMIS_OFF_HD WHERE Left(OR_NUM,1) in ('S')", gconDMIS, adOpenForwardOnly
                If Not rsOR.EOF And Not rsOR.BOF Then
                    GetLASTOR = XXX + "" + Format(Null2String(rsOR!OR_NUM), "000000")
                 End If
            End If
        ElseIf COMPANY_CODE = "DJM" Then
            If XXX = "SOA" Then
                rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where len(or_num) = 9 and left(OR_NUM,3) = '" & XXX & "'", gconDMIS, adOpenForwardOnly
                If Not rsOR.EOF And Not rsOR.BOF Then
                    GetLASTOR = XXX + Format(Null2String(rsOR!OR_NUM) + 1, "000000")
                End If
            End If
        '... JRE 06/08/2016 auto increment of OR number for CMC
        ElseIf COMPANY_CODE = "CMC" Then
            If XXX = "V" Then
                rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where Len(OR_NUM) = 6 and VAT = '1' and OR_NUM < '686335'", gconDMIS, adOpenForwardOnly
                If Not rsOR.EOF And Not rsOR.BOF Then
                    GetLASTOR = "" + Format(N2Str2Zero(rsOR!OR_NUM) + 1, "000000")
                End If
            ElseIf XXX = "NV" Then
                rsOR.Open "SELECT isnull(MAX(Right(OR_NUM,6)),0) AS OR_NUM from cmis_off_hd where Len(OR_NUM) = 6 and VAT = '0' and OR_NUM < '067806'", gconDMIS, adOpenForwardOnly
                If Not rsOR.EOF And Not rsOR.BOF Then
                    GetLASTOR = "" + Format(N2Str2Zero(rsOR!OR_NUM) + 1, "000000")
                End If
            End If
        '...
        Else
            rsOR.Open "SELECT REPLICATE('0',5-LEN(ISNULL(MAX(SUBSTRING(OR_NUM,2,len(OR_NUM))),0)+1)) + CAST(ISNULL(MAX(SUBSTRING(OR_NUM,2,len(OR_NUM))),0)+1 AS NVARCHAR(6)) as OR_NUM from CMIS_OFF_HD WHERE LEFT(OR_NUM,1) = '" & XXX & "'", gconDMIS, adOpenForwardOnly
            If Not rsOR.EOF And Not rsOR.BOF Then
                GetLASTOR = XXX + Null2String(rsOR!OR_NUM)
            End If
        End If
    End If
    Set rsOR = Nothing
    'JJE
End Function

Function SetVendorName(VVV As Variant)
    Dim rsVENDOR                                                    As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "SELECT CODE,ACCOUNTNAME as nameofvendor from ALL_ENTITY WHERE COMPLET_CODE= " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        SetVendorName = Null2String(rsVENDOR!nameofvendor)
    Else
        SetVendorName = ""
    End If
End Function

Function CHECKIFSCHED(XXX As String) As Boolean
    Dim rsCHART As ADODB.Recordset
    Set rsCHART = New ADODB.Recordset
    rsCHART.Open "SELECT ACCTCODE from AMIS_CHARTACCOUNT WHERE IS_SCHEDULE_ACCNT=1 and ACCTCODE='" & XXX & "' and TRANTYPE1 IN ('INSURANCE','LTO','CHATTEL')", gconDMIS, adOpenForwardOnly
    If Not rsCHART.EOF And Not rsCHART.BOF Then
        CHECKIFSCHED = True
    Else
        CHECKIFSCHED = False
    End If
    Set rsCHART = Nothing
End Function

Function InitializePayment() As Boolean
    Dim rsAllPayment As New ADODB.Recordset
    'JRE 07/20/16 Added Free's
    Set rsAllPayment = gconDMIS.Execute("SELECT Insurance,LTORegFee,Others,Downpayment,CHMOFEE,salesprice,Term,NetSalesPrice,Discount,FREELTO,FREEINSURANCE,FREETPL,FREECHATTEL from SMIS_SalesOrder WHERE VI_NO ='" & txtReference.Text & "'and CODE = '" & txtCUSCDE & "'")
    If Not rsAllPayment.BOF And Not rsAllPayment.EOF Then
        'AMOUNT OF PAYMENT
        If (rsAllPayment!TERM = "COD") Or (rsAllPayment!TERM = "CPO") Then
           vTerm = True
           Label37.Visible = False
           Label45.Visible = True
'           txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!salesprice))
        'JRE 06/27/2016 To less the Discount Amount in the Total Amount
            If (rsAllPayment!DownPayment) = 0 Then
                txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!NETSALESPRICE))
            Else
                txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!salesprice) - (rsAllPayment!DISCOUNT))
            End If
        'JRE
        Else
           Label37.Visible = True
           Label45.Visible = False
           vTerm = False
'        txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!downpayment))
        'JRE 06/27/2016 To less the Discount Amount in the Total Amount
            If (rsAllPayment!DownPayment) = 0 Then
                txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!salesprice) - (rsAllPayment!DISCOUNT))
            Else
                txtDownAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!DownPayment) - (rsAllPayment!DISCOUNT))
            End If
        'JRE
        End If
        
        'JRE 07/20/16 To show 0 if it is FREE in SMIS
        If (rsAllPayment!FreeTPL) = True Then
            txtTPLAmout.Text = 0
        Else
            txtTPLAmout.Text = ToDoubleNumber(NumericVal(rsAllPayment!Others))
        End If
        
        If (rsAllPayment!FREEINSURANCE) = True Then
            txtInsAmount.Text = 0
        Else
            txtInsAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!Insurance))
        End If
        
        If (rsAllPayment!FreeLTO) = True Then
            txtLTOAmount.Text = 0
        Else
            txtLTOAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!LTORegFee))
        End If
        
        If (rsAllPayment!FreeChattel) = True Then
            txtChattelAmount.Text = 0
        Else
            txtChattelAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!CHMOFEE))
        End If
'        txtTPLAmout.Text = ToDoubleNumber(NumericVal(rsAllPayment!Others))
'        txtInsAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!Insurance))
'        txtLTOAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!LTORegFee))
'        txtChattelAmount.Text = ToDoubleNumber(NumericVal(rsAllPayment!CHMOFEE))
        'JRE
        'BALANCE OF PAYMENT
        txtDownBal = ToDoubleNumber(NumericVal(txtDownAmount.Text))
        txtInsBal.Text = ToDoubleNumber(NumericVal(txtInsAmount.Text))
        txtOtherBal.Text = ToDoubleNumber(NumericVal(txtTPLAmout.Text))
        txtBalance.Text = ToDoubleNumber(NumericVal(txtAmount.Text))
        txtLTOBal.Text = ToDoubleNumber(NumericVal(txtLTOAmount.Text))
        txtChattelBal.Text = ToDoubleNumber(NumericVal(txtChattelAmount.Text))
        
        InitializePayment = True
    Else
        InitializePayment = False
    End If
End Function

Private Sub cmdSaveORDetail_Click()
    
    picDetails.Enabled = True
    If picORPayment.Visible = True Then
        If chkDownPayment.Value = 0 And chkLTORegFee.Value = 0 And chkInsurance.Value = 0 And chkTPL.Value = 0 And chkChattel.Value = 0 Then
            MsgBox "Nothing to save", vbInformation, "Customer Payment"
            Exit Sub
        Else
            Call Unapplied_Deposits(Null2String(txtCUSCDE.Text))
        End If
        
        picORPayment.Visible = False
        cmdTranSave.Value = False
    End If
    
'    Call cmdTranSave_Click
'    picORPayment.Visible = False: picORPayment.ZOrder 0
End Sub

Private Sub txtDownPay_Change()
    If NumericVal(txtDownPay.Text) <= 0 Then
        wizDigit1.TextValue = 0
    Else
        If AddorEdit = "EDIT" Then
            'wizDigit1.TextValue = ToDoubleNumber(NumericVal(TOTAL_AR_AMOUNT) + NumericVal(txtBalance.Text))
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtDownPay.Text))
        'JRE 06/27/2016
            Payment = ToDoubleNumber(NumericVal(txtDownPay.Text))
        'JRE
        Else
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtDownPay.Text))
        'JRE 06/27/2016
            Payment = ToDoubleNumber(NumericVal(txtDownPay.Text))
        'JRE
        End If
    End If
End Sub

Private Sub txtDownPay_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtInsAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    KeyAscii = 0
End Sub

Sub ClosePicOrpayment()
    If picDeposits.Visible = True Then
        picDetails.Enabled = True
        picDeposits.Visible = False
        picDeposits.ZOrder 1
    ElseIf picCreditCard.Visible = True Then
        picCreditCard.Visible = False
        picCreditCard.ZOrder 1
    Else
        picDetails.Enabled = True
        fraDetails.Enabled = True
        On_Update = False
        picDetails.ZOrder 1
        picDetails.Visible = False

        cmdDetails.ZOrder 1
        cmdDetails.Visible = False
        picDeposits.Visible = False
        picCreditCard.Visible = False
        picCreditCard.ZOrder 1

        Picture1.Enabled = True
        On Error Resume Next
        grdDetails.SetFocus
    End If
End Sub

Sub checkifpaid()
    If txtInsAmount.Text = "0.00" Or txtInsBal.Text = "0.00" Then
        chkInsurance.Enabled = False
        'chkInsurance.Value = 0
    Else
        chkInsurance.Enabled = True
        chkInsurance.Value = 0
    End If
                
    If txtDownAmount.Text = "0.00" Or txtDownBal.Text = "0.00" Then
        chkDownPayment.Enabled = False
        'chkDownPayment.Value = 0
    Else
        chkDownPayment.Enabled = True
        chkDownPayment.Value = 0
    End If
                 
    If txtLTOAmount.Text = "0.00" Or txtLTOBal.Text = "0.00" Then
        chkLTORegFee.Enabled = False
        'chkLTORegFee.Value = 0
    Else
        chkLTORegFee.Enabled = True
        chkLTORegFee.Value = 0
    End If
                
    If txtTPLAmout.Text = "0.00" Or txtOtherBal.Text = "0.00" Then
        chkTPL.Enabled = False
        'chkTPL.Value = 0
    Else
        chkTPL.Enabled = True
        chkTPL.Value = 0
    End If
    
    If txtChattelAmount.Text = "0.00" Or txtChattelBal.Text = "0.00" Then
        chkChattel.Enabled = False
        'chkTPL.Value = 0
    Else
        chkChattel.Enabled = True
        chkChattel.Value = 0
    End If
End Sub

Sub Applied_Deposits(XXX As String)
'DESCRIPTION: List of Customer Deposits
    Dim xList                                                       As ListItem
    Dim rsAppliedDep                                                As ADODB.Recordset
    Set rsAppliedDep = New ADODB.Recordset
    rsAppliedDep.Open "SELECT * from (SELECT HD.OR_DATE,HD.OR_NUM,HD.STATUS,HD.PAIDNA,HD.CUSCDE,DP.AMOUNT from CMIS_OFF_HD HD INNER JOIN CMIS_DEPOSITDT DP ON DP.OR_NUM=HD.OR_NUM)T WHERE CUSCDE ='" & XXX & "'", gconDMIS, adOpenKeyset
    lvDeposits.ListItems.Clear
    If Not rsAppliedDep.EOF And Not rsAppliedDep.BOF Then
        picDeposits.Visible = True
        picDeposits.ZOrder 0
        cmdTranSave.Enabled = False
        Do While Not rsAppliedDep.EOF
            Set xList = lvDeposits.ListItems.Add(, , Null2String(rsAppliedDep!CUSCDE))
            xList.SubItems(1) = Null2Date(Format(rsAppliedDep!OR_DATE, "mm/dd/yyyy"))
            xList.SubItems(2) = rsAppliedDep!OR_NUM
            xList.SubItems(3) = ToDoubleNumber(NumericVal(rsAppliedDep!amount))
            'xList.SubItems(4) = N2Str2Null(rsAppliedDep!APPLIED)
            'xList.SubItems(5) = Null2String(rsAppliedDep!ID_Det)
            'xList.SubItems(6) = Null2String(rsAppliedDep!Id)
            'xList.SubItems(7) = SetPaidFor(Null2String(rsAppliedDep!PAIDFOR))
            rsAppliedDep.MoveNext
        Loop
    End If
    Set rsAppliedDep = Nothing
End Sub

Function GetChartCodes(XXX As String) As String
    Dim rsCODES As ADODB.Recordset
    Set rsCODES = New ADODB.Recordset
    rsCODES.Open "SELECT CHARTCODES from CMIS_SBOOK WHERE DESCNAME = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsCODES.EOF And Not rsCODES.BOF Then
        GetChartCodes = rsCODES!CHARTCODES
    End If
    Set rsCODES = Nothing
End Function

Private Sub chkDownPayment_Click()
    If chkDownPayment.Value = 1 Then
        txtDownPay.Enabled = True
        'JJE Sum of Tagged Vehicle Payment 03/20/2013 5:58PM
        Payment = NumericVal(Payment) + NumericVal(txtDownPay)
        'JJE
    Else
        txtDownPay.Enabled = False
        txtDownPay.BackColor = &HFFFFFF
        'JRE 06/14/16 Rollback the added amount in Tagged Vehicle Payment
        Payment = NumericVal(Payment) - NumericVal(txtDownPay)
        'JRE
    End If
End Sub

Private Sub chkLTORegFee_Click()
    If chkLTORegFee.Value = 1 Then
        txtLTOPay.Enabled = True
        'JJE Sum of Tagged Vehicle Payment 03/20/2013 5:58PM
        Payment = NumericVal(Payment) + NumericVal(txtLTOPay)
        'JJE
    Else
        txtLTOPay.Enabled = False
        txtLTOPay.BackColor = &HFFFFFF
        'JRE 06/14/16 Rollback the added amount in Tagged Vehicle Payment
        Payment = NumericVal(Payment) - NumericVal(txtLTOPay)
        'JRE
    End If
End Sub

Private Sub cmdCancelPayment_Click()
On Error GoTo ErrorCode:
    On_Update = False
    grdDetails.Enabled = True
    picORPayment.ZOrder 1: picORPayment.Visible = False
    picDetails.ZOrder 0:  picDetails.Visible = False
    fraDetails.Enabled = True
    Picture1.Enabled = True
   Call ClosePicOrpayment
    StoreMemVars
      Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub chkInsurance_Click()
    If chkInsurance.Value = 1 Then
        txtInsPay.Enabled = True
        'JJE Sum of Tagged Vehicle Payment 03/20/2013 5:58PM
        Payment = NumericVal(Payment) + NumericVal(txtInsPay)
        'JJE
    Else
        txtInsPay.Enabled = False
        txtInsPay.BackColor = &HFFFFFF
        'JRE 06/14/16 Rollback the added amount in Tagged Vehicle Payment
        Payment = NumericVal(Payment) - NumericVal(txtInsPay)
        'JRE
    End If
End Sub

Private Sub chkTPL_Click()
    If chkTPL.Value = 1 Then
        txtTPLPay.Enabled = True
        'JJE Sum of Tagged Vehicle Payment 03/20/2013 5:58PM
        Payment = NumericVal(Payment) + NumericVal(txtTPLPay)
        'JJE
    Else
        txtTPLPay.Enabled = False
        txtTPLPay.BackColor = &HFFFFFF
        'JRE 06/14/16 Rollback the added amount in Tagged Vehicle Payment
        Payment = NumericVal(Payment) - NumericVal(txtTPLPay)
        'JRE
    End If
End Sub

Private Sub chkChattel_Click()
    If chkChattel.Value = 1 Then
        txtChattelPay.Enabled = True
        'JJE Sum of Tagged Vehicle Payment 03/20/2013 5:58PM
        Payment = NumericVal(Payment) + NumericVal(txtChattelPay)
        'JJE
    Else
        txtChattelPay.Enabled = False
        txtChattelPay.BackColor = &HFFFFFF
        'JRE 06/14/16 Rollback the added amount in Tagged Vehicle Payment
        Payment = NumericVal(Payment) - NumericVal(txtChattelPay)
        'JRE
    End If
End Sub

Private Sub txtInsBal_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtInsPay_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtLTOAmount_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtLTOBal_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtLTOPay_Change()
    
    If NumericVal(txtLTOPay.Text) <= 0 Then
        wizDigit1.TextValue = 0
    Else
        If AddorEdit = "EDIT" Then
            
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtLTOPay.Text))
        Else
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtLTOPay.Text))
        End If
    End If
End Sub

Private Sub txtInsPay_Change()
    
    If NumericVal(txtInsPay.Text) <= 0 Then
        wizDigit1.TextValue = 0
    Else
        If AddorEdit = "EDIT" Then
            
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtInsPay.Text))
        Else
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtInsPay.Text))
        End If
    End If
End Sub

Private Sub txtLTOPay_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtOtherBal_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtTax_LostFocus()
    'JJE
    txtTax.Text = ToDoubleNumber(txtTax.Text)
    'JJE
End Sub

Private Sub txtTPLAmout_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtTPLPay_Change()
    
    If NumericVal(txtTPLPay.Text) <= 0 Then
        wizDigit1.TextValue = 0
    Else
        If AddorEdit = "EDIT" Then
            
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtTPLPay.Text))
        Else
            wizDigit1.TextValue = ToDoubleNumber(NumericVal(txtTPLPay.Text))
        End If
    End If
End Sub

Private Sub txtInsPay_GotFocus()
    txtInsPay.BackColor = &HC0FFFF
    If txtInsPay.Text <> "" Then
        txtInsPay.Text = NumericVal(txtInsPay.Text)
    End If
End Sub
Private Sub txtTPLPay_GotFocus()
txtTPLPay.BackColor = &HC0FFFF
    If txtTPLPay.Text <> "" Then
        txtTPLPay.Text = NumericVal(txtTPLPay.Text)
    End If
End Sub
Private Sub txtDownPay_GotFocus()
    txtDownPay.BackColor = &HC0FFFF
    If txtDownPay.Text <> "" Then
        txtDownPay.Text = NumericVal(txtDownPay.Text)
    End If
End Sub
Private Sub txtLTOPay_GotFocus()
txtLTOPay.BackColor = &HC0FFFF
    If txtLTOPay.Text <> "" Then
        txtLTOPay.Text = NumericVal(txtLTOPay.Text)
    End If
End Sub
Private Sub txtDownPay_LostFocus()
txtDownPay.BackColor = &HFFFFFF
txtDownPay.Text = ToDoubleNumber(txtDownPay.Text)
End Sub

Private Sub txtDownBAl_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
       KeyAscii = 0
End Sub

Private Sub txtTPLPay_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtTPLPay_LostFocus()
txtTPLPay.BackColor = &HFFFFFF
txtTPLPay.Text = ToDoubleNumber(txtTPLPay.Text)
End Sub
Private Sub txtLTOPay_LostFocus()
txtLTOPay.BackColor = &HFFFFFF
txtLTOPay.Text = ToDoubleNumber(txtLTOPay.Text)
End Sub

Private Sub txtInsPay_LostFocus()
txtInsPay.BackColor = &HFFFFFF
txtInsPay.Text = ToDoubleNumber(txtInsPay.Text)
End Sub


