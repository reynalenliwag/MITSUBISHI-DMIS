VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A06473E6-73D7-426E-82F2-6CD4F1FA4DBE}#1.0#0"; "wizMACBut.ocx"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmAMISCustomerAROpening 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOURNAL ENTRY"
   ClientHeight    =   4755
   ClientLeft      =   10170
   ClientTop       =   3885
   ClientWidth     =   9570
   ForeColor       =   &H00DEDFDE&
   Icon            =   "CustomerAROpening.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   9570
   Begin VB.PictureBox picBatchImport 
      Height          =   1755
      Left            =   2670
      ScaleHeight     =   1695
      ScaleWidth      =   4095
      TabIndex        =   214
      Top             =   1620
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton cmdBatchPosting 
         Caption         =   "Batch Post"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   216
         Top             =   900
         Width           =   2805
      End
      Begin VB.CommandButton cmdCloseRange 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   215
         Top             =   0
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   405
         Left            =   570
         TabIndex        =   217
         Top             =   420
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   131858433
         CurrentDate     =   40603
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   405
         Left            =   2520
         TabIndex        =   218
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   131858433
         CurrentDate     =   40603
      End
      Begin VB.Label lblVoucherNo 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
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
         Left            =   2370
         TabIndex        =   224
         Top             =   1410
         Width           =   1695
      End
      Begin VB.Label lblPosting 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         TabIndex        =   223
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label47 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   0
         TabIndex        =   222
         Top             =   1410
         Width           =   4125
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   315
         Left            =   0
         TabIndex        =   221
         Top             =   0
         Width           =   4155
         _Version        =   655364
         _ExtentX        =   7329
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Batch Post"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   8388608
         GradientColorDark=   8388608
         ForeColor       =   16777215
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   60
         TabIndex        =   220
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   2190
         TabIndex        =   219
         Top             =   480
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   -30
      ScaleHeight     =   900
      ScaleWidth      =   12195
      TabIndex        =   190
      Top             =   3810
      Width           =   12195
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
         Left            =   8505
         MouseIcon       =   "CustomerAROpening.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   202
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   765
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
         Left            =   7755
         MouseIcon       =   "CustomerAROpening.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   201
         ToolTipText     =   "Print this Record"
         Top             =   45
         Width           =   765
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
         Left            =   7005
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CustomerAROpening.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   200
         ToolTipText     =   "Cancel this Transaction"
         Top             =   45
         Width           =   765
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
         Left            =   6255
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CustomerAROpening.frx":16C6
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":1818
         Style           =   1  'Graphical
         TabIndex        =   199
         ToolTipText     =   "Unpost this Transaction"
         Top             =   45
         Width           =   765
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
         Left            =   5500
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "CustomerAROpening.frx":1B5D
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":1CAF
         Style           =   1  'Graphical
         TabIndex        =   198
         ToolTipText     =   "Post this Transaction"
         Top             =   45
         Width           =   765
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
         Left            =   4755
         MouseIcon       =   "CustomerAROpening.frx":1FD4
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":2126
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
         Width           =   765
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
         Left            =   4005
         MouseIcon       =   "CustomerAROpening.frx":2482
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":25D4
         Style           =   1  'Graphical
         TabIndex        =   196
         ToolTipText     =   "Add Record"
         Top             =   45
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
         Left            =   3255
         MouseIcon       =   "CustomerAROpening.frx":28E7
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":2A39
         Style           =   1  'Graphical
         TabIndex        =   195
         ToolTipText     =   "Move to Last Record"
         Top             =   45
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
         Left            =   2505
         MouseIcon       =   "CustomerAROpening.frx":2D89
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":2EDB
         Style           =   1  'Graphical
         TabIndex        =   194
         ToolTipText     =   "Move to First Record"
         Top             =   45
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
         Left            =   1755
         MouseIcon       =   "CustomerAROpening.frx":3239
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":338B
         Style           =   1  'Graphical
         TabIndex        =   193
         ToolTipText     =   "Find a Record"
         Top             =   45
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
         Left            =   1005
         MouseIcon       =   "CustomerAROpening.frx":3685
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":37D7
         Style           =   1  'Graphical
         TabIndex        =   192
         ToolTipText     =   "Move to Next Record"
         Top             =   45
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
         Left            =   255
         MouseIcon       =   "CustomerAROpening.frx":3B2F
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":3C81
         Style           =   1  'Graphical
         TabIndex        =   191
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   765
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7710
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   203
      Top             =   3840
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
         Left            =   765
         MouseIcon       =   "CustomerAROpening.frx":3FE0
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":4132
         Style           =   1  'Graphical
         TabIndex        =   205
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   765
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
         Left            =   15
         MouseIcon       =   "CustomerAROpening.frx":4470
         MousePointer    =   99  'Custom
         Picture         =   "CustomerAROpening.frx":45C2
         Style           =   1  'Graphical
         TabIndex        =   204
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   765
      End
   End
   Begin VB.PictureBox Frame1 
      BorderStyle     =   0  'None
      Height          =   3825
      Left            =   30
      ScaleHeight     =   3825
      ScaleWidth      =   9555
      TabIndex        =   38
      Top             =   0
      Width           =   9555
      Begin VB.PictureBox picReceivable 
         BorderStyle     =   0  'None
         Height          =   3405
         Left            =   -30
         ScaleHeight     =   3405
         ScaleWidth      =   9585
         TabIndex        =   130
         Top             =   450
         Width           =   9585
         Begin VB.PictureBox picDatePosted 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   90
            ScaleHeight     =   375
            ScaleWidth      =   2925
            TabIndex        =   212
            Top             =   2970
            Width           =   2925
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Date Posted:"
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
               Height          =   210
               Left            =   0
               TabIndex        =   128
               Top             =   60
               Width           =   1245
            End
            Begin VB.Label lblDatePosted 
               BackStyle       =   0  'Transparent
               Caption         =   "02/29/2008"
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
               Height          =   255
               Left            =   1320
               TabIndex        =   213
               Top             =   60
               Width           =   1500
            End
         End
         Begin VB.PictureBox picCustomerName 
            BorderStyle     =   0  'None
            Height          =   435
            Left            =   2490
            ScaleHeight     =   435
            ScaleWidth      =   4125
            TabIndex        =   209
            Top             =   0
            Width           =   4125
            Begin VB.TextBox txtCustName 
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
               Left            =   30
               Locked          =   -1  'True
               TabIndex        =   211
               Text            =   "000226"
               Top             =   45
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
               Left            =   3780
               TabIndex        =   210
               Top             =   45
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
            TabIndex        =   206
            Top             =   2580
            Width           =   6045
         End
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
            TabIndex        =   88
            Text            =   "000000"
            Top             =   1200
            Width           =   1485
         End
         Begin VB.CheckBox chkNonVat 
            Caption         =   "Non-Vat"
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
            Left            =   1140
            TabIndex        =   186
            Top             =   1200
            Width           =   1575
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
            TabIndex        =   93
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
            TabIndex        =   94
            Top             =   930
            Width           =   1755
         End
         Begin RichTextLib.RichTextBox txtRemarks2 
            Height          =   675
            Left            =   4560
            TabIndex        =   96
            Top             =   1350
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   1191
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"CustomerAROpening.frx":4912
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
            TabIndex        =   92
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
            TabIndex        =   91
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
            Left            =   60
            TabIndex        =   87
            Text            =   "Invoice Type"
            Top             =   780
            Width           =   2970
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
            TabIndex        =   86
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
            TabIndex        =   90
            Text            =   "0.00"
            Top             =   1980
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
            TabIndex        =   89
            Text            =   "88/88/8888"
            Top             =   1590
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
            Left            =   1470
            MaxLength       =   6
            TabIndex        =   85
            Text            =   "000226"
            Top             =   45
            Width           =   1005
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
            TabIndex        =   95
            Text            =   "Invoice Type"
            Top             =   930
            Width           =   4920
         End
         Begin RichTextLib.RichTextBox txtCOBAcctNo 
            Height          =   315
            Left            =   7620
            TabIndex        =   208
            Top             =   2580
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            TextRTF         =   $"CustomerAROpening.frx":49A9
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
         Begin VB.Line Line6 
            X1              =   0
            X2              =   9570
            Y1              =   2490
            Y2              =   2490
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
            Top             =   2610
            Width           =   1545
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
            TabIndex        =   144
            Top             =   960
            Width           =   1335
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   9570
            Y1              =   450
            Y2              =   450
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
            TabIndex        =   171
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
            TabIndex        =   170
            Top             =   60
            Width           =   2775
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
            TabIndex        =   143
            Top             =   570
            Width           =   1335
         End
         Begin VB.Line Line5 
            X1              =   3090
            X2              =   3090
            Y1              =   450
            Y2              =   2370
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
            TabIndex        =   142
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
            Left            =   60
            TabIndex        =   136
            Top             =   510
            Width           =   1425
         End
         Begin VB.Line Line3 
            X1              =   6660
            X2              =   6660
            Y1              =   450
            Y2              =   -30
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
            Left            =   90
            TabIndex        =   135
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
            TabIndex        =   134
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
            TabIndex        =   133
            Top             =   2010
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
            TabIndex        =   132
            Top             =   1620
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
            TabIndex        =   131
            Top             =   1230
            Width           =   1425
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
            TabIndex        =   172
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.PictureBox picPayables 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   30
         ScaleHeight     =   1275
         ScaleWidth      =   9465
         TabIndex        =   52
         Top             =   1230
         Width           =   9465
         Begin VB.TextBox txtTaxBaseAmount 
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
            Left            =   6210
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   60
            Width           =   1665
         End
         Begin VB.TextBox txtPayCode 
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
            MaxLength       =   3
            TabIndex        =   7
            Text            =   "000226"
            Top             =   60
            Width           =   495
         End
         Begin VB.TextBox txtAmountToPay 
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
            Text            =   "0.00"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtInvoiceDate 
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
            TabIndex        =   9
            Text            =   "88/88/8888"
            Top             =   450
            Width           =   1695
         End
         Begin VB.ComboBox cboPayType 
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
            Left            =   1950
            TabIndex        =   8
            Text            =   "Cash Payment"
            Top             =   60
            Width           =   2325
         End
         Begin RichTextLib.RichTextBox txtRemarks 
            Height          =   765
            Left            =   4320
            TabIndex        =   12
            Top             =   420
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   1349
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"CustomerAROpening.frx":4A3C
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
         Begin VB.Label Label42 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "TAX BASE AMOUNT"
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
            Left            =   4380
            TabIndex        =   188
            Top             =   90
            Width           =   1845
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
            Left            =   3210
            TabIndex        =   56
            Top             =   450
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
            TabIndex        =   55
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
            TabIndex        =   54
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
            TabIndex        =   53
            Top             =   90
            Width           =   1725
         End
      End
      Begin VB.TextBox txtCode 
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
         TabIndex        =   2
         Text            =   "000226"
         Top             =   460
         Width           =   1005
      End
      Begin VB.TextBox txtJDate 
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
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   60
         Width           =   1545
      End
      Begin VB.TextBox txtVoucherNo 
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
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2490
         TabIndex        =   3
         Text            =   "cboRecvd_Desc"
         Top             =   450
         Width           =   4080
      End
      Begin VB.TextBox txtDueDate 
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
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox txtJNo 
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
         Left            =   7920
         MaxLength       =   6
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.ComboBox cboATCTAXRATE 
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
         TabIndex        =   4
         Text            =   "cboATCTAXRATE"
         Top             =   840
         Width           =   990
      End
      Begin VB.PictureBox picDisbursement 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Left            =   30
         ScaleHeight     =   1245
         ScaleWidth      =   9525
         TabIndex        =   47
         Top             =   2610
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Text            =   "000226"
            Top             =   30
            Width           =   1815
         End
         Begin RichTextLib.RichTextBox txtParticulars 
            Height          =   735
            Left            =   4380
            TabIndex        =   17
            Top             =   420
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   1296
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"CustomerAROpening.frx":4AD3
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
            TabIndex        =   119
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
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
            TabIndex        =   48
            Top             =   60
            Width           =   1935
         End
      End
      Begin VB.Label labPosted 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2490
         TabIndex        =   57
         Top             =   60
         Width           =   4065
      End
      Begin VB.Line Line2 
         X1              =   6630
         X2              =   6630
         Y1              =   1200
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   9540
         Y1              =   1200
         Y2              =   1200
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         Left            =   6480
         TabIndex        =   43
         Top             =   2490
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
         Left            =   2550
         TabIndex        =   42
         Top             =   510
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
         Top             =   870
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label txtAddress 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2490
         TabIndex        =   5
         Top             =   840
         Width           =   4065
      End
      Begin VB.Label Label41 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ATC TAX RATE"
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
         TabIndex        =   187
         Top             =   900
         Width           =   1425
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   -60
      ScaleHeight     =   1095
      ScaleWidth      =   9945
      TabIndex        =   189
      Top             =   3750
      Width           =   9945
   End
   Begin VB.PictureBox picRefCDJ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6735
      ScaleHeight     =   345
      ScaleWidth      =   2775
      TabIndex        =   184
      Top             =   930
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label RefCDJ 
         Alignment       =   2  'Center
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
         Height          =   345
         Left            =   0
         TabIndex        =   185
         Top             =   0
         Width           =   2775
      End
   End
   Begin Crystal.CrystalReport rptAP 
      Left            =   8940
      Top             =   5700
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
   Begin TabDlg.SSTab JournalTAB 
      Height          =   2775
      Left            =   30
      TabIndex        =   35
      Top             =   2580
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   4895
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "[<F3> Add &Journals]   [<Ctrl> + <J> View &Journals]   "
      TabPicture(0)   =   "CustomerAROpening.frx":4B67
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdAddJournal"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAddJournal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDetails"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "[<F4> Add &Details]   [<Ctrl> + <D> View &Details]   "
      TabPicture(1)   =   "CustomerAROpening.frx":4B83
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picPV_Detail"
      Tab(1).Control(1)=   "picPV_Entry"
      Tab(1).ControlCount=   2
      Begin VB.PictureBox picPV_Detail 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   -74940
         ScaleHeight     =   2265
         ScaleWidth      =   9375
         TabIndex        =   59
         Top             =   90
         Width           =   9405
         Begin MSComctlLib.ListView lstPV_Detail 
            Height          =   1785
            Left            =   30
            TabIndex        =   126
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
            MouseIcon       =   "CustomerAROpening.frx":4B9F
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
            Left            =   7620
            TabIndex        =   60
            Top             =   1860
            Width           =   1350
            _ExtentX        =   2381
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
            Left            =   7020
            TabIndex        =   100
            Top             =   1920
            Width           =   1275
         End
      End
      Begin VB.PictureBox picPV_Entry 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   -74820
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   71
         Top             =   600
         Width           =   9135
         Begin VB.CommandButton cmdPVCancel 
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
            Left            =   8040
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CustomerAROpening.frx":4D01
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":4E53
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   690
            Width           =   1005
         End
         Begin VB.CommandButton cmdPVSave 
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
            Left            =   7020
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CustomerAROpening.frx":5165
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":52B7
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   690
            Width           =   975
         End
         Begin VB.CommandButton cmdPVDelete 
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
            MouseIcon       =   "CustomerAROpening.frx":56F9
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":584B
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   720
            Width           =   1005
         End
         Begin MSMask.MaskEdBox txtMRR_No 
            Height          =   315
            Left            =   1950
            TabIndex        =   28
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
            TabIndex        =   31
            Top             =   330
            Width           =   1425
            _ExtentX        =   2514
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
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   315
            Left            =   7140
            TabIndex        =   72
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
            TabIndex        =   29
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
            TabIndex        =   27
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
            TabIndex        =   30
            Top             =   330
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
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
            TabIndex        =   73
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
            TabIndex        =   99
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
            TabIndex        =   98
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
            TabIndex        =   97
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
            TabIndex        =   84
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
            TabIndex        =   75
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
            TabIndex        =   74
            Top             =   390
            Width           =   1305
         End
      End
      Begin VB.PictureBox fraDetails 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   60
         ScaleHeight     =   2265
         ScaleWidth      =   9375
         TabIndex        =   58
         Top             =   90
         Width           =   9405
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   30
            Top             =   2280
         End
         Begin VB.Frame Frame3 
            Appearance      =   0  'Flat
            BackColor       =   &H00DEDFDE&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   0
            TabIndex        =   137
            Top             =   1830
            Width           =   9135
            Begin VB.PictureBox picChat 
               BackColor       =   &H00DEDFDE&
               Height          =   345
               Left            =   60
               ScaleHeight     =   285
               ScaleWidth      =   5835
               TabIndex        =   182
               Top             =   30
               Visible         =   0   'False
               Width           =   5895
               Begin VB.Label Label40 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "Warning: Sales Details Amount is not Balance with Journal Details Amount"
                  ForeColor       =   &H000000FF&
                  Height          =   255
                  Left            =   60
                  TabIndex        =   183
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
               TabIndex        =   141
               Text            =   "Text1"
               Top             =   30
               Width           =   1515
            End
            Begin VB.TextBox txtTotDebit 
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
               Left            =   6000
               MaxLength       =   14
               TabIndex        =   139
               Text            =   "Text1"
               Top             =   30
               Width           =   1485
            End
            Begin VB.TextBox txtTotCredit 
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
               Left            =   7470
               MaxLength       =   14
               TabIndex        =   138
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
               TabIndex        =   140
               Top             =   60
               Width           =   1275
            End
         End
         Begin MSComctlLib.ListView lstDetails 
            Height          =   1785
            Left            =   30
            TabIndex        =   125
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
            MouseIcon       =   "CustomerAROpening.frx":5B55
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
         Left            =   180
         ScaleHeight     =   1635
         ScaleWidth      =   9105
         TabIndex        =   61
         Top             =   600
         Width           =   9135
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
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   330
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtCredit 
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
            Left            =   7950
            MaxLength       =   10
            TabIndex        =   23
            Top             =   330
            Width           =   1100
         End
         Begin VB.TextBox txtDebit 
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
            Left            =   6780
            MaxLength       =   10
            TabIndex        =   22
            Top             =   330
            Width           =   1100
         End
         Begin VB.Frame fraComp 
            Height          =   915
            Left            =   2340
            TabIndex        =   145
            Top             =   660
            Width           =   4365
            Begin VB.TextBox txtNetAmt 
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
               Left            =   2910
               MaxLength       =   10
               TabIndex        =   21
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtTax 
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
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   20
               Top             =   510
               Width           =   1300
            End
            Begin VB.TextBox txtGrossAmt 
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
               Left            =   150
               MaxLength       =   10
               TabIndex        =   19
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
               TabIndex        =   148
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
               TabIndex        =   147
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
               TabIndex        =   146
               Top             =   240
               Width           =   1365
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   735
            Left            =   2310
            TabIndex        =   116
            Top             =   -30
            Width           =   4425
            Begin RichTextLib.RichTextBox txtAcct_Name 
               Height          =   315
               Left            =   30
               TabIndex        =   117
               Top             =   360
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393217
               BackColor       =   16777215
               Enabled         =   -1  'True
               MultiLine       =   0   'False
               Appearance      =   0
               TextRTF         =   $"CustomerAROpening.frx":5CB7
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
               TabIndex        =   118
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
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin VB.CommandButton cmdJournalDelete 
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
            MouseIcon       =   "CustomerAROpening.frx":5D4A
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":5E9C
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   720
            Width           =   1005
         End
         Begin VB.CommandButton cmdJournalSave 
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
            Left            =   7020
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CustomerAROpening.frx":61A6
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":62F8
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   690
            Width           =   975
         End
         Begin VB.CommandButton cmdJournalCancel 
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
            Left            =   8040
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "CustomerAROpening.frx":673A
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":688C
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   690
            Width           =   1005
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
            TabIndex        =   62
            Text            =   "Text1"
            Top             =   330
            Width           =   585
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
            TabIndex        =   70
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
            TabIndex        =   69
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
            TabIndex        =   68
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
            Top             =   420
            Width           =   2685
         End
      End
      Begin wizButton.cmd cmdAddJournal 
         Height          =   1785
         Left            =   120
         TabIndex        =   127
         Top             =   540
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3149
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
         MICON           =   "CustomerAROpening.frx":6B9E
      End
   End
   Begin wizButton.cmd cmdFindAccount 
      Height          =   5415
      Left            =   30
      TabIndex        =   120
      Top             =   150
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   9551
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
      MICON           =   "CustomerAROpening.frx":6BBA
   End
   Begin VB.Frame fraFindAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
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
      ForeColor       =   &H00FF0000&
      Height          =   5265
      Left            =   120
      TabIndex        =   44
      Top             =   210
      Width           =   9375
      Begin VB.TextBox txtSearch 
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
         Left            =   90
         MaxLength       =   50
         TabIndex        =   121
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
         TabIndex        =   36
         Top             =   3960
         Visible         =   0   'False
         Width           =   45
      End
      Begin MSComctlLib.ListView lstAccounts 
         Height          =   3975
         Left            =   60
         TabIndex        =   123
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
         MouseIcon       =   "CustomerAROpening.frx":6BD6
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
         TabIndex        =   155
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
         TabIndex        =   124
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         TabIndex        =   122
         Top             =   4650
         Width           =   9225
      End
   End
   Begin wizButton.cmd cmdTemplates 
      Height          =   4245
      Left            =   1140
      TabIndex        =   157
      Top             =   900
      Width           =   7335
      _ExtentX        =   12938
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
      MICON           =   "CustomerAROpening.frx":6D38
   End
   Begin VB.PictureBox picTemplates 
      Height          =   4125
      Left            =   1200
      ScaleHeight     =   4065
      ScaleWidth      =   7125
      TabIndex        =   156
      Top             =   960
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
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   60
         Width           =   6975
      End
      Begin MSComctlLib.ListView lstTemplates 
         Height          =   3165
         Left            =   30
         TabIndex        =   158
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
         MouseIcon       =   "CustomerAROpening.frx":6D54
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
         TabIndex        =   160
         Top             =   3750
         Width           =   7035
      End
   End
   Begin VB.CommandButton cmdPrinting 
      BackColor       =   &H00DEDFDE&
      Caption         =   "Command1"
      Height          =   2445
      Left            =   3420
      TabIndex        =   173
      Top             =   1800
      Width           =   2775
   End
   Begin VB.PictureBox picPrinting 
      BackColor       =   &H00DEDFDE&
      Height          =   2265
      Left            =   3510
      ScaleHeight     =   2205
      ScaleWidth      =   2535
      TabIndex        =   174
      Top             =   1890
      Width           =   2595
      Begin VB.PictureBox picPrintCheck 
         BackColor       =   &H00DEDFDE&
         Enabled         =   0   'False
         Height          =   885
         Left            =   60
         ScaleHeight     =   825
         ScaleWidth      =   2355
         TabIndex        =   181
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
            TabIndex        =   176
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
            TabIndex        =   177
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
            TabIndex        =   178
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
         TabIndex        =   179
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
         TabIndex        =   175
         Top             =   60
         Width           =   2415
      End
      Begin wizMacBut.MacBut cmdOkPrint 
         Height          =   345
         Left            =   360
         TabIndex        =   180
         Top             =   1800
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         Caption         =   "OK"
         Caption_Xpos    =   700
      End
   End
   Begin wizButton.cmd cmdShowPostRange 
      Height          =   2175
      Left            =   3480
      TabIndex        =   161
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3836
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
      MICON           =   "CustomerAROpening.frx":6EB6
   End
   Begin VB.PictureBox picShowPostRange 
      Height          =   2055
      Left            =   3540
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   162
      Top             =   1740
      Width           =   2535
      Begin wizProgBar.Prg prgPostRange 
         Height          =   285
         Left            =   90
         TabIndex        =   168
         Top             =   1650
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         Picture         =   "CustomerAROpening.frx":6ED2
         ForeColor       =   0
         BarPicture      =   "CustomerAROpening.frx":6EEE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin wizMacBut.MacBut cmdPostRange 
         Height          =   345
         Left            =   390
         TabIndex        =   165
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
         TabIndex        =   164
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
         TabIndex        =   163
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
         TabIndex        =   169
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
         TabIndex        =   167
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
         TabIndex        =   166
         Top             =   420
         Width           =   735
      End
   End
   Begin VB.PictureBox picGJ 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   30
      ScaleHeight     =   4845
      ScaleWidth      =   9525
      TabIndex        =   101
      Top             =   510
      Width           =   9555
      Begin MSComctlLib.ListView lstGJ 
         Height          =   3315
         Left            =   60
         TabIndex        =   106
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
         MouseIcon       =   "CustomerAROpening.frx":6F0A
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
         Appearance      =   0  'Flat
         BackColor       =   &H00DEDFDE&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   90
         TabIndex        =   150
         Top             =   4410
         Width           =   9135
         Begin VB.TextBox txtGJTotCredit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   153
            Text            =   "Text1"
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtGJTotDebit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   152
            Text            =   "Text1"
            Top             =   30
            Width           =   1515
         End
         Begin VB.TextBox txtGJOutBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   151
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
            TabIndex        =   154
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
         TabIndex        =   102
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
            MouseIcon       =   "CustomerAROpening.frx":706C
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":71BE
            Style           =   1  'Graphical
            TabIndex        =   81
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
            MouseIcon       =   "CustomerAROpening.frx":74D0
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":7622
            Style           =   1  'Graphical
            TabIndex        =   80
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
            MouseIcon       =   "CustomerAROpening.frx":7A64
            MousePointer    =   99  'Custom
            Picture         =   "CustomerAROpening.frx":7BB6
            Style           =   1  'Graphical
            TabIndex        =   83
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
            TabIndex        =   76
            Text            =   "Combo1"
            Top             =   330
            Width           =   2235
         End
         Begin RichTextLib.RichTextBox txtGJAccountName 
            Height          =   315
            Left            =   2340
            TabIndex        =   77
            Top             =   330
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   556
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            MultiLine       =   0   'False
            Appearance      =   0
            TextRTF         =   $"CustomerAROpening.frx":7EC0
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
            TabIndex        =   78
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
            TabIndex        =   79
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
            TabIndex        =   103
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
            TabIndex        =   104
            Text            =   "Text1"
            Top             =   330
            Width           =   855
         End
         Begin RichTextLib.RichTextBox txtGJAccountParticulars 
            Height          =   885
            Left            =   2340
            TabIndex        =   82
            Top             =   690
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   1561
            _Version        =   393217
            BackColor       =   16777215
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"CustomerAROpening.frx":7F53
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   109
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
            TabIndex        =   108
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
            TabIndex        =   107
            Top             =   60
            Width           =   2205
         End
      End
      Begin wizButton.cmd cmdGJEntry 
         Height          =   1785
         Left            =   60
         TabIndex        =   149
         Top             =   2580
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   3149
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
         MICON           =   "CustomerAROpening.frx":7FEA
      End
      Begin RichTextLib.RichTextBox txtParticulars2 
         Height          =   705
         Left            =   60
         TabIndex        =   105
         Top             =   330
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   1244
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"CustomerAROpening.frx":8006
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
         TabIndex        =   129
         Top             =   60
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmAMISCustomerAROpening"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsJournal_HD                                            As ADODB.Recordset
Dim rsJournal_Det                                           As ADODB.Recordset
Dim rsCustomer                                              As ADODB.Recordset
Dim rsCRJ_Detail                                            As ADODB.Recordset
Dim LocalAcess                                              As String
Attribute LocalAcess.VB_VarUserMemId = 1073938463
Dim kcnt                                                    As Integer
Dim TOTDEBIT                                                As Double
Dim TOTCREDIT                                               As Double
Dim TOTTAX                                                  As Double
Dim OUTBALANCE                                              As Double
Dim COMP_SJ_OUTPUT_TAX                                      As Double
Dim TOTAL_AR_AMOUNT                                         As Double
Dim AddorEdit                                               As String
Dim PrevJType                                               As String
Dim PrevJNo                                                 As String
Dim WithEvents frmNewEntity                                 As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim xEntityClass                                            As String
Dim xJOURNALTYPE                                            As String
Dim BatchPost                                               As Boolean

Function SetCOBAcctNo(XXX As String) As String
    Dim rsCOBAcctName                                       As ADODB.Recordset
    Set rsCOBAcctName = New ADODB.Recordset
    Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where Description = '" & XXX & "'")
    If Not rsCOBAcctName.EOF And Not rsCOBAcctName.BOF Then
        SetCOBAcctNo = Null2String(rsCOBAcctName!AcctCode)
    End If
End Function

Function SetPayCode(VVV As Variant)
    Dim rsPayTerm                                           As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayCode = Null2String(rsPayTerm!pay_Code)
    Else
        SetPayCode = ""
    End If
    Set rsPayTerm = Nothing
End Function

Function SetPayDesc(VVV As Variant) As String
    Dim rsPayTerm                                           As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_code,pay_desc from ALL_PayTerm where pay_code = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayDesc = Null2String(rsPayTerm!pay_desc)
    Else
        SetPayDesc = ""
    End If
    Set rsPayTerm = Nothing
End Function

Function SetPayNoDays(VVV As Variant) As Integer
    Dim rsPayTerm                                           As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_Desc,no_days from ALL_PayTerm where pay_Desc = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayNoDays = Null2String(rsPayTerm!no_Days)
    Else
        SetPayNoDays = 0
    End If
    Set rsPayTerm = Nothing
End Function

Function SetInvType(INV As Variant)
    Dim rsInvoiceType                                       As ADODB.Recordset
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invcode = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvType = Null2String(rsInvoiceType!INVTYPE)
    Else
        SetInvType = ""
    End If
    Set rsInvoiceType = Nothing
End Function

Function SetInvCode(INV As Variant)
    Dim rsInvoiceType                                       As ADODB.Recordset
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "Select invcode,invtype from ALL_InvoiceType where invtype = " & N2Str2Null(INV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        SetInvCode = Null2String(rsInvoiceType!InvCode)
    Else
        SetInvCode = ""
    End If
    Set rsInvoiceType = Nothing
End Function

Function SetCustomerCode(CCC As Variant)
    Dim rsCustomer                                          As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,custname from ALL_CustMaster_Amis where custname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    'rsCustomer.Open "Select Code,AccountName from ALL_Entity where AccountName = " & N2Str2Null(CCC) & " AND ENTITYCODE = " & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCode = Null2String(rsCustomer!CUSTCODE)
    Else
        SetCustomerCode = ""
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerCodeNew(CCC As Variant, XXX As Variant)
    Dim rsCustomer                                          As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    'rsCustomer.Open "Select custcode,custname from ALL_CustMaster_Amis where custname = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    rsCustomer.Open "Select Code,AccountName from ALL_Entity where AccountName = " & N2Str2Null(CCC) & " AND ENTITYCODE = " & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerCodeNew = Null2String(rsCustomer!CODE)
    Else
        SetCustomerCodeNew = ""
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerName(CCC As Variant)
    Dim rsCustomer                                          As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select custcode,custname from ALL_CustMaster_Amis where CustCode = " & N2Str2Null(CCC), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerName = Null2String(rsCustomer!CUSTNAME)
    Else
        SetCustomerName = ""
    End If
    Set rsCustomer = Nothing
End Function

Function SetCustomerNameNew(CCC As Variant, XXX As Variant)
    Dim rsCustomer                                          As ADODB.Recordset
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "Select Code,AccountName from ALL_Entity where Code = " & N2Str2Null(CCC) & " AND ENTITYCODE=" & N2Str2Null(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerNameNew = Null2String(rsCustomer!ACCOUNTNAME)
    Else
        SetCustomerNameNew = ""
    End If
    Set rsCustomer = Nothing
End Function

Sub FillCboAcctName()
    Dim rsCOBAcctName                                       As ADODB.Recordset
    Set rsCOBAcctName = New ADODB.Recordset
    '    If COMPANY_CODE = "HLP" Or COMPANY_CODE = "HAM" Or COMPANY_CODE = "HSP" Or COMPANY_CODE = "DGI" Then
    '        Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where (Titles IN ('1102','2102','2107','2101','2102','2104','2106','2107') AND IS_SCHEDULE_ACCNT=1) Order by AcctCode asc")
    '    Else
    Set rsCOBAcctName = gconDMIS.Execute("Select * from AMIS_ChartAccount Where (Titles IN ('1102','1103','1104') AND IS_SCHEDULE_ACCNT=1) Order by AcctCode asc")
    '    End If
    If Not rsCOBAcctName.EOF And Not rsCOBAcctName.BOF Then
        rsCOBAcctName.MoveFirst: cboCOBAcctName.Clear
        Do While Not rsCOBAcctName.EOF
            cboCOBAcctName.AddItem Null2String(rsCOBAcctName!DESCRIPTION)
            rsCOBAcctName.MoveNext
        Loop
    End If
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

Sub FindDupJNo(DDD As String)
    rsJournal_HD.Bookmark = rsFind(rsJournal_HD.Clone, "jno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Sub rsRefresh()
    Set rsJournal_HD = New ADODB.Recordset
    rsJournal_HD.Open "select * from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by VoucherNo asc", gconDMIS, adOpenKeyset
End Sub

Sub initMemvars()
    Call Get_VoucherNo

    Dim rsJDate                                             As ADODB.Recordset
    Set rsJDate = New ADODB.Recordset
    rsJDate.Open "Select * from ALL_PROFILE", gconDMIS, adOpenKeyset
    If Not rsJDate.EOF And Not rsJDate.BOF Then
        txtJDate.Text = Null2Date(rsJDate!Cut_Off_Date)
    End If
    picDatePosted.Visible = False
    RefCRJ.Visible = False
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
    cboNameofVendor.Text = ""
    txtTotDebit.Text = ZERO: txtTotCredit.Text = ZERO
    txtAmountToPay.Text = ZERO: txtOutBalance.Text = ZERO
    txtParticulars2.Locked = False
    txtParticulars.Text = "Pls Type Your Message Here!"
    txtParticulars2.Text = "Pls Type Your Message Here!"
    '---------------------------'
    'Accounts Receivable Module'
    txtCustCode.Text = ""
    txtCustName.Text = ""
    cboCustName.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate2.Text = LOGDATE
    txtInvoiceAmt.Text = ZERO
    txtRefNo.Text = ""
    txtRefDate.Text = ""
    txtRemarks2.Text = "Pls Type Your Message Here!"

    If COMPANY_CODE = "HMH" Or COMPANY_CODE = "DGI" Or COMPANY_CODE = "HCA" Or COMPANY_CODE = "HQA" Then
        picCustomerName.Visible = True
    Else
        picCustomerName.Visible = False
    End If
    '---------------------------'

    txtTotalPV_Amount.Text = ZERO
    labPosted.Caption = ""
    labPosted.Visible = False
    labOutBalance.Visible = False
    txtOutBalance.Visible = False
    'initGrid
    SendToBack

    txtCOBAcctNo.Text = ""
    cboCOBAcctName.Text = ""
End Sub

Private Sub Get_VoucherNo()
    Dim rsJournal_HDDup                                     As ADODB.Recordset
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select voucherno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' order by voucherno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtVoucherNo.Text = Format(N2Str2Zero(rsJournal_HDDup!VOUCHERNO) + 1, "000000") Else txtVoucherNo.Text = "000001"
    Set rsJournal_HDDup = New ADODB.Recordset
    Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
    If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
End Sub

Sub StoreSearch(XXX As Variant)
    rsRefresh
    rsJournal_HD.Find "VoucherNo = " & N2Str2Null(XXX)
    StoreMemVars
End Sub

Sub StoreMemVars()
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        labID.Caption = rsJournal_HD!ID
        txtJNo.Text = Null2String(rsJournal_HD!JNo)
        txtVoucherNo.Text = Null2String(rsJournal_HD!VOUCHERNO)
        txtJDate.Text = Format(Null2String(rsJournal_HD!JDATE), "DD-MMM-YY")
        txtInvoiceDate.Text = Format(Null2String(rsJournal_HD!invoicedate), "DD-MMM-YY")
        txtDueDate.Text = Format(Null2String(rsJournal_HD!DUEDATE), "DD-MMM-YY")
        txtPayCode.Text = Null2String(rsJournal_HD!paytype)
        txtTerms.Text = Null2String(rsJournal_HD!TERMS)
        cboPayType.Text = SetPayDesc(Null2String(rsJournal_HD!paytype))
        CURRENT_CUSCODE = Null2String(rsJournal_HD!CustomerCode)
        txtCustCode.Text = Null2String(rsJournal_HD!CustomerCode)
        cboCustName.Text = SetCustomerName(Null2String(rsJournal_HD!CustomerCode))
        txtCustName.Text = SetCustomerNameNew(Null2String(rsJournal_HD!CustomerCode), Null2String(rsJournal_HD!Entity_Class))
        'cboCustName.Text = Null2String(rsJournal_HD!CustomerName)
        cboInvoiceType.Text = SetInvType(Null2String(rsJournal_HD!InvoiceType))
        xEntityClass = Null2String(rsJournal_HD!Entity_Class)
        If Left(Null2String(rsJournal_HD!INVOICENO), 2) = "NV" Then
            chkNonVat.Value = 1
            txtInvoiceNo.Text = Right(Null2String(rsJournal_HD!INVOICENO), 6)
        Else
            chkNonVat.Value = 0
            txtInvoiceNo.Text = Null2String(rsJournal_HD!INVOICENO)
        End If
        txtInvoiceDate2.Text = Null2String(rsJournal_HD!invoicedate)
        txtInvoiceAmt.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!InvoiceAmt))
        'cboBankName2.Text = SetBankName(Null2String(rsJOURNAL_HD!BankCode))
        If CheckStatus(txtVoucherNo, xJOURNALTYPE) = True Then
            Set rsCRJ_Detail = New ADODB.Recordset
            SQL_STATEMENT = "Select INVOICETYPE,INVOICENO,ISNULL(CR_TYPE,'')+ '-' +ISNULL(VOUCHERNO,'') AS VOUCHERNO from AMIS_CRJ_Detail Where invoicetype = '" & SetInvCode(cboInvoiceType.Text) & "' AND InvoiceNo = " & N2Str2Null(txtInvoiceNo.Text) & "" & vbCrLf
            SQL_STATEMENT = SQL_STATEMENT & "UNION" & vbCrLf
            SQL_STATEMENT = SQL_STATEMENT & "Select INVOICETYPE,INVOICENO,ISNULL(CV_JTYPE,'')+'-'+ISNULL(VOUCHERNO,'') AS VOUCHERNO from AMIS_CV_Detail Where invoicetype = '" & SetInvCode(cboInvoiceType.Text) & "' AND InvoiceNo = " & N2Str2Null(txtInvoiceNo.Text) & ""
            Set rsCRJ_Detail = gconDMIS.Execute(SQL_STATEMENT)
            If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
                RefCRJ.Visible = True
                RefCRJ.BorderStyle = 1: RefCRJ.Caption = "Ref # " & Null2String(rsCRJ_Detail!VOUCHERNO)
            Else
                RefCRJ.BorderStyle = 0: RefCRJ.Caption = ""
            End If
        Else
            RefCRJ.BorderStyle = 0: RefCRJ.Caption = ""
        End If
        txtRefNo.Text = Null2String(rsJournal_HD!refno)
        txtRefDate.Text = Null2String(rsJournal_HD!RefDate)
        txtBankCode.Text = Null2String(rsJournal_HD!bankcode)
        txtCheckNo.Text = Null2String(rsJournal_HD!CheckNo)
        txtCheckDate.Text = Null2String(rsJournal_HD!CheckDate)
        txtParticulars.Text = Null2String(rsJournal_HD!remarks)
        txtParticulars2.Text = Null2String(rsJournal_HD!remarks)
        txtTotDebit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!Debit))
        txtTotCredit.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!Credit))
        txtOutBalance.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!OUTBALANCE))
        txtAmountToPay.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD!AMOUNTTOPAY))
        txtRemarks.Text = Null2String(rsJournal_HD!remarks)
        txtRemarks2.Text = Null2String(rsJournal_HD!remarks)
        If Null2String(rsJournal_HD!Status) = "C" Then
            labPosted.Visible = True: labPosted.Caption = "*** CANCELLED *** [" & Null2String(rsJournal_HD!USERCODE) & "]"
            cmdEdit.Enabled = False: cmdCancelCO.Enabled = False: cmdPost.Enabled = False
            cmdUnPost.Enabled = False: cmdPrint.Enabled = False
            picDatePosted.Visible = False
        ElseIf Null2String(rsJournal_HD!Status) = "P" Then
            labPosted.Visible = True: labPosted.Caption = "*** POSTED *** [" & Null2String(rsJournal_HD!USERCODE) & "]"
            cmdEdit.Enabled = False: cmdPost.Enabled = False
            cmdCancelCO.Enabled = False: cmdUnPost.Enabled = True
            picDatePosted.Visible = True
            If Null2String(rsJournal_HD!DATEPOSTED) = "" Then
                lblDatePosted = Format(rsJournal_HD!JDATE, "DD-MMM-YY")
            Else
                lblDatePosted = Format(rsJournal_HD!DATEPOSTED, "DD-MMM-YY")
            End If
        Else
            labPosted.Caption = "": labPosted.Visible = False
            cmdEdit.Enabled = True: cmdUnPost.Enabled = False
            cmdCancelCO.Enabled = True: cmdPost.Enabled = True
            cmdPrint.Enabled = False
            picDatePosted.Visible = False
        End If
        FillDetails
    Else
        MsgBox "No Such Record!": If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillDetails()
    kcnt = 0: TOTDEBIT = 0: TOTCREDIT = 0: TOTTAX = 0: OUTBALANCE = 0: COMP_SJ_OUTPUT_TAX = 0: TOTAL_AR_AMOUNT = 0
    txtTotDebit.Text = TOTDEBIT: txtTotCredit.Text = TOTCREDIT: txtOutBalance.Text = OUTBALANCE
    Set rsJournal_Det = New ADODB.Recordset
    Set rsJournal_Det = gconDMIS.Execute("select id,jtype,jno,jitemno,acct_code,acct_name,debit,credit,tax from AMIS_Journal_Det where jno = " & N2Str2Null(txtJNo.Text) & " and jtype = '" & xJOURNALTYPE & "' order by jitemno asc")
    If Not rsJournal_Det.EOF And Not rsJournal_Det.BOF Then
        txtCOBAcctNo.Text = Null2String(rsJournal_Det!ACCT_CODE)
        cboCOBAcctName.Text = Null2String(rsJournal_Det!acct_Name)
    Else
        txtCOBAcctNo.Text = ""
        cboCOBAcctName.Text = ""
        cmdPost.Enabled = False
    End If
End Sub

Sub InitCbo()
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "select custname from ALL_CustMaster_Amis order by custname asc", gconDMIS
    'rsCustomer.Open "select AccountName from ALL_Entity order by AccountName asc", gconDMIS
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        rsCustomer.MoveFirst
        cboCustName.Clear
        Do While Not rsCustomer.EOF
            cboCustName.AddItem Null2String(rsCustomer!CUSTNAME)
            rsCustomer.MoveNext
        Loop
    End If

    Dim rsInvoiceType                                       As ADODB.Recordset
    Set rsInvoiceType = New ADODB.Recordset
    rsInvoiceType.Open "select InvType from ALL_InvoiceType order by InvType asc", gconDMIS
    If Not rsInvoiceType.EOF And Not rsInvoiceType.BOF Then
        rsInvoiceType.MoveFirst
        cboInvoiceType.Clear
        Do While Not rsInvoiceType.EOF
            cboInvoiceType.AddItem Null2String(rsInvoiceType!INVTYPE)
            rsInvoiceType.MoveNext
        Loop
    End If
    Set rsInvoiceType = Nothing
    FillCboAcctName
End Sub

Sub SendToBack()
    cmdShowPostRange.Visible = False
    picShowPostRange.Visible = False
    cmdPrinting.ZOrder 1
    picPrinting.ZOrder 1
End Sub

Private Sub cboCOBAcctName_Click()
    txtCOBAcctNo.Text = SetCOBAcctNo(cboCOBAcctName.Text)
End Sub

Private Sub cmdBatchPosting_Click()
'    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then
'        '        JournalPosting = True
'        '        Exit Function
'        Exit Sub
'    End If

    If dtTo.Value < dtFrom.Value Then
        MessagePop InfoWarning, "Check Date", "Please check selected date."
        dtTo.SetFocus
        Exit Sub
    End If

    If Month(dtFrom.Value) <> Month(dtTo.Value) Then
        MessagePop InfoWarning, "Sysem Message", "Date selected should be of the same month"
        Exit Sub
    End If

    If MsgBox("Are you sure you want to perform Batch Posting?", vbQuestion + vbYesNo, "Message") = vbYes Then
        If xJOURNALTYPE <> "ADJ" And xJOURNALTYPE <> "PDJ" And xJOURNALTYPE <> "OPB" Then
            'Updated by: ACL 10202009
            '            If CheckIfOpen(xJOURNALTYPE, Trim(dtTo.Value), Year(dtTo.Value)) = False Then
            '                MsgBox "Warning: Action not authorized!", vbExclamation, "Error!"
            '                Exit Sub
            '            End If
        End If
        BatchPost = True

        Dim rsJournal_HD                                    As ADODB.Recordset
        Set rsJournal_HD = New ADODB.Recordset
        rsJournal_HD.Open "SELECT VOUCHERNO FROM AMIS_JOURNAL_HD WHERE JTYPE='" & xJOURNALTYPE & "' AND STATUS='N' AND JDATE BETWEEN '" & Format(dtFrom.Value, "mm/dd/yyyy") & "' AND '" & Format(dtTo.Value, "mm/dd/yyyy") & "' ORDER BY ID", gconDMIS, adOpenForwardOnly
        If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
            Do While Not rsJournal_HD.EOF
                lblPosting.Caption = "Posting Voucher No."
                lblVoucherNo.Caption = Null2String(rsJournal_HD!VOUCHERNO)
                Call StoreSearch(Null2String(rsJournal_HD!VOUCHERNO))
                '                    If NumericVal(txtTotDebit) <> NumericVal(txtTotCredit) Then
                '                        MsgBox "Entry is not balanced. Posting of Entry Not Allowed.", vbInformation
                '                        Exit Sub
                '                    End If
                '                    If CheckGLSLARDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                '                        Screen.MousePointer = 0
                '                        Exit Sub
                '                    End If
                '
                '                    If CheckGLSLARCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                '                        Screen.MousePointer = 0
                '                        Exit Sub
                '                    End If
                '
                '                    If CheckGLSLAPDebit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                '                        Screen.MousePointer = 0
                '                        Exit Sub
                '                    End If
                '
                '                    If CheckGLSLAPCredit(xJOURNALTYPE, txtVoucherNo.Text) = False Then
                '                        Screen.MousePointer = 0
                '                        Exit Sub
                '                    End If
                cmdPost_Click
                DoEvents
                rsJournal_HD.MoveNext
            Loop
        End If
        Set rsJournal_HD = Nothing
    End If
End Sub

Private Sub cmdCloseRange_Click()
    Picture1.Enabled = True
    JournalTAB.Enabled = True
    picBatchImport.Visible = False
    picBatchImport.ZOrder 1
    BatchPost = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    xJOURNALTYPE = ""
End Sub

Private Sub txtCustCode_Change()
    txtTerms.Text = SetCustomerTerms(txtCustCode.Text)
End Sub

Private Sub txtCustName_Change()
'    If AddorEdit = "ADD" Then
'        txtCustCode.Text = SetCustomerCode(txtCustName.Text)
'    End If
End Sub

Private Sub cboCustName_Click()
'    txtCustCode.Text = SetCustomerCode(cboCustName.Text)
End Sub

Private Sub cboCustName_GotFocus()
    VBComBoBoxDroppedDown cboCustName
End Sub

Private Sub cboInvoiceType_GotFocus()
    VBComBoBoxDroppedDown cboInvoiceType
End Sub

Private Sub cboInvoiceType_LostFocus()
    On Error Resume Next
    If COMPANY_CODE = "HGC" Then
    Else
        Dim i                                               As Integer
        For i = 0 To cboInvoiceType.ListCount
            If cboInvoiceType = UCase(cboInvoiceType.List(i)) Then
                cboInvoiceType.ListIndex = i
                Exit Sub
            Else

            End If
        Next
        cboInvoiceType = ""
    End If
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", LocalAcess) = False Then Exit Sub
    If MsgBox("Are you sure you want to Cancel this Transaction?", vbQuestion + vbYesNo, "Cancel Journal") = vbYes Then
        Screen.MousePointer = 11

        If COMPANY_CODE = "HGC" Or COMPANY_CODE = "HSB" Then
            If xJOURNALTYPE = "COB" Then
                With FrmCancelTransaction
                    .lblTransaction_type = xJOURNALTYPE
                    .LblTransactionNo = txtVoucherNo.Text
                    FrmCancelTransaction.Show
                End With
            End If
        End If

        gconDMIS.Execute "update AMIS_Journal_HD set status = 'C',USERCODE='" & LOGCODE & "',DATECANCELLED='" & LOGDATE & "' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute "update AMIS_Journal_Det set status = 'C' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)

        SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'C' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'C' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        LogAudit "C", "CUSTOMER OPENING BALANCE", txtCustCode & "-" & txtVoucherNo
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdAdd_Click()
    SendToBack
    Dim rsProfile                                           As ADODB.Recordset
    Dim AccountingMonth, AccountingYear                     As Integer
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select PeriodMonth,PeriodYear from ALL_Profile where modulename = 'AMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        AccountingMonth = rsProfile!PERIODMONTH
        AccountingYear = rsProfile!PERIODYEAR
    End If
    Dim rsDetails                                           As ADODB.Recordset
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
    FillCboAcctName
    initMemvars
    lstDetails.Enabled = False
    On Error Resume Next
    txtVoucherNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDetails.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    PrevJType = UCase(xJOURNALTYPE)
    PrevJNo = Format(txtJNo.Text, "000000")
    lstDetails.Enabled = False
    Frame1.Enabled = True: Picture1.Visible = False: Picture2.Visible = True
    labID.Caption = rsJournal_HD!ID
    If xJOURNALTYPE = "GJ" Or xJOURNALTYPE = "OPB" Or xJOURNALTYPE = "ADJ" Then txtParticulars2.Locked = False
    On Error Resume Next
    txtVoucherNo.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    frmAMISSearchCSJ.Show vbModal
End Sub

Private Sub cmdFirst_Click()
    rsJournal_HD.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    rsJournal_HD.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsJournal_HD.MoveNext
    If rsJournal_HD.EOF Then
        rsJournal_HD.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPost_Click()
    On Error GoTo ErrorCode:
    Dim str_MSG                                             As String
    str_MSG = "Error in Posting @ACL09182716350" & vbCrLf
    str_MSG = str_MSG & "Data Will Now Roll back." & vbCrLf
    str_MSG = str_MSG & "Please Contact Netspeed Software Inc." & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-5:00pm)" & vbCrLf
    str_MSG = str_MSG & "Technical log File Has been created on " & App.Path & "\" & COMPANY_CODE & "_" & LOGDATE & "log.txt" & vbCrLf
    str_MSG = str_MSG & "Please Send The Log File To nsi_dmis@yahoo.com" & vbCrLf

    gconDMIS.BeginTrans
    If JournalPosting = False Then
        str_MSG = Replace(str_MSG, "@ACL09182716350", "Customer Opening")
        MsgBox str_MSG, vbCritical, "Posting Error "
        gconDMIS.RollbackTrans
        Screen.MousePointer = 0
        Exit Sub
    End If

    gconDMIS.CommitTrans
    Screen.MousePointer = 0
    Exit Sub
ErrorCode:
    '    SaveLogFile
    ShowVBError
End Sub

Function JournalPosting() As Boolean
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Post", LocalAcess) = False Then
        JournalPosting = True
        Exit Function
    End If

    If BatchPost = True Then
        GoTo PerformBatchPosting
    End If
    If MsgBox("Are you sure you want to Post this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then

PerformBatchPosting:
        'If MsgBox("Are you sure you want to Post this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then

        'UPDATE BY: JUN --- DATE UPDATE: 11/18/2009 --- DESCRIPTION: THIS GET THE AR OF THE CUSTOMER OPENING BALANCE
        'If COMPANY_CODE <> "HGC" Then

        Call GET_AR_VOUCHERNO
        Call GET_AP_VOUCHERNO
        'End If
        'UPDATE BY: JUN
        If CheckGLSLCOB(xJOURNALTYPE, txtVoucherNo.Text) = False Then
            Screen.MousePointer = 0
            Exit Function
        End If

        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P',USERCODE='" & LOGCODE & "',DATEPOSTED='" & LOGDATE & "' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'P' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'P' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
        gconDMIS.Execute SQL_STATEMENT

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        LogAudit "P", "CUSTOMER OPENING BALANCE", txtCustCode & "-" & txtVoucherNo
    End If
    JournalPosting = True
    Exit Function
    Screen.MousePointer = 0
ErrorCode:
    'SaveLogFile
    ShowVBError
End Function

Private Sub cmdPostRange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtToVNo.Text < txtFromVNo.Text Then
        MsgBox "Error: Invalid Voucher No. Range", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    txtFromVNo.Text = Format(txtFromVNo.Text, "000000")
    txtToVNo.Text = Format(txtToVNo.Text, "000000")
    Dim rsCheckVouchers, rsCheckUnBalancedVouchers          As ADODB.Recordset
    Set rsCheckVouchers = New ADODB.Recordset
    Set rsCheckVouchers = gconDMIS.Execute("Select VoucherNo from AMIS_Journal_HD where Jtype = '" & xJOURNALTYPE & "' AND VoucherNo = '" & txtToVNo.Text & "'")
    If rsCheckVouchers.EOF And rsCheckVouchers.BOF Then
        MsgBox "Error: Voucher No. Range Exceeds Current Records Available.", vbOKOnly + vbInformation, "Please be Careful Guys!"
        Exit Sub
    End If
    Dim KIM, JOY, YZA                                       As Integer
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
                    If N2Str2Zero(rsCheckVouchers!Debit) = N2Str2Zero(rsCheckVouchers!Credit) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!JTYPE) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
                    End If
                ElseIf Null2String(rsCheckVouchers!Status) = "C" Then
                    gconDMIS.Execute "update AMIS_Journal_HD set status = 'C' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    gconDMIS.Execute "update AMIS_Journal_Det set status = 'C' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                Else
                    If N2Str2Zero(rsCheckVouchers!Debit) = N2Str2Zero(rsCheckVouchers!Credit) Then
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'P' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                    Else
                        gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where status = 'N' and jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(Format(KIM, "000000"))
                        MsgBox "Warning: Journal " & Null2String(rsCheckVouchers!JTYPE) & " " & Null2String(rsCheckVouchers!VOUCHERNO) & " is Not Balance... Posting of this Entry is Not Permitted!", vbCritical + vbOKOnly, "Unbalance Journal Entry"
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

Private Sub cmdPrevious_Click()
    rsJournal_HD.MovePrevious
    If rsJournal_HD.BOF Then
        rsJournal_HD.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim rsfindDup                                           As ADODB.Recordset

    If LTrim(RTrim(txtCustCode)) = "" Then
        MsgBox "Customer Code must not be empty", vbInformation, "Incompelete Information"
        On Error Resume Next
        txtCustName.SetFocus
        Exit Sub
    End If

    '    If COMPANY_CODE = "HGC" Then
    '    Else
    If LTrim(RTrim(cboInvoiceType)) = "" Then
        MsgBox "Invoice Type Must not be empty", vbInformation, "Incompelete Information"
        On Error Resume Next
        cboInvoiceType.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(txtInvoiceNo.Text)) = "" Then
        MsgBox "Invoice No Must not be empty", vbInformation, "Incompelete Information"
        On Error Resume Next
        txtInvoiceNo.SetFocus
        Exit Sub
    End If
    '    End If

    If IsNull(txtJNo.Text) = True Then
        MsgBox "Journal No. must not be empty"
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where jtype = '" & xJOURNALTYPE & "' and jno = '" & txtJNo.Text & "' order by jtype,jno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                '                MsgBox "Journal No. already exist!"
                '                txtVoucherNo.SetFocus
                '                Exit Sub
                Call Get_VoucherNo
            End If
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select jtype,jno from AMIS_Journal_HD where invoiceno = '" & txtInvoiceNo.Text & "' and invoicedate = '" & (txtInvoiceDate2.Text) & "' and invoicetype = '" & cboInvoiceType.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            Set rsfindDup = Nothing
        End If

    End If
    If txtJDate.Text = "" Or IsDate(txtJDate.Text) = False Then
        MsgBox "Invalid Date!", vbInformation, "System Message"
        On Error Resume Next
        txtJDate.SetFocus
        Exit Sub
    ElseIf txtInvoiceAmt.Text = "" Or NumericVal(txtInvoiceAmt.Text) = 0 Then
        MsgBox "Invoice amount is zero", vbInformation, "System Message"
        On Error Resume Next
        txtInvoiceAmt.SetFocus
        Exit Sub
    End If

    Dim J_JDATE, J_VOUCHERNO, J_JTYPE                       As String
    Dim J_JNO, J_REMARKS, J_VENDORCODE, J_CUSTOMERCODE      As String
    Dim J_DEBIT, J_CREDIT, J_OUTBALANCE, J_AMOUNTTOPAY, J_INVOICEAMT, J_BALANCE, J_AMOUNTPAID As Double
    Dim J_STATUS, J_CHECKNO                                 As String
    Dim J_INVOICEDATE, J_DUEDATE, J_PAYTYPE                 As String
    Dim J_INVOICETYPE, J_INVOICENO                          As String
    Dim J_BANKCODE                                          As String
    Dim J_REFNO, J_REFDATE                                  As String
    Dim J_TERMS, J_DEALER                                   As String
    Dim J_PAIDSTATUS, J_RECEIVESTATUS                       As String
    Dim J_TAXRATECODE                                       As String
    Dim J_ENTITYCLASS                                       As String
    Dim J_TAXBASE                                           As Double
    J_JDATE = N2Date2Null(txtJDate.Text)
    J_VOUCHERNO = N2Str2Null(txtVoucherNo.Text)
    J_JTYPE = N2Str2Null(xJOURNALTYPE)
    J_TAXRATECODE = "NULL"
    J_TAXBASE = 0
    J_INVOICEDATE = N2Str2Null(txtInvoiceDate2.Text)
    J_BALANCE = NumericVal(txtInvoiceAmt.Text)
    J_AMOUNTPAID = 0
    J_DUEDATE = N2Str2Null(txtDueDate.Text)
    J_PAYTYPE = N2Str2Null(txtPayCode.Text)
    J_JNO = N2Str2Null(txtJNo.Text)
    J_DEBIT = NumericVal(txtTotDebit.Text)
    J_CREDIT = NumericVal(txtTotCredit.Text)
    J_OUTBALANCE = NumericVal(txtOutBalance.Text)
    J_AMOUNTTOPAY = NumericVal(txtAmountToPay.Text)
    J_STATUS = "'N'"

    J_CHECKNO = N2Str2Null(txtCheckNo.Text)
    J_TERMS = N2Str2Null(txtTerms.Text)
    J_DEALER = N2Str2Null(txtDealer.Text)
    J_BANKCODE = N2Str2Null(txtBankCode.Text)

    J_VENDORCODE = "'999999'"
    J_CUSTOMERCODE = N2Str2Null(txtCustCode.Text)
    J_INVOICETYPE = N2Str2Null(SetInvCode(cboInvoiceType.Text))
    J_INVOICENO = N2Str2Null(Format(txtInvoiceNo.Text, "000000"))
    J_INVOICEAMT = NumericVal(txtInvoiceAmt.Text)
    J_REFNO = N2Str2Null(txtRefNo.Text)
    J_REFDATE = N2Date2Null(txtRefDate.Text)
    J_ENTITYCLASS = N2Str2Null(xEntityClass)
    If Trim(txtRemarks2.Text) = "Pls Type Your Message Here!" Then J_REMARKS = "NULL" Else J_REMARKS = N2Str2Null(Trim(txtRemarks2.Text))
    J_PAIDSTATUS = "'N'"
    J_RECEIVESTATUS = "'N'"
    If AddorEdit = "ADD" Then
        Dim rsDuplicateInvoice                              As ADODB.Recordset
        Set rsDuplicateInvoice = New ADODB.Recordset
        rsDuplicateInvoice.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE JTYPE IN ('COB','SJ') AND INVOICETYPE = " & J_INVOICETYPE & " AND INVOICENO=" & J_INVOICENO & "", gconDMIS, adOpenKeyset
        If Not rsDuplicateInvoice.EOF And Not rsDuplicateInvoice.BOF Then
            If COMPANY_CODE = "HMH" Then
                MsgBox "Please check Invoice Type and Invoice No." & Chr(13) & "Continue?", vbCritical + vbYesNo, "Duplicate Invoice No."
                GoTo SaveInvoice
            Else
                MessagePop InfoFriend, "Duplicate Invoice No.", "Please check Invoice Type and Invoice No."
            End If
            Exit Sub
        Else
SaveInvoice:
            Dim rsJournal_HDDup                             As ADODB.Recordset
            Set rsJournal_HDDup = New ADODB.Recordset
            Set rsJournal_HDDup = gconDMIS.Execute("select jno from AMIS_Journal_HD order by jno desc")
            If Not rsJournal_HDDup.EOF And Not rsJournal_HDDup.BOF Then txtJNo.Text = Format(N2Str2Zero(rsJournal_HDDup!JNo) + 1, "000000") Else txtJNo.Text = "000001"
            J_JNO = N2Str2Null(txtJNo.Text)
            gconDMIS.Execute "Insert into AMIS_Journal_HD" & _
                             " (jdate,voucherno,jtype,vendorcode,customercode,CustomerName,invoicedate,invoicetype,invoiceno,invoiceamt,duedate,paytype,refno,refdate,terms,dealer,amounttopay,Balance,AmountPaid,jno,debit,credit,outbalance,status,CheckNo,BankCode,remarks,PaidStatus,ReceiveStatus,USERCODE,LASTUPDATE,Entity_Class)" & _
                             " values (" & J_JDATE & ", " & J_VOUCHERNO & ", " & J_JTYPE & ", " & J_VENDORCODE & "," & J_CUSTOMERCODE & "," & N2Str2Null(txtCustName.Text) & ", " & J_INVOICEDATE & "," & J_INVOICETYPE & "," & J_INVOICENO & "," & J_INVOICEAMT & ", " & J_DUEDATE & ", " & J_PAYTYPE & "," & J_REFNO & "," & J_REFDATE & "," & J_TERMS & "," & J_DEALER & ", " & J_AMOUNTTOPAY & "," & J_BALANCE & "," & J_AMOUNTPAID & _
                             ", " & J_JNO & ", " & J_DEBIT & ", " & J_CREDIT & ", " & J_OUTBALANCE & ", " & J_STATUS & ", " & J_CHECKNO & ", " & J_BANKCODE & "," & J_REMARKS & "," & J_PAIDSTATUS & "," & J_RECEIVESTATUS & ",'" & LOGCODE & "','" & LOGDATE & "'," & J_ENTITYCLASS & ")"
            'LogAudit "A", "CUSTOMER OPENING BALANCE", txtCustCode & "-" & txtVoucherNo
            labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
            NEW_LogAudit "A", "CUSTOMER OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo
        End If
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
                         " BankCode = " & J_BANKCODE & ", " & _
                         " status = " & J_STATUS & ", PaidStatus = " & J_PAIDSTATUS & ", ReceiveStatus = " & J_RECEIVESTATUS & "," & _
                         " remarks = " & J_REMARKS & ", USERCODE = '" & LOGCODE & "', LASTUPDATE = '" & LOGDATE & "',ENTITY_CLASS=" & J_ENTITYCLASS & " " & _
                         " where id = " & labID.Caption
        labID.Caption = FindNewID(J_VOUCHERNO, "VOUCHERNO", "AMIS_JOURNAL_HD", J_JTYPE, "JTYPE")
        NEW_LogAudit "E", "CUSTOMER OPENING BALANCE", SQL_STATEMENT, labID.Caption, "", txtVoucherNo, xJOURNALTYPE, txtJNo

        gconDMIS.Execute "UPDATE AMIS_JOURNAL_DET SET" & _
                         " JTYPE = " & J_JTYPE & "," & _
                         " VOUCHERNO = " & J_VOUCHERNO & "," & _
                         " JDATE = " & J_JDATE & "," & _
                         " USERCODE = '" & LOGCODE & "'," & _
                         " LASTUPDATE = '" & LOGDATE & "'," & _
                         " JNO = " & J_JNO & _
                         " WHERE JTYPE = '" & PrevJType & "' AND JNO = '" & PrevJNo & "'"
        'LogAudit "E", "CUSTOMER OPENING BALANCE", txtCustCode & "-" & txtVoucherNo
    End If

    If Trim(txtCOBAcctNo.Text) <> "" Then
        Dim rsCOB_Journal_Det                               As ADODB.Recordset
        Set rsCOB_Journal_Det = New ADODB.Recordset
        Set rsCOB_Journal_Det = gconDMIS.Execute("Select * from AMIS_Journal_Det Where Jtype = 'COB' and JNO = " & J_JNO)
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
                             " WHERE JTYPE = 'COB' AND JNO = '" & PrevJNo & "'"
        Else
            gconDMIS.Execute "INSERT INTO AMIS_JOURNAL_DET (JITEMNO,JTYPE,JDATE,VOUCHERNO,JNO,ACCT_CODE,ACCT_NAME)" & _
                             " VALUES ('0001','COB'," & J_JDATE & "," & J_VOUCHERNO & "," & J_JNO & "," & N2Str2Null(txtCOBAcctNo.Text) & "," & N2Str2Null(cboCOBAcctName.Text) & ")"
        End If
    End If
    rsRefresh
    rsJournal_HD.Find "jno = " & J_JNO
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    MsgBox "Error:" & Err & " " & error, vbOKOnly, "Error"
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    Set frmNewEntity = New frmEntity
    frmNewEntity.LOADJOURNAL ("COB")
    frmNewEntity.Show 1
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", LocalAcess) = False Then Exit Sub
    If MsgBox("Are you sure you want to Unpost this transaction?", vbQuestion + vbYesNo, "Message") = vbYes Then
        Dim rsCRJ_Detail                                    As ADODB.Recordset
        Set rsCRJ_Detail = New ADODB.Recordset
        Set rsCRJ_Detail = gconDMIS.Execute("Select * from AMIS_CRJ_Detail where INVOICETYPE = '" & SetInvCode(cboInvoiceType.Text) & "' AND INVOICENO = '" & txtInvoiceNo.Text & "' and status <> 'C'")
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            MsgBox "Warning: This Customre Opening is already link to Cash Receipts Voucher No. " & Null2String(rsCRJ_Detail!VOUCHERNO) & vbCrLf & _
                   "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
            Exit Sub
        End If

        Set rsCRJ_Detail = New ADODB.Recordset
        Set rsCRJ_Detail = gconDMIS.Execute("Select VOUCHERNO from AMIS_CV_Detail where INVOICETYPE = '" & SetInvCode(cboInvoiceType.Text) & "' AND INVOICENO = '" & txtInvoiceNo.Text & "' and status <> 'C'")
        If Not rsCRJ_Detail.EOF And Not rsCRJ_Detail.BOF Then
            MsgBox "Warning: This Customre Opening is already link to Cash Disbursement Voucher No. " & Null2String(rsCRJ_Detail!VOUCHERNO) & vbCrLf & _
                   "         Unposting for this Journal Entry is not Allowed unless the link is deleted.", vbCritical, "WARNING!"
            Exit Sub
        End If
        Screen.MousePointer = 11
        gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)

        'UPDATED BY: JUN --- DATE UPDATED: 11/16/2009 --- DESCRIPTION: VALIDATE THE VOUCHER IF UNPOSTING IF THERE IS AN ATTACHED PAYMENT
        'If COMPANY_CODE <> "HGC" Then
        Dim xSJVOUCHERNO                                    As String
        xSJVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
        If xJOURNALTYPE = "COB" Then
            If VALIDATE_UNPOSTING = True Then
                Screen.MousePointer = 0
                Exit Sub
            Else
                gconDMIS.Execute "DELETE FROM AMIS_AR WHERE SJVOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & ""
                gconDMIS.Execute "DELETE FROM AMIS_AP WHERE VOUCHERNO = " & N2Str2Null(xSJVOUCHERNO) & ""

                '                SQL_STATEMENT = "UPDATE AMIS_AR SET STATUS = 'N' where SJVOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
                '                gconDMIS.Execute SQL_STATEMENT
                '
                '                SQL_STATEMENT = "UPDATE AMIS_AP SET STATUS = 'N' where VOUCHERNO = '" & xJOURNALTYPE + "-" + txtVoucherNo.Text & "'"
                '                gconDMIS.Execute SQL_STATEMENT
            End If
        End If
        'End If

        rsRefresh
        rsJournal_HD.Find "id = " & labID.Caption
        StoreMemVars
        LogAudit "U", "CUSTOMER OPENING BALANCE", txtCustCode & "-" & txtVoucherNo
    End If
    Screen.MousePointer = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyReturn
        MoveKeyPress KeyCode
    Case vbKeyEscape
    Case vbKeyF3
    Case vbKeyF4
    Case vbKeyF5
        cmdPost.Value = True
    Case vbKeyF6
        cmdUnPost.Value = True
    Case vbKeyF7
        cmdCancelCO.Value = True
    Case vbKeyF8
    Case vbKeyF9
    Case vbKeyF11
        SendToBack
        lblVoucherNo.Caption = ""
        Picture1.Enabled = False
        JournalTAB.Enabled = False
        picBatchImport.Visible = True
        picBatchImport.ZOrder 0
        dtFrom.Value = Format(txtJDate.Text, "mm/dd/yyyy")
        dtTo.Value = Format(txtJDate.Text, "mm/dd/yyyy")
        On Error Resume Next
    Case vbKeyF12
        '        If Null2String(rsJournal_HD!Status) = "C" Then
        '            'If ApplySecurityValidation = True Then
        '            '    If Module_Access(LOGID, "UNCANCELLED", LocalAcess) = False Then Exit Sub
        '            'End If
        '            If MsgBox("Are you sure you want to Un-Cancel this Transaction?", vbQuestion + vbYesNo, "Un-Cancel Journal") = vbYes Then
        '                Screen.MousePointer = 11
        '                gconDMIS.Execute "update AMIS_Journal_HD set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                gconDMIS.Execute "update AMIS_Journal_Det set status = 'N' where jtype = '" & xJOURNALTYPE & "' and VoucherNo = " & N2Str2Null(txtVoucherNo.Text)
        '                rsRefresh
        '                rsJournal_HD.Find "id = " & labID.Caption
        '                StoreMemVars
        '                Screen.MousePointer = 0
        '            End If
        '        End If
    Case Else
        MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Frame1.Enabled = False: SendToBack:
    Picture1.Visible = True: Picture2.Visible = False: SearchBy = "NAME": fraFindAccount.Caption = "Search Accounts by Account Description"
    picPayables.Top = 1230
    picDisbursement.Top = 1230
    picReceivable.Top = 420
    Frame1.Top = 90
    chkNonVat.Visible = False
    JournalTAB.TabEnabled(1) = False: labBankName.Visible = False: cboBankName2.Visible = False
    Me.Caption = "CUSTOMER OPENING BALANCE - DATA ENTRY"
    LocalAcess = "CUSTOMER OPENING BALANCE"
    labSupplierPayTo = "Supplier Code"
    labType.Caption = "Invoice Type": LabNo.Caption = "Invoice No."
    labDate.Caption = "Invoice Date": labAmt.Caption = "Invoice Amt."
    picGJ.Visible = False: RefCRJ.Visible = True
    picReceivable.Visible = True: picReceivable.ZOrder 0: picReceivable.Enabled = True
    picPayables.Visible = False: picPayables.ZOrder 1: picPayables.Enabled = False
    picDisbursement.Visible = False: picDisbursement.ZOrder 1: picDisbursement.Enabled = False
    labTax.Caption = "Output Tax"
    InitCbo
    initMemvars
    txtSearch.Text = "": txtSearchTemplates.Text = ""
    rsRefresh
    If Not rsJournal_HD.EOF And Not rsJournal_HD.BOF Then
        rsJournal_HD.MoveLast
    End If
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
    If labPosted.Caption <> "" Then
        If labPosted.Visible = True Then labPosted.Visible = False Else labPosted.Visible = True
    End If
End Sub

Private Sub txtAmountToPay_GotFocus()
    If Val(txtAmountToPay.Text) = 0 Then txtAmountToPay.Text = "" Else txtAmountToPay.Text = Format(txtAmountToPay.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtAmountToPay_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtAmountToPay_LostFocus()
    If txtAmountToPay.Text = "" Then txtAmountToPay.Text = "0.00" Else txtAmountToPay.Text = Format(txtAmountToPay.Text, MAXIMUM_DIGIT)
End Sub

Private Sub txtCheckDate_GotFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtCheckDate_LostFocus()
    txtCheckDate.Text = Format(txtCheckDate.Text, "DD-MMM-YY")
End Sub

Private Sub txtINV_No_GotFocus()
    If xJOURNALTYPE = "CDJ" Then
        If txtMRR_No.Text = "" Then txtMRR_No.SetFocus
    End If
End Sub

Private Sub txtINV_No_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtInvoiceAmt_GotFocus()
    txtInvoiceAmt.Text = ""
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
            txtInvoiceDate2 = ""
            txtInvoiceDate2.SetFocus
        End If
    End If
End Sub

Private Sub txtInvoiceNo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtJDate_LostFocus()
    txtJDate.Text = Format(txtJDate.Text, "DD-MMM-YY")
    If IsDate(txtJDate) = False Then
        txtJDate = ""
    End If
    If xJOURNALTYPE = "SJ" Or xJOURNALTYPE = "CRJ" Then
        txtCustName.SetFocus
    Else
        On Error Resume Next
        txtParticulars2.SetFocus
    End If
End Sub

Private Sub txtJDate_GotFocus()
    txtJDate.Text = Format(txtJDate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate_GotFocus()
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(cboPayType.Text), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtMRR_No_Change()
    If xJOURNALTYPE = "CDJ" Then
        Dim rsJournal_HD2                                   As ADODB.Recordset
        Set rsJournal_HD2 = New ADODB.Recordset
        Set rsJournal_HD2 = gconDMIS.Execute("select VoucherNo,JType,JDate,DueDate,AmountToPay from AMIS_Journal_HD where VoucherNo = '" & txtMRR_No.Text & "' and JType = 'APJ'")
        If Not rsJournal_HD2.EOF And Not rsJournal_HD2.BOF Then
            txtINV_No.Text = Null2String(rsJournal_HD2!JDATE)
            txtProd_No.Text = Null2String(rsJournal_HD2!DUEDATE)
            txtPVAmount.Text = ToDoubleNumber(N2Str2Zero(rsJournal_HD2!AMOUNTTOPAY))
        Else
            txtINV_No.Text = ""
            txtProd_No.Text = ""
            txtPVAmount.Text = ZERO
        End If
        Set rsJournal_HD2 = Nothing
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
            txtRefDate = ""
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

'Private Sub grdDetails_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        SendKeys "{TAB}"
'    End If
'End Sub

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

Private Sub txtRemarks2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRemarks2_LostFocus()
    If txtRemarks2.Text = "" Then txtRemarks2.Text = "Pls Type Your Message Here!"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstAccounts.SetFocus
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtTerms_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Sub GET_AR_VOUCHERNO()
    Dim rsAR_VOUCHER                                        As ADODB.Recordset
    Dim rsCOUNT_CODE                                        As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xENTITYCODE                                         As String
    Dim xBAL                                                As Double

    xAMOUNT_TO_PAY = 0
    xAMOUNT_PAID = 0
    xBAL = 0

    Set rsCOUNT_CODE = New ADODB.Recordset
    rsCOUNT_CODE.Open "SELECT COUNT(DISTINCT ACCT_CODE) AS COUNT_CODE FROM AMIS_JOURNAL_DET " & _
                      "WHERE VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND LEFT(ACCT_CODE,5) IN('11-02','11-03','11-04','11-06')", gconDMIS, adOpenKeyset
    If Not rsCOUNT_CODE.EOF And Not rsCOUNT_CODE.BOF Then
        If NumericVal(rsCOUNT_CODE!COUNT_CODE) > 1 Then
            Set rsAR_VOUCHER = New ADODB.Recordset
            rsAR_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,ACCT_CODE,INVOICEAMT " & _
                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03','11-04','11-06') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenKeyset
            If Not rsAR_VOUCHER.EOF And Not rsAR_VOUCHER.BOF Then
                Do While Not rsAR_VOUCHER.EOF
                    xVOUCHERNO = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE) & "-" & Null2String(rsAR_VOUCHER!VOUCHERNO))
                    xJdate = N2Str2Null(Null2String(rsAR_VOUCHER!JDATE))
                    xJType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))

                    If xJOURNALTYPE = "VPJ" Then
                        XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!VendorCode))
                        xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAR_VOUCHER!VendorCode)))
                    Else
                        XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!CustomerCode))
                        xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAR_VOUCHER!CustomerCode), Null2String(rsAR_VOUCHER!Entity_Class)))
                    End If

                    xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!INVOICENO))
                    xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!InvoiceType))
                    xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                    'xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!jtype), Null2String(rsAR_VOUCHER!ACCT_CODE))
                    xAMOUNT_TO_PAY = NumericVal(rsAR_VOUCHER!InvoiceAmt)
                    xAMOUNT_PAID = 0
                    xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                    xACCT_CODE = N2Str2Null(Null2String(rsAR_VOUCHER!ACCT_CODE))
                    xLAST_UPDATED = N2Str2Null(LOGDATE)

                    gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE) " & _
                                     "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ")"
                    rsAR_VOUCHER.MoveNext
                Loop
            End If
            Set rsAR_VOUCHER = Nothing
        Else
            Set rsAR_VOUCHER = New ADODB.Recordset
            rsAR_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,ACCT_CODE,INVOICEAMT,ENTITY_CLASS " & _
                              "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                              "WHERE LEFT(DET.ACCT_CODE,5) IN ('11-02','11-03','11-04','11-06') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & "", gconDMIS, adOpenKeyset
            If Not rsAR_VOUCHER.EOF And Not rsAR_VOUCHER.BOF Then
                xVOUCHERNO = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE) & "-" & Null2String(rsAR_VOUCHER!VOUCHERNO))
                xJdate = N2Str2Null(Null2String(rsAR_VOUCHER!JDATE))
                xJType = N2Str2Null(Null2String(rsAR_VOUCHER!JTYPE))

                If xJOURNALTYPE = "VPJ" Then
                    XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!VendorCode))
                    xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAR_VOUCHER!VendorCode)))
                Else
                    XCustomerCode = N2Str2Null(Null2String(rsAR_VOUCHER!CustomerCode))
                    xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAR_VOUCHER!CustomerCode), Null2String(rsAR_VOUCHER!Entity_Class)))
                End If

                xInvoiceNo = N2Str2Null(Null2String(rsAR_VOUCHER!INVOICENO))
                xInvoiceType = N2Str2Null(Null2String(rsAR_VOUCHER!InvoiceType))
                xInvoicedate = N2Str2Null(Null2String(rsAR_VOUCHER!invoicedate))
                'xAMOUNT_TO_PAY = GET_AR_AMOUNT(Null2String(rsAR_VOUCHER!VOUCHERNO), Null2String(rsAR_VOUCHER!jtype), Null2String(rsAR_VOUCHER!ACCT_CODE))
                xAMOUNT_TO_PAY = NumericVal(rsAR_VOUCHER!InvoiceAmt)
                xAMOUNT_PAID = 0
                xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
                xACCT_CODE = N2Str2Null(Null2String(rsAR_VOUCHER!ACCT_CODE))
                xLAST_UPDATED = N2Str2Null(LOGDATE)

                gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE,ENTITYCODE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & ",'" & xEntityClass & "')"
            End If
            Set rsAR_VOUCHER = Nothing
        End If
    End If
    Set rsCOUNT_CODE = Nothing
End Sub

Function GET_CUST_NAME(xCUSCODE As String, xENTITYCODE As String) As String
    Dim rsGET_CUST_NAME                                     As ADODB.Recordset
    Set rsGET_CUST_NAME = New ADODB.Recordset
    'rsGET_CUST_NAME.Open "SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE RTRIM(LTRIM(CUSCDE)) = '" & RTrim(LTrim(xCUSCODE)) & "'", gconDMIS, adOpenKeyset
    rsGET_CUST_NAME.Open "SELECT ACCOUNTNAME FROM ALL_ENTITY WHERE RTRIM(LTRIM(CODE)) = '" & RTrim(LTrim(xCUSCODE)) & "' AND ENTITYCODE = '" & xENTITYCODE & "'", gconDMIS, adOpenKeyset
    If Not rsGET_CUST_NAME.EOF And Not rsGET_CUST_NAME.BOF Then
        GET_CUST_NAME = Null2String(rsGET_CUST_NAME!ACCOUNTNAME)
    Else
        GET_CUST_NAME = ""
    End If
    Set rsGET_CUST_NAME = Nothing
End Function

Function GET_VEN_NAME(xVENCODE As String) As String
    Dim rsGET_VEN_NAME                                      As ADODB.Recordset
    Set rsGET_VEN_NAME = New ADODB.Recordset
    rsGET_VEN_NAME.Open "SELECT NAMEOFVENDOR FROM ALL_VENDOR WHERE  RTRIM(LTRIM(CODE)) = " & N2Str2Null(xVENCODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_VEN_NAME.EOF And Not rsGET_VEN_NAME.BOF Then
        GET_VEN_NAME = Null2String(rsGET_VEN_NAME!nameofvendor)
    Else
        GET_VEN_NAME = ""
    End If
    Set rsGET_VEN_NAME = Nothing
End Function

Function VALIDATE_UNPOSTING() As Boolean
    Dim rsVALIDATE_UNPOSTING                                As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim rsGET_CRJ_VOUCHERNO                                 As ADODB.Recordset
    xVOUCHERNO = xJOURNALTYPE & "-" & txtVoucherNo.Text
    Set rsVALIDATE_UNPOSTING = New ADODB.Recordset

    rsVALIDATE_UNPOSTING.Open "SELECT X.COMBI_LINK,X.CUSTOMERCODE,X.INVOICETYPE,X.INVOICENO,X.ACCOUNT_CODE FROM " & _
                              "( " & _
                              "SELECT RTRIM(LTRIM(CUSTOMERCODE)) + '-' + RTRIM(LTRIM(INVOICETYPE)) + '-' + RTRIM(LTRIM(INVOICENO)) + '-' + RTRIM(LTRIM(ACCOUNT_CODE)) AS COMBI_LINK, " & _
                              "RTRIM(LTRIM(CUSTOMERCODE)) AS CUSTOMERCODE,RTRIM(LTRIM(INVOICETYPE)) AS INVOICETYPE,RTRIM(LTRIM(INVOICENO)) AS INVOICENO,RTRIM(LTRIM(ACCOUNT_CODE)) AS ACCOUNT_CODE From AMIS_AR WHERE SJVOUCHERNO = '" & xVOUCHERNO & "' " & _
                              ")X  WHERE  X.COMBI_LINK IN(SELECT RTRIM(LTRIM(CUSTOMERCODE)) + '-' + RTRIM(LTRIM(INVOICETYPE)) + '-' + RTRIM(LTRIM(INVOICENO)) + '-' + RTRIM(LTRIM(ACCT_CODE)) FROM AMIS_DETAIL)", gconDMIS, adOpenKeyset
    If Not rsVALIDATE_UNPOSTING.EOF And Not rsVALIDATE_UNPOSTING.BOF Then
        Set rsGET_CRJ_VOUCHERNO = New ADODB.Recordset
        rsGET_CRJ_VOUCHERNO.Open "SELECT JTYPE,VOUCHERNO FROM AMIS_DETAIL WHERE CUSTOMERCODE = " & N2Str2Null(rsVALIDATE_UNPOSTING!CustomerCode) & " AND INVOICENO = " & N2Str2Null(rsVALIDATE_UNPOSTING!INVOICENO) & " AND INVOICETYPE = " & N2Str2Null(rsVALIDATE_UNPOSTING!InvoiceType) & " AND ACCT_CODE = " & N2Str2Null(rsVALIDATE_UNPOSTING!Account_code) & " ", gconDMIS, adOpenKeyset
        If Not rsGET_CRJ_VOUCHERNO.EOF And Not rsGET_CRJ_VOUCHERNO.BOF Then
            MessagePop InfoFriend, "INFORMATION", "You can't un-post this voucher it has a payment please see Cash Receipts Journal " & "" & Null2String(rsGET_CRJ_VOUCHERNO!JTYPE) & "" & " - " & "" & Null2String(rsGET_CRJ_VOUCHERNO!VOUCHERNO) & ""
            VALIDATE_UNPOSTING = True
        End If
        Set rsGET_CRJ_VOUCHERNO = Nothing
    Else
        VALIDATE_UNPOSTING = False
    End If
    Set rsVALIDATE_UNPOSTING = Nothing
End Function

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    txtCustCode.Text = strCode
    txtCustName.Text = strAccountName
    xEntityClass = strEntityClass
End Sub

Function CheckStatus(xVOUCHERNO As String, xJType As String) As Boolean
    Dim rsCheckStatus                                       As ADODB.Recordset
    Set rsCheckStatus = New ADODB.Recordset
    rsCheckStatus.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE STATUS = 'P' AND VOUCHERNO = '" & xVOUCHERNO & "' AND JTYPE = '" & xJType & "'", gconDMIS, adOpenKeyset
    If Not rsCheckStatus.EOF And Not rsCheckStatus.BOF Then
        CheckStatus = True
    Else
        CheckStatus = False
    End If
End Function

Private Sub txtVoucherNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 48 To 57
    Case 8
    Case Else
        KeyAscii = 0
        MsgBox "Invalid Character!", vbExclamation, "Check Entry"
        txtVoucherNo.SetFocus
    End Select
End Sub

Sub LOADJOURNAL(XXX As String)
    xJOURNALTYPE = XXX
End Sub

Function SetCustomerTerms(XXX As String)
    Set rsCustomer = New ADODB.Recordset
    rsCustomer.Open "SELECT ISNULL(CREDITDAYS,0) AS CREDITDAYS FROM ALL_CUSTOMER_TABLE WHERE CUSCDE= '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsCustomer.EOF And Not rsCustomer.BOF Then
        SetCustomerTerms = rsCustomer!CREDITDAYS
    End If
    Set rsCustomer = Nothing
End Function


Sub GET_AP_VOUCHERNO()
    Dim rsAP_VOUCHER                                        As ADODB.Recordset
    Dim xVOUCHERNO                                          As String
    Dim xJdate                                              As String
    Dim xDUEDATE                                            As String
    Dim xJType                                              As String
    Dim XCustomerCode                                       As String
    Dim xCUST_NAME                                          As String
    Dim xInvoiceNo                                          As String
    Dim xInvoiceType                                        As String
    Dim xInvoicedate                                        As String
    Dim xAMOUNT_TO_PAY                                      As Double
    Dim xAMOUNT_PAID                                        As Double
    Dim xACCT_CODE                                          As String
    Dim xLAST_UPDATED                                       As String
    Dim xBAL                                                As Double
    Dim xREFCODE                                            As String
    Dim xENTITYCODE                                         As String

    xBAL = 0
    xAMOUNT_PAID = 0
    xAMOUNT_TO_PAY = 0

    Set rsAP_VOUCHER = New ADODB.Recordset
    'rsAP_VOUCHER.Open "SELECT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,PV.INV_NO AS INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,PV.AMOUNT,ACCT_CODE,PV.REFCODE,PV.ENTITYCODE " & _
     "FROM AMIS_JOURNAL_HD HD " & _
     "INNER JOIN AMIS_JOURNAL_DET DET " & _
     "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
     "INNER JOIN AMIS_PV_DETAIL PV " & _
     "ON HD.VOUCHERNO=PV.VOUCHERNO AND HD.JTYPE=PV.JTYPE " & _
     "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-07','11-02','11-03') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenForwardOnly
    rsAP_VOUCHER.Open "SELECT DISTINCT HD.VOUCHERNO,HD.VENDORCODE,HD.JDATE,HD.JTYPE,HD.CUSTOMERCODE,HD.INVOICENO,HD.INVOICETYPE,HD.INVOICEDATE,HD.DUEDATE,CASE WHEN HD.JTYPE='COB' THEN HD.INVOICEAMT ELSE HD.AMOUNTTOPAY END AS AMOUNTTOPAY,ACCT_CODE " & _
                      "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                      "WHERE LEFT(DET.ACCT_CODE,5) IN ('21-01','21-02','21-04','21-06','21-07') AND HD.JTYPE = " & N2Str2Null(xJOURNALTYPE) & " AND HD.VOUCHERNO = " & N2Str2Null(txtVoucherNo.Text) & " AND HD.STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsAP_VOUCHER.EOF And Not rsAP_VOUCHER.BOF Then
        'Do While Not rsAP_VOUCHER.EOF
        xVOUCHERNO = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE) & "-" & Null2String(rsAP_VOUCHER!VOUCHERNO))
        xJdate = N2Str2Null(Null2String(rsAP_VOUCHER!JDATE))
        xJType = N2Str2Null(Null2String(rsAP_VOUCHER!JTYPE))
        xDUEDATE = N2Str2Null(Null2String(rsAP_VOUCHER!DUEDATE))

        If xJOURNALTYPE = "VPJ" Then
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!VendorCode))
            xCUST_NAME = N2Str2Null(GET_VEN_NAME(Null2String(rsAP_VOUCHER!VendorCode)))
        Else
            XCustomerCode = N2Str2Null(Null2String(rsAP_VOUCHER!CustomerCode))
            xCUST_NAME = N2Str2Null(GET_CUST_NAME(Null2String(rsAP_VOUCHER!CustomerCode), Null2String(rsAP_VOUCHER!Entity_Class)))
        End If

        xInvoiceNo = N2Str2Null(Null2String(rsAP_VOUCHER!INVOICENO))
        'xInvoiceNo = GetINVOICENO
        xInvoiceType = N2Str2Null(Null2String(rsAP_VOUCHER!InvoiceType))
        xInvoicedate = N2Str2Null(Null2String(rsAP_VOUCHER!invoicedate))
        'xREFCODE = N2Str2Null(Null2String(rsAP_VOUCHER!REFCODE))
        'xEntityCode = N2Str2Null(Null2String(rsAP_VOUCHER!ENTITYCODE))
        'xAMOUNT_TO_PAY = GET_AP_AMOUNT(Null2String(rsAP_VOUCHER!VOUCHERNO), Null2String(rsAP_VOUCHER!jtype), Null2String(rsAP_VOUCHER!ACCT_CODE))
        xAMOUNT_TO_PAY = NumericVal(rsAP_VOUCHER!AMOUNTTOPAY)
        xAMOUNT_PAID = 0
        xBAL = Round((xAMOUNT_TO_PAY - xAMOUNT_PAID), 2)
        xACCT_CODE = N2Str2Null(Null2String(rsAP_VOUCHER!ACCT_CODE))
        xLAST_UPDATED = N2Str2Null(LOGDATE)

        'SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE,REFCODE,ENTITYCODE) " & _
         "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & "," & xREFCODE & "," & xEntityCode & ")"
        SQL_STATEMENT = "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,DUEDATE) " & _
                        "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUST_NAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xDUEDATE & ")"
        gconDMIS.Execute SQL_STATEMENT
        'rsAP_VOUCHER.MoveNext
        'Loop
    End If
    Set rsAP_VOUCHER = Nothing
End Sub

Function CheckGLSLCOB(JTYPE As String, VOUCHERNO As String) As Boolean
    Dim rsJournalHD                                         As ADODB.Recordset
    Dim rsAR                                                As ADODB.Recordset
    Set rsJournalHD = New ADODB.Recordset
    rsJournalHD.Open "SELECT * FROM AMIS_JOURNAL_HD WHERE JTYPE = '" & JTYPE & "' AND VOUCHERNO = '" & VOUCHERNO & "'", gconDMIS, adOpenForwardOnly
    If Not rsJournalHD.EOF And Not rsJournalHD.BOF Then
        Set rsAR = New ADODB.Recordset
        rsAR.Open "SELECT SUM(AMOUNT_TOPAY) AS AMOUNTTOPAY FROM AMIS_AR WHERE LEFT(SJVOUCHERNO,3)= '" & JTYPE & "' AND RIGHT(SJVOUCHERNO,6) = '" & VOUCHERNO & "'", gconDMIS, adOpenForwardOnly
        If Not rsAR.EOF And Not rsAR.BOF Then
            If NumericVal(rsJournalHD!InvoiceAmt) = NumericVal(rsAR!AMOUNTTOPAY) Then
                CheckGLSLCOB = True
            Else
                MessagePop InfoWarning, "System Message", "Please check. GL Amount not equal to SL" & Chr(13) & "AMOUNT TO PAY => " & NumericVal(rsJournalHD!AMOUNTTOPAY) & Chr(13) & "DETAIL => " & NumericVal(rsAR!AMOUNTTOPAY)
                CheckGLSLCOB = False
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
    End If
    Set rsJournal_HD = Nothing
    Set rsAR = Nothing
End Function
