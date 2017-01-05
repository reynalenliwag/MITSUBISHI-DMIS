VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPBUTTON.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSMATReceiving 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Receiving Entry"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MATreceiving.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   11805
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   2250
      ScaleHeight     =   900
      ScaleWidth      =   12495
      TabIndex        =   63
      Top             =   5625
      Width           =   12495
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
         Left            =   8760
         MouseIcon       =   "MATreceiving.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   8040
         MouseIcon       =   "MATreceiving.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   7320
         MouseIcon       =   "MATreceiving.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelRR 
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
         Left            =   6600
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATreceiving.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Cancel this Transaction"
         Top             =   60
         Width           =   735
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
         Left            =   5880
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATreceiving.frx":1B43
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":1C95
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Unpost this Transaction"
         Top             =   60
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
         Left            =   5160
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "MATreceiving.frx":1FDA
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":212C
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Post this Transaction"
         Top             =   60
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
         Left            =   4440
         MouseIcon       =   "MATreceiving.frx":2451
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":25A3
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
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
         Left            =   3720
         MouseIcon       =   "MATreceiving.frx":28FF
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":2A51
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Add Record"
         Top             =   60
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
         Left            =   3000
         MouseIcon       =   "MATreceiving.frx":2D64
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":2EB6
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Move to Last Record"
         Top             =   60
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
         Left            =   2280
         MouseIcon       =   "MATreceiving.frx":3206
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":3358
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Move to First Record"
         Top             =   60
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
         Left            =   1560
         MouseIcon       =   "MATreceiving.frx":36B6
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":3808
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Find a Record"
         Top             =   60
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
         Left            =   840
         MouseIcon       =   "MATreceiving.frx":3B02
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":3C54
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Move to Next Record"
         Top             =   60
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
         Left            =   120
         MouseIcon       =   "MATreceiving.frx":3FAC
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":40FE
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   2445
      Left            =   2250
      TabIndex        =   19
      Top             =   3075
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid grdDetails 
         Height          =   2175
         Left            =   60
         TabIndex        =   12
         Top             =   105
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
         Enabled         =   -1  'True
         TextRTF         =   $"MATreceiving.frx":445D
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
            Enabled         =   -1  'True
            TextRTF         =   $"MATreceiving.frx":44F0
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
         TabIndex        =   52
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
         TabIndex        =   51
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
   Begin VB.Frame fraAddTran 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   4575
      TabIndex        =   34
      Top             =   1035
      Width           =   4575
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
         Left            =   3420
         MouseIcon       =   "MATreceiving.frx":4578
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":46CA
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Delete Entry"
         Top             =   3150
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
         Left            =   1470
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
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
         TabIndex        =   49
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
         Left            =   2730
         MouseIcon       =   "MATreceiving.frx":49F5
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":4B47
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Cancel Entry"
         Top             =   3150
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
         Left            =   2040
         MouseIcon       =   "MATreceiving.frx":4E85
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":4FD7
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Save Entry"
         Top             =   3150
         Width           =   705
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
         TabIndex        =   50
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
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   60
      TabIndex        =   54
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
         Top             =   390
         Value           =   -1  'True
         Width           =   1845
      End
      Begin MSComctlLib.ListView lstMATREC 
         Height          =   5085
         Left            =   60
         TabIndex        =   58
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
         MouseIcon       =   "MATreceiving.frx":5327
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
         TabIndex        =   59
         Top             =   150
         Width           =   1455
      End
   End
   Begin wizButton.cmd cmdAddTran 
      Height          =   4245
      Left            =   4500
      TabIndex        =   53
      Top             =   975
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
      MICON           =   "MATreceiving.frx":5489
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   10290
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   77
      Top             =   5655
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
         Left            =   720
         MouseIcon       =   "MATreceiving.frx":54A5
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":55F7
         Style           =   1  'Graphical
         TabIndex        =   79
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
         Left            =   30
         MouseIcon       =   "MATreceiving.frx":5935
         MousePointer    =   99  'Custom
         Picture         =   "MATreceiving.frx":5A87
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmCSMSMATReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMATREC                                           As ADODB.Recordset
Dim rsPO_HD                                            As ADODB.Recordset
Dim rsTDAYTRAN                                         As ADODB.Recordset
Dim rsMatMas                                           As ADODB.Recordset
Dim rsSupplier                                         As ADODB.Recordset
Dim rsCunter                                           As ADODB.Recordset
Dim Pcnt                                               As Integer
Dim AddorEdit                                          As String
Dim MATrec_TOTUCOST                                    As Double
Dim MATrec_TOTINVAMT                                   As Double
Dim MATrec_TOTVAT                                      As Double
Dim RR_QTY_REC                                         As Long
Dim PrevRRNo                                           As String
Dim ISNONVAT                                           As Boolean

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
        On Error Resume Next
        cboTranMatCde.SetFocus
    End If
End Sub

Private Sub cmdCancelRR_Click()
    On Error GoTo Errorcode
    If Function_Access(LOGID, "Acess_CancelEntry", "RECEIVING STORING") = False Then Exit Sub
    If MsgQuestionBox("Are you sure you want to Cancel this Transactions?", "Cancel Transactions") = True Then
        Dim PCurOnOrder, PCurTRECQTY, PCurReceipts     As Integer
        Dim PCurLast_recq, PCurTpoQty                  As Integer
        Dim rsTdaytranDup, rsPartmasDup                As ADODB.Recordset
        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno,tranqty,matcde from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno), gconDMIS
        If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
            rsTdaytranDup.MoveFirst
            Do While Not rsTdaytranDup.EOF
                Set rsPartmasDup = New ADODB.Recordset
                rsPartmasDup.Open "select matcde,onorder,trecqty,receipts,last_recq from CSMS_MatMas where matcde = " & N2Str2Null(rsTdaytranDup!MATCDE), gconDMIS
                If Not rsPartmasDup.EOF And Not rsPartmasDup.BOF Then
                    PCurOnOrder = N2Str2IntZero(rsPartmasDup!onorder) + N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurTRECQTY = N2Str2IntZero(rsPartmasDup!trecqty) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurReceipts = N2Str2IntZero(rsPartmasDup!receipts) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    PCurLast_recq = N2Str2IntZero(rsPartmasDup!last_recq) - N2Str2IntZero(rsTdaytranDup!tranqty)
                    gconDMIS.Execute "update CSMS_MatMas set" & _
                                   " onorder = " & PCurOnOrder & "," & _
                                   " trecqty = " & PCurTRECQTY & "," & _
                                   " receipts = " & PCurReceipts & "," & _
                                   " last_recq = " & PCurLast_recq & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where matcde = " & N2Str2Null(rsTdaytranDup!MATCDE)
                End If
                rsTdaytranDup.MoveNext
            Loop
        End If
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        gconDMIS.Execute "update CSMS_TdayTran set" & _
                       " status = 'C'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where tranno = " & N2Str2Null(rsMATREC!rrno) & " and trantype = 'RR'"
        rsRefresh
        On Error Resume Next
        rsMATREC.Find "id =" & labid.Caption
        StoreMemVars
    End If

Errorcode:
    ShowVBError
    Exit Sub

End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "RECEIVING STORING") = False Then Exit Sub

End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "RECEIVING STORING") = False Then Exit Sub

    Dim mmasOnOrder                                    As Integer
    If MsgQuestionBox("Are you sure you want to Post this Transactions?", "Post Transactions") = True Then
        Set rsTDAYTRAN = New ADODB.Recordset
        rsTDAYTRAN.Open "select id,itemno,trantype,tranno,matcde,tranqty,traninvamt from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS
        If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
            rsTDAYTRAN.MoveFirst
            Do While Not rsTDAYTRAN.EOF
                If N2Str2Zero(rsTDAYTRAN!TRANINVAMT) <= 0 Then
                    MsgSpeechBox "Transaction with Invoice Amount equal to Zero Encountered!"
                    Exit Sub
                End If
                rsTDAYTRAN.MoveNext
            Loop
            rsTDAYTRAN.MoveFirst
            Do While Not rsTDAYTRAN.EOF
                Set rsMatMas = New ADODB.Recordset
                rsMatMas.Open "Select matcde,onhand,trecqty,onorder,receipts from CSMS_MatMas where matcde = " & N2Str2Null(rsTDAYTRAN!MATCDE), gconDMIS
                If Not rsMatMas.EOF And Not rsMatMas.EOF Then
                    mmasOnOrder = N2Str2Zero(rsMatMas!onorder)
                    If mmasOnOrder <= 0 Then
                        mmasOnOrder = NumericVal(rsTDAYTRAN!tranqty)
                    End If
                    gconDMIS.Execute "update CSMS_MatMas set onhand =" & N2Str2Zero(rsMatMas!ONHAND) + NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(rsMatMas!trecqty) + NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " onorder = " & mmasOnOrder - NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(rsMatMas!receipts) + NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " last_recq = " & N2Str2Zero(rsTDAYTRAN!tranqty) & ", " & _
                                   " last_recd = '" & LOGDATE & "', " & _
                                   " supcode = " & N2Str2Null(Mid(txtRecvd_Code.Text, 1, 5)) & _
                                   " where matcde = " & N2Str2Null(rsMatMas!MATCDE)
                    gconDMIS.Execute "update CSMS_TdayTran set" & _
                                   " status = 'P'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTDAYTRAN!ID
                End If
                rsTDAYTRAN.MoveNext
            Loop
        End If
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " status = 'P'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        rsRefresh
        On Error Resume Next
        rsMATREC.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "RECEIVING STORING") = False Then Exit Sub
    Screen.MousePointer = 11
    rptReceiving.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptReceiving.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptReceiving, CSMS_REPORT_PATH & "rr.RPT", "{MATrec.rrno} = '" & txtRRNo.Text & "'", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdTranCancel_Click()
    SendToBack
    StoreMemVars
End Sub

Private Sub cmdTranDelete_Click()
    If MsgQuestionBox("Delete This Materials, Are you Sure?", "Delete Material Entry") = True Then
        gconDMIS.Execute "delete from CSMS_TdayTran where id = " & labDetId.Caption
        ShowDeletedMsg
    End If
    Dim cnt                                            As Integer
    Dim rsTdaytranDup                                  As ADODB.Recordset
    Set rsTdaytranDup = New ADODB.Recordset
    rsTdaytranDup.Open "select id,itemno from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS
    If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
        rsTdaytranDup.MoveFirst
        cnt = 0
        Do While Not rsTdaytranDup.EOF
            cnt = cnt + 1
            gconDMIS.Execute "update CSMS_TdayTran set itemno = " & Format(cnt, "0000") & " where id = " & rsTdaytranDup!ID
            rsTdaytranDup.MoveNext
        Loop
    End If
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        MATrec_TOTVAT = MATrec_TOTINVAMT - MATrec_TOTUCOST
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    Else
        MATrec_TOTVAT = 0
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsMATREC.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
End Sub

Private Sub cmdTranSave_Click()
    Screen.MousePointer = 11
    On Error GoTo Errorcode

    If cboTranMatCde.Text = "" Then
        MsgSpeechBox "Material Code must have a value"
        Exit Sub
    End If

    If AddorEdit = "ADD" Then
        Dim rsTDaytranClone                            As ADODB.Recordset
        Set rsTDaytranClone = New ADODB.Recordset
        rsTDaytranClone.Open "select trantype,tranno,itemno,matcde from CSMS_TdayTran where matcde = '" & cboTranMatCde.Text & "' and trantype = 'RR' and tranno =" & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS
        If Not rsTDaytranClone.EOF And Not rsTDaytranClone.BOF Then
            MsgSpeechBox "Material Code already used in this transaction!"
            Exit Sub
        End If
    End If

    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE               As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP             As String
    Dim RRTRANQTY                                      As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                      As Double
    Dim RRSTATUS, RRIN_OUT                             As String

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
        gconDMIS.Execute "insert into CSMS_TdayTran " & _
                         "(trandate,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                       " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                       " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                       " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                       " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
    Else
        gconDMIS.Execute "update CSMS_TdayTran set" & _
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
                       " where id = " & labDetId.Caption
    End If

    Dim varPmasTrecqty, varPmasOnOrder, varPmasTpoqty  As Long
    Dim rsMatMasClone                                  As ADODB.Recordset
    Set rsMatMasClone = New ADODB.Recordset
    rsMatMasClone.Open "select matcde,trecqty,receipts,onorder,tpoqty from CSMS_MatMas where matcde = " & RRSTOCK_ORD, gconDMIS
    If rsMatMasClone.EOF And rsMatMasClone.BOF Then
        gconDMIS.Execute "insert into CSMS_MatMas" & _
                         "(matcde,matdsc,date_entered)" & _
                       " values (" & N2Str2Null(cboTranMatCde.Text) & ", " & N2Str2Null(Mid(cboTranMatDsc.Text, 1, 50)) & _
                         ",'" & LOGDATE & "')"
    End If
    cleargrid grdDetails
    FillDetails
    If NumericVal(txtDS1.Text) > 0 Then
        'MATrec_TOTVAT = MATrec_TOTINVAMT - MATrec_TOTUCOST
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    Else
        MATrec_TOTVAT = 0
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsMATREC.Find "id = " & labid.Caption
    cmdTranCancel.Value = True
    If AddorEdit = "ADD" Then cmdAddTran_Click
    Screen.MousePointer = 0
    Exit Sub

Errorcode:
    ShowVBError
    Screen.MousePointer = 0
    Exit Sub
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "RECEIVING STORING") = False Then Exit Sub
    If MsgQuestionBox("Are you sure you want to Unpost this Transactions?", "Unpost Transactions") = True Then
        Set rsTDAYTRAN = New ADODB.Recordset
        rsTDAYTRAN.Open "select id,itemno,trantype,tranno,matcde,tranqty,traninvamt from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS
        If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
            rsTDAYTRAN.MoveFirst
            Do While Not rsTDAYTRAN.EOF
                Set rsMatMas = New ADODB.Recordset
                rsMatMas.Open "Select matcde,onhand,trecqty,onorder,receipts from CSMS_MatMas where matcde = " & N2Str2Null(rsTDAYTRAN!MATCDE), gconDMIS
                If Not rsMatMas.EOF And Not rsMatMas.EOF Then
                    gconDMIS.Execute "update CSMS_MatMas set onhand =" & N2Str2Zero(rsMatMas!ONHAND) - NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " trecqty = " & N2Str2Zero(rsMatMas!trecqty) - NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " onorder = " & N2Str2Zero(rsMatMas!onorder) + NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " receipts = " & N2Str2Zero(rsMatMas!receipts) - NumericVal(rsTDAYTRAN!tranqty) & ", " & _
                                   " last_recq = " & 0 & ", " & _
                                   " last_recd = NULL, " & _
                                   " supcode = NULL" & _
                                   " where matcde = " & N2Str2Null(rsMatMas!MATCDE)
                    gconDMIS.Execute "update CSMS_TdayTran set" & _
                                   " status = 'N'" & "," & _
                                   " usercode = " & N2Str2Null(LOGCODE) & "," & _
                                   " lastupdate = '" & LOGDATE & "'" & _
                                   " where id = " & rsTDAYTRAN!ID
                End If
                rsTDAYTRAN.MoveNext
            Loop
        End If
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " status = 'N'," & _
                       " usercode = " & N2Str2Null(LOGCODE) & "," & _
                       " lastupdate = '" & LOGDATE & "'" & _
                       " where id = " & labid.Caption
        rsRefresh
        On Error Resume Next
        rsMATREC.Find "id =" & labid.Caption
        StoreMemVars
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "RECEIVING STORING") = False Then Exit Sub
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
    If Function_Access(LOGID, "Acess_EDIT", "RECEIVING STORING") = False Then Exit Sub

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
    'Picture5.Visible = False
    'Dim findStr As String
    'findStr = InputSpeechBox("Please Input Transaction Number", txtRRNo.Text)
    'If findStr <> "" Then
    '   On Error GoTo ErrorCode
    '   rsMATREC.Bookmark = rsFind(rsMATREC.Clone, "rrno", findStr).Bookmark
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
    rsMATREC.Bookmark = rsFind(rsMATREC.Clone, "rrno", Format(DDD, "000000")).Bookmark
    StoreMemVars
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    rsMATREC.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    rsMATREC.MoveLast
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsMATREC.MoveNext
    If rsMATREC.EOF Then
        rsMATREC.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsMATREC.MovePrevious
    If rsMATREC.BOF Then
        rsMATREC.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim rsMATRecDup                                    As ADODB.Recordset

    If Trim(txtRRNo.Text) = "" Then
        MsgSpeechBox "RR Number must not be empty"
        On Error Resume Next
        txtRRNo.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup                              As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select rrno from CSMS_MatRec where rrno = '" & txtRRNo.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "RR Number already exist!"
                On Error Resume Next
                txtRRNo.SetFocus
                Exit Sub
            End If
            Set rsMATRecDup = New ADODB.Recordset
            rsMATRecDup.Open "select pono from CSMS_MatRec where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not rsMATRecDup.EOF And Not rsMATRecDup.BOF Then
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

    Dim NewRRCunTer                                    As String
    NewRRCunTer = NumericVal(txtRRNo.Text) + 1

    Dim VTXTRRNo, VTXTRRDate, Vcboclasscode            As String
    Dim VTXTRecvd_Code, VTXTRecvd_From, VtxtAddress    As String
    Dim VTXTTerms, VTXTPONo, VTXTPODate                As String
    Dim VTXTDRNo, VTXTINVNo                            As String
    Dim VTXTTTLRRAmt, VTXTDS1                          As Double
    Dim VTXTDS_Desc1                                   As String
    Dim VTXTDS_Amt1, VTXTNetRRAmt                      As Double
    Dim VTXTRemarks                                    As String

    Dim RRTRANDATE, RRTRANNO, RRTRANTYPE               As String
    Dim RRITEMNO, RRSTOCK_ORD, RRSTOCK_SUP             As String
    Dim RRTRANQTY                                      As Integer
    Dim RRTRANUCOST, RRTRANINVAMT                      As Double
    Dim RRIN_OUT, RRSTATUS                             As String

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
        gconDMIS.Execute "Insert into CSMS_MatRec" & _
                       " (rrno,rrdate,classcode,recvd_code,recvd_from,address,terms,pono,podate,drno,invno,ttlrramt,ds1,ds_desc1,ds_amt1,netrramt,usercode,lastupdate,remarks)" & _
                       " values (" & VTXTRRNo & ", " & VTXTRRDate & ", " & Vcboclasscode & ", " & _
                       " " & VTXTRecvd_Code & ", " & VTXTRecvd_From & ", " & VtxtAddress & ", " & VTXTTerms & _
                         ", " & VTXTPONo & ", " & VTXTPODate & ", " & VTXTDRNo & ", " & VTXTINVNo & _
                         ", " & VTXTTTLRRAmt & _
                         ", " & VTXTDS1 & ", " & VTXTDS_Desc1 & ", " & VTXTDS_Amt1 & _
                         ", " & VTXTNetRRAmt & ", " & N2Str2Null(LOGCODE) & ", '" & LOGDATE & "'," & VTXTRemarks & ")"
    Else
        gconDMIS.Execute "update CSMS_MatRec set" & _
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
                       " where id = " & labid.Caption
        gconDMIS.Execute "update CSMS_TdayTran set" & _
                       " trandate = " & VTXTRRDate & "," & _
                       " tranno = " & VTXTRRNo & _
                       " where trantype = 'RR' and tranno = '" & PrevRRNo & "'"
    End If
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "update CSMS_Cunter set nextnumber = '" & NewRRCunTer & "', lastupdate = '" & LOGDATE & "', usercode = " & N2Str2Null(LOGCODE) & " where modul = 'RR'"
    End If
    rsRefresh
    On Error Resume Next
    rsMATREC.Find "rrno = " & VTXTRRNo
    cmdCancel.Value = True
    On Error GoTo Errorcode
    If AddorEdit = "ADD" Then
        Dim rsTdaytranDup, rstdaytranDUp2              As ADODB.Recordset
        Dim varPmasTrecqty, varPmasOnOrder, varPmasOnhand As Long
        Dim rsMatMasClone                              As ADODB.Recordset
        Set rsTdaytranDup = New ADODB.Recordset
        rsTdaytranDup.Open "select trantype,tranno from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno), gconDMIS
        If rsTdaytranDup.EOF And rsTdaytranDup.BOF Then
            rsTdaytranDup.Close
            Set rstdaytranDUp2 = New ADODB.Recordset
            rstdaytranDUp2.Open "select trantype,tranno,matcde,matdsc,itemno,tranqty,traninvamt,tranucost from CSMS_TdayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsMATREC!pono), gconDMIS
            If Not rstdaytranDUp2.EOF And Not rstdaytranDUp2.BOF Then
                rstdaytranDUp2.MoveFirst
                Do While Not rstdaytranDUp2.EOF
                    RRTRANDATE = N2Date2Null(txtPODate.Text)
                    RRTRANTYPE = "'RR'"
                    RRTRANNO = N2Str2Null(rsMATREC!rrno)
                    RRITEMNO = N2Str2Null(Null2String(rstdaytranDUp2!itemno))
                    RRSTOCK_ORD = N2Str2Null(rstdaytranDUp2!MATCDE)
                    RRSTOCK_SUP = N2Str2Null(rstdaytranDUp2!MatDsc)
                    RRTRANQTY = N2Str2IntZero(rstdaytranDUp2!tranqty)
                    RRTRANINVAMT = N2Str2Zero(rstdaytranDUp2!TRANINVAMT)
                    RRTRANUCOST = N2Str2Zero(rstdaytranDUp2!TRANUCOST)
                    RRIN_OUT = "'I'"
                    RRSTATUS = "'N'"

                    gconDMIS.Execute "insert into CSMS_TdayTran " & _
                                     "(trandate,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt,lastupdate,usercode,status,in_out)" & _
                                   " values (" & RRTRANDATE & ", " & RRTRANTYPE & ", " & RRTRANNO & "," & _
                                   " " & RRITEMNO & "," & RRSTOCK_ORD & "," & _
                                   " " & RRSTOCK_SUP & ", " & RRTRANQTY & "," & _
                                   " " & RRTRANUCOST & ", " & RRTRANINVAMT & ", '" & LOGDATE & "', " & N2Str2Null(LOGCODE) & ", " & RRSTATUS & ", " & RRIN_OUT & ")"
                    rstdaytranDUp2.MoveNext
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
        MATrec_TOTVAT = MATrec_TOTINVAMT - MATrec_TOTUCOST
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = '" & "VAT" & "'," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    Else
        MATrec_TOTVAT = 0
        gconDMIS.Execute "update CSMS_MatRec set" & _
                       " ttlrramt = " & MATrec_TOTUCOST & "," & _
                       " netrramt = " & MATrec_TOTINVAMT & "," & _
                       " ds_desc1 = NULL," & _
                       " ds_amt1 = " & MATrec_TOTVAT & "," & _
                       " ds1 = " & NumericVal(txtDS1.Text) & _
                       " where id = " & labid.Caption
    End If
    rsRefresh
    On Error Resume Next
    rsMATREC.Find "rrno = " & VTXTRRNo
    StoreMemVars
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
                If Null2String(rsMATREC!Status) = "P" Then
                    MsgSpeechBox "Item(s) are Already Posted, and cannot be change..."
                ElseIf Null2String(rsMATREC!Status) = "C" Then
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
    textSearch.Text = "":    'Picture5.ZOrder 0
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    txtMaterialID.Text = ""
    initMemvars
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsMATREC = New ADODB.Recordset
    rsMATREC.Open "select * from CSMS_MatRec order by id desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub initMemvars()
    txtRRNo.Text = ""
    txtPONo.Text = ""
    Set rsCunter = New ADODB.Recordset
    rsCunter.Open "select * from CSMS_Cunter where modul = 'RR'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    If Not rsMATREC.EOF And Not rsMATREC.BOF Then
        labid.Caption = rsMATREC!ID
        txtRRNo.Text = Null2String(rsMATREC!rrno)
        txtRRDate.Text = Null2String(rsMATREC!rrdate)
        cboClasscode.Text = Null2String(rsMATREC!classcode)
        txtRecvd_Code.Text = Null2String(rsMATREC!recvd_code)
        cboRecvd_Desc.Text = SetSupdesc(Null2String(rsMATREC!recvd_code))
        txtDetails.Text = Null2String(rsMATREC!Address)
        txtTerms.Text = Null2String(rsMATREC!terms)
        txtPONo.Text = Null2String(rsMATREC!pono)
        txtPODate.Text = Null2String(rsMATREC!podate)
        txtDRNo.Text = Null2String(rsMATREC!drno)
        txtINVNo.Text = Null2String(rsMATREC!invno)
        txtDS1.Text = N2Str2IntZero(rsMATREC!ds1)
        txtDS_Desc1.Text = Null2String(rsMATREC!ds_desc1)
        txtDS_Amt1.Text = ToDoubleNumber(N2Str2Zero(rsMATREC!ds_amt1))
        txtTTLRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATREC!ttlrramt))
        txtNetRRAmt.Text = ToDoubleNumber(N2Str2Zero(rsMATREC!netrramt))
        txtRemarks.Text = ToDoubleNumber(Null2String(rsMATREC!remarks))
        If Null2String(rsMATREC!Status) = "P" Then
            labRRsted.Visible = True
            labRRsted.Caption = "POSTED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
            If LOGLEVEL = "ADM" Then cmdCancelRR.Enabled = True
        ElseIf Null2String(rsMATREC!Status) = "C" Then
            labRRsted.Visible = True
            labRRsted.Caption = "CANCELLED"
            cmdEdit.Enabled = False
            cmdPost.Enabled = False
            cmdUnPost.Enabled = False
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
    On Error GoTo Errorcode
    Pcnt = 0
    MATrec_TOTUCOST = 0
    MATrec_TOTINVAMT = 0
    MATrec_TOTVAT = 0
    RR_QTY_REC = 0
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,trantype,tranno,itemno,matcde,matdsc,tranqty,tranucost,traninvamt from CSMS_TdayTran where trantype = 'RR' and tranno = " & N2Str2Null(rsMATREC!rrno) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        Screen.MousePointer = 11
        rsTDAYTRAN.MoveFirst
        Do While Not rsTDAYTRAN.EOF
            Pcnt = Pcnt + 1
            grdDetails.AddItem rsTDAYTRAN!ID & Chr(9) & Format(Null2String(rsTDAYTRAN!itemno), "0000") & Chr(9) & _
                               Null2String(rsTDAYTRAN!MATCDE) & Chr(9) & _
                               Null2String(rsTDAYTRAN!MatDsc) & Chr(9) & _
                               N2Str2IntZero(rsTDAYTRAN!tranqty) & Chr(9) & _
                               N2Str2Zero(rsTDAYTRAN!TRANINVAMT) & Chr(9) & _
                               N2Str2Zero(rsTDAYTRAN!TRANUCOST) & Chr(9) & _
                               Format(N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUCOST), MAXIMUM_DIGIT)
            RR_QTY_REC = RR_QTY_REC + N2Str2IntZero(rsTDAYTRAN!tranqty)
            MATrec_TOTUCOST = MATrec_TOTUCOST + (N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANUCOST))
            MATrec_TOTINVAMT = MATrec_TOTINVAMT + (N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANINVAMT))
            rsTDAYTRAN.MoveNext
        Loop
        If Pcnt <> 0 Then grdDetails.RemoveItem 1
        If Null2String(rsMATREC!classcode) = "PCS" Or Null2String(rsMATREC!classcode) = "PCG" Then
            If ISNONVAT = True Then
                MATrec_TOTVAT = 0
            Else
                MATrec_TOTVAT = (MATrec_TOTUCOST * ConvertToBIRDecimalFormat(VAT_RATE)) - MATrec_TOTUCOST
            End If
        Else
            MATrec_TOTVAT = 0
        End If
        If NumericVal(MATrec_TOTVAT) <> 0 Then
            txtDS1.Text = VAT_RATE
            txtDS_Desc1.Text = "VAT"
            txtDS_Amt1.Text = ToDoubleNumber(MATrec_TOTVAT)
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

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Function Setmatdsc(ppp As String)
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select matcde,matdsc from CSMS_MatMas where matcde= '" & ppp & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        Setmatdsc = Null2String(rsMatMas!MatDsc)
    End If
End Function

Function Setmatdsc2(ppp As String)
    If ppp <> "" Then
        Set rsMatMas = New ADODB.Recordset
        rsMatMas.Open "Select id,matdsc from CSMS_MatMas where id = " & ppp, gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsMatMas.EOF And Not rsMatMas.BOF Then
            Setmatdsc2 = Null2String(rsMatMas!MatDsc)
        End If
    End If
End Function

Function Setmatcde(DDD As String)
    Set rsMatMas = New ADODB.Recordset
    rsMatMas.Open "Select id,matcde from CSMS_MatMas where id = " & DDD, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMatMas.EOF And Not rsMatMas.BOF Then
        Setmatcde = Null2String(rsMatMas!MATCDE)
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
    rsMatMas.Open "Select id,matdsc from CSMS_MatMas where ltrim(rtrim(matdsc)) = '" & LTrim(RTrim(DDD)) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
    Set rsTDAYTRAN = New ADODB.Recordset
    rsTDAYTRAN.Open "select id,itemno,matcde,matdsc,tranqty,traninvamt,tranucost from CSMS_TdayTran where id = " & ID, gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsTDAYTRAN.EOF And Not rsTDAYTRAN.BOF Then
        labDetId.Caption = rsTDAYTRAN!ID
        txtTranItemNo.Text = Null2String(rsTDAYTRAN!itemno)
        cboTranMatCde.Text = Null2String(rsTDAYTRAN!MATCDE)
        cboTranMatDsc.Text = Null2String(rsTDAYTRAN!MatDsc)
        txtTranQty.Text = N2Str2IntZero(rsTDAYTRAN!tranqty)
        txtTranINVAmt.Text = ToDoubleNumber(N2Str2Zero(rsTDAYTRAN!TRANINVAMT))
        txtUnitCost.Text = ToDoubleNumber(N2Str2Zero(rsTDAYTRAN!TRANUCOST))
        txtTranTotalAmt.Text = ToDoubleNumber(N2Str2IntZero(rsTDAYTRAN!tranqty) * N2Str2Zero(rsTDAYTRAN!TRANINVAMT))
    End If
End Function

Private Sub grdDetails_DblClick()
    Dim Fild                                           As String
    If Null2String(rsMATREC!Status) = "P" Then
        MsgSpeechBox "Item(s) are Already Posted, and cannot be change..."
    ElseIf Null2String(rsMATREC!Status) = "C" Then
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
        rsCunter.Open "select * from CSMS_Cunter where modul = 'PO'", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsCunter.EOF And Not rsCunter.BOF Then
            txtPONo.Text = Format(N2Str2Zero(rsCunter!nextnumber) - 1, "000000")
        End If
    End If
End Sub

Private Sub txtPONo_LostFocus()
    If cboClasscode.Text = "PCG" Then
        If txtPONo.Text <> "" And AddorEdit = "ADD" And Len(txtPONo.Text) > 0 Then
            Dim rsMATRecDup                            As ADODB.Recordset
            Set rsMATRecDup = New ADODB.Recordset
            rsMATRecDup.Open "select pono from CSMS_MatIss where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not rsMATRecDup.EOF And Not rsMATRecDup.BOF Then
                MsgSpeechBox "PO Number Already Received"
                Exit Sub
            End If
            Set rsPO_HD = New ADODB.Recordset
            rsPO_HD.Open "select pono,supcode,podate from PMIS_PO_Hd where pono = '" & txtPONo.Text & "'", gconDMIS
            If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
                txtRecvd_Code.Text = Null2String(rsPO_HD!SupCode)
                txtPODate.Text = Null2String(rsPO_HD!podate)
                Pcnt = 0
                MATrec_TOTUCOST = 0
                MATrec_TOTINVAMT = 0
                MATrec_TOTVAT = 0
                RR_QTY_REC = 0
                Dim rsTdaytranDup                      As ADODB.Recordset
                Set rsTdaytranDup = New ADODB.Recordset
                rsTdaytranDup.Open "select id,itemno,matcde,matdsc,tranqty,traninvamt,tranucost from CSMS_TdayTran where trantype = 'PO' and tranno = " & N2Str2Null(rsPO_HD!pono) & " order by itemno asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsTdaytranDup.EOF And Not rsTdaytranDup.BOF Then
                    Screen.MousePointer = 11
                    rsTdaytranDup.MoveFirst
                    cleargrid grdDetails
                    Do While Not rsTdaytranDup.EOF
                        Pcnt = Pcnt + 1
                        grdDetails.AddItem rsTdaytranDup!ID & Chr(9) & Null2String(Format(rsTdaytranDup!itemno, "0000")) & Chr(9) & _
                                           Null2String(rsTdaytranDup!MATCDE) & Chr(9) & _
                                           Setmatdsc(Null2String(rsTdaytranDup!MatDsc)) & Chr(9) & _
                                           N2Str2IntZero(rsTdaytranDup!tranqty) & Chr(9) & _
                                           N2Str2Zero(rsTdaytranDup!TRANINVAMT) & Chr(9) & _
                                           N2Str2Zero(rsTdaytranDup!TRANUCOST) & Chr(9) & _
                                           N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANUCOST)
                        MATrec_TOTUCOST = MATrec_TOTUCOST + (N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANUCOST))
                        MATrec_TOTINVAMT = MATrec_TOTINVAMT + (N2Str2IntZero(rsTdaytranDup!tranqty) * N2Str2Zero(rsTdaytranDup!TRANINVAMT))
                        rsTdaytranDup.MoveNext
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
    If Null2String(rsMATREC!classcode) = "PCS" Or Null2String(rsMATREC!classcode) = "PCG" Then
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
    If Null2String(rsMATREC!rrno) = "PCS" Or Null2String(rsMATREC!rrno) = "PCG" Then
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
    rsMATREC.Bookmark = rsFind(rsMATREC.Clone, "ID", lstMATREC.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstMATREC_ItemClick(ByVal item As MSComctlLib.ListItem)
    If optRRNo.Value = True Then
        rsMATREC.Bookmark = rsFind(rsMATREC.Clone, "rrno", item).Bookmark
    Else
        rsMATREC.Bookmark = rsFind(rsMATREC.Clone, "ID", lstMATREC.SelectedItem.SubItems(1)).Bookmark
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
    If KeyCode = vbKeyEscape Then: On Error Resume Next: textSearch.SetFocus
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
    Dim rsMATREC                                       As ADODB.Recordset
    lstMATREC.Enabled = False
    lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
    Set rsMATREC = New ADODB.Recordset
    Set rsMATREC = gconDMIS.Execute("select rrno,ID from CSMS_MatRec order by rrno asc")
    If Not (rsMATREC.EOF And rsMATREC.BOF) Then
        lstMATREC.Enabled = True
        Listview_Loadval Me.lstMATREC.ListItems, rsMATREC
        lstMATREC.Refresh
        lstMATREC.Enabled = True
    Else
        lstMATREC.Enabled = False
    End If
End Sub

Sub FillSearchGrid(xxx As String)
    Dim rsMATREC                                       As ADODB.Recordset
    lstMATREC.Enabled = False
    lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
    Set rsMATREC = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsMATREC = gconDMIS.Execute("select rrno, ID from CSMS_MatRec where rrno like'" & xxx & "%'")
    If Not (rsMATREC.EOF And rsMATREC.BOF) Then
        lstMATREC.Enabled = True
        Listview_Loadval Me.lstMATREC.ListItems, rsMATREC
        lstMATREC.Refresh
        lstMATREC.Enabled = True
    Else
        lstMATREC.Enabled = False
    End If
End Sub

Sub FillGrid2()
    Dim rsMATREC                                       As ADODB.Recordset
    lstMATREC.Enabled = False
    lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
    Set rsMATREC = New ADODB.Recordset
    Set rsMATREC = gconDMIS.Execute("select recvd_from, ID from CSMS_MatRec order by rrno asc")
    If Not (rsMATREC.EOF And rsMATREC.BOF) Then
        lstMATREC.Enabled = True
        Listview_Loadval Me.lstMATREC.ListItems, rsMATREC
        lstMATREC.Refresh
        lstMATREC.Enabled = True
    Else
        lstMATREC.Enabled = False
    End If
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsMATREC                                       As ADODB.Recordset
    lstMATREC.Sorted = False: lstMATREC.ListItems.Clear
    lstMATREC.Enabled = False
    Set rsMATREC = New ADODB.Recordset
    xxx = Repleys(LTrim(RTrim(xxx)))
    Set rsMATREC = gconDMIS.Execute("select recvd_from, ID from CSMS_MatRec where recvd_from like '" & xxx & "%' order by rrno asc")
    If Not (rsMATREC.EOF And rsMATREC.BOF) Then
        lstMATREC.Enabled = True
        Listview_Loadval Me.lstMATREC.ListItems, rsMATREC
        lstMATREC.Refresh
        lstMATREC.Enabled = True
    Else
        lstMATREC.Enabled = False
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstMATREC.Enabled = True And lstMATREC.ListItems.Count > 0 Then
            lstMATREC.SetFocus
        End If
    End If
End Sub

Private Sub optRONo_Click()
    lstMATREC.ColumnHeaders(1).Text = "Sup. Name"
    lstMATREC.ColumnHeaders(1).Width = 4000
    If textSearch = "" Then FillGrid2 Else FillSearchGrid2 (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub optRRNo_Click()
    lstMATREC.ColumnHeaders(1).Text = "Tran. No."
    lstMATREC.ColumnHeaders(1).Width = 2150
    If textSearch = "" Then FillGrid Else FillSearchGrid (textSearch.Text)
    On Error Resume Next
    textSearch.SetFocus
End Sub
