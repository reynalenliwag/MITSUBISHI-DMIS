VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmCMISPettyCash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Petty Cash Data Entry"
   ClientHeight    =   7620
   ClientLeft      =   270
   ClientTop       =   1740
   ClientWidth     =   12840
   ForeColor       =   &H8000000F&
   Icon            =   "PettyCash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7620
   ScaleWidth      =   12840
   Begin VB.PictureBox picLiquidate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   13020
      ScaleHeight     =   765
      ScaleWidth      =   3075
      TabIndex        =   18
      Top             =   6690
      Width           =   3105
      Begin VB.CommandButton cmdPaymentCA 
         Caption         =   "Payment of Cash Advances"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         MouseIcon       =   "PettyCash.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   390
         Width           =   3015
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Normal Liquidation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   30
         MouseIcon       =   "PettyCash.frx":0A1C
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   30
         Width           =   3015
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   3780
      ScaleHeight     =   900
      ScaleWidth      =   13425
      TabIndex        =   21
      Top             =   6690
      Width           =   13425
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
         Left            =   8235
         MouseIcon       =   "PettyCash.frx":0B6E
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":0CC0
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Exit Window"
         Top             =   45
         Width           =   705
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
         Left            =   7545
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PettyCash.frx":1026
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":1178
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Cancel this Transaction"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPOST 
         Caption         =   "Liquidate"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   6855
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PettyCash.frx":14B2
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":1604
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Liquidate"
         Top             =   45
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Tag"
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
         Left            =   6165
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "PettyCash.frx":196A
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":1ABC
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Tag"
         Top             =   45
         Width           =   705
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
         Left            =   5475
         MouseIcon       =   "PettyCash.frx":1D9B
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":1EED
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Edit Selected Record"
         Top             =   45
         Width           =   705
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
         Left            =   4785
         MouseIcon       =   "PettyCash.frx":2249
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":239B
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   4095
         MouseIcon       =   "PettyCash.frx":26AE
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":2800
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Left            =   3405
         MouseIcon       =   "PettyCash.frx":2B50
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":2CA2
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Left            =   2715
         MouseIcon       =   "PettyCash.frx":3000
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":3152
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   2025
         MouseIcon       =   "PettyCash.frx":344C
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":359E
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   1330
         MouseIcon       =   "PettyCash.frx":38F6
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":3A48
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Previous Record"
         Top             =   45
         Width           =   705
      End
      Begin VB.Label labid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   95
         Top             =   180
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   7515
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   3705
      Begin VB.TextBox textSearch 
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   1
         Text            =   "KIM LIM"
         Top             =   210
         Width           =   3525
      End
      Begin MSComctlLib.ListView lstPetty 
         Height          =   6735
         Left            =   60
         TabIndex        =   2
         Top             =   630
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   11880
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
         MouseIcon       =   "PettyCash.frx":3DA7
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.CommandButton cmdTag 
      Caption         =   "F4 - Tag/UnTag"
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
      Height          =   375
      Left            =   14220
      MouseIcon       =   "PettyCash.frx":3F09
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1740
      Width           =   1845
   End
   Begin VB.CommandButton cmdLiquidate 
      Caption         =   "F4 - Liquidate"
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
      Height          =   375
      Left            =   14190
      MouseIcon       =   "PettyCash.frx":405B
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2130
      Width           =   1845
   End
   Begin VB.CommandButton cmdReplenish 
      Caption         =   "F6 - Replenish"
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
      Height          =   375
      Left            =   14160
      MouseIcon       =   "PettyCash.frx":41AD
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2550
      Width           =   1845
   End
   Begin VB.PictureBox picStatus 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1335
      Left            =   3900
      ScaleHeight     =   1335
      ScaleWidth      =   8865
      TabIndex        =   4
      Top             =   5415
      Width           =   8865
      Begin VB.TextBox txtDistParticulars 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1110
         TabIndex        =   89
         Top             =   840
         Width           =   7695
      End
      Begin VB.TextBox txtLiquidated 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3510
         TabIndex        =   9
         Top             =   60
         Width           =   2265
      End
      Begin VB.TextBox txtLiq_Date 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7320
         TabIndex        =   8
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox txtTotalPettyCash 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   60
         Width           =   1485
      End
      Begin VB.TextBox txtLiq_Amt 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1110
         TabIndex        =   6
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox txtPetty_CashNo 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1110
         TabIndex        =   5
         Top             =   60
         Width           =   1485
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -780
         TabIndex        =   88
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Status  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1620
         TabIndex        =   14
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5430
         TabIndex        =   13
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Petty Cash  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5430
         TabIndex        =   12
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidation  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -780
         TabIndex        =   11
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PCV No.  :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -780
         TabIndex        =   10
         Top             =   90
         Width           =   1845
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   3870
      ScaleHeight     =   1365
      ScaleWidth      =   9060
      TabIndex        =   41
      Top             =   30
      Width           =   9060
      Begin VB.TextBox cboEmployee 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   420
         Left            =   2910
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   840
         Width           =   5925
      End
      Begin VB.TextBox txtEmployeeCode 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   420
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   840
         Width           =   2685
      End
      Begin VB.OptionButton optReplenish 
         Caption         =   "Show Cash Expense (Replenished)"
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
         Height          =   660
         Left            =   5970
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   90
         Width           =   2865
      End
      Begin VB.OptionButton optExpense 
         Caption         =   "Show Cash Expense (Liquidated)"
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
         Height          =   660
         Left            =   2880
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   90
         Width           =   2865
      End
      Begin VB.OptionButton optCashAdvance 
         Caption         =   "Show Cash Advance (Un-Liquidated)"
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
         Height          =   645
         Left            =   0
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   90
         Value           =   -1  'True
         Width           =   2685
      End
      Begin VB.CommandButton cmdRefresh 
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
         Height          =   420
         Left            =   8535
         MouseIcon       =   "PettyCash.frx":42FF
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":4451
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Refresh"
         Top             =   -600
         Width           =   480
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "SHOW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8175
         MouseIcon       =   "PettyCash.frx":477C
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   -630
         Width           =   855
      End
      Begin VB.TextBox txtPetty_Code 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00701E2A&
         Height          =   390
         Left            =   2400
         TabIndex        =   45
         Top             =   -660
         Width           =   765
      End
      Begin VB.ComboBox cboPetty_Type 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00973640&
         Height          =   390
         Left            =   3210
         TabIndex        =   42
         Text            =   "cboPetty_Type"
         Top             =   -660
         Width           =   4875
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PCV Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   60
         TabIndex        =   50
         Top             =   -600
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         TabIndex        =   49
         Top             =   -570
         Width           =   2115
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2250
         TabIndex        =   48
         Top             =   -480
         Width           =   195
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1410
         TabIndex        =   47
         Top             =   -510
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   11220
      ScaleHeight     =   885
      ScaleWidth      =   1980
      TabIndex        =   33
      Top             =   6690
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
         Left            =   795
         MouseIcon       =   "PettyCash.frx":48CE
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":4A20
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Cancel"
         Top             =   45
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
         Left            =   105
         MouseIcon       =   "PettyCash.frx":4D5E
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":4EB0
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Save this Record"
         Top             =   45
         Width           =   705
      End
   End
   Begin VB.PictureBox picPettyCashEntry 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   4350
      ScaleHeight     =   2865
      ScaleWidth      =   7875
      TabIndex        =   51
      Top             =   2115
      Width           =   7905
      Begin VB.ComboBox cboReplenishment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   330
         Left            =   1710
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox txtParticulars 
         BackColor       =   &H00C0FFFF&
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
         Height          =   675
         Left            =   1710
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   54
         Top             =   1110
         Width           =   4575
      End
      Begin VB.CommandButton cmdDeletePetty 
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
         Left            =   150
         MouseIcon       =   "PettyCash.frx":5200
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":5352
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1875
         Width           =   705
      End
      Begin VB.TextBox txtPCF_NUMBER 
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
         Height          =   345
         Left            =   2760
         TabIndex        =   57
         Top             =   1890
         Width           =   2025
      End
      Begin VB.TextBox txtOriginal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   345
         Left            =   6390
         TabIndex        =   55
         Text            =   "0.00"
         Top             =   690
         Width           =   1455
      End
      Begin VB.TextBox txtPetty_Date 
         BackColor       =   &H00C0FFFF&
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
         Height          =   345
         Left            =   150
         TabIndex        =   52
         Top             =   690
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelPetty 
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
         Left            =   7050
         MouseIcon       =   "PettyCash.frx":567D
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":57CF
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1890
         Width           =   705
      End
      Begin VB.CommandButton cmdSavePetty 
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
         Left            =   6360
         MouseIcon       =   "PettyCash.frx":5B0D
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":5C5F
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label Label20 
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
         Left            =   180
         TabIndex        =   87
         Top             =   1200
         Width           =   1635
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   315
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Width           =   7875
         _Version        =   655364
         _ExtentX        =   13891
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   " Petty Cash Entry Box"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorLight=   12632256
         GradientColorDark=   4210752
         ForeColor       =   16777215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "PCV No.  :"
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
         Left            =   1680
         TabIndex        =   64
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label labPettyID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   2130
         TabIndex        =   63
         Top             =   -360
         Width           =   1185
      End
      Begin VB.Label labelOrgi 
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
         Left            =   6420
         TabIndex        =   62
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Replenishment"
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
         Left            =   1740
         TabIndex        =   61
         Top             =   390
         Width           =   1785
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   180
         TabIndex        =   60
         Top             =   390
         Width           =   1185
      End
   End
   Begin VB.PictureBox picBreakDown 
      BackColor       =   &H00C0C0C0&
      Height          =   5595
      Left            =   4050
      ScaleHeight     =   5535
      ScaleWidth      =   8445
      TabIndex        =   66
      Top             =   900
      Width           =   8505
      Begin VB.TextBox txtParticularsBD 
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
         Height          =   855
         Left            =   1380
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   71
         Top             =   3270
         Width           =   6915
      End
      Begin VB.TextBox txtCashAdvance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   4050
         TabIndex        =   75
         Top             =   4770
         Width           =   1845
      End
      Begin VB.CommandButton cmdDeleteBreakDown 
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
         Left            =   135
         MouseIcon       =   "PettyCash.frx":5FAF
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":6101
         Style           =   1  'Graphical
         TabIndex        =   74
         ToolTipText     =   "Delete Entry"
         Top             =   4635
         Width           =   705
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   60
         ScaleHeight     =   435
         ScaleWidth      =   8235
         TabIndex        =   77
         Top             =   4170
         Width           =   8235
         Begin VB.TextBox txtBDPCF_NUMBER 
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
            TabIndex        =   72
            Top             =   30
            Width           =   1965
         End
         Begin VB.TextBox txtTotalCashAdvance 
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
            Left            =   6480
            Locked          =   -1  'True
            TabIndex        =   73
            Text            =   "0.00"
            Top             =   30
            Width           =   1605
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "PCV No.  :"
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
            Left            =   -780
            TabIndex        =   79
            Top             =   60
            Width           =   1845
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount   :"
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
            Left            =   4530
            TabIndex        =   78
            Top             =   60
            Width           =   1845
         End
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "Insert"
         Height          =   345
         Left            =   7620
         TabIndex        =   70
         Top             =   330
         Width           =   675
      End
      Begin VB.TextBox txtBDPetty_Date 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   60
         TabIndex        =   67
         Top             =   330
         Width           =   1455
      End
      Begin VB.ComboBox cboBDReplenishment 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   1590
         TabIndex        =   68
         Text            =   "cboBDReplenishment"
         Top             =   330
         Width           =   4455
      End
      Begin VB.TextBox txtBDPetty_Cash 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6090
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   300
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdBreakDown 
         Height          =   2445
         Left            =   60
         TabIndex        =   81
         Top             =   750
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   4313
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Date           | Code  |    Replenishment            |       Account Code      |  Amount      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "PettyCash.frx":642C
      End
      Begin VB.CommandButton cmdCancelBreakDown 
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
         Left            =   7470
         MouseIcon       =   "PettyCash.frx":6746
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":6898
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Cancel Entry"
         Top             =   4635
         Width           =   705
      End
      Begin VB.CommandButton cmdSaveBreakDown 
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
         Left            =   6765
         MouseIcon       =   "PettyCash.frx":6BD6
         MousePointer    =   99  'Custom
         Picture         =   "PettyCash.frx":6D28
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Save Entry"
         Top             =   4635
         Width           =   705
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars :"
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
         Left            =   120
         TabIndex        =   86
         Top             =   3330
         Width           =   1845
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount to be Liquidated."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1410
         TabIndex        =   85
         Top             =   4860
         Width           =   2460
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         TabIndex        =   84
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Replenishment"
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
         Left            =   1590
         TabIndex        =   83
         Top             =   60
         Width           =   1785
      End
      Begin VB.Label Label17 
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
         Left            =   6120
         TabIndex        =   82
         Top             =   60
         Width           =   1185
      End
   End
   Begin VB.PictureBox picNORMAL_LIQUIDATE 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   4350
      ScaleHeight     =   1215
      ScaleWidth      =   8040
      TabIndex        =   36
      Top             =   1680
      Visible         =   0   'False
      Width           =   8070
      Begin VB.OptionButton optBREAKDOWN 
         Caption         =   "BREAKDOWN"
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   375
         Width           =   2565
      End
      Begin VB.OptionButton optNORMAL 
         Caption         =   "NORMAL"
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
         Left            =   2730
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   375
         Width           =   2565
      End
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
         Left            =   5385
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   375
         Width           =   2565
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   315
         Left            =   0
         TabIndex        =   40
         Top             =   -30
         Width           =   8055
         _Version        =   655364
         _ExtentX        =   14208
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "  Select Button to liquidate cash advances ..."
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
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
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   3810
      ScaleHeight     =   4095
      ScaleWidth      =   9015
      TabIndex        =   3
      Top             =   1305
      Width           =   9015
      Begin MSFlexGridLib.MSFlexGrid grdPetty 
         Height          =   3915
         Left            =   30
         TabIndex        =   94
         Top             =   90
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   16777215
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   "  Date           | Code       |    Replenishment                |   Amount   | T   | R   | Balance         "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "PettyCash.frx":7078
      End
   End
End
Attribute VB_Name = "frmCMISPettyCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Employee                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim PrevPettyCash                                                     As Double
Dim GridToLiquidate                                                   As Long
Dim TotalBreakDownCA                                                  As Double

Sub DisabledOut(COND As Boolean)
    fraDetails.Enabled = COND
    Picture1.Enabled = COND
End Sub

Function SetPettyTypeDesc(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_CBOOK Where BOOK = 'J' and CODE = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetPettyTypeDesc = Null2String(rsSBOOK!DESCNAME)
    End If
    Set rsSBOOK = Nothing
End Function

Function SetEmployeeCode(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_vw_Vemployee Where BOOK = 'I' and DESCNAME = '" & XXX & "'")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetEmployeeCode = Null2String(rsSBOOK!Code)
    End If
    Set rsSBOOK = Nothing
End Function

'Function SetPettyType(XXX As Variant)
    'Dim rsSBOOK As ADODB.Recordset
    'Set rsSBOOK = New ADODB.Recordset
    'Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'E' and CODE = '"  &   XXX  &  "'")
    'If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
    '   SetPettyType = Null2String(rsSBOOK!DESCNAME)
    'End If
    'Set rsSBOOK = Nothing
'End Function

Function SetPettyCode(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_SBOOK Where BOOK = 'E' and DESCNAME = " & N2Str2Null(XXX))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetPettyCode = Null2String(rsSBOOK!Code)
    End If
    Set rsSBOOK = Nothing
End Function

Function SetReplenishCode(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_CBOOK Where BOOK = 'J' and DESCNAME = " & N2Str2Null(XXX))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetReplenishCode = Null2String(rsSBOOK!Code)
    End If
    Set rsSBOOK = Nothing
End Function

Function SetReplenishAcctCode(XXX As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select * from CMIS_CBOOK Where BOOK = 'J' and DESCNAME = " & N2Str2Null(XXX))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetReplenishAcctCode = Null2String(rsSBOOK!CHARTCODES)
    End If
    Set rsSBOOK = Nothing
End Function

Sub BreakDown()
    InitBDGrid
    txtTotalCashAdvance.Text = "0.00": TotalBreakDownCA = 0
    picNORMAL_LIQUIDATE.ZOrder 1: picNORMAL_LIQUIDATE.Visible = False
    'cmdBreakDown.Visible = True: cmdBreakDown.ZOrder 0
    picBreakDown.Visible = True: picBreakDown.ZOrder 0
    On Error Resume Next
    txtBDPetty_Date.SetFocus
    txtBDPetty_Date.Text = LOGDATE
    Dim rsPCF_NUMBER                                                  As ADODB.Recordset
    Set rsPCF_NUMBER = New ADODB.Recordset
    
'orig code: for HPI para automatoc ang PCV no.
'    Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '001' order by PCF_NUMBER DESC")
'    If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
'        txtBDPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
'    Else
'        txtBDPCF_NUMBER.Text = "000001"
'    End If
'    Set rsPCF_NUMBER = Nothing
    
    
    If COMPANY_CODE = "HPI" Then
        
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '002' order by PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtBDPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)), "000000")
        Else
            txtBDPCF_NUMBER.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing
            
    Else
    
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '001' order by PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtBDPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
        Else
            txtBDPCF_NUMBER.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing

    End If
    


End Sub

Sub rsRefresh()
    Set Employee = New ADODB.Recordset
    Call Employee.Open("Select * from CMIS_vw_Vemployee Where BOOK = 'I' order by DESCNAME asc", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub

Sub initMemvars()
    txtEmployeeCode.Text = ""
    cboEmployee.Text = ""
End Sub

Sub InitPettyMemVars()
    cmdDeletePetty.Visible = False
    AddorEdit = "ADD"
    txtPetty_Date.Text = LOGDATE
    cboReplenishment.ListIndex = -1
    txtParticulars.Text = ""
    txtOriginal.Text = "0.00"
    Dim rsPCF_NUMBER                                                  As ADODB.Recordset
    If txtPetty_Code.Text = "001" Then
        Set rsPCF_NUMBER = New ADODB.Recordset
        
        If COMPANY_CODE = "HPI" Then
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '002' ORDER BY PCF_NUMBER DESC")
                   If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
        
        Else
            
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '001' ORDER BY PCF_NUMBER DESC")
            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
        
        End If
    
        Set rsPCF_NUMBER = Nothing
    Else
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '002' ORDER BY PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtPCF_NUMBER.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER) + 1, "000000")
        Else
            txtPCF_NUMBER.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing
    End If
    On Error Resume Next
    txtPetty_Date.SetFocus
End Sub

Sub FillCboPettyType()
    Dim rsSBook2                                                      As ADODB.Recordset
    Set rsSBook2 = New ADODB.Recordset
    Set rsSBook2 = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where BOOK = 'E' order by DESCNAME ASC")
    If Not rsSBook2.EOF And Not rsSBook2.BOF Then
        Combo_Loadval cboPetty_Type, rsSBook2
    End If
    Set rsSBook2 = Nothing
End Sub

Sub FillCboReplenishment()
    Dim rsSBook2                                                      As ADODB.Recordset
    Set rsSBook2 = New ADODB.Recordset
    Set rsSBook2 = gconDMIS.Execute("Select DESCNAME from CMIS_CBOOK Where BOOK = 'J' order by DESCNAME ASC")
    If Not rsSBook2.EOF And Not rsSBook2.BOF Then
        Combo_Loadval cboReplenishment, rsSBook2
    End If
    Set rsSBook2 = New ADODB.Recordset
    Set rsSBook2 = gconDMIS.Execute("Select DESCNAME from CMIS_CBOOK Where BOOK = 'J' order by DESCNAME ASC")
    If Not rsSBook2.EOF And Not rsSBook2.BOF Then
        Combo_Loadval cboBDReplenishment, rsSBook2
    End If
    Set rsSBook2 = Nothing
End Sub

Sub StoreMemVars()
    If Not Employee.EOF And Not Employee.BOF Then
        txtEmployeeCode.Text = Null2String(Employee!Code)
        cboEmployee.Text = Null2String(Employee!DESCNAME)
        cboPetty_Type.ListIndex = -1
        InitGrid
        txtPetty_CashNo.Text = "": txtLiq_Amt.Text = ""
        txtLiq_Date.Text = "": txtLiquidated.Text = ""
    End If
End Sub

Sub StoreDetails()
    Dim TOTAL_PETTY_CASH                                              As Double
    Dim i                                                             As Integer
    Dim Tag, Repl                                                     As String

    cleargrid grdPetty: InitGrid: i = 0: TOTAL_PETTY_CASH = 0
    Dim rsPETTY                                                       As ADODB.Recordset
    Set rsPETTY = New ADODB.Recordset
    If txtPetty_Code.Text = "002" Then
        Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where (LIQUIDATED = 0 OR LIQUIDATED IS NULL) AND EMPLOYEE = '" & txtEmployeeCode.Text & "' and petty_code = '" & txtPetty_Code.Text & "' order by PETTY_DATE DESC,PCF_NUMBER DESC")
    ElseIf txtPetty_Code.Text = "001" Then
        If optExpense.Value = True Then
            Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where (REPLENISH = 0) AND EMPLOYEE = '" & txtEmployeeCode.Text & "' and petty_code = '" & txtPetty_Code.Text & "' order by PETTY_DATE DESC,PCF_NUMBER DESC")
        Else
            Set rsPETTY = gconDMIS.Execute("Select * from CMIS_Petty Where (REPLENISH = 1) AND EMPLOYEE = '" & txtEmployeeCode.Text & "' and petty_code = '" & txtPetty_Code.Text & "' order by PETTY_DATE DESC,PCF_NUMBER DESC")
        End If
    End If
    If Not rsPETTY.EOF And Not rsPETTY.BOF Then
        rsPETTY.MoveFirst: Screen.MousePointer = 11
        Do While Not rsPETTY.EOF
            i = i + 1
            If Null2Bool(rsPETTY!Tag) = True Then Tag = " T" Else Tag = ""
            If Null2Bool(rsPETTY!REPLENISH) = True Then Repl = " T" Else Repl = ""
            grdPetty.AddItem Format(Null2String(rsPETTY!PETTY_DATE), "MM/DD/YYYY") & Chr(9) & Null2String(rsPETTY!Petty_type) & Chr(9) & SetPettyTypeDesc(Null2String(rsPETTY!Petty_type)) & Chr(9) & ToDoubleNumber(Null2String(rsPETTY!original)) & Chr(9) & Tag & Chr(9) & Repl & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTY!PETTY_CASH)) & Chr(9) & rsPETTY!Id
            If i = 1 Then grdPetty.RemoveItem 1
            TOTAL_PETTY_CASH = TOTAL_PETTY_CASH + N2Str2Zero(rsPETTY!PETTY_CASH)
            rsPETTY.MoveNext
        Loop
        Screen.MousePointer = 0: txtTotalPettyCash.Text = ToDoubleNumber(TOTAL_PETTY_CASH)
    End If
    Set rsPETTY = Nothing
End Sub

Sub StorePettyMemVars(XXX As Variant)
    Dim rsPetty2                                                      As ADODB.Recordset
    Set rsPetty2 = New ADODB.Recordset
    Set rsPetty2 = gconDMIS.Execute("Select * from CMIS_Petty where id = " & XXX)
    If Not rsPetty2.EOF And Not rsPetty2.BOF Then
        AddorEdit = "EDIT": cmdDeletePetty.Visible = True
        labPettyID.Caption = rsPetty2!Id
        txtPetty_Date.Text = Null2Date(rsPetty2!PETTY_DATE)
        txtParticulars.Text = Null2String(rsPetty2!Particulars)
        If SetPettyTypeDesc(rsPetty2!Petty_type) <> "" Then
            cboReplenishment.Text = SetPettyTypeDesc(rsPetty2!Petty_type)
        Else
            cboReplenishment.ListIndex = -1
        End If
        If Null2Bool(rsPetty2!liquid) = True Then
            txtOriginal.Text = ToDoubleNumber(N2Str2Zero(rsPetty2!PETTY_CASH))
            PrevPettyCash = N2Str2Zero(rsPetty2!PETTY_CASH)
        Else
            txtOriginal.Text = ToDoubleNumber(N2Str2Zero(rsPetty2!original))
            PrevPettyCash = N2Str2Zero(rsPetty2!original)
        End If
        txtPCF_NUMBER.Text = Format(Null2String(rsPetty2!PCF_NUMBER), "000000")
    End If
    Set rsPetty2 = Nothing
End Sub

Sub InitGrid()
    cleargrid grdPetty
    grdPetty.FormatString = "  Date           | Code  |    Replenishment                |   Amount        | T   | R   | Balance         "
    grdPetty.ColWidth(7) = 1
End Sub

Sub InitBDGrid()
    cleargrid grdBreakDown
    grdBreakDown.FormatString = "  Date           | Code  |    Replenishment            |       Account Code      |  Amount      "
    grdBreakDown.ColWidth(5) = 1
End Sub

Sub FillGrid()
    Dim Employee2                                                     As ADODB.Recordset
    lstPetty.Sorted = False: lstPetty.ListItems.Clear
    lstPetty.Enabled = False
    Set Employee2 = New ADODB.Recordset
    Set Employee2 = gconDMIS.Execute("select DESCNAME,CODE ,ID from CMIS_vw_Vemployee where BOOK = 'I' order by DESCNAME asc")
    If Not (Employee2.EOF And Employee2.BOF) Then
        lstPetty.Enabled = True
        Listview_Loadval Me.lstPetty.ListItems, Employee2
        lstPetty.Refresh
        lstPetty.Enabled = True
    Else
        lstPetty.Enabled = False
    End If

    Set Employee2 = Nothing
End Sub

Sub FillSearchGrid(XXX As String)
    Dim Employee2                                                     As ADODB.Recordset
    lstPetty.Sorted = False: lstPetty.ListItems.Clear
    lstPetty.Enabled = False
    Set Employee2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    Set Employee2 = gconDMIS.Execute("select DESCNAME,CODE ,ID from CMIS_vw_Vemployee where BOOK = 'I' and DESCNAME like '" & XXX & "%' order by DESCNAME asc")
    If Not (Employee2.EOF And Employee2.BOF) Then
        lstPetty.Enabled = True
        Listview_Loadval Me.lstPetty.ListItems, Employee2
        lstPetty.Refresh
        lstPetty.Enabled = True
    Else
        lstPetty.Enabled = False
    End If

    Set Employee2 = Nothing
End Sub

Private Sub cboBDReplenishment_GotFocus()
    VBComBoBoxDroppedDown cboBDReplenishment
End Sub

Private Sub cboEmployee_Click()
    txtEmployeeCode.Text = SetEmployeeCode(cboEmployee.Text)
End Sub

Private Sub cboEmployee_GotFocus()
    'VBComBoBoxDroppedDown cboEmployee
End Sub

Private Sub cboEmployee_LostFocus()
    'If txtEmployeeCode.Text <> "" Then cboEmployee.Enabled = False
End Sub

Private Sub cboPetty_Type_Click()
    txtPetty_Code.Text = SetPettyCode(cboPetty_Type.Text)
End Sub

Private Sub cboReplenishment_GotFocus()
    VBComBoBoxDroppedDown cboReplenishment
     If COMPANY_CODE = "HPI" Then
            Dim rsPCF_NUMBER                                                  As ADODB.Recordset
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '002' ORDER BY PCF_NUMBER DESC")

            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER))
            Else
                txtPCF_NUMBER.Text = "000001"
            End If

     End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_Add", "TRANSACTION PETTY CASH ENTRY") = False Then Exit Sub

    If txtEmployeeCode.Text <> "" And txtPetty_Code.Text <> "" Then
        'cmdPettyCashEntry.Visible = True
        'cmdPettyCashEntry.ZOrder 0
        picPettyCashEntry.Visible = True
        picPettyCashEntry.ZOrder 0
        PrevPettyCash = 0
        InitPettyMemVars
    Else
        MsgBox "Select Employee and Petty Cash Type...", vbInformation, "Message"
    End If
End Sub

Private Sub cmdCancelBreakDown_Click()
    'cmdBreakDown.Visible = False: cmdBreakDown.ZOrder 1
    picBreakDown.Visible = False: picBreakDown.ZOrder 1
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "TRANSACTION PETTY CASH ENTRY") = False Then Exit Sub

    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:
    'Exit Sub
    'ErrorCode:
    '    ShowVBError
    'LogAudit "C", "PETTY CASH", txtPetty_Code
End Sub

Private Sub cmdCancelPetty_Click()
    AddorEdit = ""
    'cmdPettyCashEntry.Visible = False: cmdPettyCashEntry.ZOrder 1
    picPettyCashEntry.Visible = False: picPettyCashEntry.ZOrder 1
    
End Sub

Private Sub cmdDeleteBreakDown_Click()

    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:
    'Exit Sub
    'ErrorCode:
    '    ShowVBError
End Sub

Private Sub cmdDeletePetty_Click()
    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If MsgBox("Delete this Entry... Are you Sure?", vbQuestion + vbYesNo) = vbYes Then
        SQL_STATEMENT = "delete from CMIS_Petty Where id = " & labPettyID.Caption
        gconDMIS.Execute SQL_STATEMENT
                
        If txtPetty_Code.Text = "002" Then              'UN-LIQUIDATE
            NEW_LogAudit "XX", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid.Caption, "", "PCV NO: " & txtPCF_NUMBER & " - " & txtParticulars, "", labPettyID.Caption
        ElseIf txtPetty_Code.Text = "001" Then          'LIQUIDATE
            NEW_LogAudit "XX", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid.Caption, "", "PCV NO: " & txtPCF_NUMBER & " - " & cboReplenishment, "", labPettyID.Caption
        Else                                            'REPLENISH
        
        
        End If
                
        If txtPetty_Code.Text = "001" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " EXPENSE = EXPENSE - " & NumericVal(PrevPettyCash) & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If txtPetty_Code.Text = "002" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " ADVANCES = ADVANCES - " & NumericVal(PrevPettyCash) & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        
        ShowDeletedMsg
        StoreDetails
        cmdCancelPetty_Click
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", "TRANSACTION PETTY CASH ENTRY") = False Then Exit Sub

    If txtEmployeeCode.Text <> "" And txtPetty_Code.Text <> "" Then
        grdPetty.Col = 7
        If grdPetty.Text <> "" Then
            'If txtLiquidated.Text = "Liquidated" Then
            '   MsgBox "Petty Cash has been Liquidated... Edit Denied...", vbInformation, "Message"
            '   Exit Sub
            'End If
            'If txtLiquidated.Text = "Liquidation" Then
            '   MsgBox "This is a Liquidation... Edit Denied...", vbInformation, "Message"
            '   Exit Sub
            'End If
            If LOGLEVEL <> "ADMIN" Then
                If txtLiquidated.Text = "Replenish" Then
                    MsgBox "Expense has been Replenish... Edit Denied...", vbInformation, "Message"
                    Exit Sub
                End If
            End If
            'cmdPettyCashEntry.Visible = True: cmdPettyCashEntry.ZOrder 0
            picPettyCashEntry.Visible = True: picPettyCashEntry.ZOrder 0
            StorePettyMemVars grdPetty.Text
        End If
    Else
        MsgBox "Select Employee and Petty Cash Type...", vbInformation, "Message"
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    fraDetails.ZOrder 0
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    Employee.MoveFirst
    StoreMemVars
End Sub

Private Sub cmdInsert_Click()

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    grdBreakDown.Row = 1: grdBreakDown.Col = 1
    If grdBreakDown.Text = "" Then
        grdBreakDown.AddItem Format(txtBDPetty_Date.Text, "MM/DD/YYYY") & Chr(9) & SetReplenishCode(cboBDReplenishment.Text) & Chr(9) & cboBDReplenishment.Text & Chr(9) & SetReplenishAcctCode(cboReplenishment.Text) & Chr(9) & ToDoubleNumber(txtBDPetty_Cash.Text)
        grdBreakDown.RemoveItem 1
    Else
        grdBreakDown.AddItem Format(txtBDPetty_Date.Text, "MM/DD/YYYY") & Chr(9) & SetReplenishCode(cboBDReplenishment.Text) & Chr(9) & cboBDReplenishment.Text & Chr(9) & SetReplenishAcctCode(cboReplenishment.Text) & Chr(9) & ToDoubleNumber(txtBDPetty_Cash.Text)
    End If
    Dim ChatBino                                                      As Double
    TotalBreakDownCA = 0
    For ChatBino = 1 To grdBreakDown.Rows - 1
        grdBreakDown.Row = ChatBino: grdBreakDown.Col = 4
        TotalBreakDownCA = TotalBreakDownCA + NumericVal(grdBreakDown.Text)
    Next
    txtTotalCashAdvance.Text = ToDoubleNumber(TotalBreakDownCA)
    txtBDPetty_Date.Text = LOGDATE
    cboBDReplenishment.ListIndex = -1
    txtBDPetty_Cash.Text = "0.00"
    On Error Resume Next
    txtBDPetty_Date.SetFocus

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdLast_Click()
    Employee.MoveLast
    StoreMemVars
End Sub

Private Sub cmdLiquidate_Click()
    grdPetty.Col = 7:
    If grdPetty.Text <> "" Then
        GridToLiquidate = grdPetty.Text
        picLiquidate.ZOrder 0
        picLiquidate.Visible = True
        On Error Resume Next
        cmdNormal.SetFocus
    End If
End Sub

Private Sub cmdNext_Click()
    Employee.MoveNext
    If Employee.EOF Then
        Employee.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdNormal_Click()
    picLiquidate.Visible = False: picLiquidate.ZOrder 1
    picNORMAL_LIQUIDATE.ZOrder 0: picNORMAL_LIQUIDATE.Visible = True
    On Error Resume Next
    optBREAKDOWN.SetFocus
End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "TRANSACTION PETTY CASH ENTRY") = False Then Exit Sub

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If txtLiquidated.Text = "Liquidated" Then
        MsgBox "Advances is Fully Liquidated... Access Denied...", vbInformation, "Message"
        Exit Sub
    End If
    'If txtLiquidated.Text = "Liquidation" Then
    '   MsgBox "This is a Liquidation... Access Denied...", vbInformation, "Message"
    '   Exit Sub
    'End If
    If txtLiquidated.Text = "Replenish" Then
        'MsgBox "Expense has been Replenish... Access Denied...", vbInformation, "Message"
        If MsgBox("Expense has been Replenish... Would you like to Cancel this Replenishment?", vbQuestion + vbYesNo, "Already Replenished") = vbYes Then
            Dim GridID                                                As Long
            Dim Karabs                                                As Integer
            Dim PettyAmount                                           As Double
            For Karabs = 1 To grdPetty.Rows - 1
                grdPetty.Col = 7
                If grdPetty.Text <> "" Then
                    GridID = grdPetty.Text
                    grdPetty.Col = 6: PettyAmount = NumericVal(grdPetty.Text)
                    grdPetty.Col = 5
                    If grdPetty.Text = " T" Then
                        grdPetty.Col = 4
                        If grdPetty.Text <> " T" Then
                            grdPetty.Col = 5: grdPetty.Text = "": grdPetty.Col = 4: grdPetty.Text = " T"
                            SQL_STATEMENT = "update CMIS_Petty Set Tag = 1, Replenish = 0 Where id = " & GridID
                            gconDMIS.Execute SQL_STATEMENT
                            'NEW_LogAudit "LQ", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, GridID, "", "", "", ""
                            If txtPetty_Code.Text = "001" Then
                                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                                " EXPENSE = EXPENSE + " & PettyAmount & "," & _
                                                " REPLENISH = REPLENISH - " & PettyAmount & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                            End If
                        End If
                    End If
                End If
            Next
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    'MODIFIED BY KIM: FML 7/17/2007 (REMARKS TO DIRECTLY MAKE THE LUQUIDATION NORMAL)
    'If txtPetty_Code.Text = "002" Then cmdLiquidate.Value = True
    grdPetty.Col = 7:
    GridToLiquidate = grdPetty.Text
    If txtPetty_Code.Text = "002" Then
        grdPetty.Col = 6: txtCashAdvance.Text = NumericVal(grdPetty.Text)
        Call BreakDown
    End If
    If txtPetty_Code.Text = "001" Then cmdReplenish.Value = True

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdPrevious_Click()
    Employee.MovePrevious
    If Employee.BOF Then
        Employee.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
    If txtPetty_Code.Text = "001" Then
        cmdTag.Value = True
    End If
End Sub

Private Sub cmdRefresh_Click()
    Picture1.Enabled = False
    cboEmployee.Enabled = True
    'fraDetails.Enabled = True
    On Error Resume Next
    cboEmployee.SetFocus
    VBComBoBoxDroppedDown cboEmployee
End Sub

Private Sub cmdReplenish_Click()

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Dim GridID                                                        As Long
    Dim Karabs                                                        As Integer
    Dim PettyAmount                                                   As Double
    For Karabs = 1 To grdPetty.Rows - 1
        grdPetty.Col = 7
        If grdPetty.Text <> "" Then
            GridID = grdPetty.Text
            grdPetty.Col = 6: PettyAmount = NumericVal(grdPetty.Text)
            grdPetty.Col = 4
            If grdPetty.Text = " T" Then
                grdPetty.Col = 5
                If grdPetty.Text <> " T" Then
                    grdPetty.Col = 4: grdPetty.Text = "": grdPetty.Col = 5: grdPetty.Text = " T"
                    gconDMIS.Execute ("update CMIS_Petty Set Tag = 0, Replenish = 1 Where id = " & GridID)
                    If txtPetty_Code.Text = "001" Then
                        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                        " EXPENSE = EXPENSE - " & PettyAmount & "," & _
                                        " REPLENISH = REPLENISH + " & PettyAmount & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                    End If
                End If
            End If
        End If
    Next
    LogAudit "R", "PETTY CASH - REPLENISH", txtPetty_Code
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSaveBreakDown_Click()

    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    Dim vLIQEmployee                                                  As String
    Dim vLIQPetty_Code                                                As String
    Dim vLIQPetty_type                                                As String
    Dim vLIQAccount_cd                                                As String
    Dim vLIQpetty_date                                                As String
    Dim vLIQdatecreate                                                As String
    Dim vLIQtimecreate                                                As String
    Dim vLIQPCF_NUMBER                                                As String

    Dim vLiq_Amt                                                      As Double
    Dim vLiq_Date                                                     As String
    Dim vLIQParticulars                                               As String

    Dim vLIQpetty_cash                                                As Double
    Dim vLIQoriginal                                                  As Double

    Dim NoManIsPerfect                                                As Long

    Dim rsPettyDup                                                    As ADODB.Recordset
    Set rsPettyDup = New ADODB.Recordset
    Set rsPettyDup = gconDMIS.Execute("Select * from CMIS_Petty Where ID = " & GridToLiquidate)
    If Not rsPettyDup.EOF And Not rsPettyDup.BOF Then
        vLIQEmployee = N2Str2Null(rsPettyDup!Employee)
        vLIQPetty_Code = "'001'"
        vLIQPetty_type = N2Str2Null(rsPettyDup!Petty_type)
        vLIQAccount_cd = N2Str2Null(rsPettyDup!Account_cd)
        vLIQpetty_date = N2Str2Null(rsPettyDup!PETTY_DATE)
        vLIQdatecreate = N2Str2Null(rsPettyDup!DATECREATE)
        vLIQtimecreate = N2Str2Null(rsPettyDup!TimeCreate)
        vLIQPCF_NUMBER = N2Str2Null(rsPettyDup!PCF_NUMBER)
        vLIQpetty_cash = N2Str2Zero(rsPettyDup!PETTY_CASH)

        vLiq_Amt = NumericVal(txtTotalCashAdvance.Text)
        vLiq_Date = N2Date2Null(txtBDPetty_Date)
        vLIQParticulars = N2Str2Null(txtParticularsBD)

        gconDMIS.Execute ("update CMIS_Petty Set " & _
                        " petty_cash = 0," & _
                        " Liq_Amt = " & vLiq_Amt & "," & _
                        " Liq_Date = " & vLiq_Date & "," & _
                        " Particulars = " & vLIQParticulars & "," & _
                        " Liquidated = 1" & _
                        " where ID = " & GridToLiquidate)
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                        " ADVANCES = ADVANCES - " & vLIQpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    End If

    For NoManIsPerfect = 1 To grdBreakDown.Rows - 1
        grdBreakDown.Row = NoManIsPerfect
        vLIQEmployee = N2Str2Null(txtEmployeeCode.Text)
        vLIQpetty_date = N2Date2Null(txtBDPetty_Date.Text)
        vLIQPetty_Code = "'001'"
        grdBreakDown.Col = 0: vLIQpetty_date = N2Date2Null(grdBreakDown.Text)
        grdBreakDown.Col = 2: vLIQPetty_type = N2Str2Null(SetReplenishCode(grdBreakDown.Text))
        grdBreakDown.Col = 3: vLIQAccount_cd = N2Str2Null(SetReplenishAcctCode(grdBreakDown.Text))
        grdBreakDown.Col = 4: vLIQpetty_cash = NumericVal(grdBreakDown.Text)
        vLIQdatecreate = "'" & LOGDATE & "'"
        vLIQtimecreate = "'" & Time & "'"
        vLIQPCF_NUMBER = N2Str2Null(txtBDPCF_NUMBER.Text)
        vLIQoriginal = 0
        gconDMIS.Execute ("Insert into CMIS_Petty " & _
                          "(Employee,Petty_Code,Petty_type,Account_cd,petty_date,Particulars,petty_cash,datecreate,timecreate,PCF_NUMBER,LiquidType,Liquidated,Liquid,Replenish,Original)" & _
                        " values (" & vLIQEmployee & ",'001'," & vLIQPetty_type & "," & vLIQAccount_cd & "," & vLIQpetty_date & "," & vLIQParticulars & "," & vLIQpetty_cash & "," & vLIQdatecreate & "," & vLIQtimecreate & "," & vLIQPCF_NUMBER & ",'1',0,1,0,0)")
        gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                        " EXPENSE = EXPENSE + " & vLIQpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    Next
    'cmdBreakDown.Visible = False: cmdBreakDown.ZOrder 1
    picBreakDown.Visible = False: picBreakDown.ZOrder 1
    StoreDetails

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSavePetty_Click()
    'updating code:    JAA - 07112007
    'On Error GoTo ErrorCode:

    Dim vEmployee                                                     As String
    Dim vPetty_Code                                                   As String
    Dim vPetty_type                                                   As String
    Dim vAccount_cd                                                   As String
    Dim vpetty_date                                                   As String
    Dim vparticulars                                                  As String
    Dim vdatecreate                                                   As String
    Dim vtimecreate                                                   As String
    Dim vPCF_NUMBER                                                   As String

    Dim vpetty_cash                                                   As Double
    Dim voriginal                                                     As Double

    vEmployee = N2Str2Null(txtEmployeeCode.Text)
    vPetty_Code = N2Str2Null(txtPetty_Code.Text)
    vPetty_type = N2Str2Null(SetReplenishCode(cboReplenishment.Text))
    vAccount_cd = N2Str2Null(SetReplenishAcctCode(cboReplenishment.Text))
    vpetty_date = N2Date2Null(txtPetty_Date.Text)
    vparticulars = N2Str2Null(txtParticulars.Text)
    vdatecreate = "'" & LOGDATE & "'"
    vtimecreate = "'" & Time & "'"
    vPCF_NUMBER = N2Str2Null(txtPCF_NUMBER.Text)
    vpetty_cash = NumericVal(txtOriginal.Text)
    voriginal = NumericVal(txtOriginal.Text)

    Dim rsCheckPettyPosition                                          As ADODB.Recordset
    Set rsCheckPettyPosition = New ADODB.Recordset
    Set rsCheckPettyPosition = gconDMIS.Execute("Select * from CMIS_Cash_Pos where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
    If Not rsCheckPettyPosition.EOF And Not rsCheckPettyPosition.BOF Then
        Dim Cash_Pos_Remaining_Petty_Fund                             As Double
        Cash_Pos_Remaining_Petty_Fund = N2Str2Zero(rsCheckPettyPosition!FUND) - (N2Str2Zero(rsCheckPettyPosition!REPLENISH) + N2Str2Zero(rsCheckPettyPosition!EXPENSE) + N2Str2Zero(rsCheckPettyPosition!ADVANCES))
        If vpetty_cash > Cash_Pos_Remaining_Petty_Fund Then
            MsgBox "Warning: Fund is Lower than Requested Petty Cash!", vbCritical, "Not Allowed!"
            Exit Sub
        End If
    End If

    If AddorEdit = "ADD" Then
        If vPetty_Code = "001" Then
            If vPetty_type = "" Then
                MsgBox "You can't save this transaction if this field is empty. Please select", vbCritical, "Not Allowed"
                cboReplenishment.SetFocus
                Exit Sub
            ElseIf vAccount_cd = "" Then
                MsgBox "You can't save this transaction if this field is empty. Please select", vbCritical, "Not Allowed"
                cboReplenishment.SetFocus
                Exit Sub
            End If
        End If
        SQL_STATEMENT = "Insert into CMIS_Petty " & _
                        "(Employee,Petty_Code,Petty_type,Account_cd,petty_date,Particulars,petty_cash,datecreate,timecreate,original,PCF_NUMBER)" & _
                      " values (" & vEmployee & "," & vPetty_Code & "," & vPetty_type & "," & vAccount_cd & "," & vpetty_date & "," & vparticulars & "," & vpetty_cash & "," & vdatecreate & "," & vtimecreate & "," & voriginal & "," & vPCF_NUMBER & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        If txtPetty_Code.Text = "002" Then
            NEW_LogAudit "AA", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid, "", "PCV NO: " & txtPCF_NUMBER & " - " & txtParticulars, "", ""
        ElseIf txtPetty_Code.Text = "001" Then
            NEW_LogAudit "AA", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid.Caption, "", "PCV NO: " & txtPCF_NUMBER & " - " & cboReplenishment, "", ""
        Else
        End If

        If txtPetty_Code.Text = "001" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " EXPENSE = EXPENSE + " & vpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If txtPetty_Code.Text = "002" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " ADVANCES = ADVANCES + " & vpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CMIS_Petty Set " & _
                      " Employee = " & vEmployee & "," & _
                      " Petty_Code = " & vPetty_Code & "," & _
                      " Petty_type = " & vPetty_type & "," & _
                      " Account_cd = " & vAccount_cd & "," & _
                      " petty_date = " & vpetty_date & "," & _
                      " Particulars = " & vparticulars & "," & _
                      " datecreate = " & vdatecreate & "," & _
                      " timecreate = " & vtimecreate & "," & _
                      " petty_cash = " & vpetty_cash & "," & _
                      " PCF_NUMBER = " & vPCF_NUMBER & "," & _
                      " original = " & voriginal & _
                      " where ID = " & labPettyID.Caption
        gconDMIS.Execute SQL_STATEMENT
        If txtPetty_Code.Text = "002" Then
            NEW_LogAudit "EE", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid, "", "PCV NO: " & txtPCF_NUMBER & " - " & txtParticulars, "", labPettyID.Caption
        ElseIf txtPetty_Code.Text = "001" Then
            NEW_LogAudit "EE", "TRANSACTION PETTY CASH ENTRY", SQL_STATEMENT, labid.Caption, "", "PCV NO: " & txtPCF_NUMBER & " - " & cboReplenishment, "", labPettyID.Caption
        Else
        End If
        
        If txtPetty_Code.Text = "001" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " EXPENSE = (EXPENSE - " & PrevPettyCash & ") + " & vpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        If txtPetty_Code.Text = "002" Then
            gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                            " ADVANCES = (ADVANCES - " & PrevPettyCash & ") + " & vpetty_cash & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        End If
        ShowSuccessFullyUpdated
    End If
    
    StoreDetails
    cmdCancelPetty_Click

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdShow_Click()
    Picture1.Enabled = True
    If optCashAdvance.Value = True Then
        txtPetty_Code.Text = "002"
        cmdPOST.Caption = "Liquidate"
        cmdPOST.Enabled = True
        cmdAdd.Enabled = True
        cmdEdit.Enabled = True
        cmdPrint.Enabled = False
        Label20.Top = 390
        Label20.Left = 1740
        txtParticulars.Top = 690
        txtParticulars.Height = 1065
        Label11.Top = 1230
        cboReplenishment.Enabled = False
        cboReplenishment.ListIndex = -1
        cboReplenishment.ZOrder 1
        StoreDetails
    ElseIf optExpense.Value = True Then
        'cmdPOST.Caption = "Replenish"
        cmdPOST.Enabled = False
        txtPetty_Code.Text = "001"
        cmdAdd.Enabled = True
        cmdEdit.Enabled = True
        cmdPrint.Enabled = False

        Label20.Left = 180
        Label20.Top = 1200
        txtParticulars.Top = 1110
        txtParticulars.Height = 675
        Label11.Top = 390
        cboReplenishment.Enabled = True
        cboReplenishment.ListIndex = -1
        cboReplenishment.ZOrder 1

        StoreDetails
    ElseIf optReplenish.Value = True Then
        txtPetty_Code.Text = "001"
        cmdPOST.Caption = ""
        cmdPOST.Enabled = False
        cmdPrint.Enabled = False
        cmdAdd.Enabled = False
        cmdEdit.Enabled = False
        StoreDetails
    End If
End Sub

Private Sub cmdTag_Click()
    Dim GridID                                                        As Long
    grdPetty.Col = 7
    If grdPetty.Text <> "" Then
        GridID = grdPetty.Text
        grdPetty.Col = 5:
        If grdPetty.Text = " T" Then
            MsgBox "Already Replenished!", vbInformation, "Message"
        Else
            grdPetty.Col = 4:
            If grdPetty.Text = " T" Then
                grdPetty.Text = ""
                gconDMIS.Execute ("update CMIS_Petty Set Tag = 0 Where id = " & GridID)
            Else
                grdPetty.Text = " T"
                gconDMIS.Execute ("update CMIS_Petty Set Tag = 1 Where id = " & GridID)
            End If
        End If
    End If
End Sub

Private Sub Command1_Click()
    Dim rsPettyUpdate As ADODB.Recordset
    Set rsPettyUpdate = New ADODB.Recordset
    Dim xAccount_CD As String
    Dim xAcctCode As String
    Dim xAcctCodeRep As String
    rsPettyUpdate.Open "SELECT DISTINCT AMIS_CHARTACCOUNT.ACCTCODE,ACCOUNT_CD,AMIS_CHARTACCOUNT.DESCRIPTION FROM CMIS_PETTY INNER JOIN AMIS_CHARTACCOUNT ON AMIS_CHARTACCOUNT.ACCTCODE=CMIS_PETTY.ACCOUNT_CD INNER JOIN AMIS_CHARTACCOUNT ON AMIS_CHARTACCOUNT.DESCRIPTION=AMIS_CHARTACCOUNT.DESCRIPTION WHERE CMIS_PETTY.ACCOUNT_CD NOT IN (SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE AMIS_CHARTACCOUNT.ACCTCODE=CMIS_PETTY.ACCOUNT_CD) AND CMIS_PETTY.ACCOUNT_CD IS NOT NULL ORDER BY CMIS_PETTY.ACCOUNT_CD", gconDMIS, adOpenKeyset
    'rsPettyUpdate.Open "SELECT * FROM CMIS_PETTYHAI WHERE ACCOUNT_CD IS NOT NULL", gconDMIS, adOpenKeyset
    'If Not rsPettyUpdate.EOF And Not rsPettyUpdate.BOF Then
    Do While Not rsPettyUpdate.EOF
        xAccount_CD = Null2String(rsPettyUpdate!Account_cd)
        xAcctCodeRep = Replace(Null2String(rsPettyUpdate!AcctCode), "-", "")
        xAcctCode = Null2String(rsPettyUpdate!AcctCode)
        gconDMIS.Execute ("UPDATE CMIS_PETTY SET ACCOUNT_CD = '" & xAcctCode & "',PETTY_TYPE='" & xAcctCodeRep & "' where ACCOUNT_CD = '" & xAccount_CD & "' AND PETTY_TYPE IS NOT NULL")
'        gconDMIS.Execute ("UPDATE CMIS_PETTYHAI SET PETTY_TYPE='" & xAcctCode & "' FROM CMIS_PETTYHAI WHERE ACCOUNT_CD='" & xAccount_Cd & "' AND ACCOUNT_CD IS NOT NULL")
        rsPettyUpdate.MoveNext
    Loop
    'End If
    Set rsPettyUpdate = Nothing
End Sub

Private Sub Command2_Click()
    Dim rsUpdate As ADODB.Recordset
    Set rsUpdate = New ADODB.Recordset
    Dim xAccount_CD As String
    rsUpdate.Open "SELECT DISTINCT PETTY_TYPE,REPLACE(ACCOUNT_CD,'-','') ACCOUNT_CD FROM CMIS_PETTY WHERE ACCOUNT_CD IS NULL AND PETTY_TYPE IS NOT NULL", gconDMIS, adOpenKeyset
    Do While Not rsUpdate.EOF
        xAccount_CD = Left(rsUpdate!Petty_type, 2) + "-" + Mid(rsUpdate!Petty_type, 3, 5) + "-" + Right(rsUpdate!Petty_type, 2)
        gconDMIS.Execute ("UPDATE CMIS_PETTY SET ACCOUNT_CD ='" & xAccount_CD & "' WHERE PETTY_TYPE='" & rsUpdate!Petty_type & "'")
        rsUpdate.MoveNext
    Loop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
                Unload frmALL_AuditInquiry
                 
                frmALL_AuditInquiry.Show
                frmALL_AuditInquiry.ZOrder 0
                frmALL_AuditInquiry.Caption = "Audit Inquiry (TRANSACTION PETTY CASH ENTRY)"
                Call frmALL_AuditInquiry.DisplayHistory(labid, "TRANSACTION PETTY CASH ENTRY")
            'End If
            
        Case vbKeyF2
            cmdAdd.Value = True
        Case vbKeyF3
            cmdEdit.Value = True
        Case vbKeyF4
            If txtPetty_Code.Text = "001" Then cmdTag.Value = True
            If txtPetty_Code.Text = "002" Then cmdLiquidate.Value = True
        Case vbKeyF6
            cmdReplenish.Value = True
        Case vbKeyEscape
            If picPettyCashEntry.Visible = True Then
                'cmdPettyCashEntry.Visible = False
                'cmdPettyCashEntry.ZOrder 1
                picPettyCashEntry.Visible = False
                picPettyCashEntry.ZOrder 1
                PrevPettyCash = 0
            End If
            If picBreakDown.Visible = True Then
                'cmdBreakDown.Visible = False: cmdBreakDown.ZOrder 1
                picBreakDown.Visible = False: picBreakDown.ZOrder 1
            End If
            If picLiquidate.Visible = True Then
                picLiquidate.ZOrder 1
                picLiquidate.Visible = False
            End If
        Case vbKeyF11
            Shell "calc.exe"
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Dim rsProfile                                                     As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing
    CenterMe frmMain, Me, 1: initMemvars
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'cmdPettyCashEntry.Visible = False: cmdPettyCashEntry.ZOrder 1
    picPettyCashEntry.Visible = False: picPettyCashEntry.ZOrder 1
    PrevPettyCash = 0: FillCboPettyType: FillCboReplenishment
    'cmdBreakDown.Visible = False: cmdBreakDown.ZOrder 1
    picBreakDown.Visible = False: picBreakDown.ZOrder 1
    picLiquidate.ZOrder 1: picLiquidate.Visible = False: textSearch.Text = ""
    cmdCancelCO.Enabled = False: cmdCancelCO.Caption = "": cmdCancelCO.Picture = LoadPicture("")
    rsRefresh
    StoreMemVars
    cmdTag.Visible = False: cmdReplenish.Visible = False: cmdLiquidate.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub grdPetty_Click()
    grdPetty.Col = 7
    If grdPetty.Text <> "" Then
        Dim rsPetty2                                                  As ADODB.Recordset
        Set rsPetty2 = New ADODB.Recordset
        Set rsPetty2 = gconDMIS.Execute("Select * from CMIS_Petty where id = " & grdPetty.Text)
        If Not rsPetty2.EOF And Not rsPetty2.BOF Then
            txtPetty_CashNo.Text = Null2String(rsPetty2!PCF_NUMBER)
            txtLiq_Amt.Text = ToDoubleNumber(N2Str2Zero(rsPetty2!LIQ_AMT))
            txtLiq_Date.Text = Null2String(rsPetty2!LIQ_DATE)
            txtDistParticulars.Text = Null2String(rsPetty2!Particulars)
            If Null2Bool(rsPetty2!REPLENISH) = True Then
                txtLiquidated.Text = "Replenish"
            Else
                If Null2Bool(rsPetty2!liquid) = True Then
                    txtLiquidated.Text = "Liquidation"
                Else
                    If Null2Bool(rsPetty2!liquidated) = True Then
                        txtLiquidated.Text = "Liquidated"
                    Else
                        txtLiquidated.Text = ""
                    End If
                End If
            End If
        Else
            txtPetty_CashNo.Text = ""
            txtLiq_Amt.Text = ""
            txtLiq_Date.Text = ""
            txtLiquidated.Text = ""
        End If
    Else
        txtPetty_CashNo.Text = ""
        txtLiq_Amt.Text = ""
        txtLiq_Date.Text = ""
        txtLiquidated.Text = ""
    End If
End Sub

Private Sub grdPetty_DblClick()
    If txtLiquidated.Text <> "Liquidation" And txtLiquidated.Text <> "Liquidated" Then
        cmdEdit.Value = True
    End If
End Sub

'SEARCH MODULE
Private Sub lstPetty_GotFocus()
    'Employee.Bookmark = rsFind(Employee.Clone, "id", lstPetty.SelectedItem.ListSubItems(2)).Bookmark
    'StoreMemvars
    
    If lstPetty.ListItems.Count = 0 Then Exit Sub
    
    Dim Index As Integer
    Dim Item As ListItem
    
    Index = lstPetty.SelectedItem.Index
    Employee.MoveFirst
    Employee.Find "ID = " & lstPetty.ListItems(Index).ListSubItems(2) & ""
    labid.Caption = lstPetty.ListItems(Index).ListSubItems(2)
    StoreMemVars
End Sub

Private Sub lstPetty_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'Employee.Bookmark = rsFind(Employee.Clone, "ID", lstPetty.SelectedItem.ListSubItems(2)).Bookmark
    'StoreMemvars
    'cmdShow.Value = True
    
    Employee.MoveFirst
    Employee.Find "ID = " & Item.ListSubItems(2) & ""
    labid.Caption = Item.ListSubItems(2)
    Call StoreMemVars
    Call cmdShow_Click
End Sub

Private Sub lstPetty_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPetty
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

Private Sub lstPetty_DblClick()
    'cmdEdit.Value = True

End Sub

Private Sub lstPetty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus

    End If
End Sub

Private Sub optBREAKDOWN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        InitBDGrid
        txtTotalCashAdvance.Text = "0.00": TotalBreakDownCA = 0
        picNORMAL_LIQUIDATE.ZOrder 1: picNORMAL_LIQUIDATE.Visible = False
        'cmdBreakDown.Visible = True: cmdBreakDown.ZOrder 0
        picBreakDown.Visible = True: picBreakDown.ZOrder 0
        On Error Resume Next
        txtBDPetty_Date.SetFocus
        txtBDPetty_Date.Text = LOGDATE
        Dim rsPCF_NUMBER                                              As ADODB.Recordset
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '001' order by PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtBDPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
        Else
            txtBDPCF_NUMBER.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing
    End If
End Sub

Private Sub optCANCEL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        picNORMAL_LIQUIDATE.ZOrder 1
        picNORMAL_LIQUIDATE.Visible = False
        picStatus.Enabled = False
    End If
End Sub

Private Sub optCashAdvance_Click()
    txtPetty_Code.Text = "002"
    Call cmdShow_Click
End Sub

Private Sub optExpense_Click()
    txtPetty_Code.Text = "001"
    Call cmdShow_Click
End Sub

Private Sub optNORMAL_KeyDown(KeyCode As Integer, Shift As Integer)
    picLiquidate.Visible = False: picLiquidate.ZOrder 1
    picNORMAL_LIQUIDATE.Visible = False: picNORMAL_LIQUIDATE.ZOrder 1
    Dim rsPCF_NUMBER                                                  As ADODB.Recordset
    If txtPetty_Code.Text = "001" Then
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '001' ORDER BY PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtPetty_CashNo.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER) + 1, "000000")
        Else
            txtPetty_CashNo.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing
    Else
        Set rsPCF_NUMBER = New ADODB.Recordset
        Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty WHERE PETTY_CODE = '002' ORDER BY PCF_NUMBER DESC")
        If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
            txtPetty_CashNo.Text = Format(NumericVal(rsPCF_NUMBER!PCF_NUMBER) + 1, "000000")
        Else
            txtPetty_CashNo.Text = "000001"
        End If
        Set rsPCF_NUMBER = Nothing
    End If
    picStatus.Enabled = True: txtPetty_CashNo.Enabled = True
    txtLiq_Amt.Enabled = True: txtLiquidated.Enabled = False
    txtTotalPettyCash.Enabled = False: txtLiq_Date.Enabled = True
    On Error Resume Next

    txtPetty_CashNo.SetFocus
    txtLiq_Date.Text = LOGDATE
End Sub

Private Sub optReplenish_Click()
    txtPetty_Code.Text = "001"
    Call cmdShow_Click
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstPetty.ListItems.Count > 0 And lstPetty.Enabled = True Then: lstPetty.SetFocus
    End If
End Sub

Private Sub txtBDPetty_Cash_GotFocus()
    If NumericVal(txtBDPetty_Cash.Text) = 0 Then txtBDPetty_Cash.Text = "" Else txtBDPetty_Cash.Text = NumericVal(txtBDPetty_Cash.Text)
End Sub

Private Sub txtBDPetty_Cash_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtBDPetty_Cash_LostFocus()
    txtBDPetty_Cash.Text = ToDoubleNumber(txtBDPetty_Cash.Text)
End Sub

Private Sub txtLiq_Date_GotFocus()
    If IsDate(txtLiq_Date.Text) = True Then txtLiq_Date.Text = Format(txtLiq_Date.Text, "MM/DD/YYYY") Else txtLiq_Date.Text = ""
End Sub

Private Sub txtLiq_Date_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrorCode
    If KeyCode = vbKeyReturn Then
        If MsgBox("Save Liquidation?", vbQuestion + vbYesNo) = vbYes Then
            Dim rsPettyDup                                            As ADODB.Recordset
            Set rsPettyDup = New ADODB.Recordset
            Set rsPettyDup = gconDMIS.Execute("Select * from CMIS_Petty Where ID = " & GridToLiquidate)
            If Not rsPettyDup.EOF And Not rsPettyDup.BOF Then
                Dim vLIQEmployee                                      As String
                Dim vLIQPetty_Code                                    As String
                Dim vLIQPetty_type                                    As String
                Dim vLIQAccount_cd                                    As String
                Dim vLIQpetty_date                                    As String
                Dim vLIQdatecreate                                    As String
                Dim vLIQtimecreate                                    As String
                Dim vLIQPCF_NUMBER                                    As String
                Dim vLiq_Amt                                          As Double
                Dim vLiq_Date                                         As String
                Dim vLIQpetty_cash                                    As Double

                vLIQEmployee = N2Str2Null(rsPettyDup!Employee)
                vLIQPetty_Code = "'001'"
                vLIQPetty_type = N2Str2Null(rsPettyDup!Petty_type)
                vLIQAccount_cd = N2Str2Null(rsPettyDup!Account_cd)
                vLIQpetty_date = N2Str2Null(rsPettyDup!PETTY_DATE)
                vLIQdatecreate = N2Str2Null(rsPettyDup!DATECREATE)
                vLIQtimecreate = N2Str2Null(rsPettyDup!TimeCreate)
                vLIQPCF_NUMBER = N2Str2Null(rsPettyDup!PCF_NUMBER)
                'vLIQpetty_cash = N2Str2Zero(rsPettyDup!petty_cash)
                vLIQpetty_cash = NumericVal(txtLiq_Amt.Text)
                vLiq_Amt = NumericVal(txtLiq_Amt.Text)
                vLiq_Date = N2Date2Null(txtLiq_Date.Text)

                gconDMIS.Execute ("update CMIS_Petty Set " & _
                                " petty_cash = 0," & _
                                " Liq_Amt = " & vLiq_Amt & "," & _
                                " Liq_Date = " & vLiq_Date & "," & _
                                " Liquidated = 1" & _
                                " where ID = " & GridToLiquidate)
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                " ADVANCES = ADVANCES - " & N2Str2Zero(rsPettyDup!PETTY_CASH) & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")

                gconDMIS.Execute ("Insert into CMIS_Petty " & _
                                  "(Employee,Petty_Code,Petty_type,Account_cd,petty_date,petty_cash,datecreate,timecreate,PCF_NUMBER,LiquidType,Liquidated,Liquid,Replenish,Original)" & _
                                " values (" & vLIQEmployee & "," & vLIQPetty_Code & "," & vLIQPetty_type & "," & vLIQAccount_cd & "," & vLIQpetty_date & "," & vLIQpetty_cash & "," & vLIQdatecreate & "," & vLIQtimecreate & "," & vLIQPCF_NUMBER & ",'1',0,1,0,0)")
                gconDMIS.Execute ("update CMIS_Cash_Pos Set" & _
                                " EXPENSE = EXPENSE + " & vLiq_Amt & " where CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
                StoreDetails
                picStatus.Enabled = False
            End If
        End If
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub txtLiq_Date_LostFocus()
    If IsDate(txtLiq_Date.Text) = True Then txtLiq_Date.Text = Format(txtLiq_Date.Text, "DD-MMM-YY") Else txtLiq_Date.Text = ""
End Sub

Private Sub txtOriginal_GotFocus()
    If NumericVal(txtOriginal.Text) = 0 Then txtOriginal.Text = "" Else txtOriginal.Text = NumericVal(txtOriginal.Text)
End Sub

Private Sub txtOriginal_LostFocus()
    txtOriginal.Text = ToDoubleNumber(txtOriginal.Text)
End Sub

Private Sub txtPetty_Code_Change()
    If txtPetty_Code.Text = "001" Then
        cmdPOST.Caption = "Replenish"
    Else
        cmdPOST.Caption = "Liquidate"
    End If
End Sub

Private Sub txtPetty_Date_GotFocus()
    If IsDate(txtPetty_Date.Text) = True Then
        txtPetty_Date.Text = Format(txtPetty_Date.Text, "MM/DD/YYYY")
    Else
        txtPetty_Date.Text = ""
    End If
End Sub

Private Sub txtPetty_Date_LostFocus()
    Dim rsPCF_NUMBER                                                  As ADODB.Recordset
    If IsDate(txtPetty_Date.Text) = True Then
        txtPetty_Date.Text = Format(txtPetty_Date.Text, "DD-MMM-YYYY")
        If txtPetty_Code.Text = "001" Then
            Set rsPCF_NUMBER = New ADODB.Recordset
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '001' order by PCF_NUMBER DESC")
            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
            Set rsPCF_NUMBER = Nothing
        Else
            Set rsPCF_NUMBER = New ADODB.Recordset
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '002' order by PCF_NUMBER DESC")
            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
            Set rsPCF_NUMBER = Nothing
        End If
    Else
        txtPetty_Date.Text = ""
        If txtPetty_Code.Text = "001" Then
            Set rsPCF_NUMBER = New ADODB.Recordset
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '001' order by PCF_NUMBER DESC")
            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
            Set rsPCF_NUMBER = Nothing
        Else
            Set rsPCF_NUMBER = New ADODB.Recordset
            Set rsPCF_NUMBER = gconDMIS.Execute("Select PCF_NUMBER from CMIS_Petty where PETTY_CODE = '002' order by PCF_NUMBER DESC")
            If Not rsPCF_NUMBER.EOF And Not rsPCF_NUMBER.BOF Then
                txtPCF_NUMBER.Text = Format(NumericVal(Null2String(rsPCF_NUMBER!PCF_NUMBER)) + 1, "000000")
            Else
                txtPCF_NUMBER.Text = "000001"
            End If
            Set rsPCF_NUMBER = Nothing
        End If
    End If
End Sub

